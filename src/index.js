import * as XLSX from 'xlsx';

const PROJECT_ID = 'work-schedule-1f2e1';

function loadServiceAccount(env) {
  const raw = env?.FIREBASE_SERVICE_ACCOUNT;
  if (!raw) {
    throw new Error('Missing FIREBASE_SERVICE_ACCOUNT secret');
  }
  try {
    return typeof raw === 'string' ? JSON.parse(raw) : raw;
  } catch (e) {
    throw new Error('FIREBASE_SERVICE_ACCOUNT is not valid JSON: ' + e.message);
  }
}

// ─── Auth ───────────────────────────────────────────────────────────

async function getAccessToken(serviceAccount) {
  const now = Math.floor(Date.now() / 1000);
  
  // CRITICAL FIX: Add 'kid' (key ID) to header - Google requires this
  const jwtHeader = btoa(JSON.stringify({ 
    alg: 'RS256', 
    typ: 'JWT',
    kid: serviceAccount.private_key_id  // ← This was missing!
  }));
  
  // CRITICAL FIX: Add 'sub' (subject) to claims - Google requires this
  const jwtClaimSet = btoa(JSON.stringify({
    iss: serviceAccount.client_email,
    sub: serviceAccount.client_email,  // ← This was missing!
    scope: 'https://www.googleapis.com/auth/datastore',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  }));

  const unsignedToken = `${jwtHeader}.${jwtClaimSet}`;

  const pemHeader = "-----BEGIN PRIVATE KEY-----";
  const pemFooter = "-----END PRIVATE KEY-----";
  const pemContents = serviceAccount.private_key.substring(
    pemHeader.length,
    serviceAccount.private_key.length - pemFooter.length - 1
  );
  const binaryDer = Uint8Array.from(atob(pemContents), c => c.charCodeAt(0));

  const key = await crypto.subtle.importKey(
    'pkcs8',
    binaryDer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['sign']
  );

  const signature = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    key,
    new TextEncoder().encode(unsignedToken)
  );

  const signatureBase64 = btoa(String.fromCharCode(...new Uint8Array(signature)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  const jwt = `${unsignedToken}.${signatureBase64}`;

  const tokenResponse = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
  });

  if (!tokenResponse.ok) {
    const errorText = await tokenResponse.text();
    throw new Error(`Token exchange failed: ${tokenResponse.status} - ${errorText}`);
  }

  const tokenData = await tokenResponse.json();
  return tokenData.access_token;
}

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ─── Firestore helpers ──────────────────────────────────────────────

async function batchWriteFirestore(accessToken, writes) {
  const batchUrl = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:commit`;

  const response = await fetch(batchUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ writes })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Batch write failed: ${response.status} - ${errorText}`);
  }
  return await response.json();
}

async function listCollectionDocs(accessToken, collectionPath, filterFn) {
  const listUrl = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/${collectionPath}?pageSize=1000`;
  const docs = [];
  let pageToken = '';

  do {
    const url = pageToken ? `${listUrl}&pageToken=${pageToken}` : listUrl;
    const res = await fetch(url, { headers: { 'Authorization': `Bearer ${accessToken}` } });
    const data = await res.json();

    if (data.documents) {
      for (const doc of data.documents) {
        if (!filterFn || filterFn(doc)) docs.push(doc.name);
      }
    }
    pageToken = data.nextPageToken || '';
  } while (pageToken);

  return docs;
}

async function runBatchedWrites(accessToken, deletes, creates, label) {
  const BATCH_SIZE = 500;
  let delIdx = 0, createIdx = 0, batchNum = 1, total = 0;

  while (delIdx < deletes.length || createIdx < creates.length) {
    const writes = [];
    while (delIdx < deletes.length && writes.length < BATCH_SIZE) {
      writes.push({ delete: deletes[delIdx++] });
    }
    while (createIdx < creates.length && writes.length < BATCH_SIZE) {
      writes.push(creates[createIdx++]);
    }
    if (writes.length === 0) break;

    console.log(`📦 [${label}] Batch ${batchNum}: ${writes.length} ops`);
    await batchWriteFirestore(accessToken, writes);
    total += writes.length;
    batchNum++;
    if (delIdx < deletes.length || createIdx < creates.length) await sleep(100);
  }
  console.log(`✅ [${label}] ${deletes.length} deleted, ${creates.length} created (${total} total ops)`);
}

// ─── File type detection ────────────────────────────────────────────

function classifyWorkbook(workbook, fileName) {
  const fn = fileName.toLowerCase();

  // Filename-based IOCS detection
  if (fn.includes('iocs') || fn.includes('daily_iocs') || fn.includes('daily iocs')) {
    return { kind: 'iocs' };
  }

  // Content-based IOCS detection: sheets with "ASSIGNED DCT" header
  const sheetNames = workbook.SheetNames;
  let looksLikeIocs = false;
  for (const sheetName of sheetNames) {
    if (sheetName === 'BLANK' || sheetName === 'Sheet2') continue;
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    if (rows.length < 2) continue;
    for (let i = 0; i < Math.min(3, rows.length); i++) {
      const rowStr = (rows[i] || []).map(c => String(c || '').toUpperCase()).join('|');
      if (rowStr.includes('ASSIGNED DCT') || rowStr.includes('IOCS READINGS')) {
        looksLikeIocs = true;
        break;
      }
    }
    if (looksLikeIocs) break;
  }
  if (looksLikeIocs) return { kind: 'iocs' };

  // Schedule file — determine state
  if (fn.includes('utah') || /\but\b/.test(fn) || fn.startsWith('ut')) {
    return { kind: 'schedule', state: 'Utah' };
  }
  if (fn.includes('nevada') || /\bnv\b/.test(fn) || fn.startsWith('nv')) {
    return { kind: 'schedule', state: 'Nevada' };
  }
  return { kind: 'schedule', state: 'Nevada' };
}

// ─── Schedule parser ────────────────────────────────────────────────

function normalizePersonName(name) {
  const upper = name.trim().toUpperCase();
  const nameMap = {
    'JASON': 'JRA', 'JASON ALVAREZ': 'JRA',
    'JERRY': 'JA', 'JERRY ANGLO': 'JA',
    'JEFF': 'JN', 'JEFF NIZNICK': 'JN',
    'GRACE': 'GA', 'GRACE AGRESOR': 'GA',
    'SHEILA': 'SV', 'SHEILA VELASQUEZ': 'SV',
    'RYAN': 'RR', 'RYAN REYNOSO': 'RR',
    'CHRISTIAN': 'CA', 'CHRISTIAN ALBERT': 'CA',
    'KRISTINE': 'KK', 'KRISTINE KIESLING': 'KK', 'BUDD': 'KK'
  };
  return nameMap[upper] || upper;
}

function parseScheduleExcel(bytes, state) {
  const workbook = XLSX.read(bytes, { type: 'array' });
  const scheduleData = [];
  let totalRowsParsed = 0;

  for (const sheetName of workbook.SheetNames) {
    console.log(`[SCHEDULE/${state}] Sheet: ${sheetName}`);
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    let currentDate = null;
    let rowsInSheet = 0;

    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      if (!row || row.length === 0) continue;

      const firstCell = String(row[0] || '').trim();

      if (firstCell.match(/^(MON|TUE|WED|THU|FRI|SAT|SUN)/i)) {
        const dateMatch = firstCell.match(/(\d{1,2})-(\d{1,2})-(\d{2})/);
        if (dateMatch) {
          const mm = dateMatch[1].padStart(2, '0');
          const dd = dateMatch[2].padStart(2, '0');
          const yy = dateMatch[3];
          currentDate = `20${yy}-${mm}-${dd}`;
        }
        continue;
      }

      if (firstCell === 'ZIP CODES' || firstCell === 'NEVADA' || firstCell === 'UTAH' || firstCell === 'TEST SCHEDULE') continue;

      if (currentDate && row.length >= 6) {
        const testName = String(row[0] || '').trim();
        const zip = String(row[1] || '').trim();
        const site = String(row[2] || '').trim();
        const type = String(row[3] || '').trim();
        const testId = String(row[4] || '').trim();
        const tech = String(row[5] || '').trim();

        if (tech && tech !== 'TECH(S)' && testName && testName !== 'ZIP CODES') {
          const normalizedPerson = normalizePersonName(tech);
          totalRowsParsed++;
          if (totalRowsParsed <= 3) {
            console.log(`   ${tech} → ${normalizedPerson} | ${testName} | ${currentDate}`);
          }
          scheduleData.push({
            date: currentDate,
            person: normalizedPerson,
            test: testName,
            zipCode: zip,
            testId: testId,
            location: site,
            mep: type,
            state
          });
          rowsInSheet++;
        }
      }
    }
    console.log(`   Added ${rowsInSheet} rows`);
  }

  return scheduleData;
}

async function processScheduleFile(fileName, bytes, state, accessToken) {
  console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  console.log(`📅 Processing SCHEDULE for ${state}: ${fileName}`);
  console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

  const scheduleData = parseScheduleExcel(bytes, state);
  console.log(`Parsed ${scheduleData.length} schedule entries`);

  if (scheduleData.length === 0) {
    return { fileName, kind: 'schedule', state, count: 0, status: 'skipped' };
  }

  const docsToDelete = await listCollectionDocs(
    accessToken,
    'schedule/current/rows',
    doc => doc.name.includes('/schedule/current/rows/') && doc.fields?.state?.stringValue === state
  );
  console.log(`🗑️  Found ${docsToDelete.length} old ${state} schedule docs to delete`);

  const creates = scheduleData.map((entry, idx) => ({
    update: {
      name: `projects/${PROJECT_ID}/databases/(default)/documents/schedule/current/rows/doc_${Date.now()}_${idx}_${Math.floor(Math.random() * 10000)}`,
      fields: {
        date: { stringValue: String(entry.date || '') },
        person: { stringValue: String(entry.person || '') },
        test: { stringValue: String(entry.test || '') },
        zipCode: { stringValue: String(entry.zipCode || '') },
        testId: { stringValue: String(entry.testId || '') },
        location: { stringValue: String(entry.location || '') },
        state: { stringValue: String(entry.state || '') },
        mep: { stringValue: String(entry.mep || '') },
        time: { stringValue: '' }
      }
    }
  }));

  await runBatchedWrites(accessToken, docsToDelete, creates, `SCHEDULE/${state}`);

  await fetch(`https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/schedule/current`, {
    method: 'PATCH',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      fields: {
        [`${state.toLowerCase()}UpdatedAt`]: { timestampValue: new Date().toISOString() },
        [`${state.toLowerCase()}Count`]: { integerValue: scheduleData.length.toString() },
        [`${state.toLowerCase()}Filename`]: { stringValue: fileName }
      }
    })
  });

  return { fileName, kind: 'schedule', state, count: scheduleData.length, status: 'success' };
}

// ─── IOCS parser ────────────────────────────────────────────────────

function normalizeIocsDate(raw) {
  if (!raw && raw !== 0) return '';
  const s = String(raw).trim();
  if (!s) return '';

  const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const mm = slashMatch[1].padStart(2, '0');
    const dd = slashMatch[2].padStart(2, '0');
    let yy = slashMatch[3];
    if (yy.length === 2) yy = '20' + yy;
    return `${yy}-${mm}-${dd}`;
  }

  const dashMatch = s.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
  if (dashMatch) {
    const mm = dashMatch[1].padStart(2, '0');
    const dd = dashMatch[2].padStart(2, '0');
    let yy = dashMatch[3];
    if (yy.length === 2) yy = '20' + yy;
    return `${yy}-${mm}-${dd}`;
  }

  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s);
    if (serial > 25000 && serial < 80000) {
      const utcDays = serial - 25569;
      const utcValue = utcDays * 86400 * 1000;
      const d = new Date(utcValue);
      if (!isNaN(d.getTime())) {
        return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
      }
    }
  }

  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
  } catch {}

  return s;
}

function detectIocsState(location) {
  const loc = String(location || '').toUpperCase();
  if (loc.includes(' UT ') || loc.endsWith(' UT') || loc.includes('UTAH')) return 'Utah';
  return 'Nevada';
}

function parseIocsExcel(bytes) {
  const workbook = XLSX.read(bytes, { type: 'array' });
  const allEntries = [];

  for (const sheetName of workbook.SheetNames) {
    if (sheetName === 'BLANK' || sheetName === 'Sheet2') continue;

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

    let rowsInSheet = 0;
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[1]) continue;

      const assignedDct = String(row[11] || '').trim().toUpperCase();
      if (!assignedDct || assignedDct === 'ASSIGNED DCT') continue;
      if (!/^[A-Z][A-Z\s.\-]*[A-Z]$/.test(assignedDct)) continue;
      if (!assignedDct.includes(' ') && assignedDct.length < 2) continue;

      const rawDate = String(row[1] || '').trim();
      if (!/\d/.test(rawDate)) continue;

      const location = String(row[3] || '').trim();
      const entry = {
        weekSheet: sheetName,
        date: normalizeIocsDate(row[1]),
        financeCode: String(row[2] || ''),
        location,
        ein: String(row[4] || ''),
        employeeName: String(row[5] || '').trim(),
        rd: String(row[6] || '').trim(),
        bt: String(row[7] || ''),
        et: String(row[8] || ''),
        rt: String(row[10] || ''),
        dct: assignedDct,
        state: detectIocsState(location)
      };

      if (allEntries.length < 3) {
        console.log(`[IOCS] Sample: ${entry.dct} | ${entry.employeeName} | ${entry.date} | ${entry.state} | RT ${entry.rt}`);
      }

      allEntries.push(entry);
      rowsInSheet++;
    }
    if (rowsInSheet > 0) console.log(`[IOCS] Sheet ${sheetName}: ${rowsInSheet} entries`);
  }

  return allEntries;
}

async function processIocsFile(fileName, bytes, accessToken) {
  console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  console.log(`📊 Processing IOCS file: ${fileName}`);
  console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

  const iocsEntries = parseIocsExcel(bytes);
  console.log(`Parsed ${iocsEntries.length} IOCS entries`);

  if (iocsEntries.length === 0) {
    return { fileName, kind: 'iocs', count: 0, status: 'skipped' };
  }

  const statesInFile = [...new Set(iocsEntries.map(e => e.state))];
  console.log(`States represented: ${statesInFile.join(', ')}`);

  const docsToDelete = await listCollectionDocs(
    accessToken,
    'iocs',
    doc => {
      if (!doc.name.includes('/iocs/')) return false;
      const state = doc.fields?.state?.stringValue;
      return statesInFile.includes(state);
    }
  );
  console.log(`🗑️  Found ${docsToDelete.length} old IOCS docs to delete`);

  const creates = iocsEntries.map((entry, idx) => ({
    update: {
      name: `projects/${PROJECT_ID}/databases/(default)/documents/iocs/doc_${Date.now()}_${idx}_${Math.floor(Math.random() * 10000)}`,
      fields: {
        weekSheet: { stringValue: String(entry.weekSheet || '') },
        date: { stringValue: String(entry.date || '') },
        financeCode: { stringValue: String(entry.financeCode || '') },
        location: { stringValue: String(entry.location || '') },
        ein: { stringValue: String(entry.ein || '') },
        employeeName: { stringValue: String(entry.employeeName || '') },
        rd: { stringValue: String(entry.rd || '') },
        bt: { stringValue: String(entry.bt || '') },
        et: { stringValue: String(entry.et || '') },
        rt: { stringValue: String(entry.rt || '') },
        dct: { stringValue: String(entry.dct || '') },
        state: { stringValue: String(entry.state || '') },
        uploadedAt: { timestampValue: new Date().toISOString() }
      }
    }
  }));

  await runBatchedWrites(accessToken, docsToDelete, creates, 'IOCS');

  await fetch(`https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/schedule/current`, {
    method: 'PATCH',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      fields: {
        iocsUpdatedAt: { timestampValue: new Date().toISOString() },
        iocsCount: { integerValue: iocsEntries.length.toString() },
        iocsFilename: { stringValue: fileName }
      }
    })
  });

  return { fileName, kind: 'iocs', count: iocsEntries.length, states: statesInFile, status: 'success' };
}

// ─── MIME attachment decode ─────────────────────────────────────────

function decodeAttachmentPart(part) {
  const headerEndIndex = part.indexOf('\r\n\r\n');
  if (headerEndIndex === -1) return null;

  const dataSection = part.substring(headerEndIndex + 4);
  const base64Data = dataSection.split('--')[0].replace(/\r\n/g, '').replace(/\s/g, '');

  try {
    const binaryString = atob(base64Data);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);
    return bytes;
  } catch (e) {
    console.error('Failed to decode base64:', e.message);
    return null;
  }
}

// ─── Main handler ───────────────────────────────────────────────────

export default {
  async email(message, env, ctx) {
    try {
      const serviceAccount = loadServiceAccount(env);

      const rawEmail = await new Response(message.raw).text();

      const contentType = message.headers.get('content-type') || '';
      const boundaryMatch = contentType.match(/boundary="?([^";]+)"?/);
      if (!boundaryMatch) {
        console.error('❌ NO MIME BOUNDARY');
        return;
      }

      const boundary = boundaryMatch[1];
      const parts = rawEmail.split(`--${boundary}`);

      const attachments = [];
      for (const part of parts) {
        if (part.includes('Content-Disposition: attachment') &&
            (part.includes('.xlsx') || part.includes('.xls'))) {
          const fileNameMatch = part.match(/filename="?([^"\r\n]+)"?/);
          if (!fileNameMatch) continue;
          const fileName = fileNameMatch[1];
          attachments.push({ part, fileName });
          console.log(`✅ Attachment: ${fileName} (${part.length} bytes)`);
        }
      }

      if (attachments.length === 0) {
        console.error('❌ NO EXCEL ATTACHMENTS');
        return;
      }

      console.log(`\n📬 Email has ${attachments.length} attachment(s)`);

      const accessToken = await getAccessToken(serviceAccount);
      console.log('✅ Got access token');

      const results = [];

      for (const { part, fileName } of attachments) {
        try {
          const bytes = decodeAttachmentPart(part);
          if (!bytes) {
            results.push({ fileName, status: 'decode_failed' });
            continue;
          }

          const workbook = XLSX.read(bytes, { type: 'array' });
          const classification = classifyWorkbook(workbook, fileName);

          console.log(`\n🔍 ${fileName} → ${classification.kind}${classification.state ? ' (' + classification.state + ')' : ''}`);

          let result;
          if (classification.kind === 'iocs') {
            result = await processIocsFile(fileName, bytes, accessToken);
          } else if (classification.kind === 'schedule') {
            result = await processScheduleFile(fileName, bytes, classification.state, accessToken);
          } else {
            result = { fileName, status: 'unknown_type' };
          }
          results.push(result);
        } catch (err) {
          console.error(`❌ Error processing ${fileName}:`, err.message);
          console.error('Stack:', err.stack);
          results.push({ fileName, status: 'error', error: err.message });
        }
      }

      console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
      console.log(`📊 SUMMARY:`);
      for (const r of results) {
        if (r.status === 'success') {
          if (r.kind === 'iocs') {
            console.log(`   ✅ IOCS: ${r.count} entries [${(r.states || []).join(', ')}] (${r.fileName})`);
          } else {
            console.log(`   ✅ ${r.state} schedule: ${r.count} entries (${r.fileName})`);
          }
        } else {
          console.log(`   ❌ ${r.fileName}: ${r.status}${r.error ? ' — ' + r.error : ''}`);
        }
      }
      console.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

    } catch (error) {
      console.error('❌ Top-level error:', error);
      console.error('Stack:', error.stack);
    }
  }
};
