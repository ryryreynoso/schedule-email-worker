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
  
  const jwtHeader = btoa(JSON.stringify({ 
    alg: 'RS256', 
    typ: 'JWT',
    kid: serviceAccount.private_key_id
  }));
  
  const jwtClaimSet = btoa(JSON.stringify({
    iss: serviceAccount.client_email,
    sub: serviceAccount.client_email,
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
    while (createIdx < creates.length && writes.length < BATCH_SIZE) {
      writes.push(creates[createIdx++]);
    }
    while (delIdx < deletes.length && writes.length < BATCH_SIZE) {
      writes.push({ delete: deletes[delIdx++] });
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

async function writeWorkerLog(accessToken, event) {
  if (!accessToken) return;

  const docId = `mail_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
  const fields = {
    createdAt: { timestampValue: new Date().toISOString() },
    status: { stringValue: String(event.status || '') },
    subject: { stringValue: String(event.subject || '') },
    from: { stringValue: String(event.from || '') },
    attachmentCount: { integerValue: String(event.attachmentCount || 0) },
    attachmentNames: { arrayValue: { values: (event.attachmentNames || []).map(name => ({ stringValue: String(name) })) } },
    message: { stringValue: String(event.message || '') }
  };

  if (event.results) {
    fields.resultsJson = { stringValue: JSON.stringify(event.results).slice(0, 5000) };
  }

  try {
    await batchWriteFirestore(accessToken, [{
      update: {
        name: `projects/${PROJECT_ID}/databases/(default)/documents/scheduleEmailLogs/${docId}`,
        fields
      }
    }]);
  } catch (err) {
    console.error('⚠️ Failed to write worker diagnostic log:', err.message);
  }
}

// ─── File type detection ────────────────────────────────────────────

function normalizeFileName(fileName) {
  return String(fileName || '')
    .toLowerCase()
    .replace(/%20/g, ' ')
    .replace(/[_\-.()[\]]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function detectStateFromFileName(fileName) {
  const normalized = normalizeFileName(fileName);
  const tokens = normalized.split(' ').filter(Boolean);

  if (tokens.includes('utah') || tokens.includes('ut')) return 'Utah';
  if (tokens.includes('nevada') || tokens.includes('nv')) return 'Nevada';

  return null;
}

function classifyWorkbook(workbook, fileName) {
  const fn = fileName.toLowerCase();

  // Filename-based IOCS detection
  if (fn.includes('iocs') || fn.includes('daily_iocs') || fn.includes('daily iocs')) {
    return { kind: 'iocs' };
  }

  // Content-based IOCS detection: sheets with "ASSIGNED DCT" or "TEST DAY" header.
  // Some master workbooks have title/instruction rows before the actual header.
  const sheetNames = workbook.SheetNames;
  let looksLikeIocs = false;
  for (const sheetName of sheetNames) {
    if (sheetLooksBlank(sheetName)) continue;
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    if (rows.length < 2) continue;
    for (let i = 0; i < Math.min(25, rows.length); i++) {
      const rowStr = (rows[i] || []).map(c => String(c || '').toUpperCase()).join('|');
      if (rowStr.includes('ASSIGNED DCT') || rowStr.includes('TEST DAY') || rowStr.includes('IOCS READINGS')) {
        looksLikeIocs = true;
        break;
      }
    }
    if (looksLikeIocs) break;
  }
  if (looksLikeIocs) return { kind: 'iocs' };

  // Schedule file — determine state
  const detectedState = detectStateFromFileName(fileName);
  if (!detectedState) {
    throw new Error(`Could not detect schedule state from filename "${fileName}". Include UT/Utah or NV/Nevada in the filename.`);
  }
  return { kind: 'schedule', state: detectedState };
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

  const scheduleMetaFields = {
    [`${state.toLowerCase()}UpdatedAt`]: { timestampValue: new Date().toISOString() },
    [`${state.toLowerCase()}Count`]: { integerValue: scheduleData.length.toString() },
    [`${state.toLowerCase()}Filename`]: { stringValue: fileName }
  };
  const scheduleUpdateMask = new URLSearchParams();
  Object.keys(scheduleMetaFields).forEach(field => scheduleUpdateMask.append('updateMask.fieldPaths', field));
  const scheduleMetaResponse = await fetch(`https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/schedule/current?${scheduleUpdateMask}`, {
    method: 'PATCH',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: scheduleMetaFields })
  });
  if (!scheduleMetaResponse.ok) {
    throw new Error(`Schedule metadata update failed: ${scheduleMetaResponse.status} ${await scheduleMetaResponse.text()}`);
  }

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

function formatTime(t) {
  if (!t) return '';
  // If it's already HH:MM format, return as-is
  if (typeof t === 'string' && t.match(/^\d{1,2}:\d{2}/)) return t;
  
  // If it's a Date object from Excel
  if (typeof t === 'object' && t.getHours) {
    return `${String(t.getHours()).padStart(2, '0')}:${String(t.getMinutes()).padStart(2, '0')}`;
  }
  
  // If it's an Excel serial number for time (0-1 range)
  if (typeof t === 'number' && t >= 0 && t < 1) {
    const totalMinutes = Math.round(t * 24 * 60);
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }
  
  return String(t);
}

function normalizeHeader(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

function findIocsHeader(rows) {
  const maxRowsToScan = Math.min(rows.length, 25);

  for (let rowIndex = 0; rowIndex < maxRowsToScan; rowIndex++) {
    const row = rows[rowIndex] || [];
    const headers = row.map(normalizeHeader);
    const indexOf = (...needles) => headers.findIndex(header => needles.some(needle => header === needle || header.includes(needle)));

    const utah = {
      dct: indexOf('TECH', 'DCT'),
      date: indexOf('TEST DAY'),
      finance: indexOf('FINANCE NO', 'FINANCE NUMBER', 'FINANCE'),
      office: indexOf('OFFICE'),
      ein: indexOf('TEST ID', 'EIN'),
      employee: indexOf('EMPLOYEE'),
      rd: indexOf('ROSTER DES', 'ROSTER'),
      bt: indexOf('EMP START TIME', 'START TIME', 'BT'),
      et: indexOf('EMP END TIME', 'END TIME', 'ET'),
      rt: indexOf('READ TIME', 'RT')
    };

    if (utah.date >= 0 && utah.employee >= 0 && utah.dct >= 0) {
      return { format: 'Utah', rowIndex, columns: utah, headers: row };
    }

    const nevada = {
      date: indexOf('TEST DATE', 'DATE'),
      finance: indexOf('FINANCE'),
      office: indexOf('OFFICE', 'LOCATION'),
      ein: indexOf('EIN', 'TEST ID'),
      employee: indexOf('EMPLOYEE'),
      rd: indexOf('ROSTER DES', 'ROSTER'),
      bt: indexOf('BT', 'BEGIN TIME', 'START TIME'),
      et: indexOf('ET', 'END TIME'),
      rt: indexOf('RT', 'READ TIME'),
      dct: indexOf('ASSIGNED DCT', 'DCT')
    };

    if (nevada.dct >= 0 && nevada.date >= 0 && nevada.employee >= 0) {
      return { format: 'Nevada', rowIndex, columns: nevada, headers: row };
    }
  }

  return null;
}

function getCell(row, index) {
  return index >= 0 ? row[index] : '';
}

function sheetLooksBlank(sheetName) {
  return /^(BLANK|Sheet2|Sheet3)$/i.test(String(sheetName || '').trim());
}

function parseIocsExcel(bytes) {
  const workbook = XLSX.read(bytes, { type: 'array' });
  const allEntries = [];

  for (const sheetName of workbook.SheetNames) {
    if (sheetLooksBlank(sheetName)) continue;

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const headerInfo = findIocsHeader(rows);

    if (!headerInfo) {
      console.log(`[IOCS] Sheet ${sheetName}: no IOCS header found`);
      continue;
    }

    let rowsInSheet = 0;
    const { format, rowIndex, columns } = headerInfo;
    console.log(`[IOCS] Sheet ${sheetName}: detected ${format} header at row ${rowIndex + 1}`);

    if (format === 'Utah') {
      for (let i = rowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;
        
        const techName = String(getCell(row, columns.dct) || '').trim().toUpperCase();
        const testDate = getCell(row, columns.date);
        const financeNo = String(getCell(row, columns.finance) || '');
        const office = String(getCell(row, columns.office) || '').trim();
        const testId = String(getCell(row, columns.ein) || '');
        const employeeName = String(getCell(row, columns.employee) || '').trim();
        const rosterDes = String(getCell(row, columns.rd) || '').trim();
        const empStartTime = getCell(row, columns.bt);
        const empEndTime = getCell(row, columns.et);
        const readTime = getCell(row, columns.rt);
        
        if (!employeeName || !techName || !normalizeIocsDate(testDate)) continue;
        
        const entry = {
          weekSheet: sheetName,
          date: normalizeIocsDate(testDate),
          financeCode: financeNo,
          location: office,
          ein: testId,
          employeeName: employeeName,
          rd: rosterDes,
          bt: formatTime(empStartTime),
          et: formatTime(empEndTime),
          rt: formatTime(readTime),
          dct: techName,
          state: 'Utah'
        };
        
        if (allEntries.length < 3) {
          console.log(`[IOCS/Utah] Sample: ${entry.dct} | ${entry.employeeName} | ${entry.date} | RT ${entry.rt}`);
        }
        
        allEntries.push(entry);
        rowsInSheet++;
      }
      
    } else {
      for (let i = rowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const assignedDct = String(getCell(row, columns.dct) || '').trim().toUpperCase();
        if (!assignedDct || assignedDct === 'ASSIGNED DCT') continue;
        if (!/^[A-Z][A-Z\s.\-]*[A-Z]$/.test(assignedDct)) continue;
        if (!assignedDct.includes(' ') && assignedDct.length < 2) continue;

        const rawDate = String(getCell(row, columns.date) || '').trim();
        if (!/\d/.test(rawDate)) continue;

        const location = String(getCell(row, columns.office) || '').trim();
        const entry = {
          weekSheet: sheetName,
          date: normalizeIocsDate(getCell(row, columns.date)),
          financeCode: String(getCell(row, columns.finance) || ''),
          location,
          ein: String(getCell(row, columns.ein) || ''),
          employeeName: String(getCell(row, columns.employee) || '').trim(),
          rd: String(getCell(row, columns.rd) || '').trim(),
          bt: String(getCell(row, columns.bt) || ''),
          et: String(getCell(row, columns.et) || ''),
          rt: String(getCell(row, columns.rt) || ''),
          dct: assignedDct,
          state: 'Nevada'
        };

        if (allEntries.length < 3) {
          console.log(`[IOCS/Nevada] Sample: ${entry.dct} | ${entry.employeeName} | ${entry.date} | RT ${entry.rt}`);
        }

        allEntries.push(entry);
        rowsInSheet++;
      }
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

  const metaFields = {
    iocsUpdatedAt: { timestampValue: new Date().toISOString() },
    iocsCount: { integerValue: iocsEntries.length.toString() },
    iocsFilename: { stringValue: fileName }
  };
  for (const state of statesInFile) {
    const stateKey = state.toLowerCase();
    const stateCount = iocsEntries.filter(entry => entry.state === state).length;
    metaFields[`${stateKey}IocsUpdatedAt`] = { timestampValue: new Date().toISOString() };
    metaFields[`${stateKey}IocsCount`] = { integerValue: stateCount.toString() };
    metaFields[`${stateKey}IocsFilename`] = { stringValue: fileName };
  }

  const iocsUpdateMask = new URLSearchParams();
  Object.keys(metaFields).forEach(field => iocsUpdateMask.append('updateMask.fieldPaths', field));
  const iocsMetaResponse = await fetch(`https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/schedule/current?${iocsUpdateMask}`, {
    method: 'PATCH',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields: metaFields })
  });
  if (!iocsMetaResponse.ok) {
    throw new Error(`IOCS metadata update failed: ${iocsMetaResponse.status} ${await iocsMetaResponse.text()}`);
  }

  return { fileName, kind: 'iocs', count: iocsEntries.length, states: statesInFile, status: 'success' };
}

// ─── MIME attachment decode ─────────────────────────────────────────

function getPartHeader(part, headerName) {
  const headerEndIndex = part.search(/\r?\n\r?\n/);
  const headerText = headerEndIndex === -1 ? part : part.slice(0, headerEndIndex);
  const unfolded = headerText.replace(/\r?\n[ \t]+/g, ' ');
  const prefix = `${headerName.toLowerCase()}:`;
  const line = unfolded.split(/\r?\n/).find(l => l.toLowerCase().startsWith(prefix));
  return line ? line.slice(line.indexOf(':') + 1).trim() : '';
}

function decodeQuotedPrintable(value) {
  const normalized = String(value || '')
    .replace(/=\r?\n/g, '')
    .replace(/=([0-9A-F]{2})/gi, (_m, hex) => String.fromCharCode(parseInt(hex, 16)));
  return new TextEncoder().encode(normalized);
}

function decodeAttachmentPart(part) {
  const headerEndIndex = part.search(/\r?\n\r?\n/);
  if (headerEndIndex === -1) return null;

  const separator = part.match(/\r?\n\r?\n/)[0];
  const dataSection = part.substring(headerEndIndex + separator.length);
  const transferEncoding = getPartHeader(part, 'content-transfer-encoding').toLowerCase();

  try {
    if (!transferEncoding || transferEncoding.includes('base64')) {
      const base64Data = dataSection.replace(/\s/g, '');
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);
      return bytes;
    }
    if (transferEncoding.includes('quoted-printable')) {
      return decodeQuotedPrintable(dataSection);
    }
    return new TextEncoder().encode(dataSection);
  } catch (e) {
    console.error(`Failed to decode attachment (${transferEncoding || 'base64'}):`, e.message);
    return null;
  }
}

function decodeMimeHeaderValue(value) {
  return String(value || '').replace(/=\?([^?]+)\?([BQbq])\?([^?]+)\?=/g, (_match, charset, encoding, encoded) => {
    if (!/^utf-?8$/i.test(charset)) return encoded;
    try {
      if (encoding.toUpperCase() === 'B') {
        const binary = atob(encoded);
        return new TextDecoder('utf-8').decode(Uint8Array.from(binary, c => c.charCodeAt(0)));
      }
      return decodeURIComponent(encoded.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, '%$1'));
    } catch {
      return encoded;
    }
  });
}

function extractAttachmentFileName(part) {
  const unfolded = part.replace(/\r?\n[ \t]+/g, ' ');
  const filenameStar = unfolded.match(/filename\*\s*=\s*(?:UTF-8''|utf-8'')?([^;\r\n]+)/i);
  if (filenameStar) {
    try {
      return decodeURIComponent(filenameStar[1].trim().replace(/^"|"$/g, ''));
    } catch {
      return filenameStar[1].trim().replace(/^"|"$/g, '');
    }
  }

  const filename = unfolded.match(/filename\s*=\s*"([^"]+)"/i) || unfolded.match(/filename\s*=\s*([^;\r\n]+)/i);
  if (filename) return decodeMimeHeaderValue(filename[1].trim().replace(/^"|"$/g, ''));

  const nameStar = unfolded.match(/name\*\s*=\s*(?:UTF-8''|utf-8'')?([^;\r\n]+)/i);
  if (nameStar) {
    try {
      return decodeURIComponent(nameStar[1].trim().replace(/^"|"$/g, ''));
    } catch {
      return nameStar[1].trim().replace(/^"|"$/g, '');
    }
  }

  const name = unfolded.match(/\bname\s*=\s*"([^"]+)"/i) || unfolded.match(/\bname\s*=\s*([^;\r\n]+)/i);
  return name ? decodeMimeHeaderValue(name[1].trim().replace(/^"|"$/g, '')) : '';
}

function getHeaderFromRawEmail(rawEmail, headerName) {
  const lines = String(rawEmail || '').split(/\r?\n/);
  const headerLines = [];

  for (const line of lines) {
    if (line === '') break;
    if (/^[ \t]/.test(line) && headerLines.length > 0) {
      headerLines[headerLines.length - 1] += ` ${line.trim()}`;
    } else {
      headerLines.push(line);
    }
  }

  const prefix = `${headerName.toLowerCase()}:`;
  const match = headerLines.find(line => line.toLowerCase().startsWith(prefix));
  return match ? match.slice(match.indexOf(':') + 1).trim() : '';
}

function getMimeBoundary(contentType) {
  const match = String(contentType || '').match(/boundary=(?:"([^"]+)"|'([^']+)'|([^;\s]+))/i);
  return match ? (match[1] || match[2] || match[3] || '').trim() : '';
}

function extractExcelAttachments(rawEmail, boundary) {
  const topLevelParts = boundary ? rawEmail.split(`--${boundary}`) : [];
  const anyBoundaryParts = rawEmail.split(/\r?\n--[^\r\n]+(?:--)?\r?\n/g);

  const collectFromParts = (parts) => {
    const candidates = [];
    const seen = new Set();

    for (const part of parts) {
      const fileName = extractAttachmentFileName(part);
      if (!/\.(xlsx|xls)$/i.test(fileName)) continue;

      const contentType = getPartHeader(part, 'content-type').toLowerCase();
      const disposition = getPartHeader(part, 'content-disposition').toLowerCase();
      const transferEncoding = getPartHeader(part, 'content-transfer-encoding').toLowerCase();
      const isSpreadsheet =
        contentType.includes('spreadsheet') ||
        contentType.includes('excel') ||
        contentType.includes('octet-stream') ||
        /\.(xlsx|xls)$/i.test(fileName);

      if (!isSpreadsheet) continue;

      const key = `${fileName}:${part.length}:${transferEncoding}`;
      if (seen.has(key)) continue;
      seen.add(key);
      candidates.push({ part, fileName, disposition: disposition || '(none)' });
    }
    return candidates;
  };

  const exactMatches = collectFromParts(topLevelParts);
  return exactMatches.length > 0 ? exactMatches : collectFromParts(anyBoundaryParts);
}

// ─── Main handler ───────────────────────────────────────────────────

export default {
  async email(message, env, ctx) {
    let accessToken = '';
    let subject = '';
    let from = '';
    try {
      const serviceAccount = loadServiceAccount(env);
      accessToken = await getAccessToken(serviceAccount);
      console.log('✅ Got access token');

      const rawEmail = await new Response(message.raw).text();
      subject = decodeMimeHeaderValue(message.headers.get('subject') || getHeaderFromRawEmail(rawEmail, 'subject') || '');
      from = decodeMimeHeaderValue(message.headers.get('from') || getHeaderFromRawEmail(rawEmail, 'from') || '');

      const headerContentType = message.headers.get('content-type') || message.headers.get('Content-Type') || '';
      const rawContentType = getHeaderFromRawEmail(rawEmail, 'content-type');
      const contentType = headerContentType || rawContentType;
      const boundary = getMimeBoundary(contentType);

      if (!boundary) {
        console.error('❌ NO MIME BOUNDARY');
        console.error(`Header content-type: ${headerContentType || '(empty)'}`);
        console.error(`Raw content-type: ${rawContentType || '(empty)'}`);
        await writeWorkerLog(accessToken, {
          status: 'no_boundary',
          subject,
          from,
          message: `Header content-type: ${headerContentType || '(empty)'} | Raw content-type: ${rawContentType || '(empty)'}`
        });
        return;
      }

      const attachments = extractExcelAttachments(rawEmail, boundary);
      for (const attachment of attachments) {
        console.log(`✅ Attachment: ${attachment.fileName} (${attachment.part.length} bytes, disposition: ${attachment.disposition})`);
      }

      if (attachments.length === 0) {
        console.error('❌ NO EXCEL ATTACHMENTS');
        await writeWorkerLog(accessToken, {
          status: 'no_excel_attachments',
          subject,
          from,
          message: `Boundary ${boundary}; raw length ${rawEmail.length}`
        });
        return;
      }

      console.log(`\n📬 Email has ${attachments.length} attachment(s)`);

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

      await writeWorkerLog(accessToken, {
        status: results.some(r => r.status === 'success') ? 'processed' : 'failed',
        subject,
        from,
        attachmentCount: attachments.length,
        attachmentNames: attachments.map(a => a.fileName),
        results
      });

      try {
        await message.forward('ryryreynoso@gmail.com');
        console.log('📨 Forwarded processed email to ryryreynoso@gmail.com');
      } catch (forwardError) {
        console.error('⚠️ Upload processing finished, but forwarding failed:', forwardError.message);
      }

    } catch (error) {
      console.error('❌ Top-level error:', error);
      console.error('Stack:', error.stack);
      await writeWorkerLog(accessToken, {
        status: 'top_level_error',
        subject,
        from,
        message: error.message
      });
    }
  }
};
