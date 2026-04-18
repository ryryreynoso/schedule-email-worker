import * as XLSX from 'xlsx';

const SERVICE_ACCOUNT = {
  type: "service_account",
  project_id: "work-schedule-1f2e1",
  private_key_id: "85f6c481b5a6e7a157fa55fb352202c1a1d3f14b",
  private_key: "-----BEGIN PRIVATE KEY-----\nMIIEugIBADANBgkqhkiG9w0BAQEFAASCBKQwggSgAgEAAoIBAQC0rL7L5622ubEy\nRqN2bb4QXLJwvdxio0ak3iXK78cSHZNPbNnjqg+nKmxo+XIbE0jHnCvS6nKtFKEt\n6RhQsVYkF6FjkvQIz/Io1M1HbXT5QfbBQl4VQqTac/hLSk+94oGI9AdhNXmukguP\nIXdf7Ai7VqnhwXESSoZ2VMhiISpBqRm2NY9fWttTchExEcFS17Vv6pHtfS6nVhfB\nIL/JxtHrKWyq4tMZAzhvMMMvh5EIGeFRaSMkm04iFgQyBOxYR8W4ufgBBm4v2FSw\neY7lSgWDHQ4ln579fSrJm0DzbZoYHi9GpBGpoXwPpjeUNrpNaCO2TYja9Ye+ztM0\ni46ZQMlvAgMBAAECggEAGXf0brKaRpTHNBVzEMxdgNmgVHY0bTnOSTVWJypFFKT9\n/AdAcRsAQ99IvZopie7eUTe5kcIiJ4st7AEyFUXk+lzTukFKjJIg9RKHsWxDk1Ny\nIOy7GHd23PiAur70ngl3mxeokVJuybtP99LkN1hYJBhjsC2K7o96LuS6ro01ngr8\nlNlaZlU1bpy4biDKD8ovKZE6OFoUIUJzG4zWaLsrEDUIPkMZh1I4nr9pXMpfwaJw\nPaCz2o5MBZUWdQ7NlOstV4R/JlCjmlfPqHZYutFy+V/jYpyZUuwYXD81a44ZdaIL\nSfwwVPrp8FP+IXNYwULq2sJGqUI1gfvAis1AjhClCQKBgQDkbHDTxP2CDoakiEzo\nU71SQBF11o7ArYGdqPArzAv14gqXpoeOyQkNwwZUa/79v+shikuNYEiK5ZHM5BLU\niOC67zE+QkK29CVZmhdjn2CQ0Bjd9pPn5mfoArB3HkH+yx/M8FnKgiVdj1LwzFIx\nzpLGLh3bNCQXIJGgXwo+ECtpIwKBgQDKfJnPuGnvk50Tez5YIEQQyEq6q7E0FpiV\nhyweKLYp5F+dxbiTBtGgrwc/EXaMWqRIV9mIaZXVhP7FEWjCiKecI0/O1DWLRoEJ\neI2H5VqapjVosLtf6EcFdAJ4xGZ2rxgExEdjIyVcAxyzmPALvmuB7yof40Izb9zc\no3vh9hdxRQKBgBev1xNevhsafoPZToBZDqzUz/q0QSFh3KsItb8U7biVtBt9vVjl\nJ/cxXhqrCEov+KYFvUfv0BX3MGNa00kO2J8J5sVaAakPMEBWZk6CXHUn3yxFQZku\nn1/Dx6DSlm1hiu6pjeYeENne3u7xgSSBE19RsO7mPUfYrMFAmcNN0fKZAn8a5HGJ\nJPTs3K3/6F5fVem0UOWb5TGjuVyKf2lcmAuZhLsuORRKcp1kudo8hhU4jtFCymgZ\ntewwb3lmsuk27O9VzVrMHWL/HF4G4/voEI33/BsbzF0WX8MO9lldsLfrC1YlS+wv\nPnu3vLITKDy5UpD0sM7nbUddjX3Hz+6kFAsJAoGAHz6uDJC3cacZ2PPQeYEqQMWy\ndNRAq7DByPozZndl1ApsO2dbyDQRAgMppFAmZqfUr6vg8iffaDSxdncA3iniqRrZ\nMsGGaXJN2CKFbGMsy9g6mZmKu3rqecimATM8fJ6PK+P3wBRmw+21+jHomeIVwrCs\nsEMPDDZR1+jBfnCSgYQ=\n-----END PRIVATE KEY-----\n",
  client_email: "firebase-adminsdk-fbsvc@work-schedule-1f2e1.iam.gserviceaccount.com",
  token_uri: "https://oauth2.googleapis.com/token"
};


// Normalize tech names to initials for Nevada
function normalizePersonName(name) {
  const upper = name.trim().toUpperCase();
  
  // Map full names to initials
  const nameMap = {
    'JASON': 'JRA',
    'JASON ALVAREZ': 'JRA',
    'JERRY': 'JA', 
    'JERRY ANGLO': 'JA',
    'JEFF': 'JN',
    'JEFF NIZNICK': 'JN',
    'GRACE': 'GA',
    'GRACE AGRESOR': 'GA',
    'SHEILA': 'SV',
    'SHEILA VELASQUEZ': 'SV',
    'RYAN': 'RR',
    'RYAN REYNOSO': 'RR',
    'CHRISTIAN': 'CA',
    'CHRISTIAN ALBERT': 'CA',
    'KRISTINE': 'KK',
    'KRISTINE KIESLING': 'KK',
    'BUDD': 'KK'
  };
  
  return nameMap[upper] || upper;
}

async function getAccessToken() {
  const jwtHeader = btoa(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const now = Math.floor(Date.now() / 1000);
  const jwtClaimSet = btoa(JSON.stringify({
    iss: SERVICE_ACCOUNT.client_email,
    scope: 'https://www.googleapis.com/auth/datastore',
    aud: SERVICE_ACCOUNT.token_uri,
    exp: now + 3600,
    iat: now
  }));
  
  const unsignedToken = `${jwtHeader}.${jwtClaimSet}`;
  
  const pemHeader = "-----BEGIN PRIVATE KEY-----";
  const pemFooter = "-----END PRIVATE KEY-----";
  const pemContents = SERVICE_ACCOUNT.private_key.substring(
    pemHeader.length,
    SERVICE_ACCOUNT.private_key.length - pemFooter.length - 1
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
  
  const tokenResponse = await fetch(SERVICE_ACCOUNT.token_uri, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
  });
  
  const tokenData = await tokenResponse.json();
  return tokenData.access_token;
}

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// NEW: Batch write function using Firestore's commit API
async function batchWriteFirestore(projectId, accessToken, writes) {
  const batchUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:commit`;
  
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

export default {
  async email(message, env, ctx) {
    try {
      const rawEmail = await new Response(message.raw).text();
      
      const contentType = message.headers.get('content-type') || '';
      const boundaryMatch = contentType.match(/boundary="?([^";]+)"?/);
      if (!boundaryMatch) return;
      
      const boundary = boundaryMatch[1];
      const parts = rawEmail.split(`--${boundary}`);
      
      let excelPart = null;
      let fileName = null;
      
      for (const part of parts) {
        if (part.includes('Content-Disposition: attachment') && 
            (part.includes('.xlsx') || part.includes('.xls'))) {
          excelPart = part;
          const fileNameMatch = part.match(/filename="?([^"\r\n]+)"?/);
          if (fileNameMatch) fileName = fileNameMatch[1];
          break;
        }
      }
      
      if (!excelPart || !fileName) return;
      
      const headerEndIndex = excelPart.indexOf('\r\n\r\n');
      if (headerEndIndex === -1) return;
      
      const dataSection = excelPart.substring(headerEndIndex + 4);
      let base64Data = dataSection.split('--')[0].replace(/\r\n/g, '').replace(/\s/g, '');
      
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      
      const workbook = XLSX.read(bytes, { type: 'array' });
      
      let state = 'Nevada';
      if (fileName.toLowerCase().includes('utah') || fileName.toLowerCase().includes('ut')) {
        state = 'Utah';
      } else if (fileName.toLowerCase().includes('nevada') || fileName.toLowerCase().includes('nv')) {
        state = 'Nevada';
      }
      
      console.log(`Processing ${state} file: ${fileName}`);
      console.log(`Total sheets: ${workbook.SheetNames.length}`);
      
      const scheduleData = [];
      
      for (const sheetName of workbook.SheetNames) {
        console.log(`Processing sheet: ${sheetName}`);
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
              console.log(`Found date: ${currentDate}`);
            }
            continue;
          }
          
          if (firstCell === 'ZIP CODES' || firstCell === 'NEVADA' || firstCell === 'UTAH' || firstCell === 'TEST SCHEDULE') {
            continue;
          }
          
          if (currentDate && row.length >= 6) {
            const testName = String(row[0] || '').trim();
            const zip = String(row[1] || '').trim();
            const site = String(row[2] || '').trim();
            const type = String(row[3] || '').trim();
            const testId = String(row[4] || '').trim();
            const tech = String(row[5] || '').trim();
            
            if (tech && tech !== 'TECH(S)' && testName && testName !== 'ZIP CODES') {
              const normalizedPerson = normalizePersonName(tech);
              scheduleData.push({
                date: currentDate,
                person: normalizedPerson,
                test: testName,
                zipCode: zip,
                testId: testId,
                location: site,
                mep: type,
                state: state
              });
              rowsInSheet++;
            }
          }
        }
        console.log(`Sheet ${sheetName}: Added ${rowsInSheet} rows`);
      }
      
      console.log(`Total entries parsed: ${scheduleData.length}`);
      
      if (scheduleData.length === 0) return;
      
      const accessToken = await getAccessToken();
      const projectId = 'work-schedule-1f2e1';
      
      // STEP 1: Get all existing documents for this state
      console.log(`📋 Fetching existing ${state} documents...`);
      const listUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current/rows?pageSize=1000`;
      
      let docsToDelete = [];
      let pageToken = '';
      
      do {
        const url = pageToken ? `${listUrl}&pageToken=${pageToken}` : listUrl;
        const listResponse = await fetch(url, {
          method: 'GET',
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        const listData = await listResponse.json();
        
        if (listData.documents) {
          for (const doc of listData.documents) {
            if (doc.name.includes('/schedule/current/rows/') && 
                doc.fields?.state?.stringValue === state) {
              docsToDelete.push(doc.name);
            }
          }
        }
        
        pageToken = listData.nextPageToken || '';
      } while (pageToken);
      
      console.log(`🗑️  Found ${docsToDelete.length} old ${state} documents to delete`);
      
      // STEP 2: Prepare batch writes (deletes + creates)
      // Firestore allows 500 operations per batch
      const BATCH_SIZE = 500;
      let totalOperations = 0;
      
      // Process in batches
      let deleteIndex = 0;
      let createIndex = 0;
      let batchNumber = 1;
      
      while (deleteIndex < docsToDelete.length || createIndex < scheduleData.length) {
        const writes = [];
        
        // Add deletes to this batch (up to BATCH_SIZE total operations)
        while (deleteIndex < docsToDelete.length && writes.length < BATCH_SIZE) {
          writes.push({
            delete: docsToDelete[deleteIndex]
          });
          deleteIndex++;
        }
        
        // Add creates to this batch (if space remaining)
        while (createIndex < scheduleData.length && writes.length < BATCH_SIZE) {
          const entry = scheduleData[createIndex];
          
          writes.push({
            update: {
              name: `projects/${projectId}/databases/(default)/documents/schedule/current/rows/doc_${Date.now()}_${createIndex}_${Math.floor(Math.random() * 10000)}`,
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
          });
          createIndex++;
        }
        
        // Execute this batch
        if (writes.length > 0) {
          console.log(`📦 Batch ${batchNumber}: Processing ${writes.length} operations...`);
          await batchWriteFirestore(projectId, accessToken, writes);
          totalOperations += writes.length;
          console.log(`✅ Batch ${batchNumber} complete (${totalOperations} total operations)`);
          batchNumber++;
          
          // Small delay between batches to avoid rate limits
          if (deleteIndex < docsToDelete.length || createIndex < scheduleData.length) {
            await sleep(100);
          }
        }
      }
      
      console.log(`✅ All batches complete: Deleted ${docsToDelete.length}, Created ${scheduleData.length}`);
      
      // STEP 3: Update metadata document
      await fetch(`https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current`, {
        method: 'PATCH',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          fields: {
            [`${state.toLowerCase()}UpdatedAt`]: { timestampValue: new Date().toISOString() },
            [`${state.toLowerCase()}Count`]: { integerValue: scheduleData.length.toString() },
            [`${state.toLowerCase()}Filename`]: { stringValue: fileName }
          }
        })
      });
      
      console.log(`🎉 Upload complete for ${state}: ${scheduleData.length} documents`);
      
    } catch (error) {
      console.error('❌ Error:', error);
      console.error('Stack:', error.stack);
    }
  }
};
