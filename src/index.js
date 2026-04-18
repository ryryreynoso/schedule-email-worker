import * as XLSX from 'xlsx';

const SERVICE_ACCOUNT = {
  type: "service_account",
  project_id: "work-schedule-1f2e1",
  private_key_id: "85f6c481b5a6e7a157fa55fb352202c1a1d3f14b",
  private_key: "-----BEGIN PRIVATE KEY-----\nMIIEugIBADANBgkqhkiG9w0BAQEFAASCBKQwggSgAgEAAoIBAQC0rL7L5622ubEy\nRqN2bb4QXLJwvdxio0ak3iXK78cSHZNPbNnjqg+nKmxo+XIbE0jHnCvS6nKtFKEt\n6RhQsVYkF6FjkvQIz/Io1M1HbXT5QfbBQl4VQqTac/hLSk+94oGI9AdhNXmukguP\nIXdf7Ai7VqnhwXESSoZ2VMhiISpBqRm2NY9fWttTchExEcFS17Vv6pHtfS6nVhfB\nIL/JxtHrKWyq4tMZAzhvMMMvh5EIGeFRaSMkm04iFgQyBOxYR8W4ufgBBm4v2FSw\neY7lSgWDHQ4ln579fSrJm0DzbZoYHi9GpBGpoXwPpjeUNrpNaCO2TYja9Ye+ztM0\ni46ZQMlvAgMBAAECggEAGXf0brKaRpTHNBVzEMxdgNmgVHY0bTnOSTVWJypFFKT9\n/AdAcRsAQ99IvZopie7eUTe5kcIiJ4st7AEyFUXk+lzTukFKjJIg9RKHsWxDk1Ny\nIOy7GHd23PiAur70ngl3mxeokVJuybtP99LkN1hYJBhjsC2K7o96LuS6ro01ngr8\nlNlaZlU1bpy4biDKD8ovKZE6OFoUIUJzG4zWaLsrEDUIPkMZh1I4nr9pXMpfwaJw\nPaCz2o5MBZUWdQ7NlOstV4R/JlCjmlfPqHZYutFy+V/jYpyZUuwYXD81a44ZdaIL\nSfwwVPrp8FP+IXNYwULq2sJGqUI1gfvAis1AjhClCQKBgQDkbHDTxP2CDoakiEzo\nU71SQBF11o7ArYGdqPArzAv14gqXpoeOyQkNwwZUa/79v+shikuNYEiK5ZHM5BLU\niOC67zE+QkK29CVZmhdjn2CQ0Bjd9pPn5mfoArB3HkH+yx/M8FnKgiVdj1LwzFIx\nzpLGLh3bNCQXIJGgXwo+ECtpIwKBgQDKfJnPuGnvk50Tez5YIEQQyEq6q7E0FpiV\nhyweKLYp5F+dxbiTBtGgrwc/EXaMWqRIV9mIaZXVhP7FEWjCiKecI0/O1DWLRoEJ\neI2H5VqapjVosLtf6EcFdAJ4xGZ2rxgExEdjIyVcAxyzmPALvmuB7yof40Izb9zc\no3vh9hdxRQKBgBev1xNevhsafoPZToBZDqzUz/q0QSFh3KsItb8U7biVtBt9vVjl\nJ/cxXhqrCEov+KYFvUfv0BX3MGNa00kO2J8J5sVaAakPMEBWZk6CXHUn3yxFQZku\nn1/Dx6DSlm1hiu6pjeYeENne3u7xgSSBE19RsO7mPUfYrMFAmcNN0fKZAn8a5HGJ\nJPTs3K3/6F5fVem0UOWb5TGjuVyKf2lcmAuZhLsuORRKcp1kudo8hhU4jtFCymgZ\ntewwb3lmsuk27O9VzVrMHWL/HF4G4/voEI33/BsbzF0WX8MO9lldsLfrC1YlS+wv\nPnu3vLITKDy5UpD0sM7nbUddjX3Hz+6kFAsJAoGAHz6uDJC3cacZ2PPQeYEqQMWy\ndNRAq7DByPozZndl1ApsO2dbyDQRAgMppFAmZqfUr6vg8iffaDSxdncA3iniqRrZ\nMsGGaXJN2CKFbGMsy9g6mZmKu3rqecimATM8fJ6PK+P3wBRmw+21+jHomeIVwrCs\nsEMPDDZR1+jBfnCSgYQ=\n-----END PRIVATE KEY-----\n",
  client_email: "firebase-adminsdk-fbsvc@work-schedule-1f2e1.iam.gserviceaccount.com",
  token_uri: "https://oauth2.googleapis.com/token"
};

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
            const rt = String(row[0] || '').trim();
            const zip = String(row[1] || '').trim();
            const site = String(row[2] || '').trim();
            const test = String(row[3] || '').trim();
            const iocs = String(row[4] || '').trim();
            const tech = String(row[5] || '').trim();
            
            if (tech && tech !== 'TECH(S)' && rt !== 'ZIP CODES') {
              scheduleData.push({
                date: currentDate,
                person: tech,
                test: rt,
                zipCode: zip,
                testId: iocs,
                location: site,
                mep: test,
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
      
      const listUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current/rows?pageSize=1000`;
      
      let deletedCount = 0;
      let pageToken = '';
      
      do {
        const url = pageToken ? `${listUrl}&pageToken=${pageToken}` : listUrl;
        const listResponse = await fetch(url, {
          method: 'GET',
          headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        const listData = await listResponse.json();
        
        if (listData.documents) {
          const deletePromises = [];
          for (const doc of listData.documents) {
            if (doc.name.includes('/schedule/current/rows/') && 
                doc.fields?.state?.stringValue === state) {
              deletePromises.push(
                fetch(`https://firestore.googleapis.com/v1/${doc.name}`, {
                  method: 'DELETE',
                  headers: { 'Authorization': `Bearer ${accessToken}` }
                })
              );
            }
          }
          
          if (deletePromises.length > 0) {
            await Promise.all(deletePromises);
            deletedCount += deletePromises.length;
            console.log(`Deleted ${deletePromises.length} documents (total: ${deletedCount})`);
          }
        }
        
        pageToken = listData.nextPageToken || '';
      } while (pageToken);
      
      console.log(`✅ Deletion complete: ${deletedCount} old documents for ${state}`);
      
      const BATCH_SIZE = 50;
      let uploadedCount = 0;
      
      for (let i = 0; i < scheduleData.length; i += BATCH_SIZE) {
        const batch = scheduleData.slice(i, i + BATCH_SIZE);
        const uploadPromises = batch.map(entry => {
          const docData = {
            fields: {
              date: { stringValue: entry.date },
              person: { stringValue: entry.person },
              test: { stringValue: entry.test },
              zipCode: { stringValue: entry.zipCode },
              testId: { stringValue: entry.testId },
              location: { stringValue: entry.location },
              state: { stringValue: entry.state },
              mep: { stringValue: entry.mep || '' },
              time: { stringValue: '' }
            }
          };
          
          return fetch(`https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current/rows`, {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            },
            body: JSON.stringify(docData)
          });
        });
        
        await Promise.all(uploadPromises);
        uploadedCount += batch.length;
        console.log(`Uploaded ${uploadedCount}/${scheduleData.length} documents`);
        
        if (i + BATCH_SIZE < scheduleData.length) {
          await sleep(100);
        }
      }
      
      // Update metadata document - THIS IS THE KEY PART!
      await fetch(`https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current`, {
        method: 'PATCH',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          fields: {
            [`${state.toLowerCase()}UpdatedAt`]: { timestampValue: new Date().toISOString() },
            [`${state.toLowerCase()}Count`]: { integerValue: uploadedCount.toString() },
            [`${state.toLowerCase()}Filename`]: { stringValue: fileName }
          }
        })
      });
      
      console.log(`✅ Upload complete: ${uploadedCount} documents for ${state}`);
      
    } catch (error) {
      console.error('❌ Error:', error);
    }
  }
};
