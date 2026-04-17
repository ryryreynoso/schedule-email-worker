import * as XLSX from 'xlsx';

// Service account credentials (will be set as environment variable)
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
  
  // Import private key
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
  
  // Exchange JWT for access token
  const tokenResponse = await fetch(SERVICE_ACCOUNT.token_uri, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
  });
  
  const tokenData = await tokenResponse.json();
  return tokenData.access_token;
}

export default {
  async email(message, env, ctx) {
    try {
      // Get raw email
      const rawEmail = await new Response(message.raw).text();
      
      // Find boundary
      const contentType = message.headers.get('content-type') || '';
      const boundaryMatch = contentType.match(/boundary="?([^";]+)"?/);
      if (!boundaryMatch) return;
      
      const boundary = boundaryMatch[1];
      const parts = rawEmail.split(`--${boundary}`);
      
      // Find Excel attachment
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
      
      // Extract base64 data
      const headerEndIndex = excelPart.indexOf('\r\n\r\n');
      if (headerEndIndex === -1) return;
      
      const dataSection = excelPart.substring(headerEndIndex + 4);
      let base64Data = dataSection.split('--')[0].replace(/\r\n/g, '').replace(/\s/g, '');
      
      // Decode base64
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      
      // Parse Excel
      const workbook = XLSX.read(bytes, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      
      // Detect state from filename
      let state = 'Nevada';
      if (fileName.toLowerCase().includes('utah') || fileName.toLowerCase().includes('ut')) {
        state = 'Utah';
      } else if (fileName.toLowerCase().includes('nevada') || fileName.toLowerCase().includes('nv')) {
        state = 'Nevada';
      }
      
      // Parse schedule data
      const scheduleData = [];
      let currentDate = null;
      
      for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;
        
        const firstCell = String(row[0] || '').trim();
        
        // Check if this is a date header row
        if (firstCell.match(/^(MON|TUE|WED|THU|FRI|SAT|SUN)/i)) {
          const dateMatch = firstCell.match(/(\d{2}-\d{2}-\d{2})/);
          if (dateMatch) {
            const dateParts = dateMatch[1].split('-');
            currentDate = `20${dateParts[2]}-${dateParts[0]}-${dateParts[1]}`;
          }
          continue;
        }
        
        // Skip header rows
        if (firstCell === 'ZIP CODES' || firstCell === 'NEVADA' || firstCell === 'TEST SCHEDULE') {
          continue;
        }
        
        // Parse data rows
        if (currentDate && row.length >= 6) {
          const zip = String(row[1] || '').trim();
          const site = String(row[2] || '').trim();
          const test = String(row[3] || '').trim();
          const iocs = String(row[4] || '').trim();
          const tech = String(row[5] || '').trim();
          const rt = String(row[0] || '').trim();
          
          if (tech && tech !== 'TECH(S)') {
            scheduleData.push({
              date: currentDate,
              tech: tech,
              test: test,
              zip: zip,
              iocs: iocs,
              rt: rt,
              state: state
            });
          }
        }
      }
      
      if (scheduleData.length === 0) return;
      
      // Get access token
      const accessToken = await getAccessToken();
      
      // Upload to Firestore
      const projectId = 'work-schedule-1f2e1';
      
      // Delete existing documents for this state
      const listUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:runQuery`;
      const queryResponse = await fetch(listUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          structuredQuery: {
            from: [{ collectionId: 'schedules' }],
            where: {
              fieldFilter: {
                field: { fieldPath: 'state' },
                op: 'EQUAL',
                value: { stringValue: state }
              }
            }
          }
        })
      });
      
      const queryData = await queryResponse.json();
      
      // Delete each existing document
      if (queryData && Array.isArray(queryData)) {
        for (const item of queryData) {
          if (item.document?.name) {
            await fetch(`https://firestore.googleapis.com/v1/${item.document.name}`, {
              method: 'DELETE',
              headers: { 'Authorization': `Bearer ${accessToken}` }
            });
          }
        }
      }
      
      // Add new documents
      for (const entry of scheduleData) {
        const docData = {
          fields: {
            date: { stringValue: entry.date },
            tech: { stringValue: entry.tech },
            test: { stringValue: entry.test },
            zip: { stringValue: entry.zip },
            iocs: { stringValue: entry.iocs },
            rt: { stringValue: entry.rt },
            state: { stringValue: entry.state }
          }
        };
        
        await fetch(`https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedules`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(docData)
        });
      }
      
      // SUCCESS - email accepted silently
      
    } catch (error) {
      // Silent failure - email still accepted
      console.error('Error:', error);
    }
  }
};
