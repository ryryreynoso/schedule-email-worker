import * as XLSX from 'xlsx';

export default {
  async email(message, env, ctx) {
    let debugInfo = [];
    
    try {
      debugInfo.push('Email received');
      debugInfo.push(`From: ${message.from}`);
      debugInfo.push(`Subject: ${message.headers.get('subject')}`);
      
      // Get raw email as text
      const rawEmail = await new Response(message.raw).text();
      debugInfo.push(`Raw email size: ${rawEmail.length}`);
      
      // Find boundary from Content-Type header
      const contentType = message.headers.get('content-type') || '';
      const boundaryMatch = contentType.match(/boundary="?([^";]+)"?/);
      
      if (!boundaryMatch) {
        message.setReject(`DEBUG: No boundary found in content-type. ${debugInfo.join(' | ')}`);
        return;
      }
      
      const boundary = boundaryMatch[1];
      debugInfo.push(`Boundary: ${boundary}`);
      
      // Split email by boundary
      const parts = rawEmail.split(`--${boundary}`);
      debugInfo.push(`Email parts: ${parts.length}`);
      
      // Find Excel attachment
      let excelPart = null;
      let fileName = null;
      
      for (const part of parts) {
        if (part.includes('Content-Disposition: attachment') && 
            (part.includes('.xlsx') || part.includes('.xls'))) {
          excelPart = part;
          
          // Extract filename
          const fileNameMatch = part.match(/filename="?([^"\r\n]+)"?/);
          if (fileNameMatch) {
            fileName = fileNameMatch[1];
          }
          break;
        }
      }
      
      if (!excelPart || !fileName) {
        message.setReject(`DEBUG: No Excel attachment found. ${debugInfo.join(' | ')}`);
        return;
      }
      
      debugInfo.push(`Excel found: ${fileName}`);
      
      // Extract base64 data
      const lines = excelPart.split('\r\n');
      let base64Data = '';
      let inData = false;
      
      for (const line of lines) {
        if (line.trim() === '') {
          inData = true;
          continue;
        }
        if (inData && !line.startsWith('--')) {
          base64Data += line.trim();
        }
      }
      
      debugInfo.push(`Base64 length: ${base64Data.length}`);
      
      // Decode base64 to ArrayBuffer
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      
      debugInfo.push(`ArrayBuffer size: ${bytes.length}`);
      
      // Parse Excel
      const workbook = XLSX.read(bytes, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      debugInfo.push(`Parsed ${rawData.length} rows`);

      // Detect state
      let state = 'Nevada';
      
      if (fileName.toLowerCase().includes('utah') || fileName.toLowerCase().includes('ut')) {
        state = 'Utah';
      } else if (fileName.toLowerCase().includes('nevada') || fileName.toLowerCase().includes('nv')) {
        state = 'Nevada';
      } else {
        const sampleTech = rawData[1]?.[findColumnIndex(rawData[0], 'TECH')];
        if (sampleTech && sampleTech.length > 3) {
          state = 'Utah';
        }
      }

      debugInfo.push(`State: ${state}`);

      // Find columns
      const headers = rawData[0] || [];
      const dateCol = findColumnIndex(headers, 'DATE');
      const techCol = findColumnIndex(headers, 'TECH');
      const testCol = findColumnIndex(headers, 'TEST');
      const zipCol = findColumnIndex(headers, 'ZIP');
      const iocsCol = findColumnIndex(headers, 'IOCS');
      const rtCol = findColumnIndex(headers, 'RT');

      if (dateCol === -1 || techCol === -1) {
        message.setReject(`DEBUG: Missing columns DATE=${dateCol} TECH=${techCol}. ${debugInfo.join(' | ')}`);
        return;
      }

      // Parse schedule data
      const scheduleData = [];
      for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length === 0) continue;

        const dateValue = row[dateCol];
        const tech = row[techCol];
        
        if (!dateValue || !tech) continue;

        const date = parseExcelDate(dateValue);
        if (!date) continue;

        scheduleData.push({
          date: date,
          tech: String(tech).trim(),
          test: testCol !== -1 ? String(row[testCol] || '').trim() : '',
          zip: zipCol !== -1 ? String(row[zipCol] || '').trim() : '',
          iocs: iocsCol !== -1 ? String(row[iocsCol] || '').trim() : '',
          rt: rtCol !== -1 ? String(row[rtCol] || '').trim() : '',
          state: state
        });
      }

      debugInfo.push(`Schedule entries: ${scheduleData.length}`);

      if (scheduleData.length === 0) {
        message.setReject(`DEBUG: No valid data. ${debugInfo.join(' | ')}`);
        return;
      }

      // Upload to Firebase
      const firebaseUrl = 'https://work-scheduler-1-default-rtdb.firebaseio.com';
      
      const deleteUrl = `${firebaseUrl}/schedules/${state}.json`;
      const deleteResponse = await fetch(deleteUrl, { method: 'DELETE' });
      debugInfo.push(`Delete: ${deleteResponse.ok}`);

      const uploadUrl = `${firebaseUrl}/schedules/${state}.json`;
      const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(scheduleData)
      });

      debugInfo.push(`Upload status: ${uploadResponse.status}`);
      
      if (!uploadResponse.ok) {
        const errorText = await uploadResponse.text();
        message.setReject(`DEBUG: Upload failed ${uploadResponse.status} ${errorText}. ${debugInfo.join(' | ')}`);
        return;
      }

      // SUCCESS - but reject with success message so we can see it
      message.setReject(`SUCCESS: Uploaded ${scheduleData.length} ${state} entries. ${debugInfo.join(' | ')}`);

    } catch (error) {
      message.setReject(`DEBUG ERROR: ${error.message} at ${error.stack}. ${debugInfo.join(' | ')}`);
    }
  }
};

function findColumnIndex(headers, keyword) {
  const searchTerm = keyword.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();
    if (header.includes(searchTerm)) {
      return i;
    }
  }
  return -1;
}

function parseExcelDate(value) {
  if (!value) return null;

  if (typeof value === 'string') {
    const cleaned = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
      return cleaned;
    }
    const date = new Date(cleaned);
    if (!isNaN(date.getTime())) {
      return formatDate(date);
    }
  }

  if (typeof value === 'number') {
    const date = new Date((value - 25569) * 86400 * 1000);
    return formatDate(date);
  }

  if (value instanceof Date) {
    return formatDate(value);
  }

  return null;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}
