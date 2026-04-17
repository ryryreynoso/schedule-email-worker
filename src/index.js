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
        message.setReject(`DEBUG: No boundary found. ${debugInfo.join(' | ')}`);
        return;
      }
      
      const boundary = boundaryMatch[1];
      debugInfo.push(`Boundary: ${boundary}`);
      
      // Split email by boundary
      const parts = rawEmail.split(`--${boundary}`);
      debugInfo.push(`Parts: ${parts.length}`);
      
      // Find Excel attachment
      let excelPart = null;
      let fileName = null;
      
      for (const part of parts) {
        if (part.includes('Content-Disposition: attachment') && 
            (part.includes('.xlsx') || part.includes('.xls'))) {
          excelPart = part;
          
          const fileNameMatch = part.match(/filename="?([^"\r\n]+)"?/);
          if (fileNameMatch) {
            fileName = fileNameMatch[1];
          }
          break;
        }
      }
      
      if (!excelPart || !fileName) {
        message.setReject(`DEBUG: No Excel attachment. ${debugInfo.join(' | ')}`);
        return;
      }
      
      debugInfo.push(`Excel: ${fileName}`);
      
      // Extract base64 data
      const headerEndIndex = excelPart.indexOf('\r\n\r\n');
      if (headerEndIndex === -1) {
        message.setReject(`DEBUG: No header end found. ${debugInfo.join(' | ')}`);
        return;
      }
      
      const dataSection = excelPart.substring(headerEndIndex + 4);
      
      let base64Data = dataSection
        .split('--')[0]
        .replace(/\r\n/g, '')
        .replace(/\s/g, '');
      
      debugInfo.push(`Base64 len: ${base64Data.length}`);
      
      // Decode base64
      let bytes;
      try {
        const binaryString = atob(base64Data);
        bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
      } catch (e) {
        message.setReject(`DEBUG: Base64 decode failed: ${e.message}. ${debugInfo.join(' | ')}`);
        return;
      }
      
      debugInfo.push(`Bytes: ${bytes.length}`);
      
      // Parse Excel
      const workbook = XLSX.read(bytes, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      debugInfo.push(`Rows: ${rawData.length}`);

      // Show actual headers for debugging
      const headers = rawData[0] || [];
      debugInfo.push(`Headers: ${JSON.stringify(headers).substring(0, 200)}`);

      // Detect state
      let state = 'Nevada';
      
      if (fileName.toLowerCase().includes('utah') || fileName.toLowerCase().includes('ut')) {
        state = 'Utah';
      } else if (fileName.toLowerCase().includes('nevada') || fileName.toLowerCase().includes('nv')) {
        state = 'Nevada';
      } else {
        const sampleTech = rawData[1]?.[findColumnIndex(headers, 'TECH')];
        if (sampleTech && sampleTech.length > 3) {
          state = 'Utah';
        }
      }

      debugInfo.push(`State: ${state}`);

      // Find columns
      const dateCol = findColumnIndex(headers, 'DATE');
      const techCol = findColumnIndex(headers, 'TECH');
      const testCol = findColumnIndex(headers, 'TEST');
      const zipCol = findColumnIndex(headers, 'ZIP');
      const iocsCol = findColumnIndex(headers, 'IOCS');
      const rtCol = findColumnIndex(headers, 'RT');

      if (dateCol === -1 || techCol === -1) {
        message.setReject(`DEBUG: Missing cols DATE=${dateCol} TECH=${techCol}. ${debugInfo.join(' | ')}`);
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

      debugInfo.push(`Entries: ${scheduleData.length}`);

      if (scheduleData.length === 0) {
        message.setReject(`DEBUG: No valid data. ${debugInfo.join(' | ')}`);
        return;
      }

      // Upload to Firebase
      const firebaseUrl = 'https://work-scheduler-1-default-rtdb.firebaseio.com';
      
      const deleteUrl = `${firebaseUrl}/schedules/${state}.json`;
      await fetch(deleteUrl, { method: 'DELETE' });

      const uploadUrl = `${firebaseUrl}/schedules/${state}.json`;
      const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(scheduleData)
      });

      debugInfo.push(`Upload: ${uploadResponse.status}`);
      
      if (!uploadResponse.ok) {
        message.setReject(`DEBUG: Upload failed ${uploadResponse.status}. ${debugInfo.join(' | ')}`);
        return;
      }

      // SUCCESS
      message.setReject(`SUCCESS: ${scheduleData.length} ${state} entries uploaded. ${debugInfo.join(' | ')}`);

    } catch (error) {
      message.setReject(`ERROR: ${error.message}. ${debugInfo.join(' | ')}`);
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
