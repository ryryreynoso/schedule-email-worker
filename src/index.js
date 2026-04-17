import * as XLSX from 'xlsx';

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
      
      if (!excelPart || !fileName) {
        message.setReject('No Excel file');
        return;
      }
      
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
      
      // Find header row (look for row containing "DATE" and "TECH")
      let headerRowIndex = -1;
      for (let i = 0; i < Math.min(10, rawData.length); i++) {
        const row = rawData[i] || [];
        const rowStr = row.join('|').toUpperCase();
        if (rowStr.includes('DATE') && rowStr.includes('TECH')) {
          headerRowIndex = i;
          break;
        }
      }
      
      if (headerRowIndex === -1) {
        message.setReject('Header row not found');
        return;
      }
      
      const headers = rawData[headerRowIndex];
      
      // Detect state from filename
      let state = 'Nevada';
      if (fileName.toLowerCase().includes('utah') || fileName.toLowerCase().includes('ut')) {
        state = 'Utah';
      } else if (fileName.toLowerCase().includes('nevada') || fileName.toLowerCase().includes('nv')) {
        state = 'Nevada';
      }
      
      // Find columns
      const dateCol = findColumnIndex(headers, 'DATE');
      const techCol = findColumnIndex(headers, 'TECH');
      const testCol = findColumnIndex(headers, 'TEST');
      const zipCol = findColumnIndex(headers, 'ZIP');
      const iocsCol = findColumnIndex(headers, 'IOCS');
      const rtCol = findColumnIndex(headers, 'RT');
      
      if (dateCol === -1 || techCol === -1) {
        message.setReject(`Cols not found DATE=${dateCol} TECH=${techCol}`);
        return;
      }
      
      // Parse schedule data (start from row after headers)
      const scheduleData = [];
      for (let i = headerRowIndex + 1; i < rawData.length; i++) {
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
      
      if (scheduleData.length === 0) {
        message.setReject('No data found');
        return;
      }
      
      // Upload to Firebase
      const firebaseUrl = 'https://work-scheduler-1-default-rtdb.firebaseio.com';
      
      await fetch(`${firebaseUrl}/schedules/${state}.json`, { method: 'DELETE' });
      
      const uploadResponse = await fetch(`${firebaseUrl}/schedules/${state}.json`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(scheduleData)
      });
      
      if (!uploadResponse.ok) {
        message.setReject(`Upload failed ${uploadResponse.status}`);
        return;
      }
      
      // SUCCESS
      message.setReject(`SUCCESS: ${scheduleData.length} ${state} entries`);
      
    } catch (error) {
      message.setReject(`ERROR: ${error.message}`);
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
