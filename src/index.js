import * as XLSX from 'xlsx';

export default {
  async email(message, env, ctx) {
    try {
      console.log('Email received from:', message.from);
      console.log('Subject:', message.headers.get('subject'));
      
      // Get Excel attachment
      const attachments = [...message.attachments];
      const excelAttachment = attachments.find(att => 
        att.name.endsWith('.xlsx') || att.name.endsWith('.xls')
      );

      if (!excelAttachment) {
        console.log('No Excel file found in email');
        return;
      }

      console.log('Excel file found:', excelAttachment.name);

      // Read attachment as ArrayBuffer
      const arrayBuffer = await streamToArrayBuffer(excelAttachment);
      
      // Parse Excel
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      console.log('Parsed rows:', rawData.length);

      // Detect state (Nevada or Utah)
      const fileName = excelAttachment.name.toLowerCase();
      let state = 'Nevada'; // Default
      
      if (fileName.includes('utah') || fileName.includes('ut')) {
        state = 'Utah';
      } else if (fileName.includes('nevada') || fileName.includes('nv')) {
        state = 'Nevada';
      } else {
        // Try to detect from data (Utah uses full names, Nevada uses initials)
        const sampleTech = rawData[1]?.[findColumnIndex(rawData[0], 'TECH')];
        if (sampleTech && sampleTech.length > 3) {
          state = 'Utah';
        }
      }

      console.log('Detected state:', state);

      // Find column indices
      const headers = rawData[0] || [];
      const dateCol = findColumnIndex(headers, 'DATE');
      const techCol = findColumnIndex(headers, 'TECH');
      const testCol = findColumnIndex(headers, 'TEST');
      const zipCol = findColumnIndex(headers, 'ZIP');
      const iocsCol = findColumnIndex(headers, 'IOCS');
      const rtCol = findColumnIndex(headers, 'RT');

      if (dateCol === -1 || techCol === -1) {
        console.error('Could not find DATE or TECH columns');
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

      console.log('Parsed schedule entries:', scheduleData.length);

      if (scheduleData.length === 0) {
        console.error('No valid schedule data found in Excel file');
        return;
      }

      // Upload to Firebase
      const firebaseUrl = 'https://work-scheduler-1-default-rtdb.firebaseio.com';
      
      // Delete old state data
      const deleteUrl = `${firebaseUrl}/schedules/${state}.json`;
      const deleteResponse = await fetch(deleteUrl, { method: 'DELETE' });
      console.log('Deleted old data:', deleteResponse.ok);

      // Upload new data
      const uploadUrl = `${firebaseUrl}/schedules/${state}.json`;
      const uploadResponse = await fetch(uploadUrl, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(scheduleData)
      });

      if (!uploadResponse.ok) {
        throw new Error(`Firebase upload failed: ${uploadResponse.status}`);
      }

      console.log('✅ Upload successful:', scheduleData.length, state, 'entries');

    } catch (error) {
      console.error('❌ Error processing email:', error);
    }
  }
};

// Helper: Find column index by searching for keyword in headers
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

// Helper: Parse Excel date to YYYY-MM-DD
function parseExcelDate(value) {
  if (!value) return null;

  // If it's already a string date
  if (typeof value === 'string') {
    const cleaned = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) {
      return cleaned;
    }
    // Try parsing common formats
    const date = new Date(cleaned);
    if (!isNaN(date.getTime())) {
      return formatDate(date);
    }
  }

  // If it's an Excel serial number
  if (typeof value === 'number') {
    const date = new Date((value - 25569) * 86400 * 1000);
    return formatDate(date);
  }

  // If it's already a Date object
  if (value instanceof Date) {
    return formatDate(value);
  }

  return null;
}

// Helper: Format date as YYYY-MM-DD
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Helper: Convert stream to ArrayBuffer
async function streamToArrayBuffer(stream) {
  const reader = stream.getReader();
  const chunks = [];
  
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }
  
  const totalLength = chunks.reduce((acc, chunk) => acc + chunk.length, 0);
  const result = new Uint8Array(totalLength);
  let offset = 0;
  
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  
  return result.buffer;
}
