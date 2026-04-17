import * as XLSX from 'xlsx';

export default {
  async email(message, env, ctx) {
    let debugInfo = [];
    
    try {
      debugInfo.push('Email received');
      debugInfo.push(`From: ${message.from}`);
      debugInfo.push(`Subject: ${message.headers.get('subject')}`);
      
      // Get Excel attachment - FIXED: message.attachments is async iterable
      const attachments = [];
      for await (const att of message.attachments) {
        attachments.push(att);
      }
      
      debugInfo.push(`Total attachments: ${attachments.length}`);
      
      if (attachments.length > 0) {
        attachments.forEach((att, i) => {
          debugInfo.push(`Attachment ${i + 1}: ${att.name}`);
        });
      }
      
      const excelAttachment = attachments.find(att => 
        att.name.endsWith('.xlsx') || att.name.endsWith('.xls')
      );

      if (!excelAttachment) {
        message.setReject(`DEBUG: No Excel file found. ${debugInfo.join(' | ')}`);
        return;
      }

      debugInfo.push(`Excel found: ${excelAttachment.name}`);

      // Read attachment
      const arrayBuffer = await excelAttachment.arrayBuffer();
      debugInfo.push(`ArrayBuffer size: ${arrayBuffer.byteLength}`);
      
      // Parse Excel
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      debugInfo.push(`Parsed ${rawData.length} rows`);

      // Detect state
      const fileName = excelAttachment.name.toLowerCase();
      let state = 'Nevada';
      
      if (fileName.includes('utah') || fileName.includes('ut')) {
        state = 'Utah';
      } else if (fileName.includes('nevada') || fileName.includes('nv')) {
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
        message.setReject(`DEBUG: Missing columns. ${debugInfo.join(' | ')}`);
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
        message.setReject(`DEBUG: Upload failed ${uploadResponse.status}. ${debugInfo.join(' | ')}`);
        return;
      }

      // SUCCESS - but reject with success message so we can see it
      message.setReject(`SUCCESS: Uploaded ${scheduleData.length} ${state} entries. ${debugInfo.join(' | ')}`);

    } catch (error) {
      message.setReject(`DEBUG ERROR: ${error.message}. ${debugInfo.join(' | ')}`);
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
