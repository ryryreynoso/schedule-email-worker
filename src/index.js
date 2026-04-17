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
      
      if (scheduleData.length === 0) {
        message.setReject('No data parsed');
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
        const errorText = await uploadResponse.text();
        message.setReject(`Upload failed: ${uploadResponse.status} ${errorText.substring(0, 50)}`);
        return;
      }
      
      // SUCCESS - email is accepted (no rejection)
      
    } catch (error) {
      message.setReject(`ERROR: ${error.message}`);
    }
  }
};
