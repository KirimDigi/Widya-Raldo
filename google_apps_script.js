const SPREADSHEET_ID = '1Ys5ip0-x0rFiIR16Z2efyVmUpAALfXPHCLdezZ3HLqY';
const SHEET_NAME = 'Sheet1';

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    const rows = [];
    
    // Process all rows
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Skip header row if present in row 0
      if (i === 0 && row[0] && row[0].toString().toLowerCase().trim().indexOf('timestamp') !== -1) {
        continue;
      }
      
      // If row has no timestamp and no name and no ucapan, skip
      if (!row[0] && !row[1] && !row[2]) {
        continue;
      }
      
      let ts = row[0];
      if (ts instanceof Date) {
        ts = ts.toLocaleString('id-ID');
      } else if (ts) {
        ts = ts.toString();
      } else {
        ts = '';
      }
      
      const record = {
        'timestamp': ts,
        'nama': row[1] ? row[1].toString().trim() : '',
        'ucapan': row[2] ? row[2].toString().trim() : '',
        'kehadiran': row[3] ? row[3].toString().trim() : '',
        'jumlah_tamu': row[4] ? row[4].toString().trim() : ''
      };
      rows.push(record);
    }
    
    // Return newest first
    rows.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    let postData;
    if (e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (err) {
        postData = e.parameter;
      }
    } else {
      postData = e.parameter;
    }
    
    const timestamp = new Date();
    const namaTamu = postData['nama'] || postData['nama tamu'] || '';
    const ucapan = postData['ucapan'] || '';
    const konfirmasi = postData['kehadiran'] || postData['konfirmasi kehadiran'] || '';
    const jumlahTamu = postData['jumlah_tamu'] || postData['jumlah tamu'] || '1';
    
    sheet.appendRow([timestamp, namaTamu, ucapan, konfirmasi, jumlahTamu]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
