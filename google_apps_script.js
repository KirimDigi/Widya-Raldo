const SPREADSHEET_ID = '1Ys5ip0-x0rFiIR16Z2efyVmUpAALfXPHCLdezZ3HLqY';
const SHEET_NAME = 'Sheet1';

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const record = {};
      headers.forEach((header, index) => {
        let val = row[index];
        if (header === 'timestamp' && val instanceof Date) {
          val = val.toLocaleString('id-ID');
        }
        record[header] = val;
      });
      rows.push(record);
    }
    
    // Return newest comments first
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
    const namaTamu = postData['nama tamu'] || postData['nama'] || '';
    const ucapan = postData['ucapan'] || '';
    const konfirmasi = postData['konfirmasi kehadiran'] || postData['kehadiran'] || '';
    const jumlahTamu = postData['jumlah tamu'] || postData['jumlah_tamu'] || '1';
    
    sheet.appendRow([timestamp, namaTamu, ucapan, konfirmasi, jumlahTamu]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
