function doGet() {
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function saveToSheet(data) {
  try {
    // Validate input data
    if (!data || !data.name || !data.email || !data.address || !data.productType) {
      throw new Error('ข้อมูลไม่ครบถ้วน กรุณากรอกข้อมูลให้ครบทุกช่อง');
    }

    // Try to open spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Data');
    if (!sheet) {
      // If sheet doesn't exist, create it
      ss.insertSheet('Data');
      sheet = ss.getSheetByName('Data');
      sheet.appendRow(["Name", "Email", "Address", "Product Type"]);
    }
    
    // Append data
    sheet.appendRow([
      data.name.trim(),
      data.email.trim(),
      data.address.trim(),
      data.productType.trim()
    ]);
    
    return {
      success: true,
      message: 'บันทึกข้อมูลสำเร็จ'
    };

  } catch (error) {
    Logger.log('Error in saveToSheet: ' + error.toString());
    return {
      success: false,
      message: error.message || 'เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง'
    };
  }
}
