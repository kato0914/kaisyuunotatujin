 // 顧客マスタと各月の請求書をロードする
function loadSheets() {
  const folder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');  // フォルダIDを指定
  const files = [];  // ファイルを格納する配列
  const mimeType = MimeType.GOOGLE_SHEETS;  // スプレッドシートのMIMEタイプ

  // フォルダ内の全てのファイルを取得し、スプレッドシートファイルのみを選別
  const allFiles = folder.getFiles();
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    if (file.getMimeType() === mimeType) {
      files.push(file);
    }
  }

  if (files.length === 0) {
    Logger.log("Error: No billing files found in folder");
    return { files: null, masterSheet: null };
  }

  // 顧客マスタファイルを取得
  const masterFiles = DriveApp.getFilesByName("設置者マスタ_消費者");
  if (!masterFiles.hasNext()) {
    Logger.log("Error: Master file not found");
    return { files: files, masterSheet: null };
  }

  const masterFile = masterFiles.next();  // ここでイテレータを使い切る
  const masterSheet = SpreadsheetApp.open(masterFile).getActiveSheet();  // 顧客マスタシート

  Logger.log("Billing files found: " + files.length);
  Logger.log("Master file: " + masterFile.getName());

  return { files: files, masterSheet: masterSheet };
}
