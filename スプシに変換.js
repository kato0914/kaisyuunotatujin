function convertAllExcelsInFolder() {
  const sourceFolderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';
  const destFolderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const destFolder = DriveApp.getFolderById(destFolderId);

  // 保存先フォルダ内のファイルを完全に削除
  // deleteAllFilesInFolder(destFolder);

  const files = sourceFolder.getFiles();
  let convertedFiles = []; // 変換したスプレッドシートのリスト

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();

    // Excel形式のファイルのみ処理
    if (fileName.endsWith('.xls') || fileName.endsWith('.xlsx')) {
      const spreadsheetFile = convertExcelToSpreadsheet(file); // スプレッドシートに変換
      convertedFiles.push(spreadsheetFile); // 変換したファイルをリストに追加
    }
  }

}

// 1つのエクセルファイルをGoogleスプレッドシートに変換する関数
function convertExcelToSpreadsheet(file) {
  try {
    const originalFileName = file.getName(); // 元のファイル名を取得

    // ExcelファイルをGoogleスプレッドシートとして作成
    const excelData = Drive.Files.copy(
      {
        title: originalFileName.replace(/\.[^/.]+$/, ""), // 拡張子を取り除いた名前
        mimeType: MimeType.GOOGLE_SHEETS
      },
      file.getId()
    );

    // 変換されたスプレッドシートのIDを返す
    return excelData.id;
  } catch (error) {
    console.error("エラーが発生しました:", error);
    throw error;
  }
}

// 指定されたフォルダID内の全てのファイルを完全に削除する関数
function deleteAllFilesInFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId); // フォルダIDからフォルダを取得
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    file.setTrashed(true); // ゴミ箱に移動
  }
}
