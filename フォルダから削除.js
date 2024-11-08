function moveMatchingFiles() {
  const sourceFolderId = '1iR7MnkKH4zVsZhB9ugDwYkQLY-wfgotI';  // 元フォルダのID
  const destFolderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';    // 移動先フォルダのID

  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  const destFolder = DriveApp.getFolderById(destFolderId);
  const files = destFolder.getFiles();
  const regexPattern = /\.xls$/;  // `.xls`拡張子のファイルをマッチ

  // フォルダ内のファイルを順に確認
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName(); // ファイル名を取得

    // 正規表現に一致する場合はファイルを移動
    if (regexPattern.test(fileName)) {
      file.moveTo(sourceFolder);                // ファイルを移動先フォルダに移動
      Logger.log("Moved file: " + fileName);  // 移動したファイル名をログに記録
    }
  }
}