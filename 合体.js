function runDeleteRenameAndUpload() {

  // フォルダIDを指定してファイル削除関数を呼び出し
  //Logger.log("Starting deleteAllFilesInFolder process...");
  //deleteAllFilesInFolder('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');
  //Logger.log("deleteAllFilesInFolder process completed.");

  //Logger.log("Starting moveMatchingFilesToDestFolder process...");
  //moveMatchingFilesToDestFolder();  
  //Logger.log("moveMatchingFilesToDestFolder process completed.");

  Logger.log("Starting moveFile process...");
  moveFile();
  Logger.log("moveFile process completed.");
  
  Logger.log("Starting convertAllExcelsInFolder process...");
  convertAllExcelsInFolder();
  Logger.log("convertAllExcelsInFolder process completed.");

  Logger.log("Starting moveMatchingFiles process...");
  moveMatchingFiles()
  Logger.log("moveMatchingFiles process completed.");

    // 2. renameMultipleFiles() を実行
  Logger.log("Starting renameMultipleFiles process...");
  renameMultipleFiles();
  Logger.log("renameMultipleFiles process completed.");

  Logger.log("Starting insertArrearsListIntoDocuments process...");
  insertArrearsListIntoDocuments();
  Logger.log("insertArrearsListIntoDocuments process completed.");
}
