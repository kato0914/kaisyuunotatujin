// 監視したいフォルダのIDを指定します
var folderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';  // フォルダIDをここに設定してください（sourceFolder）

// 以前のファイルリストを保存するプロパティ
var properties = PropertiesService.getScriptProperties();

// 正規表現でチェックするファイル名のパターン
var fileNamePattern = /^(0?[1-9]|1[0-2])月_請求締切合計D\.xls$/;

var sourceFolder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');
var destFolder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');

function main() {
  // 最初の実行でトリガーを設定
  setTrigger();
  
  // ファイルのチェックを実行
  checkForNewFiles();
}

function checkForNewFiles() {
  // フォルダを取得
  
  // フォルダ内の全ファイルを取得
  var files = sourceFolder.getFiles();
  
  // 現在のファイル名リストを取得
  var currentFileNames = [];
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    currentFileNames.push(fileName);
  }
  
  // 一致するファイルを移動
  //moveMatchingFilesToDestFolder(sourceFolder, destFolder);
  
  // 以前のファイル名リストを取得
  var previousFileNames = properties.getProperty('fileNames');
  
  // 初回実行時はファイル名リストを保存して終了
  if (!previousFileNames) {
    properties.setProperty('fileNames', JSON.stringify(currentFileNames));
    return;
  }
  
  // JSON形式のファイル名リストを配列に変換
  previousFileNames = JSON.parse(previousFileNames);
  
  // 新しいファイルをチェック
  var newFiles = currentFileNames.filter(function(fileName) {
    return previousFileNames.indexOf(fileName) === -1 && fileNamePattern.test(fileName);
  });
  
  // 条件を満たす新しいファイルがあれば全ての処理を実行
  if (newFiles.length > 0) {
    Logger.log('指定のパターンに一致する新しいファイルが追加されました: ' + newFiles);
    
    // runDeleteRenameAndUpload関数を実行
    runDeleteRenameAndUpload(); 
  }
  
  // 現在のファイル名リストを保存
  properties.setProperty('fileNames', JSON.stringify(currentFileNames));
}

function moveMatchingFilesToDestFolder(sourceFolder, destFolder) {

  var sourceFolder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');
  var destFolder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');

  // フォルダが正しく取得できているかを確認
  if (!sourceFolder) {
    Logger.log('Source folder is undefined. Please check the folder ID.');
    return;
  }
  
  if (!destFolder) {
    Logger.log('Destination folder is undefined. Please check the folder ID.');
    return;
  }

  // フォルダ内のファイルを取得
  var files = sourceFolder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    
    // ファイル名がパターンに一致する場合に移動
    if (fileNamePattern.test(fileName)) {
      file.moveTo(destFolder);
      Logger.log(fileName + 'を移動しました');
    }
  }
}


function setTrigger() {
  // トリガーがすでに存在するか確認
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === 'checkForNewFiles';
  });
  
  // トリガーが存在しない場合に設定
  if (!triggerExists) {
    ScriptApp.newTrigger('checkForNewFiles')
      .timeBased()
      .everyMinutes(1)  // 1分ごとに実行
      .create();
    Logger.log('トリガーを設定しました。');
  } else {
    Logger.log('トリガーはすでに設定されています。');
  }
}

function moveFile() {
  var fileNameToMove = "before6MonthsBillingDeadlineSum.xls"; // 移動するファイル名
  var sourceFolderId = '1iR7MnkKH4zVsZhB9ugDwYkQLY-wfgotI';  // 元のフォルダID
  var destFolderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';    // 移動先フォルダID

  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var destFolder = DriveApp.getFolderById(destFolderId);
  var files = destFolder.getFiles();
  var fileFound = false; // ファイルが見つかったかどうかのフラグ

  Logger.log("Searching for file to move: " + fileNameToMove); // 検索開始をログに出力
  
  // フォルダ内のファイルを検索
  while (files.hasNext()) {
    var file = files.next();
    
    // ファイル名が一致する場合に移動
    if (file.getName() === fileNameToMove) {
      Logger.log("Found file: " + file.getName()); // ファイルが見つかったことをログに出力
      file.moveTo(sourceFolder);  // ファイルを移動先フォルダに移動
      Logger.log("Moved " + fileNameToMove); // ログに出力
      fileFound = true; // ファイルが見つかったフラグを更新
      break; // 一度見つけたら終了
    }
  }
}

function renameMultipleFiles() {
  var folderId = '1McruNM0jy0uJTWP29oy2LPecB-fg0-ow';  // フォルダIDをここに設定してください（destFolder）
  var renameMap = {
    "before5MonthsBillingDeadlineSum": "before6MonthsBillingDeadlineSum",
    "before4MonthsBillingDeadlineSum": "before5MonthsBillingDeadlineSum",
    "before3MonthsBillingDeadlineSum": "before4MonthsBillingDeadlineSum",
    "before2MonthsBillingDeadlineSum": "before3MonthsBillingDeadlineSum",
    "before1MonthsBillingDeadlineSum": "before2MonthsBillingDeadlineSum",
    "currentMonthBillingDeadlineSum": "before1MonthsBillingDeadlineSum"
  };

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  // 変更前のファイル名を順番に処理
  for (var oldFileName in renameMap) {
    var newFileName = renameMap[oldFileName]; // 新しいファイル名を取得
    var fileFound = false; // ファイルが見つかったかどうかのフラグ

    Logger.log("Searching for file: " + oldFileName); // 検索開始をログに出力
    
    // フォルダ内のファイルを検索
    while (files.hasNext()) {
      var file = files.next();
      
      // ファイル名が一致する場合に変更
      if (file.getName() === oldFileName) {
        Logger.log("Found file: " + file.getName()); // ファイルが見つかったことをログに出力
        file.setName(newFileName); // ファイル名を変更
        Logger.log("Renamed " + oldFileName + " to " + newFileName); // ログに出力
        fileFound = true; // ファイルが見つかったフラグを更新
        break; // 一度見つけたら次のファイル名に移る
      }
    }
    
    // ファイルリストを再取得する必要があるため、最初に戻す
    files = folder.getFiles();

    // ファイルが見つからなかった場合の処理
    if (!fileFound) {
      Logger.log("File not found: " + oldFileName); // ファイルが見つからない場合のログ
      Logger.log("Check if the file exists in the specified folder and verify the file name.");
    }
  }

  // 月が異なるファイル名を変更
  var monthPattern = /^[1-9]|1[0-2]月_請求締切合計D\.xls$/; // 1月から12月までのファイル名をマッチするパターン
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    
    if (monthPattern.test(fileName)) {
      Logger.log("Found month file: " + fileName); // ファイルが見つかったことをログに出力
      file.setName("currentMonthBillingDeadlineSum"); // ファイル名を変更
      Logger.log("Renamed " + fileName + " to currentMonthBillingDeadlineSum"); // ログに出力
      break;
    }
  }
}
