// スプレッドシートに未納者リストを書き込む
function writeArrearsList() {
  const arrearsList = createArrearsList();

  // 現在の日付を取得し、フォーマット
  const todayDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
  const spreadsheetName = "未納者リスト" + todayDate;

  // スプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create(spreadsheetName);
  const sheet = spreadsheet.getActiveSheet();

  // ヘッダー行
  const headers = ["設置者CD", "今月未納者名", "郵便番号", "住所_1", "住所_2", "電話番号", 
                   "未納月数", "今月繰越額", "1ヶ月前繰越額", "2ヶ月前繰越額", 
                   "3ヶ月前繰越額", "4ヶ月前繰越額", "5ヶ月前繰越額", "6ヶ月前繰越額", 
                   "処理日", "ID"];
  sheet.appendRow(headers);

  // データ行
  arrearsList.forEach(item => {
    sheet.appendRow([
      item["設置者CD"], item["今月未納者名"], item["郵便番号"], item["住所_1"], item["住所_2"], item["電話番号"],
      item["未納月数"], item["今月繰越額"], item["1ヶ月前繰越額"], item["2ヶ月前繰越額"],
      item["3ヶ月前繰越額"], item["4ヶ月前繰越額"], item["5ヶ月前繰越額"], item["6ヶ月前繰越額"],
      item["処理日"], item["ID"]
    ]);
  });

  // 指定フォルダにスプレッドシートを移動
  const folderId = "1T6QEGuJzRMZEkFXasupayFl3Huqinlr0";
  const folder = DriveApp.getFolderById(folderId);
  const file = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);  // ルートフォルダから削除

  Logger.log("Spreadsheet created with ID: " + spreadsheet.getId());
}
