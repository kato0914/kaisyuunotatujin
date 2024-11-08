// ドキュメントIDを未納月数ごとに指定
const oneMonthDocId = '1CcGF5OLIm4HVof0HETsnN2pKgI4uY9aHZ7TrUtwVGhs';
const twoMonthsDocId = '12DMPnXmc3dXDpJPWVfFLd2uT_hzq_WFRquafqjwmb-k';
const threeMonthsPlusDocId = '1Tt6PTv_17UpW49srSDNbCEynL3_5K_0obj14l1IIU74';

// 保存先フォルダID
const oneMonthFolderId = '1LONjf3ZNZPl5rAlzupnsPllLG6Yjgzmv';
const twoMonthsFolderId = '1zOyWbeJlbo7ULyIyIERF3pjOHaoZjzbi';
const threeMonthsPlusFolderId = '1eBjNUAYLNe0WwH4igLDbkevGU_MBT7Co';


// 未納者リストを指定フォルダに文書として保存
function insertArrearsListIntoDocuments() {

  writeArrearsList();

  // 既存のファイルを完全削除
  deleteFilesPermanently();

  const arrearsList = createArrearsList();

  arrearsList.forEach(item => {
    let docId;
    let folderId;
    let placeholders;
    let dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + 7);  // 振込期限を1週間後に設定
    const formattedDueDate = Utilities.formatDate(dueDate, "Asia/Tokyo", "M月d日") + "（" + getJapaneseDayOfWeek(dueDate) + "）";

    // カンマ区切りの今月繰越額を作成
    const formattedAmount = item["今月繰越額"].toLocaleString();

    // 未納月数によってドキュメントIDとフォルダID、プレースホルダーを設定
    if (item["未納月数"] === 1) {
      docId = oneMonthDocId;
      folderId = oneMonthFolderId;
      placeholders = {
        '{{郵便番号}}': item["郵便番号"],
        '{{住所_1}}': item["住所_1"],
        '{{今月未納者名}}': item["今月未納者名"],
        '{{今月繰越額}}': formattedAmount,
        '{{未納月数}}': item["未納月数"]
      };
    } else if (item["未納月数"] === 2) {
      docId = twoMonthsDocId;
      folderId = twoMonthsFolderId;
      placeholders = {
        '{{郵便番号}}': item["郵便番号"],
        '{{住所_1}}': item["住所_1"],
        '{{今月未納者名}}': item["今月未納者名"],
        '{{今月繰越額}}': formattedAmount,
        '{{未納月数}}': item["未納月数"],
        '{{振込期限}}': formattedDueDate
      };
    } else if (item["未納月数"] >= 3) {
      docId = threeMonthsPlusDocId;
      folderId = threeMonthsPlusFolderId;
      placeholders = {
        '{{郵便番号}}': item["郵便番号"],
        '{{住所_1}}': item["住所_1"],
        '{{今月未納者名}}': item["今月未納者名"],
        '{{今月繰越額}}': formattedAmount,
        '{{未納月数}}': item["未納月数"],
        '{{振込期限}}': formattedDueDate
      };
    } else {
      return;  // 未納月数が0以下の場合は処理をスキップ
    }

    // 現在の日付を取得し、フォーマット
    const todayDate = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");

    // ファイル名を作成
    const fileName = `${item["未納月数"]}ヶ月未納_${item["今月未納者名"]}_${todayDate}`;

    // 指定のドキュメントをコピーして新規作成
    const copiedDoc = DriveApp.getFileById(docId).makeCopy(fileName, DriveApp.getFolderById(folderId));
    const copiedDocId = copiedDoc.getId();
    const doc = DocumentApp.openById(copiedDocId);
    const body = doc.getBody();

    // プレースホルダーを置換
    for (const [key, value] of Object.entries(placeholders)) {
      body.replaceText(key, value);
    }

    Logger.log(`Document created for ${item["今月未納者名"]} in folder with ID: ${folderId}`);
  });
}

function deleteFilesPermanently() {
  [oneMonthFolderId, twoMonthsFolderId, threeMonthsPlusFolderId].forEach(folderId => {
    clearFolderContents(folderId);
  });
}

// 指定フォルダのファイルを完全に削除する
function clearFolderContents(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.setTrashed(true);
  }
}

// 曜日を日本語に変換
function getJapaneseDayOfWeek(date) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return days[date.getDay()];
}