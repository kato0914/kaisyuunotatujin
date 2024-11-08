// 顧客マスタとフォルダ内の請求ファイルをロードする関数
function loadSheetsFromFolder() {
  const folder = DriveApp.getFolderById('1McruNM0jy0uJTWP29oy2LPecB-fg0-ow');  // 特定のフォルダID
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

  const masterFileId = "1MG0u82Tufwtj_QpxoKkUyWZqoahn0J3BjU1EDrZoyUc";  // マスターファイルのID
  const masterFile = DriveApp.getFileById(masterFileId);  // IDでファイルを取得
  if (!masterFile) {
    Logger.log("Error: Master file not found");
    return { files: files, masterSheet: null };
  }

  const masterSheet = SpreadsheetApp.open(masterFile).getActiveSheet();  // 顧客マスタシート

  Logger.log("Billing files found: " + files.length);
  Logger.log("Master file: " + masterFile.getName());

  return { files: files, masterSheet: masterSheet };
}

// 顧客マスタのデータを取得する関数
function getMasterData(masterSheet) {
  let data = masterSheet.getDataRange().getValues();
  
  Logger.log("Master Sheet Rows: " + data.length);
  
  let masterData = {};
  
  data.forEach(row => {
    masterData[row[0]] = {
      "設置者名": row[1],
      "郵便番号": row[3],
      "住所_1": row[4],
      "住所_2": row[5],
      "電話番号": row[6]
    };
  });
  
  Logger.log("Master data count: " + Object.keys(masterData).length);  // 顧客データの件数を確認
  
  return masterData;
}

// スプレッドシートの内容を読み込む（特定のファイルのみ処理）
function extractBillingDataFromFolder() {
  const { files, masterSheet } = loadSheetsFromFolder();
  if (!files) {
    Logger.log("Error: Billing file not found in folder");
    return;
  }
  if (!masterSheet) {
    Logger.log("Error: Master file not found");
    return;
  }

  let masterData = getMasterData(masterSheet);

  // "currentMonthBillingDeadlineSum"という名前のファイルを1つ見つける
  const targetFile = files.find(file => file.getName() === "currentMonthBillingDeadlineSum");
  
  if (!targetFile) {
    Logger.log("Error: 'currentMonthBillingDeadlineSum' file not found in folder");
    return;
  }

  const sheet = SpreadsheetApp.open(targetFile).getActiveSheet();
  Logger.log("Opened file: " + targetFile.getName());

  let rows = sheet.getDataRange().getValues();
  Logger.log("Number of rows in file: " + rows.length);

  // 繰越額が1以上のデータのみをフィルタリング
  let allBillingData = rows.slice(1).map(row => ({
    "設置者CD": row[7],
    "繰越額": row[12],
    "請求月": row[2],
    "今回請求額": row[16]
  })).filter(bill => bill["繰越額"] > 0);

  Logger.log("Total billing data after extraction: " + allBillingData.length);
  return { allBillingData, masterData };
}




// 各月の請求ファイルをロードする
function loadPreviousMonthSheets() {
  const previousFiles = [];
  for (let i = 1; i <= 6; i++) {
    const fileName = `before${i}MonthsBillingDeadlineSum`;  // 修正箇所: テンプレートリテラルを使用
    const files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      previousFiles.push(files.next());
    } else {
      Logger.log(`Error: ${fileName} not found`);  // 修正箇所: Logger.logのエラーメッセージ
    }
  }
  return previousFiles;
}


// スプレッドシートの内容を読み込む（過去6ヶ月分）
function extractPreviousBillingData() {
  const previousFiles = loadPreviousMonthSheets();
  let previousBillingData = [];

  previousFiles.forEach(file => {
    const sheet = SpreadsheetApp.open(file).getActiveSheet();
    let rows = sheet.getDataRange().getValues();

    previousBillingData = previousBillingData.concat(rows.slice(1).map(row => ({
      "設置者CD": row[7],
      "繰越額": row[12]  // 繰越額が格納されている列
    })));
  });

  Logger.log("Total previous billing data: " + previousBillingData.length);
  return previousBillingData;
}

// 未納者リストを作成する
function createArrearsList() {
  const { allBillingData, masterData } = extractBillingDataFromFolder();
  const previousBillingData = extractPreviousBillingData();
  
  let arrearsList = [];
  
  let arrearsMap = {};
  
  // 今月の未納データをグループ化
  allBillingData.forEach(bill => {
    const customerId = bill["設置者CD"];
    if (!arrearsMap[customerId]) {
      arrearsMap[customerId] = {
        "設置者CD": customerId,
        "繰越額": [],
        "今月未納者名": masterData[customerId]["設置者名"],
        "郵便番号": masterData[customerId]["郵便番号"],
        "住所_1": masterData[customerId]["住所_1"],
        "住所_2": masterData[customerId]["住所_2"],
        "電話番号": masterData[customerId]["電話番号"],
        "処理日": Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"),
        "ID": generateID()  
      };
    }
    arrearsMap[customerId]["繰越額"].push(bill["繰越額"]);
  });
  
  // 過去6ヶ月のデータを繰越額に追加
  previousBillingData.forEach(bill => {
    const customerId = bill["設置者CD"];
    if (arrearsMap[customerId]) {
      arrearsMap[customerId]["繰越額"].push(bill["繰越額"]);
    }
  });
  
  // 最終的な未納者リストを作成
  Object.keys(arrearsMap).forEach(customerId => {
    let data = arrearsMap[customerId];
    
    // 未納月数を計算（今月の繰越額から数えて1以上の連続した値の数をカウント）
    let unpaidMonthsCount = 0;
    for (let i = 0; i < data["繰越額"].length; i++) {
      if (data["繰越額"][i] >= 1) {
        unpaidMonthsCount++;
      } else {
        break;  // 連続していない場合はカウントを終了
      }
    }

    let arrearsData = {
      "設置者CD": data["設置者CD"],
      "今月未納者名": data["今月未納者名"],
      "郵便番号": data["郵便番号"],
      "住所_1": data["住所_1"],
      "住所_2": data["住所_2"],
      "電話番号": data["電話番号"],
      "未納月数": unpaidMonthsCount,  // 修正された未納月数
      "今月繰越額": data["繰越額"][0] || 0,
      "1ヶ月前繰越額": data["繰越額"][1] || 0,
      "2ヶ月前繰越額": data["繰越額"][2] || 0,
      "3ヶ月前繰越額": data["繰越額"][3] || 0,
      "4ヶ月前繰越額": data["繰越額"][4] || 0,
      "5ヶ月前繰越額": data["繰越額"][5] || 0,
      "6ヶ月前繰越額": data["繰越額"][6] || 0,
      "処理日": data["処理日"],
      "ID": data["ID"]
    };
    
    arrearsList.push(arrearsData);
  });
  
  //Logger.log("Arrears List Content: " + JSON.stringify(arrearsList));
  return arrearsList;
}


// IDを生成する関数
function generateID() {
  return Math.floor(Math.random() * 1000000);  // 一意な連番IDを生成
}
