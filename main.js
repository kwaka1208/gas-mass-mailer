// 設定シート定義
const settings = {
  sheetName: '設定',
  fromName: 'B1',
  fromAddress: 'B2',
  titleText: 'B3',
  documentUrl: 'B4',
  sendListTop: 'A2',
}

// 送信先設定シート定義
const sendList = {
  sheetName: '送信先リスト',
  execute: 0,
  to_address: 1,
  cc_address: 2,
  bcc_address: 3,
  insert1: 4,
  insert2: 5,
  insert3: 6,
  attachment: 7,
  maxAttachments: 3,
  lastColumn: 9
}

const allowedMimeTypes = [
  MimeType.PDF,
  // MimeType.MICROSOFT_EXCEL,
  // MimeType.JPEG,
  // MimeType.PNG
];


function main() {
  const panel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.sheetName);
  const ss_send_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sendList.sheetName);

  const doc = DocumentApp.openByUrl(panel.getRange(settings.documentUrl).getValue());
  const docText = doc.getBody().getText();
  const subject = panel.getRange(settings.titleText).getValue(); // Subject

  /*
    送信元アドレス確認
  */
  var ui = SpreadsheetApp.getUi(); // スプレッドシートのUIを取得
  var senderAddress = Session.getActiveUser().getEmail(); // 現在のユーザーのメールアドレスを取得
  if (panel.getRange(settings.fromAddress).getValue() != '') {
    senderAddress = panel.getRange(settings.fromAddress).getValue()
  }

  // 確認ダイアログを表示
  var response = ui.alert(
    '送信元アドレスの確認',
    '送信者のメールアドレスは ' + senderAddress + ' です。メールを送信しますか？',
    ui.ButtonSet.YES_NO);
  
  // NOなら処理を中断
  if (response == ui.Button.NO) {
    ui.alert('メール送信を中断しました。');
    return
  }

  let startRow = ss_send_list.getRange(settings.sendListTop).getRow();
  let lastRow = ss_send_list.getLastRow()
  Logger.log(`${startRow}から${lastRow}まで`);
  Logger.log('件名： ' + subject);

  const values = ss_send_list.getRange(startRow, 1, lastRow - startRow + 1, sendList.lastColumn + 1).getValues();
  Logger.log(values);

  let count = 0
  for(let i = 0; i < (lastRow - startRow + 1); i++) {
    if (values[i][sendList.execute] == false) continue // チェックが入ってなければスキップ 
    let insert1 = values[i][sendList.insert1]; //宛先
    let insert2 = values[i][sendList.insert2]; //宛先
    let insert3 = values[i][sendList.insert3]; //宛先
    let mailAddress = values[i][sendList.to_address]; // メールアドレス
    if (mailAddress == "") break;

    // 本文作成
    let body = docText
    Logger.log(body);
    body = body.replace(/{埋込1}/g, insert1)
    body = body.replace(/{埋込2}/g, insert2)
    body = body.replace(/{埋込3}/g, insert3)

    // 添付ファイル
    var attachments = [];
    for (j = 0; j < sendList.maxAttachments; j++) {
      fileId = extractFileId(values[i][sendList.attachment + j])
      if (fileId == null) continue
      var attachFile = DriveApp.getFileById(fileId);
      var mimeType = attachFile.getMimeType();
      Logger.log(`ファイル名：${attachFile.getName()}`);
      if (allowedMimeTypes.includes(mimeType)) {
        attachments.push(attachFile.getAs(mimeType));
      } else {
        Logger.log('添付ファイルのMIMEタイプが許可されていません: ' + attachFile.getName());
        ui.alert('エラー: 添付ファイル "' + attachFile.getName() + '" の形式が許可されていません。');
        ui.alert(`${count}件のメール送信を実行しました。`)
        return; // 許可されないMIMEタイプがある場合、処理を中止
      }
    }

    var options = {
      cc: values[i][sendList.cc_address],                 // CC送信先メールアドレス
      bcc: values[i][sendList.bcc_address],               // BCC送信先メールアドレス
      name: panel.getRange(settings.fromName).getValue(), // 送信者名
      from: senderAddress,
      attachments: attachments                            // 添付ファイル
    } 
    Logger.log('送信先: ' + mailAddress)

    // メール送信
    GmailApp.sendEmail(mailAddress, subject, body, options)

    count++
    Logger.log(`startRow + i: ${startRow + i}`)
    Logger.log(`sendList.execute + 1: ${sendList.execute + 1}`)
    ss_send_list.getRange(startRow + i, sendList.execute + 1).setValue(false)
    Logger.log('送信完了: ' + mailAddress)
  }
  ui.alert(`${count}件のメール送信を実行しました。`)
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('拡張メニュー')
    .addItem('メール送信', 'main')
    .addToUi();
}

// URLからファイルIDを抽出する関数
function extractFileId(fileUrl) {
  var regex = /\/d\/([a-zA-Z0-9_-]+)/; // ファイルIDを抽出する正規表現
  var matches = fileUrl.match(regex);
  return matches ? matches[1] : null; // ファイルIDが見つかった場合は返し、なければnullを返す
}
