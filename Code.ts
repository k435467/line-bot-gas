const spreadSheetID = "to replace this with your spread sheet id";
const CHANNEL_ACCESS_TOKEN = "to replace this with your channel access token";
const replyUrl = "https://api.line.me/v2/bot/message/reply";

const manualMsg = `指令:

bot o
- (output) 顯示回報統整結果

bot b
- (back) 顯示返營統整結果

bot m
- (missing) 顯示缺少人數與學號

bot c
- (clear) 清除紀錄的回報訊息，機器人在新的一天(UTC)收到任意訊息時會自動清除

bot s <學號> <姓名> <電話>
- (set) 設定姓名電話，設定後可使用下面兩個指令 e.g.
bot s 1 xxx 09xxxxxxxxx

bot s <學號> <返營方式>
- (set) 設定返營方式為 北車/板橋/新竹/自行 e.g.
bot s 1 北車

bot i <學號> <體溫> [其他]
- (input) [其他]為選填，如果沒有填寫則預設為目前人在家 e.g.
bot i 1 36.1 在外吃晚餐中
陪同人員: 父母`;

const numOfPeople = 13;
const len = numOfPeople + 1;
const keywordUsrMsgStart = "級職姓名";
const cmds = {
  outResult: "bot o",
  outBackMethod: "bot b",
  outMissing: "bot m",
  clear: "bot c",
  set: "bot s ",
  shortInput: "bot i ",
};

const rangeColForName = 1;
const rangeColForPhoneNumber = 2;
const rangeColForReportMsg = 3;
const rangeColForBackMethod = 4;

const arrColForName = rangeColForName - 1;
const arrColForPhoneNumber = rangeColForPhoneNumber - 1;
const arrColForReportMsg = rangeColForReportMsg - 1;
const arrColForBackMethod = rangeColForBackMethod - 1;

function datesAreOnSameDay(
  a: Date | GoogleAppsScript.Base.Date,
  b: Date | GoogleAppsScript.Base.Date
): boolean {
  return (
    a.getUTCFullYear() === b.getUTCFullYear() &&
    a.getUTCMonth() === b.getUTCMonth() &&
    a.getUTCDate() === b.getUTCDate()
  );
}

function reply(token: string, text: string, text2?: string): void {
  let messagesValue = [
    {
      type: "text",
      text: text,
    },
  ];
  if (text2 !== undefined) {
    messagesValue.push({
      type: "text",
      text: text2,
    });
  }
  UrlFetchApp.fetch(replyUrl, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    method: "post",
    payload: JSON.stringify({
      replyToken: token,
      messages: messagesValue,
    }),
  });
  return;
}

function getMissingList(data: any[][]): number[] {
  let missingId: number[] = [];
  if (data[0].length > arrColForReportMsg) {
    for (let i = 0; i < numOfPeople; ++i) {
      if (data[i][arrColForReportMsg] === "") {
        missingId.push(i + 1);
      }
    }
  } else {
    missingId = Array.from(Array(numOfPeople).keys(), (v) => v + 1);
  }
  return missingId;
}

function getResultText(data: any[][]): string {
  let replyText = "";
  if (data[0].length > arrColForReportMsg) {
    for (let i = 0; i < numOfPeople; ++i) {
      let str = data[i][arrColForReportMsg];
      if (str !== "") {
        if (replyText !== "") {
          replyText += "\n\n";
        }
        replyText += str;
      }
    }
    if (replyText === "") {
      replyText = "There are no reports.";
    }
  }
  return replyText;
}

function doPost(e: any): void {
  // ---------------------------
  // Extract usrMsg and token
  // ---------------------------

  const requestContents = JSON.parse(e.postData.contents);
  const replyToken: string = requestContents.events[0].replyToken;
  const usrMsg: string = requestContents.events[0].message.text;

  if (typeof replyToken === "undefined") {
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // ---------------------------
  // Clear reports if expired
  // ---------------------------

  const curDate = new Date();
  const lastModifyDate = DriveApp.getFileById(spreadSheetID).getLastUpdated();
  if (!datesAreOnSameDay(curDate, lastModifyDate)) {
    sheet.getRange(1, rangeColForReportMsg, numOfPeople).clear();
  }

  // --------------
  // Process cmd
  // --------------

  switch (usrMsg) {
    case cmds.outResult: {
      // ---------------
      // Reply result
      // ---------------

      reply(replyToken, getResultText(sheet.getDataRange().getValues()));
      return;
    }
    case cmds.outBackMethod: {
      // --------------------
      // Reply back method
      // --------------------

      let replyText = "";
      let data = sheet.getDataRange().getValues();
      if (data[0].length > arrColForBackMethod) {
        for (let i = 0; i < numOfPeople; ++i) {
          let str = data[i][arrColForBackMethod];
          if (str !== "") {
            if (replyText !== "") {
              replyText += "\n\n";
            }
            replyText += str;
          }
        }
        reply(replyToken, replyText);
      }
      return;
    }
    case cmds.outMissing: {
      // ---------------------
      // Reply missing list
      // ---------------------

      let missingId = getMissingList(sheet.getDataRange().getValues());

      let replyText = missingId.length + " missing";
      if (missingId.length > 0) {
        replyText += "\n";
        replyText += missingId.join(", ");
      }
      reply(replyToken, replyText);
      return;
    }
    case cmds.clear: {
      // ----------------------------
      // Clear all report messages
      // ----------------------------

      sheet.getRange(1, rangeColForReportMsg, numOfPeople).clear();
      reply(replyToken, "Clear complete.");
      return;
    }
    case "man":
    case "help":
    case "說明":
    case "指令": {
      reply(replyToken, manualMsg);
      return;
    }
  }

  if (usrMsg.startsWith(cmds.set)) {
    // ----------------------------------------------
    // Set name and phone number / Set back method
    // ----------------------------------------------

    let msgArr = usrMsg.split(" ");
    if (msgArr.length === 5) {
      // bot s <id> <name> <phone number>

      let id = parseInt(msgArr[2]);
      let name = msgArr[3];
      let phoneNumber = msgArr[4];
      if (!isNaN(id)) {
        sheet.getRange(id, rangeColForName).setValue(name);
        sheet.getRange(id, rangeColForPhoneNumber).setValue(phoneNumber);
        SpreadsheetApp.flush();
        reply(replyToken, "Set complete.");
        return;
      }
    } else if (msgArr.length === 4) {
      // bot s <id> <back method>

      let id = parseInt(msgArr[2]);
      let method = msgArr[3];
      if (!isNaN(id)) {
        let name = sheet.getRange(id, rangeColForName).getValue();
        let phoneNumber = sheet.getRange(id, rangeColForPhoneNumber).getValue();
        let str = `二兵 ${("000" + id).slice(-3)} ${name}
電話：0${phoneNumber}
無發燒感冒症狀，無接觸返國人員、無軍紀案件、無味覺、嗅覺遺失及不明原因腹瀉，`;
        if (method === "自行") {
          str += "自行返營";
        } else {
          str += "17:30在" + method + "搭專車返營";
        }
        sheet.getRange(id, rangeColForBackMethod).setValue(str);
        SpreadsheetApp.flush();
        reply(replyToken, str);
        return;
      }
    }
    reply(replyToken, "Wrong format!");
    return;
  } else if (usrMsg.startsWith(cmds.shortInput)) {
    // -----------------------------------------
    // Short input - bot i <id> <temp> [misc]
    // -----------------------------------------

    let arr = usrMsg.match(/^bot i (\d+) (\d+(\.\d+)?)/);
    let id = parseInt(arr[1]);
    if (!isNaN(id)) {
      let name = sheet.getRange(id, rangeColForName).getValue();
      let phoneNumber = sheet.getRange(id, rangeColForPhoneNumber).getValue();

      let temp = arr[2];
      let misc = "";
      if (usrMsg.length > arr[0].length + 1) {
        misc = usrMsg.slice(arr[0].length + 1);
      } else {
        misc = "目前人在家";
      }
      let reportMsg = `級職姓名：新兵 ${("000" + id).slice(-3)} ${name}
電話：0${phoneNumber}
${misc}
體溫：${temp}`;

      // Whether to reply summary
      let replySummary = false;
      let missingId = getMissingList(sheet.getDataRange().getValues());
      if (missingId.length === 1 && missingId.includes(id)) {
        replySummary = true;
      }

      sheet.getRange(id, rangeColForReportMsg).setValue(reportMsg);
      SpreadsheetApp.flush();
      if (replySummary) {
        reply(replyToken, reportMsg, getResultText(sheet.getDataRange().getValues()));
      } else {
        reply(replyToken, reportMsg);
      }
      return;
    }
    reply(replyToken, "Wrong format!");
    return;
  } else if (usrMsg.startsWith(keywordUsrMsgStart)) {
    // ---------------------------------------------------
    // Update the sheet when the report msg is received
    // ---------------------------------------------------

    let preNumMissing = getMissingList(sheet.getDataRange().getValues()).length;
    const usrMsgArr = usrMsg.split("\n\n");
    for (let msg of usrMsgArr) {
      let id = parseInt(msg.match(/(\d+)/i)[0]);
      sheet.getRange(id, rangeColForReportMsg).setValue(msg);
    }
    SpreadsheetApp.flush();
    let curNumMissing = getMissingList(sheet.getDataRange().getValues()).length;
    if (curNumMissing === 0 && curNumMissing < preNumMissing) {
      reply(replyToken, getResultText(sheet.getDataRange().getValues()));
    }
    return;
  }
}
