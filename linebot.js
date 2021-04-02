const CHANNEL_ACCESS_TOKEN = "to replace this with your channel access token";
const replyUrl = "https://api.line.me/v2/bot/message/reply";

const manualMsg = `指令:

bot o
- (output) 顯示統整結果

bot m
- (missing) 顯示缺少人數與學號

bot c
- (clear) 清除紀錄的回報訊息

bot s <學號> <姓名> <電話>
- (set) e.g.
bot s 1 xxx 09xxxxxxxxx

bot i <學號> <體溫> [其他]
- (input) [其他]為選填，如果沒有填寫則預設為目前人在家 e.g.
bot i 1 36.1 在外吃晚餐中
陪同人員: 父母`;

const numOfPeople = 13;
const len = numOfPeople + 1;
const keywordUsrMsgStart = "級職姓名";
const cmds = {
  outResult: "bot o",
  outMissing: "bot m",
  clear: "bot c",
  setNamePhone: "bot s ",
  shortInput: "bot i ",
};

const colForName = 1;
const colForPhoneNumber = 2;
const colForReportMsg = 3;

function reply(token, text) {
  UrlFetchApp.fetch(replyUrl, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    method: "post",
    payload: JSON.stringify({
      replyToken: token,
      messages: [
        {
          type: "text",
          text: text,
        },
      ],
    }),
  });
}

function doPost(e) {
  // ---------------------------
  // Extract usrMsg and token
  // ---------------------------

  const requestContents = JSON.parse(e.postData.contents);
  const replyToken = requestContents.events[0].replyToken;
  const usrMsg = requestContents.events[0].message.text;

  if (typeof replyToken === "undefined") {
    return;
  }

  // --------------
  // Process cmd
  // --------------

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  switch (usrMsg) {
    case cmds.outResult: {
      // ---------------
      // Reply result
      // ---------------

      let replyText = "";
      let data = sheet.getDataRange().getValues();
      if (data[0].length > colForReportMsg - 1) {
        for (let i = 0; i < len - 1; ++i) {
          let str = data[i][colForReportMsg - 1];
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

      let data = sheet.getDataRange().getValues();
      let missingId = [];
      if (data[0].length > colForReportMsg - 1) {
        for (let i = 0; i < len - 1; ++i) {
          if (data[i][colForReportMsg - 1] === "") {
            missingId.push(i + 1);
          }
        }
      } else {
        missingId = Array.from(Array(13).keys()).map((v) => v + 1);
      }

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

      sheet.getRange(1, colForReportMsg, numOfPeople).clear();
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

  if (usrMsg.startsWith(cmds.setNamePhone)) {
    // ---------------------------------------------------------------
    // Set name and phone number - bot s <id> <name> <phone number>
    // ---------------------------------------------------------------

    let msgArr = usrMsg.split(" ");
    if (msgArr.length === 5) {
      let id = parseInt(msgArr[2]);
      let name = msgArr[3];
      let phoneNumber = msgArr[4];
      if (!isNaN(id)) {
        sheet.getRange(id, colForName).setValue(name);
        sheet.getRange(id, colForPhoneNumber).setValue(phoneNumber);
        SpreadsheetApp.flush();
        reply(replyToken, "Set complete.");
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
      let name = sheet.getRange(id, colForName).getValue();
      let phoneNumber = sheet.getRange(id, colForPhoneNumber).getValue();

      let temp = arr[2];
      let misc = "";
      if (usrMsg.length > arr[0].length + 1) {
        misc = usrMsg.slice(arr[0].length + 1);
      } else {
        misc = "目前人在家";
      }
      let reportMsg = `級職姓名：新兵${("000" + id).slice(-3)}${name}
電話：0${phoneNumber}
${misc}
體溫：${temp}`;
      sheet.getRange(id, colForReportMsg).setValue(reportMsg);
      SpreadsheetApp.flush();
      reply(replyToken, reportMsg);
      return;
    }
    reply(replyToken, "Wrong format!");
    return;
  } else if (usrMsg.startsWith(keywordUsrMsgStart)) {
    // ---------------------------------------------------
    // Update the sheet when the report msg is received
    // ---------------------------------------------------

    const usrMsgArr = usrMsg.split("\n\n");
    for (let msg of usrMsgArr) {
      let id = parseInt(msg.match(/(\d+)/i)[0]);
      sheet.getRange(id, colForReportMsg).setValue(msg);
    }
    SpreadsheetApp.flush();
    return;
  }
}
