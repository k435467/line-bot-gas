const CHANNEL_ACCESS_TOKEN = "to replace this with your channel access token";
const replyUrl = "https://api.line.me/v2/bot/message/reply";

const numOfPeople = 13;
const len = numOfPeople + 1;
const cmds = { outResult: "bot o", outMissing: "bot m", reset: "bot r" };
const keywordUsrMsgStart = "級職姓名";

function doPost(e) {
  // --------------------------------------------
  // Extract usrMsg and token from the request
  // --------------------------------------------

  const requestContents = JSON.parse(e.postData.contents);
  const replyToken = requestContents.events[0].replyToken;
  const usrMsg = requestContents.events[0].message.text;

  if (typeof replyToken === "undefined") {
    return;
  }

  // --------------
  // Process cmd
  // --------------

  let replyText = "";
  let scriptProperties = PropertiesService.getScriptProperties();

  switch (usrMsg) {
    case cmds.outResult: {
      // ---------------------------------------------------------------
      // Set replyText as result when received a cmd to output result
      // ---------------------------------------------------------------

      for (let i = 1; i < len; ++i) {
        let msg = scriptProperties.getProperty(i.toString());
        if (msg !== null) {
          if (replyText !== "") {
            replyText += "\n\n";
          }
          replyText += msg;
        }
      }
      break;
    }
    case cmds.outMissing: {
      // ----------------------------------------------------------------------
      // Set replyText as missing list when received a cmd to output missing
      // ----------------------------------------------------------------------

      let missingId = [];
      for (let i = 1; i < len; ++i) {
        let msg = scriptProperties.getProperty(i.toString());
        if (msg == null) {
          missingId.push(i);
        }
      }

      replyText = missingId.length + " missing\n" + missingId.join(", ");
      break;
    }
    case cmds.reset: {
      // -----------------------------------------------------
      // Reset the persistence when received a cmd to reset
      // -----------------------------------------------------

      scriptProperties.deleteAllProperties();
      replyText = "Reset complete.";
      break;
    }
    case "help":
    case "說明":
    case "指令": {
      replyText = "指令:";
      for (let key in cmds) {
        replyText += "\n" + cmds[key] + ": " + key;
      }
      break;
    }
  }

  if (usrMsg.startsWith(keywordUsrMsgStart)) {
    // ------------------------------------------------------
    // Update the persistence when received a report msg
    // ------------------------------------------------------

    const usrMsgArr = usrMsg.split("\n\n");
    for (let msg of usrMsgArr) {
      let id = parseInt(msg.match(/(\d+)/i)[0]);
      scriptProperties.setProperty(id.toString(), msg);
    }
  }

  if (replyText !== "") {
    // --------
    // Reply
    // --------

    UrlFetchApp.fetch(replyUrl, {
      headers: {
        "Content-Type": "application/json; charset=UTF-8",
        Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN,
      },
      method: "post",
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [
          {
            type: "text",
            text: replyText,
          },
        ],
      }),
    });
  }
}
