const CHANNEL_ACCESS_TOKEN =
  "to replace this with your channel access token";
const replyUrl = "https://api.line.me/v2/bot/message/reply";

const numOfPeople = 13;
const len = numOfPeople + 1;
const cmds = { outResult: "bot o", outMissing: "bot m", reset: "bot r" };
const keywordForCmd = ["指令", "命令", "cmd", "help"];
const keywordUsrMsgStart = "級職姓名";
const propKey = "data";

function doPost(e) {
  // init persistence

  if (PropertiesService.getScriptProperties().getProperty(propKey) === null) {
    let initArr = [];
    for (let i = 0; i < len; i++) {
      initArr.push("");
    }
    PropertiesService.getScriptProperties().setProperty(propKey, JSON.stringify(initArr));
  }

  // extract usrMsg and token from the request

  const msg = JSON.parse(e.postData.contents);
  const replyToken = msg.events[0].replyToken;
  const usrMsg = msg.events[0].message.text;

  if (typeof replyToken === "undefined") {
    return;
  }

  let replyText = "";
  if (usrMsg.startsWith(keywordUsrMsgStart)) {
    // update the persistence

    let data = JSON.parse(PropertiesService.getScriptProperties().getProperty(propKey));

    const usrMsgArr = usrMsg.split("\n\n");
    for (let m of usrMsgArr) {
      let id = parseInt(m.match(/(\d+)/i)[0]);
      data[id] = m;
    }

    PropertiesService.getScriptProperties().setProperty(propKey, JSON.stringify(data));
  } else if (usrMsg.startsWith(cmds.outResult)) {
    // set replyText as result

    const data = JSON.parse(PropertiesService.getScriptProperties().getProperty(propKey));
    for (let i = 1; i < len; i++) {
      if (data[i] !== "") {
        if (replyText !== "") {
          replyText += "\n\n";
        }

        replyText += data[i];
      }
    }
  } else if (usrMsg.startsWith(cmds.outMissing)) {
    // set replyText as missing list

    const data = JSON.parse(PropertiesService.getScriptProperties().getProperty(propKey));
    let missingId = [];
    for (let i = 1; i < len; i++) {
      if (data[i] === "") {
        missingId.push(i);
      }
    }

    replyText = missingId.length + " missing\n" + missingId.join(", ");
  } else if (usrMsg.startsWith(cmds.reset)) {
    // reset the persistence

    PropertiesService.getScriptProperties().deleteProperty(propKey);

    replyText = "Reset complete.";
  } else {
    for (let keyword of keywordForCmd) {
      if (usrMsg.startsWith(keyword)) {
        // set replyText as commands

        replyText = "指令:";
        for (let key in cmds) {
          let value = cmds[key];
          replyText += "\n" + value + ": " + key;
        }
        break;
      }
    }
  }

  // reply

  if (replyText !== "") {
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
