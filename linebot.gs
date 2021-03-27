const CHANNEL_ACCESS_TOKEN =
  "replace your channel access token here";
const replyUrl = "https://api.line.me/v2/bot/message/reply";

const numOfPeople = 13;
const len = numOfPeople + 1;
const keywordUsrMsgStart = "級職姓名";
const keywordForOutResult = "bot o";
const keywordForOutState = "bot s";
const keywordForReset = "bot r";
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

  // extract data from the request

  const msg = JSON.parse(e.postData.contents);
  const replyToken = msg.events[0].replyToken;
  const usrMsg = msg.events[0].message.text;

  if (typeof replyToken === "undefined") {
    return;
  }

  let replyText = "";
  if (usrMsg.startsWith(keywordUsrMsgStart)) {
    // update persistence

    let data = JSON.parse(PropertiesService.getScriptProperties().getProperty(propKey));
    let id = parseInt(usrMsg.match(/(\d+)/i)[0]);
    data[id] = usrMsg;
    PropertiesService.getScriptProperties().setProperty(propKey, JSON.stringify(data));
  } else if (usrMsg.startsWith(keywordForOutResult)) {
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
  } else if (usrMsg.startsWith(keywordForOutState)) {
    // set replyText as state

    const data = JSON.parse(PropertiesService.getScriptProperties().getProperty(propKey));
    let missingId = [];
    for (let i = 1; i < len; i++) {
      if (data[i] === "") {
        missingId.push(i);
      }
    }

    replyText = missingId.length + " missing\n" + missingId.join(", ");
  } else if (usrMsg.startsWith(keywordForReset)) {
    // reset the persistence

    PropertiesService.getScriptProperties().deleteProperty(propKey);

    replyText = "Reset complete.";
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
