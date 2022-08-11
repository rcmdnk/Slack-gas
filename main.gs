var users = new Map();
var bots = new Map();
var nMessages = 0;

function run(){
  getUsers();
  const channels = getChannels();
  nMessages = 0;
  channels.forEach(function(name, id){
    if(CHANNEL_INCLUDE.length > 0 && ! (CHANNEL_INCLUDE.includes(name))) return;
    if(CHANNEL_EXCLUDE.includes(name)) return;
    console.log('Retrieving messages from ' + name);
    nMessagesPre = nMessages;
    retrieveMessages(id, name, nMessages);
    console.log(nMessages - nMessagesPre + ' messages were retrieved from channel: ' + name);
    if(nMessages >= MAX_MESSAGES)return;
  });
}

function getTimeZone(){
  if(TIME_ZONE === null) return Session.getScriptTimeZone();
  return TIME_ZONE;
}


function request(method, payload={}, urlPrefix='https://slack.com/api/', jsonParse=true) {
  const url = urlPrefix + method;
  const options = {
    headers: {
      Authorization: 'Bearer ' + SLACK_TOKEN
    },
    payload: payload
  };
  const response = UrlFetchApp.fetch(url, options);
  let result = response;
  if(jsonParse){
    result = JSON.parse(response.getContentText());
    if(!result.ok){
      let error = result;
      if('error' in result){
        error = result.error;
      }
      console.log('API Error for ' + method + ': ' + error);
    }
  }
  return result;
}

function paginate(method, payload){
  let results = [];
  payload['limit'] = PIGINATE_LIMIT;
  payload['cursor'] = '';
  while(true){
    const result = request(method, payload);
    if(!result.ok) break;
    results.push(result);
    if(!('response_metadata' in result)) break;
    payload['cursor'] = result['response_metadata']['next_cursor']
    if(payload['cursor'] != ""){
      continue;
    }
    break;
  }
  return results;
}

function getUsers(){
  const method = 'users.list';
  let payload = {};
  const results = paginate(method, payload);
  results.forEach(function(result){
    result.members.forEach(function(f){
      users.set(f['id'], f['name']);
    });
  });
}

function getChannels(){
  const method = 'conversations.list';
  let payload = {
      exclude_archived: CHANNEL_EXCLUDE_ARCHIVED,
      types: CHANNEL_TYPES
  };
  const results = paginate(method, payload);
  const channels = new Map();
  results.forEach(function(result){
    result.channels.forEach(function(channel){
      let name = '';
      if('name' in channel){
        name = channel['name']
      }else if('user' in channel){
        name = getUser(channel['user'], users);
      }
      if(name == '') name = channel['id'];
      channels.set(channel['id'], name);
    });
  });
  return channels;
}

function getMessages(channelId, oldest=0){
  // `limit` does not work for history? (it seems always 100)
  const method = 'conversations.history';
  let payload = {
    channel: channelId,
    oldest: oldest,
    inclusive: true,
  }
  const results = paginate(method, payload);
  let messages = [];
  results.forEach(function(result){
    messages = messages.concat(result.messages);
  })
  messages.sort(function(a, b){
    return a.ts - b.ts;
  })
  return messages;
}

function getReplies(channelId, ts){
  const method = 'conversations.replies';
  let payload = {
    channel: channelId,
    ts: ts,
  }
  const results = paginate(method, payload);
  let messages = [];
  results.forEach(function(result){
    messages = messages.concat(result.messages);
  })
  messages = messages.filter(function(message){
    return message.ts != ts;
  })
  messages.sort(function(a, b){
    return a.ts - b.ts;
  })
  return messages;
}

function getSheet(sheetName, cols=[], timeFormat='yyyy/MM/dd HH:mm:ss', makeSheet=true) {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(getTimeZone());
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet) {
    if(!makeSheet){
      return null;
    }
    sheet = ss.insertSheet(sheetName);
    // remain 1 additional row, to frozen first row
    // (need additional rows to fix rows)
    sheet.deleteRows(1, sheet.getMaxRows() - 2);
    const nCols = cols.length != 0 ? cols.length: 1;
    sheet.deleteColumns( 1, sheet.getMaxColumns() - nCols);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
    sheet.getRange('A:A').setNumberFormat(timeFormat);
    sheet.setFrozenRows(1);
  }else if(UPDATE_COLUMN_NAME){
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  }
  return sheet;
}

function fillValues(sheet, data, sortColumn=1){
  if (data.length == 0)return;
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setNumberFormat("@").setRichTextValues(data);
  sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).sort(sortColumn);
}


function getDate(unixtime=null, timezone=null, format='yyyy/MM/dd HH:mm:ss'){
  if (!timezone) timezone = getTimeZone();
  if (!unixtime) {
    return Utilities.formatDate(new Date(), timezone, format);
  }
  return Utilities.formatDate(new Date(unixtime * 1000), timezone, format);
}

function getUser(user){
  if(users.has(user)){
    user = '@' + users.get(user);
  }
  if(user === null)user = '';
  return user;
}

function getBot(botId){
  if(bots.has('botId')) return bots[botId];

  const method = 'bots.info';
  let payload = {
    bot: botId
  };
  const results = request(method, payload);
  if('bot' in results && 'name' in results.bot){
    bots.set(botId, results.bot['name']);
    return results.bot['name'];
  }
  return botId;
}

function parseMessage(message){
  if(message === null)message = "";
  message = message
  .replace(/&lt;/g, "<")
  .replace(/&gt;/g, ">")
  .replace(/&amp;/g, "&");
  users.forEach(function(name, id){
    const reg = new RegExp("<@" + id + ">", "g");
    message = message.replace(reg, '@' + name)
  });
  return message;
}

function getDriveFolder(name=null, folder=null){
  if(folder === null){
    const ss = SpreadsheetApp.getActive();
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    folder = parents.next();
  }
  if(name !== null){
    let folders = folder.getFoldersByName(name);
    if(folders.hasNext()){
      folder = folders.next();
    }else{
      folder = folder.createFolder(name);
    }
  }
  return folder;
}

function download(url, fileName, folder){
  if(REMOVE_SAME_NAME_FILES){
    const files = folder.getFilesByName(fileName);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
  }
  const result = request(url, {}, '', false);
  const file = folder.createFile(result);
  file.setName(fileName);
  return file.getUrl();
}

function saveMessage(message, folder){
  const ts = message.ts;
  const fileName = ts + '.json';
  if(REMOVE_OLD_MESSAGE){
    const files = folder.getFilesByName(fileName);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
  }
  const file = folder.createFile(fileName, JSON.stringify(message));
  file.setName(fileName);
  return file.getUrl();
}

function getMessageData(messages, timesEdited, channelName){
  const channelFolder = getDriveFolder(channelName);
  const downloadFolder = getDriveFolder('files', channelFolder);
  const messageFolder = getDriveFolder('messages', channelFolder);

  const data = new Array();
  const threadTsArray = new Array();
  messages.forEach(function(message){
    if(nMessages >= MAX_MESSAGES) return;
    if(message.type != "message"){
      console.log("Non 'message' type message in channel: ", name, '\n', message);
    }else{
      let editedTs = '';
      if('edited' in message){
        editedTs = message.edited.ts;
      }
      let threadTs = '';
      if('thread_ts' in message){
        threadTs = message['thread_ts'];
        threadTsArray.push(threadTs);
      }
      if(timesEdited.find(e => message.ts == e[0] && (editedTs === null || editedTs == e[1])))return;

      const datetime = getDate(message.ts);

      let user = '';
      if('username' in message){
        user = message['username'];
      }else if('user' in message){
        user = getUser(message.user);
      }else if('bot_id' in message){
        user = getBot(message.bot_id);
      }

      let text = '';
      if('attachments' in message){
        message['attachments'].forEach(function(attachment){
          if('text' in attachment){
            if(text != '') text == '\n';
            text += attachment['text'];
          }
        });
      }
      if('text' in message){
        if(text != '') text == '\n';
        text += message.text;
      }
      text = parseMessage(text);
      let messageText = SpreadsheetApp.newRichTextValue().setText(text);
      let files = new Array();
      if('files' in message){
        if(text != "")text += "\n";
        text += "Files: ";
        message.files.forEach(function(file){
          let fileUrl = '';
          if(SAVE_FILE) fileUrl = download(file.url_private_download, file.name, downloadFolder);
          if(!text.endsWith("Files: ")) text += ', ';
          files.push([file.name, fileUrl, text.length, text.length + file.name.length]);
          text += file.name
        })
        messageText = SpreadsheetApp.newRichTextValue().setText(text);
        if(SAVE_FILE){
          files.forEach(function(f){
            messageText = messageText.setLinkUrl(f[2], f[3], f[1]);
          });
        }
      }
      messageText = messageText.build();

      let messageTs = SpreadsheetApp.newRichTextValue().setText(message.ts);
      if(SAVE_MESSAGE_JSON){
        const messageUrl = saveMessage(message, messageFolder);
        messageTs = messageTs.setLinkUrl(0, message.ts.length, messageUrl);
      }
      messageTs = messageTs.build()

      data.push([
        SpreadsheetApp.newRichTextValue().setText(datetime).build(),
        SpreadsheetApp.newRichTextValue().setText(user).build(),
        messageText,
        SpreadsheetApp.newRichTextValue().setText(threadTs).build(),
        messageTs,
        SpreadsheetApp.newRichTextValue().setText(editedTs).build()]);
    }
    nMessages += 1;
  });
  return [data, threadTsArray];
}

function fillMessages(messages, timesEdited, channelName, sheet){
  let threadTs = new Array();
  if(FILL_EACH){
    messages.forEach(function(message){
      if(nMessages >= MAX_MESSAGES) return;
      let ret = getMessageData([message], timesEdited, channelName);
      let data = ret[0];
      if(ret[1].length == 1) threadTs.push(ret[1][0]);
      fillValues(sheet, data, 5);
    });
  }else{
    let ret = getMessageData(messages, timesEdited, channelName);
    let data = ret[0];
    threadTs = ret[1];
    fillValues(sheet, data, 5);
  }
  return threadTs;
}

function retrieveMessages(id, name){
  const cols = ['Datetime (' + getTimeZone() + ')', 'User/Bot', 'Message', 'ThreadTS', 'UnixTime', 'Edited'];
  const sheet = getSheet(name, cols);

  sheet.setColumnWidth(1, DATETIME_COLUMN_WIDTH);
  sheet.setColumnWidth(2, USER_COLUMN_WIDTH);
  sheet.setColumnWidth(3, MESSAGE_COLUMN_WIDTH);
  sheet.setColumnWidth(4, THREAD_TS_COLUMN_WIDTH);
  sheet.setColumnWidth(5, UNIXTIME_COLUMN_WIDTH);
  sheet.setColumnWidth(6, EDITED_TS_COLUMN_WIDTH);

  let oldest = '0';
  if(!FULL_CHECK){
    const times = sheet.getRange('E:E').getValues().flat();
    if(times.length > 1){
      oldest = times[times.length - 1];
    }
  }
  if(COVERAGE !== null && COVERAGE != 0){
    const coverageStart = Date.now() / 1000. - parseFloat(COVERAGE)
    if(coverageStart > parseFloat(oldest)){
      oldest = String(coverageStart);
    }
  }
  let messages = getMessages(id, oldest);
  const timesEdited = sheet.getRange('E:F').getValues();
  let threadTs = fillMessages(messages, timesEdited, name, sheet);
  if(nMessages >= MAX_MESSAGES) return nMessages;

  if(CHECK_THREAD_TS_IN_SHEET){
    let threadTsInSheet = sheet.getRange('D:D').getValues().filter(function(ts){
      return ts != '';
    });;
    threadTsInSheet.shift()
    threadTs = [...new Set(threadTs.concat(threadTsInSheet))].map(function(ts){
      return String(ts);
    });
    oldest = null;
    if(THREAD_TS_COVERAGE !== null && COVERAGE != 0){
      oldest = Date.now() / 1000. - parseFloat(THREAD_TS_COVERAGE);
    }else if(COVERAGE !== null && COVERAGE != 0){
      oldest = Date.now() / 1000. - parseFloat(COVERAGE);
    }
    if(oldest !== null){
      threadTs = threadTs.filter(function(ts){
        return parseFloat(ts) >= oldest;
      });
    }
  }
  threadTs.forEach(function(ts){
    if(nMessages >= MAX_MESSAGES) return;
    replies = getReplies(id, ts);
    fillMessages(replies, timesEdited, name, sheet);
  });
}
