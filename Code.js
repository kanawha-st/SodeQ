var properties = PropertiesService.getScriptProperties(); 
var twitter = null;
//  -------------- TWITTER SETTINGS ------------------
function getTwitterService() {
  var clientId = TWIITER_CLIENT_ID;
  var clientSecret = TWITTER_CLIENT_SECRET;
  var redirectUri = 'https://script.google.com/macros/d/1ATf8DeIMF7GN9PeBR56qoGH6bQRfMfst42CuD9V_cPgRbQo_jhnn9Jnh/usercallback';
  var scope = 'tweet.read,tweet.write';
  
  var service = OAuth2.createService('twitter')
    .setAuthorizationBaseUrl('https://twitter.com/i/oauth2/authorize')
    .setTokenUrl('https://api.twitter.com/2/oauth2/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope(scope)
    .setRedirectUri(redirectUri)
    .setGrantType('authorization_code');
  
  return service;
}

function authCallback(request) {
  var service = getTwitterService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can now post tweets.');
  } else {
    return HtmlService.createHtmlOutput('Authentication failed.');
  }
}
function doGet() {
  var service = getTwitterService();
  var authorizationUrl = service.getAuthorizationUrl();
  var html = HtmlService.createHtmlOutput('<a href="' + authorizationUrl + '">Authorize</a>');
  return html;
}

//  -------------- END ------------------

// following program is taken from https://qiita.com/Panda_Program/items/31f331fd4c2f3cfab333=====
function isHoliday(today){ 
  const calendars = CalendarApp.getCalendarsByName('日本の祝日');
  const count = calendars[0].getEventsForDay(today).length;
  return count === 1;
}
// ==== https://qiita.com/Panda_Program/items/31f331fd4c2f3cfab333 =====

function saveAsSpreadSheet(title, url) {
  Logger.log("SAVING " + title + " from " + url);
  var options = {
    title: title,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{id: DATA_FOLDER_ID}]
  };
  Drive.Files.insert(
    options,
    UrlFetchApp.fetch(url).getBlob());
  var files = DriveApp.getFolderById(DATA_FOLDER_ID).getFilesByName(title);
  Logger.log(files);
  return files.next();
}

function sendMail(address, title, message) {
  GmailApp.sendEmail(
    address,
    title,
    message,
    {
      from: Session.getActiveUser().getEmail(),
      name: 'SodeQツイッターシステム'
    }
  );
}

function tweet(text) {
  if(text.length>140){
    text = text.substring(0, 139) + '…';
  }
  let service = newTwitterV2();
  var response = service.postMessage(text);
  Logger.log(response);
  /* if(response.getResponseCode()>300){
    Logger.log(response.getResponseCode());
    Logger.log(response.getContentText());
    throw "twitter error\n" + response.getContentText();
  } */
  Logger.log(text);
}

function test() {
  //Logger.log(getYomi("テスト中"))
  //Logger.log(getMenusFromDate(new Date('2025-03-10'), false));
  //authorize();
  //reset();
//  var saved_file = saveAsSpreadSheet("SodeQテスト", DOWNLOAD_URL);
//  processSpreadSheet(saved_file);
//  Logger.log(saved_file);
  tweet("テストです");
//  Logger.log(getMenusFromDate(new Date('2021-04-15'), true));
//  Logger.log(bentoDays());
//  Logger.log(getYomi("茎わかめサラダ和風味"))
}

function moveToLast(ary, words) {
  words.forEach( w => {
    if(ary.indexOf(w) >= 0){
      ary = ary.filter(a => a!==w).concat([w]);
    }
  });
  return ary; 
}

// Yahoo Version
function getYomi(text) { 
  var endpoint = "https://jlp.yahooapis.jp/MAService/V2/parse?appid=" + YAHOO_API_ID;
  
  // Prepare the request parameters
  
  var payload = {
    "id": Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd'),
    "jsonrpc": "2.0",
    "method": "jlp.furiganaservice.furigana",
    "params": {
      "q": text,
    }
  };
  var headers = {
        "Content-Type": "application/json",
  };
  try {
    var response = UrlFetchApp.fetch(endpoint, {
      'method': 'post',
      'headers': headers,
      'payload': JSON.stringify(payload)
    });
    var response_json = JSON.parse(response.getContentText())
    return response_json['result']['word'].reduce((x, y) => x + y['furigana'], "")
  } catch(e) {
    Logger.log("Yahooテキスト解析エラー")
    Logger.log(e);
    return text;
  }

}

function getMenusFromDate(date, exact_date) {
  function F(d){ return Utilities.formatDate(d, 'JST', 'yyMM'); }
  function D(r){
    if(r.slice(3).every(function(s){ return s == ""; })){ return null; }
    else{return [r[1]].concat(r.slice(3));}
  }
  function disp(r, s){
    function apnd(_s){ return _s + s; }
    return [r[0]].concat(r.slice(1).map(apnd));
  }
  
  function getMenu(d, EorJ) {
    var file = null;
    try { file = DriveApp.getFolderById(DATA_FOLDER_ID).getFilesByName(F(d) + EorJ).next(); }
    catch(e) { return null; }
    var sheetE = SpreadsheetApp.openById(file.getId());
    var data = sheetE.getSheets()[0].getDataRange().getValues();
    var today = Utilities.formatDate(d, 'JST', 'yyMMdd');
    
    for(var i=1; i<data.length; i++) {
      var r = data[i].filter(function(s){ return s != ''; });
      if(r.length <= 4){ continue; } // Too few menu
      var d = Utilities.formatDate(new Date(r[1]), 'JST', 'yyMMdd');
      if(today == d) {
        return D(r);
      }
      else if(today < d) { 
        return exact_date ? null : D(r);
      }
    }
    if(!exact_date){ // NOT IN THIS MONTH
      date.setMonth(date.getMonth() + 1, 1);
      return getMenu(date, EorJ)
    }
    return null;
  }
    
  var menuE = getMenu(date, "E");
  var menuJ = getMenu(date, "J");
  if(!menuE && !menuJ){ return null; } // BOTH INVALID/NO LUNCH
  else if(menuE && menuJ){ //BOTH VALID
    if(menuE[0].getTime() == menuJ[0].getTime()){
      var menus = [menuE[0]];
      for(var i = 1; i < menuJ.length; i++) {
        var mJ = menuJ[i].trim();
        if(!menuE[i]){ menus.push(mJ + "㊥"); continue}
        var mE = menuE[i].trim();
        if(mE == "" && mJ == ""){ /* do nothing */ }
        else if(mJ === mE || getYomi(mJ) === getYomi(mE)){ menus.push(mJ); }
        else { menus.push(mJ + "㊥/" + mE + "㋛"); }
      }
      return moveToLast(menus,['牛乳','ごはん'])
    }
    else if(menuE[0].getTime() < menuJ[0].getTime()){return disp(moveToLast(menuE, ['牛乳','ごはん']), "㋛"); }
    else { return disp(moveToLast(menuJ, ['牛乳','ごはん']), "㊥");} 
  }
  if(menuE){ return disp(moveToLast(menuE, ['牛乳','ごはん']), "㋛"); }
  return disp(moveToLast(menuJ, ['牛乳','ごはん']), "㊥");
}
  

function dailyTweet() {
  var text = "";
  let bentoDays = getBentoDays();

  // check yomi
  if(getYomi('テスト中')!=='てすとちゅう'){
    sendMail(ADMIN_EMAIL, "【SodeQ】ワーニング", 'ひらがな読み機能でエラーが発生しました。');
  }

  try {
    var d = new Date();
    if(bentoDays[Utilities.formatDate(d, "JST", 'yyyy-MM-dd')]) {
      text = "🍱🍱🍱今日はお弁当の日!!🍱🍱🍱\n\n"
    } else {
      var menuToday = getMenusFromDate(d, true);
      if(menuToday) {
        text = "本日のメニュー\n▫" + menuToday.slice(1).join("\n▫") + "\n\n";
      }
    }
    if(text.length > 0){ 
      d.setDate(d.getDate() + 1);
      if(bentoDays[Utilities.formatDate(d, "JST", 'yyyy-MM-dd')]){
        text += "🍱🍱🍱明日はお弁当!!🍱🍱🍱";
      } else {
        var menuNext = getMenusFromDate(new Date(d), false);
          if( menuNext ){
            if(Utilities.formatDate(menuNext[0], 'JST', "MM月dd日") === Utilities.formatDate(d, 'JST', "MM月dd日")){ text += "明日"; }
            else { text += "次回(" + Utilities.formatDate(menuNext[0], 'JST', "MM月dd日") + ")"; }
            text += "は\n【" + menuNext.slice(1).join("、") + "】の予定です。"
          }
      }
      Logger.log(text);
      Logger.log(text.length);
      tweet(text);
    }
  } catch(e) {
    sendMail(ADMIN_EMAIL, "【SodeQ】エラー発生", 'ツイート機能でエラーが発生しました。\n' + e);
    throw e;
  }
}

function forceTweet() {
  var text = "あけましておめでとうございます。今年もSodeQをよろしくお願いします。\n"
  var d = new Date();
  var menuNext = getMenusFromDate(new Date(d), false);
  if( menuNext ){
    if(Utilities.formatDate(menuNext[0], 'JST', "MM月dd日") === Utilities.formatDate(d, 'JST', "MM月dd日")){ text += "明日"; }
    else { text += "次回(" + Utilities.formatDate(menuNext[0], 'JST', "MM月dd日") + ")"; }
    text += "の給食は\n【" + menuNext.slice(1).join("、") + "】の予定です。"
  }
  Logger.log(text);
  Logger.log(text.length);
  tweet(text);
}

function getExcelFiles() {
  try {
    var src = UrlFetchApp.fetch(DOWNLOAD_URL).getContentText();
    var divs = Parser.data(src).from('<div class="file_excel">').to('</div>').iterate();
    var M = ['', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    
    divs.map(function(s) {
      var jtitle = Parser.data(s).from('.xlsx">').to('  [').build();
      for(var i=0; i<10; i++){ jtitle = jtitle.replace("０１２３４５６７８９".charAt(i), String(i)); }
      var year = (jtitle.indexOf("元年") > -1) ? 1 : parseInt(/令和(\d+)年/.exec(jtitle)[1]);
      year += 18;
      var month = parseInt(/(\d+)月/.exec(jtitle)[1]);
      var type = (jtitle.indexOf("小学校") > -1) ? "E":"J"; 
      return { 
        title: year + M[month] + type,
        href:  Parser.data(s).from('<a href="').to('">').build()
      };
    }).forEach( function(d) {
      var files = DriveApp.getFolderById(DATA_FOLDER_ID).getFilesByName(d.title);
      if(!files.hasNext()) {
        saveAsSpreadSheet(d.title, "https://www.city.sodegaura.lg.jp" + d.href);
      }
    });
  } catch(e) {
    sendMail(ADMIN_EMAIL, "【SodeQ】エラー発生", 'ファイルダウンロード機能でエラーが発生しました。\n' + e);    
  }
}

function getBentoDays() {
    let sheet = SpreadsheetApp.openById('1bg0zP4HbEqQZmwavYm_ZJ1UzUdjEQwk5IMeCzNRvCVQ').getSheetByName('弁当日');
    let rows = sheet.getDataRange().getValues();
    let ret = {};
    rows.forEach(r => {ret[Utilities.formatDate(r[0],"JST","yyyy-MM-dd")]=true;});
    return ret;
}


function bentoCheck() {
  try {
    let sheet = SpreadsheetApp.openById('1bg0zP4HbEqQZmwavYm_ZJ1UzUdjEQwk5IMeCzNRvCVQ').getSheetByName('弁当日');
    let rows = sheet.getDataRange().getValues();
    let today = new Date();
    let nextweek = new Date(); nextweek.setDate(nextweek.getDate() + 7);
    rows.some(r => {
      let d = new Date(r[0]);
      if( today < d && d < nextweek){
        let text = `📢📢来週の${Utilities.formatDate(d, "JST", "M月d日")}は『お弁当の日』です。\nお買い物は済んでますか？\nフォロワーの袖ケ浦のお父さん・お母さん達にも届くようリツイートを！！📢📢`;
        Logger.log(text);
        tweet(text)
        return true;
      }
    });
  } catch(e) {
    sendMail(ADMIN_EMAIL, "【SodeQ】エラー発生", '弁当日ツイート機能でエラーが発生しました。\n' + e);
  }

}
