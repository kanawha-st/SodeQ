var twitter = getTwitter(); //TwitterWebService.getInstance(TWITTER_ID,TWITTER_LOGIN);
var properties = PropertiesService.getScriptProperties(); 
//  -------------- TWITTER SETTINGS ------------------
// 認証
function authorize() {
  twitter.authorize();
}

// 認証解除
function reset() {
  twitter.reset();
}

// 認証後のコールバック
function authCallback(request) {
  return twitter.authCallback(request);
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
      name: 'ツイッターシステム'
    }
  );
}

function tweet(text) {
  var twitter = getTwitter();
  var service = twitter.getService();
  if(text.length>140){
    text = text.substring(0, 139) + '…';
  }
  
  var response = service.fetch('https://api.twitter.com/1.1/statuses/update.json', {
    method: 'post',
    payload: {status: text}
  });
}

function test() {
//  var saved_file = saveAsSpreadSheet("SodeQテスト", DOWNLOAD_URL);
//  processSpreadSheet(saved_file);
//  Logger.log(saved_file);
  tweet("テストです");
}



function getMenusFromDate(date, exact_date) {
  function F(d){ return Utilities.formatDate(d, 'JST', 'yyMM'); }
  function D(r){
    if(r.slice(3).every(function(s){ return s == ""; })){ return null; }
    else { return [r[1]].concat(r.slice(3)); }
  }
  function disp(r, s){
    function apnd(_s){ return _s + s; }
    return [r[0]].concat(r.slice(1).map(apnd));
  }
  
  function getMenu(d, EorJ) {
    var file = DriveApp.getFolderById(DATA_FOLDER_ID).getFilesByName(F(d) + EorJ).next();
    if(!file){ return null; }
    var sheetE = SpreadsheetApp.openById(file.getId());
    var data = sheetE.getSheets()[0].getDataRange().getValues();
    var today = Utilities.formatDate(d, 'JST', 'yyMMdd');
    
    for(var i=1; i<data.length; i++) {
      var r = data[i].filter(function(s){ return s != ''; });
      if(r.length <= 3){ continue; }
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
    var menus = [menuE[0]];
    for(var i = 1; i < menuJ.length; i++) {
      var mE = menuE[i].trim();
      var mJ = menuJ[i].trim();
      if(mE == "" && mJ == ""){ /* do nothing */ }
      else if(mJ == mE){ menus.push(mJ); }
      else { menus.push(mJ + "㊥/" + mE + "㋛"); }
    }
    return menus
  }
  if(menuE){ return disp(menuE, "㋛"); }
  return disp(menuE, "㊥");
}
  

function dailyTweet() {
  var text = "";
  try {
    var d = new Date();
    var menuToday = getMenusFromDate(d, true);
    if(menuToday){ 
      text = "本日（" + Utilities.formatDate(menuToday[0], 'JST', "MM月dd日") +  "）の給食メニューです。\n▫" + menuToday.slice(1).join("\n▫") + "\n\n";
      d.setDate(d.getDate() + 1);
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
  } catch(e) {
    sendMail(ADMIN_EMAIL, "【SodeQ】エラー発生", 'ツイート機能でエラーが発生しました。\n' + e);
  }
}

function getExcelFiles() {
  try {
    var src = UrlFetchApp.fetch(DOWNLOAD_URL).getContentText();
    var divs = Parser.data(src).from('<div class="file_excel">').to('</div>').iterate();
    var M = ['', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    
    divs.map(function(s) {
      var jtitle = Parser.data(s).from('.xlsx">').to('  [').build();
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
