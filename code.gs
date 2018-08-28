function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "せんべろネットの情報を取得",
      functionName : "myFunction"
    }
  ];
  sheet.addMenu("スクリプト実行", entries);
  //メインメニュー部分に[スクリプト実行]メニューを作成して、
  //下位項目のメニューを設定している
};
function getMainSheet() {
  if (getMainSheet.memoSheet) { return getMainSheet.memoSheet; }

  getMainSheet.memoSheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  return getMainSheet.memoSheet;
}
function dataClear() {
  const rowStartData = 4;
  const rowEndData = 100;
  const colURL = 12;
  const sheetData = getMainSheet();
  sheetData.getRange(rowStartData, 2,rowEndData,colURL).clear();
}
function myFunction() {
  //const book = SpreadsheetApp.getActiveSpreadsheet();
  //const sheetData = book.getSheetByName("シート1");
  const sheetData = getMainSheet();

  const colID = 1;
  const colNAME = 2;
  const colOPENING = 3;
  const colHOLIDAY = 4;
  const colCHARGE  = 7;
  const colTAG = 8
  const colSTATION_ON_FOOT = 10;
  const colTABELOG = 11;
  const colURL = 12;

  const rowHeader = 3;
  const rowStartData = 4;
  const rowEndData = 100;

  const colListUrl = 3;
  const rowListUrl = 1;
  var rowRestaurantCount = 0;
    
  // ヘッダ描画
  const header = ['順番','店名','営業時間','定休日','ドリンク最安','おすすめおつまみ','チャージ','タグ','駅徒歩','備考','食べログ','せんべろネット'];
  for(var j = 0;j < header.length;j += 1){
    sheetData.getRange(rowHeader,j+1).setValue(header[j]);
  }
  // 店リストの取得
  const restListUrl = sheetData.getRange(rowListUrl, colListUrl).getValue();
  const restLisResponse = UrlFetchApp.fetch(restListUrl);
  var html = restLisResponse.getContentText('UTF-8');
  for (var i = rowStartData; i <= rowEndData; i += 1) {
    var searchTag = '<div class="entry-content">';
    var index = html.indexOf(searchTag)
    if (index !== -1) {
      var html = html.substring(index + searchTag.length); 
      
      var startSerchTag = '<h3><a href="';
      var endSerchTag = '" class="entry-title entry-title-link"';
      var setUrl = html.substring(html.indexOf(startSerchTag)+startSerchTag.length, html.indexOf(endSerchTag));
      if(setUrl.indexOf("restaurant") !== -1){
        sheetData.getRange(i, colURL).setValue(setUrl);
      }
      else{
        i -= 1;
      }
    }
    else{
      rowRestaurantCount = i;
      break;
    }
  }

  //各店の詳細情報取得
  
  for (var i = rowStartData; i < rowRestaurantCount; i += 1) {
    Utilities.sleep(1000);
    var url = sheetData.getRange(i, colURL).getValue();
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText('UTF-8');
    
    
    // charge
    var searchStartTag = '※チャージ：'
    var searchEndTag = '<br />'
    var index = html.indexOf(searchStartTag);
    if(index !== -1){
      html = html.substring(index+searchStartTag.length); 
      sheetData.getRange(i,colCHARGE).setValue(html.substring(0,html.indexOf(searchEndTag)));
    }
    
    // 食べログURL
    var searchStartTag = '<div><strong><a href="'
    var searchEndTag = '" target="_blank"'
    var index = html.indexOf(searchStartTag);
    var tabelogFlag = false;
    if(index !== -1){
      html = html.substring(index+searchStartTag.length); 
      sheetData.getRange(i,colTABELOG).setValue(html.substring(0,html.indexOf(searchEndTag)));
      tabelogFlag = true;
    }
    // 店名
    if(tabelogFlag == true){
      var searchStartTag1 = '" target="_blank" rel="noopener">'
      var searchStartTag2 = '" target="_blank">'
      var searchEndTag = '</a></strong><br />'
      var index = html.indexOf(searchStartTag1);
      if(index !== -1){
        html = html.substring(index+searchStartTag1.length);     
        sheetData.getRange(i,colNAME).setValue(html.substring(0,html.indexOf(searchEndTag)));
      }
      else{
        var index = html.indexOf(searchStartTag2);
        if(searchStartTag2 !== -1){
          html = html.substring(index+searchStartTag2.length);     
          sheetData.getRange(i,colNAME).setValue(html.substring(0,html.indexOf(searchEndTag)));
        }
      }
    }
    // tag
    var searchStartTag = '<!--カテゴリ一覧表示-->';
    var index = html.indexOf(searchStartTag);
    if(index !== -1){
      html = html.substring(index+searchStartTag.length); 
    }
    var searchStartTag = 'rel="tag">';
    var searchEndTag = '</a>';
    var setTags = new Array(0);
    while(1){
      var index = html.indexOf(searchStartTag);
      if(index !== -1){
        html = html.substring(index + searchStartTag.length); 
        setTags.push(html.substring(0,html.indexOf(searchEndTag)));
        
      }
      else{
        break;
      }
    }
    sheetData.getRange(i,colTAG).setValue(setTags.join('\n'))
    
    // 食べログ情報取得
    if(tabelogFlag){
      var url = sheetData.getRange(i, colTABELOG).getValue();
      info = getTabeLogInfo(url);
      if(info['営業時間']){
        sheetData.getRange(i,colOPENING).setValue(info['営業時間']);
      }
      if(info['定休日']){
        sheetData.getRange(i,colHOLIDAY).setValue(info['定休日']);
      }
    }
  }
}
function getTabeLogInfo(url){
  var info = {};
  var response = UrlFetchApp.fetch(url);
  var html = response.getContentText('UTF-8');
  // 店名
  /*
  var searchStartTag = '/<th>店名<\/th>\s+<td>/'
  var searchEndTag = '</td>'
  var index = html.search(searchStartTag);
  if(index !== -1){
    html = html.substring(index+searchStartTag.length); 
    info['店名'] = html.substring(0,html.indexOf(searchEndTag));
  }
  */
  // ジャンル
  // 予約・問い合わせ
  // 予約可否
  // 住所
  // 交通手段
  // 営業時間
  var searchStartTag = '<p>'
  var searchEndTag = '</p>'
  var index = html.indexOf('<th>営業時間</th>');
  if(index !== -1){
    html = html.substring(index); 
    index = html.indexOf(searchStartTag);
    if(index !== -1){
      info['営業時間'] = html.substring(index+searchStartTag.length,html.indexOf(searchEndTag));
      info['営業時間'] = info['営業時間'].replace(/(<br>|<br \/>)/gi,'\n');
    }
  }
  // 定休日
  var searchStartTag = '<p>'
  var searchEndTag = '</p>'
  var index = html.search('<th>定休日</th>');
  if(index !== -1){
    html = html.substring(index); 
    index = html.indexOf(searchStartTag);
    if(index !== -1){
      info['定休日'] = html.substring(index+searchStartTag.length,html.indexOf(searchEndTag));
    }
  }
  // 予算
  return info;
}
