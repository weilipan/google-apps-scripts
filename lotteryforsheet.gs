function onOpen(e) {
  let cache = CacheService.getScriptCache(); //利用cache暫存變數
  cache.put('ltime',null);
  cache.put('num',null);
  cache.put('lotterylist',null);
  cache.put('lotteryliststring',null);
  SpreadsheetApp.getUi()
      .createMenu('書展抽獎程式')
      .addItem('抽獎活動開始', 'toLottery')
      .addItem('顯示得獎名單','showLottery')
      .addItem('寫入得獎名單','recordLottery')
      .addToUi();
}

function toLottery(){
  // 1. SpreadsheetApp -> Spreadsheet
  let sheet = SpreadsheetApp.getActiveSheet()
  ui=SpreadsheetApp.getUi();
  let cache = CacheService.getScriptCache(); //利用cache暫存變數
  let num=0; //用來紀錄要抽出多少得獎人數
  let item_range = sheet.getRange('1:1').getValues(); //'1:1'指的是第一列，標題列

  let items = []; //儲存標題用
    
  for (let idx in item_range[0]) {
    //確認有幾個欄位，一直數到沒有標題就跳出 
    if (item_range[0][idx] == '') {
      break;
    }
    items.push(item_range[0][idx]); //將標題紀錄下來
  }
  //共有幾筆資料 
  let row_number=sheet.getLastRow();
  
  // 共有幾個欄位
  let column_number = items.length;
 
  //將資料的範圍存到cache中，之後可以取用。
  cache.put('row_number',row_number);
  cache.put('column_number',column_number);

  let response = ui.prompt('本次抽獎人數設定', '請問這是本次書展的第幾次抽獎?：', ui.ButtonSet.OK );
  if (response.getSelectedButton()==ui.Button.OK){
    let ltime=response.getResponseText();
    ui.alert('第'+ltime+'次抽獎作業開始....');
    ltime=parseInt(ltime);
    cache.put('ltime',ltime);
  }
  response = ui.prompt('本次抽獎人數設定', '請輸入本次要抽出的人數：', ui.ButtonSet.OK );
  if (response.getSelectedButton()==ui.Button.OK){
    num=response.getResponseText();
    ui.alert('本次抽獎預計抽出'+num+'名幸運得主');
    num=parseInt(num);
    cache.put('num',num);
  }
  //設定資料要處理的欄位
  let range = sheet.getRange(2,1,row_number,column_number);

  let data = [];

  // 將所有資料取出，是一個二維陣列
  let values = range.getValues();
  // transform into JSON-like array
  for (let row = 0; row < row_number; ++row) {
    let row_object = {};
    if (values[row][column_number-1]==''){ //如果尚未得獎才納入，是否中獎要看最後一個欄位。
      for (let col = 0; col < column_number; ++col) {
        let item = items[col];
        row_object[item] = values[row][col];
      }
      data.push(row_object);

    }
  }

  shuffle(data);//將資料打亂

  msg='本次活動得獎名單如下：\n'
  data.slice(0,num).forEach(function(x,i){ //只取中獎的前N筆資料，如果這次取5名，則取亂數排序資料列中的前五名做為中獎名單。
    let temp={} //將中獎名單存起來
    temp['夢駝林帳號']=x['夢駝林帳號'];
    temp['姓名']=x['姓名'];
    temp['好書推薦(ISBN)']=x['好書推薦(ISBN)'];
    temp['好書推薦(書名)']=x['好書推薦(書名)'];
    msg=msg+(i+1).toString()+'. '+x['夢駝林帳號']+' '+showname(x['姓名'])+' '+x['好書推薦(書名)']+x['好書推薦(ISBN)']+' '+'\n';
  })
  msg=msg+'恭喜以上得獎的人員！！獲得價值600元以下書展現場展示個人指定書籍乙冊。'
  //顯示得獎訊息。 
  ui.alert(msg);

  //將得獎名單暫存在cache中
  
  cache.put('lotterylist',JSON.stringify(data.slice(0,num))); //已經轉成字串了
  cache.put('lotteryliststring',msg);

}
function shuffle(array) { //亂數排序
  for (let i = array.length - 1; i > 0; i--) {
    let j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}
function showLottery(){ //將得獎名單展示出來
  let cache = CacheService.getScriptCache(); //利用cache暫存變數
  result=cache.get('lotterylist');
  resultstring=cache.get('lotteryliststring');
  ui=SpreadsheetApp.getUi();
  if (result==null){
    ui.alert('目前尚未有得獎名單')
  }
  else{
    ui.alert(resultstring);
  }
}
//將結果寫回sheet的功能。
function recordLottery(){
  let sheet = SpreadsheetApp.getActiveSheet()
  ui=SpreadsheetApp.getUi();
  let cache = CacheService.getScriptCache(); //利用cache暫存變數
  let lottery=cache.get('lotterylist');//直接將串列由cache端下載下來使用
  
  if (lottery==null){
    ui.alert('目前還沒有得獎名單喔!');
    return;
  }

  let row_number=parseInt(cache.get('row_number'));
  let column_number=parseInt(cache.get('column_number'));
  let ltime=parseInt(cache.get('ltime'));
  let num=parseInt(cache.get('num'));
  // let id_list=cache.get('id_list').split(','); //學號列表
  let id_list=checklist();  //得獎人名單列表
  let range = sheet.getRange(2,1,row_number,column_number);
  let values = range.getValues();
  for (let row = 0; row < row_number; ++row) {
    if (id_list.indexOf(values[row][2])!=-1){//對照是不是得獎人員
      sheet.getRange(row+2,column_number).setValue(ltime); 
      //row+2是因為標題要加1，表格定位是由1開始不像程式定位由0開始，所以總共要加2
    }
  }
  SpreadsheetApp.flush();
  ui.alert('設定完成！')
}

function showname(name){
  // 姓名遮罩
  leng=name.length;
  let temp=name.slice(0,1);
  if (leng>2){ //如果名字長度大於2才要oo
    for (let i=0;i<leng-2;i++)
    {
      temp=temp+'O';
    }
    temp=temp+name.slice(-1);
  }
  else{
    temp=temp+'O';
  }
  return temp;
}
function checklist(){ //取出得獎人員名單
  ui=SpreadsheetApp.getUi();
  let cache = CacheService.getScriptCache(); //利用cache暫存變數
  let lottery=JSON.parse(cache.get('lotterylist'));//直接將串列由cache端下載下來使用，將它轉回物件，接著用物件走訪的方式處理
  let namelist=[]
  for (let i=0;i<lottery.length;i++){
    namelist.push(lottery[i]['夢駝林帳號'])
  }
  return namelist;
}