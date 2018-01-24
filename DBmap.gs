function myFunction() {
  //Browser.msgBox("Hello World");
  getTableList();
  //onEdit(1);
  
}

/* フォームに指定されたシート名と画面名に一致するテーブル一覧を取得する */
function getTableList() {
  /* スプレットシート豆知識　～　値取得編 ～
     1. スプレッドシートのオブジェクトを取得
     2. シートのオブジェクトを取得
     3. セル範囲を指定したオブジェクトを取得
     4. オブジェクトの内容を取得・変更
     
     作業フロー(関数等実装時の引数定義時に考慮する項目)
     1. どのシートから値取得
     　　　> シート名または、ID、アクティブシートのどこから値取得が必要か?
     2. どの行から取得?
          > どの範囲までが必要(何行必要か)?
     3. どの列から取得?
     　　　> どの範囲?
     4. どの行から表示?
     5. どの列から表示?
     6. 1~5が決まったら、定義した要件に基づき作成する     
  */
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 2の手順。シート名を指定せず、現在開いている(スクリプトを実行している)シートを取得
  var activeSheet = spreadsheet.getActiveSheet();
  // 「テスト」シートの「B2」にシート名がある想定
  var sheetNameSelector = activeSheet.getRange("B2");
  //var relativeSheetNameSelector = activeSheet.getRange("A2");
  var sheetName = sheetNameSelector.getValues();
  //  A2 !== B2の条件を作るためA2の値を取得
  //var relativeSheetName = relativeSheetNameSelector.getValue();
  var frontMasterSheetName = activeSheet.getRange("N3");
  var adminMasterSheetName = activeSheet.getRange("N4");
  var getFrontMasterSheetName = frontMasterSheetName.getValue();
  var getAdminMasterSheetName = adminMasterSheetName.getValue();
  var relativeFrontSheet;
  var relativeAdminSheet;
  // シート名によって、テーブル名の位置が異なる
  var screenNameSelector;
  if (sheetName == "入出力一覧_フロント") {
    screenNameSelector = activeSheet.getRange("A3");
     
    /* 画面マスターからのシート値取得(フロント)*/
    relativeAdminSheet = spreadsheet.getSheetByName(getAdminMasterSheetName);
    var relativeAdminScreenRow = getRowForValueName(relativeAdminSheet, 1, "画面");    
    var inputAdminRow  = getRowForValueName(relativeAdminSheet, 1, "入力");
    var outputAdminRow = getRowForValueName(relativeAdminSheet, 1, "出力");
    var tableAdminRow = getRowForValueName(relativeAdminSheet, 1, "テーブル名");
  } else if (sheetName == "入出力一覧_管理画面") {
    screenNameSelector = activeSheet.getRange("B3");
     
     /* マスターシートからの値取得(フロント)*/
     relativeFrontSheet = spreadsheet.getSheetByName(getFrontMasterSheetName);
     var relativeFrontScreenRow = getRowForValueName(relativeFrontSheet, 1, "画面");
     var inputFrontRow  = getRowForValueName(relativeFrontSheet, 1, "入力");
     var outputFrontRow = getRowForValueName(relativeFrontSheet, 1, "出力");
     var tableFrontRow = getRowForValueName(relativeFrontSheet, 1, "テーブル名");
  }else {
    return Browser.msgBox("Error");
  }
  var screenName = screenNameSelector.getValue();
  
  // 2の手順。シート名を指定して、シートを取得
  var dataSheet = spreadsheet.getSheetByName(sheetName);
  //var relativeDataSheet = spreadsheet.getSheetByName(relativeSheetName);
  /* 1行目で「画面」となっている列を取得 */
  var screenRow = getRowForValueName(dataSheet, 1, "画面"); 
  /* 「画面」列の中から、画面名が記載されている行を取得 */
  var screenColumn = getColumnForValueName(dataSheet, screenRow, screenName);
  /* 上記画面名と違う画面名が現れる行を取得 */
  var anotherScreenColumn = getLastColumnForValueName(dataSheet, screenColumn, screenRow, screenName);
  /* 1行目で「テーブル名」となっている列を取得 */
  var tableRow = getRowForValueName(dataSheet, 1, "テーブル名");
  /* 1行目で「入力」となっている列を取得 */
  var inputRow  = getRowForValueName(dataSheet, 1, "入力");
  /* 1行目で「出力」となっている列を取得 */
  var outputRow = getRowForValueName(dataSheet, 1, "出力");
  /* 表示を一回削除する */
  clearValues(activeSheet);  
  /* アクティブシートにdataSheetの内容を書き込む */
  var defaultDisplayPosition = 10;
  setDataSheetValues(activeSheet, dataSheet, screenColumn, tableRow, anotherScreenColumn, 2, defaultDisplayPosition);
  setDataSheetValues(activeSheet, dataSheet, screenColumn, inputRow, anotherScreenColumn, 3, defaultDisplayPosition);
  setDataSheetValues(activeSheet, dataSheet, screenColumn, outputRow, anotherScreenColumn, 4, defaultDisplayPosition);
  /* 構造設計
     1.選択画面のテーブル一覧を取得する
     2.全画面のテーブル一覧を取得する
     3.「1」で取得したテーブルを1つ1つ確認し、全テーブルの一覧から一致する物を取得する
     4.選択画面のテーブルと、取得したテーブルの「入力」と「出力」の関係により処理を変更する
     4-1.選択画面のテーブルが「出力」の場合：取得したテーブルの「入力一覧」と「出力一覧」を表示する
     4-2.選択画面のテーブルが「入力」の場合：取得したテーブルの「出力一覧」のみ表示する
     4-3.選択画面のテーブルが「入力」「出力」の場合：取得したテーブルの「入力一覧」と「出力一覧」を表示する
     */
  // 選択画面のテーブル一覧取得
  var tablesNames = dataSheet.getRange(screenColumn, tableRow, anotherScreenColumn - screenColumn).getValues();
  // 全テーブルの一覧取得
  var allTables = dataSheet.getRange(1, tableRow, dataSheet.getLastRow()).getValues();
  // 選択画面のテーブルと、全テーブルで一致した物を配列で保存する
  // key : table名、 value : 一致したカラムが存在する行番号
  var tableIndexs = {};
  for(var i=1;i<=tablesNames.length;i++){
    var tableName = tablesNames[i-1][0];
    var indexs = [];
    for (var allTableIndex=1; allTableIndex<= allTables.length; allTableIndex++) {
      if (allTables[allTableIndex-1][0] === tableName) {
        indexs.push(allTableIndex);
      }
    }
    tableIndexs[tableName] = indexs;
  }
  // 選択画面のテーブルが入力と出力のどちらに丸が付いているか判断するためのアルゴリズム
  /* 
    1.選択画面の「出力列(outputRow)」にある値を取得(getValues())し、出力一覧の配列を作成する。
    2.選択したテーブルと、対応するテーブルの出力に〇が付いている相関図を作成する。
    3.「1」で作成した出力一覧で「〇」が付いていたら、対応するテーブルの入力に〇が付いている相関図を作成する。
    */
  
  // 相関図を作成するアルゴリズム
  /*
    1.「入力or出力ー入力」「入力ー出力」の相関関係があるか判断する
    2.ある場合、対象の画面名が記載されているカラム位置を取得する
    3.画面名を現在のシートの該当箇所に記載する
  */
  var isInputValues = dataSheet.getRange(screenColumn, outputRow, anotherScreenColumn - screenColumn).getValues();
  var isInputDisplayFlg = false;
  var isOutputValues = dataSheet.getRange(screenColumn, inputRow, anotherScreenColumn - screenColumn).getValues();
  var isOutputDisplayFlg = false;
  var isInAndOutputDisplayFlg = false;
  var isInAndInputDisplayFlg = false;
  var outputListDisplayPosition = defaultDisplayPosition + 1 + tablesNames.length;
  var inAndOutputListDisplayPosition = outputListDisplayPosition + 1 + tablesNames.length;
  //var inAndInputListDisplayPosition = inAndOutputListDisplayPosition + 1 + tablesNames.length;
  var relativeTableOutputValue;
  var relativeTableInputValue;
  var relativeAnotherTableOutputValue;
  var relativeAnotherTableInputValue;
  //  管理画面とフロント画面の入出力”〇”
  var relativeAdminTableOutputValue;
  var relativeAdminTableInputValue;
  var relativeFrontTableOutputValue;
  var relativeFrontTableInputValue;
  var activedCellSheet = activeSheet.getActiveCell().getColumn();
  var color = "#FFFF00";
 
  /*登録・参照の関係性表示*/
  setRelation(activeSheet, defaultDisplayPosition, outputListDisplayPosition, inAndOutputListDisplayPosition);
  
  for(var i=0;i<isInputValues.length;i++){
    var tableName = tablesNames[i][0];
    var inputDisplayRow = 5;
    var outputDisplayRow = 5;
    var inAndOutputDisplayRow = 5;
    var inAndInputDisplayRow = 5;
    var setTableNumRow = 5;
    var indexs = tableIndexs[tableName];
    for (var allTableIndex=0; allTableIndex<indexs.length; allTableIndex++) {
      // 選択テーブル(出力or入力)と、対応するテーブル(出力)の相関図作成
      // 対応するテーブルの「入力」に〇がついているか確認する
      relativeTableInputValue = dataSheet.getRange(indexs[allTableIndex], inputRow).getValue();
      //relativeAnotherTableInputValue = relativeDataSheet.getRange(indexs[allTableIndex], inputAnotherRow).getValue();
      if(relativeAdminSheet) {      
         relativeAdminTableInputValue = relativeAdminSheet.getRange(indexs[allTableIndex], inputAdminRow).getValue();        
      } else if(relativeFrontSheet){      
         relativeFrontTableInputValue = relativeFrontSheet.getRange(indexs[allTableIndex], inputFrontRow).getValue();
      }
      
      
     // if (relativeTableInputValue === "〇" && relativeAnotherTableInputValue === "〇") 
      if (relativeTableInputValue === "〇") {
        // 「入力or出力ー出力」の相関関係がある場合の処理
        setDataSheetValues(activeSheet, dataSheet, indexs[allTableIndex], screenRow, indexs[allTableIndex]+1, inputDisplayRow++, defaultDisplayPosition + i);        
        //選択値がフロント側か管理画面側か
        if (relativeAdminTableInputValue=== "〇") {
          var setColorAdmin = setDataSheetValues(activeSheet, relativeAdminSheet, indexs[allTableIndex], relativeAdminScreenRow, indexs[allTableIndex]+1, inputDisplayRow++, defaultDisplayPosition + i);
    Logger.log(setColorAdmin);
          for (var colorColumn =5; colorColumn<inputDisplayRow; colorColumn++){
            var range =activeSheet.getRange(defaultDisplayPosition +i,colorColumn,1,activedCellSheet);
              
              range.setBackground(color);
          }
        } else if(relativeFrontTableInputValue=== "〇") {
          setDataSheetValues(activeSheet, relativeFrontSheet, indexs[allTableIndex], relativeFrontScreenRow, indexs[allTableIndex]+1, inputDisplayRow++, defaultDisplayPosition + i);
        }        
      }
      
      if (isInputValues[i][0] === "〇") {
        isInputDisplayFlg = true;
        // 選択テーブル(出力)と、対応するテーブル(入力)の相関図作成
        relativeTableOutputValue = dataSheet.getRange(indexs[allTableIndex], outputRow).getValue(); 
         if(relativeAdminSheet) {      
           relativeAdminTableOutputValue = relativeAdminSheet.getRange(indexs[allTableIndex], outputAdminRow).getValue();
         } else if(relativeFrontSheet){      
           relativeFrontTableOutputValue = relativeFrontSheet.getRange(indexs[allTableIndex], outputFrontRow).getValue();
         }
        
        if (relativeTableOutputValue === "〇") {
          // 「出力ー入力」の相関関係がある場合の処理
          setDataSheetValues(activeSheet, dataSheet, indexs[allTableIndex], screenRow, indexs[allTableIndex]+1, outputDisplayRow++, outputListDisplayPosition + i);
          //setDataSheetValues(activeSheet, relativeDataSheet, indexs[allTableIndex], relativeScreenRow, indexs[allTableIndex]+1, outputDisplayRow++, outputListDisplayPosition + i);
          
          if (relativeAdminTableOutputValue=== "〇") {          
            setDataSheetValues(activeSheet, relativeAdminSheet, indexs[allTableIndex], relativeAdminScreenRow, indexs[allTableIndex]+1, outputDisplayRow++, outputListDisplayPosition + i); 
          } else if(relativeFrontTableOutputValue=== "〇") {
            setDataSheetValues(activeSheet, relativeFrontSheet, indexs[allTableIndex], relativeFrontScreenRow, indexs[allTableIndex]+1, outputDisplayRow++, outputListDisplayPosition + i);      
          }       
          
        } 
      }
      if (isInputValues[i][0] === "〇" && isOutputValues[i][0] !== "〇") {
        // 「参照 + 登録」の場合の処理 
        isOutputDisplayFlg = true;
        relativeTableOutputValue = dataSheet.getRange(indexs[allTableIndex], outputRow).getValue();
        if(relativeAdminSheet) {      
           relativeAdminTableOutputValue = relativeAdminSheet.getRange(indexs[allTableIndex], outputAdminRow).getValue();
         } else if(relativeFrontSheet){      
           relativeFrontTableOutputValue = relativeFrontSheet.getRange(indexs[allTableIndex], outputFrontRow).getValue();
         }        
        if (relativeTableOutputValue === "〇") {
          setDataSheetValues(activeSheet, dataSheet, indexs[allTableIndex], screenRow, indexs[allTableIndex]+1, inAndOutputDisplayRow++, inAndOutputListDisplayPosition + i);
          //setDataSheetValues(activeSheet, relativeDataSheet, indexs[allTableIndex], relativeScreenRow, indexs[allTableIndex]+1, inAndOutputDisplayRow++, inAndOutputListDisplayPosition + i);
           if (relativeAdminTableOutputValue=== "〇") {          
            setDataSheetValues(activeSheet, relativeAdminSheet, indexs[allTableIndex], relativeAdminScreenRow, indexs[allTableIndex]+1, inAndOutputDisplayRow++, inAndOutputListDisplayPosition + i); 
          } else if(relativeFrontTableOutputValue=== "〇") {
            setDataSheetValues(activeSheet, relativeFrontSheet, indexs[allTableIndex], relativeFrontScreenRow, indexs[allTableIndex]+1, inAndOutputDisplayRow++, inAndOutputListDisplayPosition + i);      
          }
          
          
        }
      }
    }
  }
     
  
  if (isInputDisplayFlg) {
    // 出力一覧を表示していた場合、対応するテーブル一覧も表示する
    setDataSheetValues(activeSheet, dataSheet, screenColumn, tableRow, anotherScreenColumn, 2, outputListDisplayPosition);
    setDataSheetValues(activeSheet, dataSheet, screenColumn, inputRow, anotherScreenColumn, 3, outputListDisplayPosition);
    setDataSheetValues(activeSheet, dataSheet, screenColumn, outputRow, anotherScreenColumn, 4, outputListDisplayPosition);
  }
  if (isOutputDisplayFlg) {
    setDataSheetValues(activeSheet, dataSheet, screenColumn, tableRow, anotherScreenColumn, 2, inAndOutputListDisplayPosition);
    setDataSheetValues(activeSheet, dataSheet, screenColumn, inputRow, anotherScreenColumn, 3, inAndOutputListDisplayPosition);
    setDataSheetValues(activeSheet, dataSheet, screenColumn, outputRow, anotherScreenColumn, 4, inAndOutputListDisplayPosition); 
  }
}

/* 特定の行に存在する値が、どの列にあるか取得する */
function getRowForValueName (sheet, column, valueName) {  
  // column行の1列目のセル ~ 同じcolumn行(1)の最終列目のセルの範囲を捜索対象とする
  var searchRow = sheet.getRange(column, 1, 1, sheet.getLastColumn()).getValues();
  for(var i=1;i<searchRow[column-1].length;i++){
    if(searchRow[0][i-1] === valueName){
      return i;
    }
  }
  return 0;
}

/* 特定の列に存在する値が、どの行にあるか取得する */
function getColumnForValueName (sheet, row, valueName) {  
  // 1行目のrowのセル ~ 最終行目のrowのセルの範囲を捜索対象とする
  var searchRow = sheet.getRange(1, row, sheet.getLastRow()).getValues();
  for(var i=1;i<searchRow.length;i++){
    if(searchRow[i-1][0] === valueName){
      return i;
    }
  }
  return 0;
}

/* 特定の列にある特定の値が、別の値に変化する行を取得する */
function getLastColumnForValueName (sheet, column, row, valueName) {  
  // 1行目のrowのセル ~ 最終行目のrowのセルの範囲を捜索対象とする
  var searchRow = sheet.getRange(1, row, sheet.getLastRow()).getValues();
  // 最初に見つかった行数から検索を開始し、同じ名前がどこまで続くかを判断する
  for(var i=column;i<searchRow.length;i++){
    if(searchRow[i-1][0] !== valueName){   
      return i;
    }
  }
  return 0;
}

function clearValues(activeSheet){
  // 現在のシートの「B10」以降の列の書き込みをクリアする
  activeSheet.getRange(10, 2, activeSheet.getLastRow(), activeSheet.getLastColumn()).clear();
  activeSheet.getRange(10, 1, activeSheet.getLastRow(), activeSheet.getLastColumn()).clear();
  var clearSelectFunc = activeSheet.getRange("B2").getValue();

  switch (clearSelectFunc) {
    case "入出力一覧_管理画面":
      activeSheet.getRange("A3").clear();
      break;
    case "入出力一覧_フロント":
      activeSheet.getRange("B3").clear();
      break;
  }
  
}

function setDataSheetValues(activeSheet, dataSheet,screenColumn, tableRow, anotherScreenColumn, insertRow, insertColumn){
  var tablesNames = dataSheet.getRange(screenColumn, tableRow, anotherScreenColumn - screenColumn).getValues();
  for(var i=0;i<tablesNames.length;i++){
    // 現在のシートの「insertRow列のinsertColumn行目」から書き込みを開始
    activeSheet.getRange(insertColumn+i, insertRow).setValue(tablesNames[i][0]);
  }   
}
/* 登録・参照の関係を表示する */
function setRelation (sheet, def, out, inAndOut, inAndIn) {
  // 使用する名称を格納
  var arrName = ["登録-登録","登録ー参照","参照ー登録","入力ー入力"];  
  if (def) {
      sheet.getRange(def, 1).setValue(arrName[0]);
  } 
  if (out) {
      sheet.getRange(out, 1).setValue(arrName[1]);
  }
  if (inAndOut) {
      sheet.getRange(inAndOut, 1).setValue(arrName[2]);
  }
//  if (inAndIn) {
//      sheet.getRange(inAndIn, 1).setValue(arrName[3]);
//  }  
  
  
//  var arrName = [
//                 ["登録-登録","登録ー参照","参照ー登録"],
//                 ["def","out","inAndOut"]
//                ];
   
//  for (var i=0; i<arrName[0].length; i++) {
//  //Logger.getLog(arrName[0]);
//    if (arrName[1][i]=== "def"){
//      sheet.getRange(def, 1).setValue(arrName[0]);
//    }
//    if (arrName[1][i]=== "out") {
//      sheet.getRange(out, 1).setValue(arrName[0][i]);
//    
//    }
//    if (arrName[1][i] === "inAndOut") {
//      sheet.getRange(inAndOut, 1).setValue(arrName[0][i]);
//    }    
//  }


}

function onEdit(e){
   var r1 = e.range;
   var r2 = SpreadsheetApp.getActiveSheet().getRange("A1");
   r2.setValue("Edited: " + r1.getA1Notation() + ", value = " + r1.getValue());
   var color = "#FFFF00";
   var getColor = r2.setBackground(color);
   var colorWhite = "#FFFFFF"; 
   var turnColorFlg = 0;
   
}
