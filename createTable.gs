


/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //var entries = [{
  //  name : "Read Data",
  //  functionName : "readRows"
  //}];
  //spreadsheet.addMenu("Script Center Menu", entries);
  

   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menuEntries = [];
   // When the user clicks on "addMenuExample" then "Menu Entry 1", the function function1 is
   // executed.
   menuEntries.push({name: "createTableスクリプト（mysql）", functionName: "mySqlCreate"});
   menuEntries.push({name: "createTableスクリプト（sqlServer）", functionName: "sqlServerCreate"});
   //menuEntries.push(null); // line separator     
   ss.addMenu("自動処理", menuEntries);  
};



function mySqlCreate() {
    
  var outputString = "";
  
  outputString += setmySqlCreateTableMain();
 
  outputString += "            " + setmySqlCreatePrimaryKeyMain();
  
  Browser.msgBox(outputString);
}

function sqlServerCreate() {
    
  var outputString = "";
  
  outputString += setSqlCreateTableMain();
 
  outputString += "            " + setSqlCreatePrimaryKeyMain();
  
  Browser.msgBox(outputString);
}

//----------------------------------------------------------------------
//↓↓↓↓テーブル作成詳細(mysql)↓↓↓↓
//----------------------------------------------------------------------
function setmySqlCreateTableMain() {

  var firstFlg = true;  //初めの行はカンマが必要ない為  
  var setSql = "";  //sql文を保持する用
  //シートとアプリ
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  //【SQL文作成】テーブル名セット  
  setSql += setmySqlCreateTableStart(sheet.getRange(3, 5).getValue());
  
  //A6から縦にデータが存在チェックをする
  var rowNum = 6;
  var colNum = 1;
  while(true)
  {
    
    var colName     = sheet.getRange(rowNum, 4).getValue();
    var dataType    = sheet.getRange(rowNum, 6).getValue();
    var dataSize    = sheet.getRange(rowNum, 7).getValue();
    var dataNull    = sheet.getRange(rowNum, 8).getValue();
    var dataDefault = sheet.getRange(rowNum, 9).getValue();    
    
    if(firstFlg == true)
    {
      setSql += " "
    }
    else
    {
      setSql += " ,"
    }
    
 
        
    setSql += setmySqlCreateTableMiddle(colName, dataType, dataSize, dataNull, dataDefault);
      
    rowNum++;
    firstFlg = false;
    //No行が空の場合にブレイク
    if(sheet.getRange(rowNum, colNum).getValue() == "")
    {
      setSql += setmySqlCreateTableEnd();
      break;
    }
  }  
  
  return setSql;
}


function setmySqlCreateTableStart(tableName) {

  var strSql = "create table " + tableName + " ( ";
  
  return strSql;
}

function setmySqlCreateTableMiddle(colName, dataType, dataSize, dataNull, dataDefault) {

  var strSql = "";
  var strSqlDataSize = "";
  if(dataSize != "" && dataSize != "-")
  {
    if(dataType == "numeric")
    {      
      strSqlDataSize = "(" + dataSize + ")";
    }
    else if(dataType == "int")
    {
      //none
    }
    else
    {
      strSqlDataSize = "(" + dataSize + ")";
    }
  }
  
  strSql = " " + colName + " " + dataType + strSqlDataSize + " " + dataNull;
    
  return strSql;
  
}


function setmySqlCreateTableEnd() {
  
  var strSql = ");";
  
  return strSql;  
}
//----------------------------------------------------------------------
//↑↑↑↑テーブル作成詳細(mysql)↑↑↑↑
//----------------------------------------------------------------------

//----------------------------------------------------------------------
//↓↓↓↓テーブル主キー作成詳細(mysql)↓↓↓↓
//----------------------------------------------------------------------
function setmySqlCreatePrimaryKeyMain() {

  var firstFlg = true;  //初めの行はカンマが必要ない為  
  var setSql = "";  //sql文を保持する用
  //シートとアプリ
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  //【SQL文作成】テーブル名セット  
  setSql += setmySqlCreatePrimaryKeyStart(sheet.getRange(3, 5).getValue());
  
  //A6から縦にデータが存在チェックをする
  var rowNum = 6;
  var colNum = 1;
  while(true)
  {
    
    var colName     = sheet.getRange(rowNum, 4).getValue();
    var primarykey  = "";
    
    if(sheet.getRange(rowNum, 2).getValue() == "●")
    {
      
      if(firstFlg == true)
      {
        setSql += " "
      }
      else
      {
        setSql += " ,"
      }     
    
      setSql += setmySqlCreatePrimaryKeyMiddle(colName);
    }
      
    rowNum++;
    firstFlg = false;
    //No行が空の場合にブレイク
    if(sheet.getRange(rowNum, colNum).getValue() == "")
    {
      setSql += setmySqlCreatePrimaryKeyEnd();
      break;
    }
  }  
  
  return setSql;
}


function setmySqlCreatePrimaryKeyStart(tableName) {

  var strSql = "alter table " + tableName + " add primary key ( ";
  
  return strSql;
}

function setmySqlCreatePrimaryKeyMiddle(colName) {

  var strSql = "";  
  
  strSql = " " + colName + " ";
    
  return strSql;  
}


function setmySqlCreatePrimaryKeyEnd() {
  
  var strSql = ");";
  
  return strSql;  
}
//----------------------------------------------------------------------
//↑↑↑↑テーブル主キー作成詳細(mysql)↑↑↑↑
//----------------------------------------------------------------------



//----------------------------------------------------------------------
//↓↓↓↓テーブル作成詳細(sqlServer)↓↓↓↓
//----------------------------------------------------------------------
function setSqlCreateTableMain() {

  var firstFlg = true;  //初めの行はカンマが必要ない為  
  var setSql = "";  //sql文を保持する用
  //シートとアプリ
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  //【SQL文作成】テーブル名セット  
  setSql += setSqlCreateTableStart(sheet.getRange(3, 5).getValue());
  
  //A6から縦にデータが存在チェックをする
  var rowNum = 6;
  var colNum = 1;
  while(true)
  {
    
    var colName     = sheet.getRange(rowNum, 4).getValue();
    var dataType    = sheet.getRange(rowNum, 6).getValue();
    var dataSize    = sheet.getRange(rowNum, 7).getValue();
    var dataNull    = sheet.getRange(rowNum, 8).getValue();
    var dataDefault = sheet.getRange(rowNum, 9).getValue();    
    
    if(firstFlg == true)
    {
      setSql += " "
    }
    else
    {
      setSql += " ,"
    }
    
 
        
    setSql += setSqlCreateTableMiddle(colName, dataType, dataSize, dataNull, dataDefault);
      
    rowNum++;
    firstFlg = false;
    //No行が空の場合にブレイク
    if(sheet.getRange(rowNum, colNum).getValue() == "")
    {
      setSql += setSqlCreateTableEnd();
      break;
    }
  }  
  
  return setSql;
}


function setSqlCreateTableStart(tableName) {

  var strSql = "create table " + tableName + " ( ";
  
  return strSql;
}

function setSqlCreateTableMiddle(colName, dataType, dataSize, dataNull, dataDefault) {

  var strSql = "";
  var strSqlDataSize = "";
  if(dataSize != "" && dataSize != "-")
  {
    if(dataType == "numeric")
    {      
      strSqlDataSize = "(" + dataSize + ")";
    }
    else if(dataType == "int")
    {
      //処理なし
    }
    else
    {
      strSqlDataSize = "(" + dataSize + ")";
    }
  }
  
  strSql = " " + colName + " " + dataType + strSqlDataSize + " " + dataNull;
    
  return strSql;
  
}


function setSqlCreateTableEnd() {
  
  var strSql = ");";
  
  return strSql;  
}
//----------------------------------------------------------------------
//↑↑↑↑テーブル作成詳細↑↑↑↑(sqlServer)
//----------------------------------------------------------------------

//----------------------------------------------------------------------
//↓↓↓↓テーブル主キー作成詳細(sqlServer)↓↓↓↓
//----------------------------------------------------------------------
function setSqlCreatePrimaryKeyMain() {

  var firstFlg = true;  //初めの行はカンマが必要ない為  
  var setSql = "";  //sql文を保持する用
  //シートとアプリ
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  //【SQL文作成】テーブル名セット  
  setSql += setSqlCreatePrimaryKeyStart(sheet.getRange(3, 5).getValue());
  
  //A6から縦にデータが存在チェックをする
  var rowNum = 6;
  var colNum = 1;
  while(true)
  {
    
    var colName     = sheet.getRange(rowNum, 4).getValue();
    var primarykey  = "";
    
    if(sheet.getRange(rowNum, 2).getValue() == "●")
    {
      
      if(firstFlg == true)
      {
        setSql += " "
      }
      else
      {
        setSql += " ,"
      }     
    
      setSql += setSqlCreatePrimaryKeyMiddle(colName);
    }
      
    rowNum++;
    firstFlg = false;
    //No行が空の場合にブレイク
    if(sheet.getRange(rowNum, colNum).getValue() == "")
    {
      setSql += setSqlCreatePrimaryKeyEnd();
      break;
    }
  }  
  
  return setSql;
}


function setSqlCreatePrimaryKeyStart(tableName) {

  var strSql = "alter table " + tableName + " add CONSTRAINT  PK_Name_" + Math.floor(Math.random() * 10000) + Math.floor(Math.random() * 100) + " primary key CLUSTERED ( ";
  
  return strSql;
}

function setSqlCreatePrimaryKeyMiddle(colName) {

  var strSql = "";  
  
  strSql = " " + colName + " ";
    
  return strSql;  
}


function setSqlCreatePrimaryKeyEnd() {
  
  var strSql = ");";
  
  return strSql;  
}
//----------------------------------------------------------------------
//↑↑↑↑テーブル主キー作成詳細↑↑↑↑
//----------------------------------------------------------------------



