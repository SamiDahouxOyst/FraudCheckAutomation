/**
 * @author Dahoux Sami sami.dahoux@oyst.com
 * @date 24 Mars 2019
 * @brief Fraude check automation
 * @details 
 * Processes transaction to validate and displays the suspected ones that require manual check.
 *
 * Checks transactions using Oyst transactions fraud scoring documented here :
 * https://confluence.oyst.io/display/PT/Fraud+Management+and+Fraud+Scoring
 * 
 * NB : Use that code on the following spreadsheet : 
 * https://docs.google.com/spreadsheets/d/1GenNWP31LY-DIBrpREzE5uSR9K6Bsh_kE_Fnp2XNw_4/edit#gid=478485810
 * 
 */

NUMVERIFY_API_TOKEN = "72d8593e02b78207760b1f7a95fe542f";

/**
 * @brief Columns positions in sheets
 */ 

var rawCols = {
  createdAt:"A",
  merchant:"C",
  amount:"F",
  phoneNumber:"I",
  email:"H",
  address:"M",
  fullName:"O",
  merchantId:"R",
  orderId:"S"
  
};

var processedCols = {
  merchant: "A",
  amount:"B",
  phoneNumber:"D",
  email:"C",
  fullName:"E",
  orderId:"F",
  mailRank:"G",
  phoneRank:"H"
};

/**
 * @brief AOV sorted by merchantId
 */
var merchantAOV = {
  "f660bf1e-2333-4fe3-9473-b0fce7c0019d":80, // juliendorecel
  "1e7b2b33-cf87-41f2-a3c1-41e8ee3402ae":30, // louis-herboristerie
  "3c406813-17f0-4ff7-8622-16afb45df12b":40, // vetomalin
  "70cd921f-17a8-4846-962d-b177690d491d":60, // peyrouse-hair-shop
  "85f3b517-dd80-48a2-bf1f-5dcfd4eb634b":50, // lactolerance
}

/**
 * @brief strength of fraud scoring in terms of amount scoring, in ]0, 1].
 */
var transactionAmountTol = 0.5;

var document;

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Rank transactions', functionName: 'rankTransactions_'},
    {name: 'Sort transactions', functionName: 'sortTransactions_'}
  ];
  spreadsheet.addMenu('Actions', menuItems);
}

function runDemoRanked() {
  document = SpreadsheetApp.getActiveSpreadsheet();
  var data = _extractAll();
  
  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  _insertIsBlacklistedMerchant(data);
  
  for(var i = 0; i < data.count; i++) {
    Logger.log(data.createdAt[i]);
    Logger.log(data.mailRank[i]);
    Logger.log(data.phoneRank[i]);
    Logger.log(data.isBlacklisted[i]);
  }
  fillProcessedSheet(data, "Ranked");
  document.setActiveSheet(document.getSheetByName("Ranked"));
}

function runDemoSorted() {
  document = SpreadsheetApp.getActiveSpreadsheet();
  var extData = _extractSuspected();
  
  for(var i = 0; i < extData.count; i++) {
    Logger.log(extData.createdAt[i]);
    Logger.log(extData.mailRank[i]);
    Logger.log(extData.phoneRank[i]);
    Logger.log(extData.isBlacklisted[i]);
  }
  fillProcessedSheet(extData, "Sorted");
  document.setActiveSheet(document.getSheetByName("Sorted"));
}

/**
 * @brief Ranks transactions and fill Ranked sheet
 * @details Fill Raw sheet with transactions so this function can extract data and process it
 * 
 */
function rankTransactions_() {
  document = SpreadsheetApp.getActiveSpreadsheet();
  var data = _extractAll();
  
  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  _insertIsBlacklistedMerchant(data);
  
  fillProcessedSheet(data, "Ranked");
  document.setActiveSheet(document.getSheetByName("Ranked"));
}

/**
 * @brief Ranks transactions and fill Sorted sheet with suspected ones
 * @details Fill Raw sheet with transactions so this function can extract data and process it
 * The Sorted and Ranked sheets have the same data format but Sorted contains only suspected transactions,
 * theses ones require manual check.
 * 
 */
function sortTransactions_() {
  document = SpreadsheetApp.getActiveSpreadsheet();
  var data = _extractSuspected();
  
  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  _insertIsBlacklistedMerchant(data);

  fillProcessedSheet(data, "Sorted");
  document.setActiveSheet(document.getSheetByName("Sorted"));
}

/**
 * @brief extract values from Blacklist Email sheet
 * @return array with mail address sufix from the mail blacklist
 */
function getMailBlacklist() {
  var sheet = document.getSheetByName("Blacklist Email");
  return _extractValues(sheet, "A", 4742, "String");
}

/**
 * @brief extract values from Blacklist Phone sheet
 * @return object arrray with country prefix, starting range and ending range from the phone blacklist
 */
function getPhoneBlacklist() {
  var sheet = document.getSheetByName("Blacklist phone");
  
  return {
    count: 70,
    countryPrefix:_extractValues(sheet, "A", 70, "Number"),
    rangeStart:_extractValues(sheet, "C", 70, "Number"),
    rangeEnd:_extractValues(sheet, "D", 70, "Number")
  };
}

/**
 * @brief extract values from Blacklist Merchant sheet
 * @return array with merchant ids from the merchant blacklist
 */
function getMerchantBlacklist() {
  var sheet = document.getSheetByName("Blacklist merchant");  
  return _extractValues(sheet, "A", 1, "String");
}

/**
 * @brief Fills a sheet with processed data
 * @param data object array containing transactions rows
 * @param name string containing the name of the sheet to be filled
 * @details 
 * data object must define the same properties as processedCols object.
 * name string must be the name of an existing sheet.
 *
 */
function fillProcessedSheet(data, name) {
  var sheet = document.getSheetByName(name);
  
  Object.keys(processedCols).forEach(function(key){
    if(key == "count") 
      return;
    
    range = sheet.getRange(processedCols[key] + "2:" + processedCols[key] + (data.count + 1));
    range.setValues(_prepareForRange(data[key]));
  });
}

/**
 * @brief extracts all transactions from Raw sheet
 * @return object array containing values of rows
 */
function _extractAll() {
  var sheet = document.getSheetByName("Raw");
  var count = _countTransactions();
  
  return {
    count:count,
    createdAt:_extractValues(sheet, rawCols.createdAt, count, "Date"),
    fullName:_extractValues(sheet, rawCols.fullName, count, "String"),
    merchant:_extractValues(sheet, rawCols.merchant, count, "String"), 
    amount:_extractValues(sheet, rawCols.amount, count, "Currency"), 
    phoneNumber:_extractValues(sheet,rawCols.phoneNumber, count, "String"), 
    email:_extractValues(sheet, rawCols.email, count, "String"), 
    address:_extractValues(sheet, rawCols.address, count, "String"),
    merchantId:_extractValues(sheet, rawCols.merchantId, count, "String"),
    orderId:_extractValues(sheet, rawCols.orderId, count, "String")
  };
}

/**
 * @brief extracts supected transactions from Raw sheet
 * @details 
 * The values are filtered from the output of _extractAll function acording to fraud scoring. 
 * Only supsicious transactions are returned.
 * @return object array containing values of rows
 */
function _extractSuspected() {
  var data = _extractAll();
  var extData = {
    count:0,
    merchant:[],
    createdAt:[],
    amount:[],
    phoneNumber:[],
    email:[],
    address:[],
    fullName:[],
    orderId:[],
    merchantId:[],
    phoneRank:[],
    mailRank:[],
    isBlacklisted:[]
  }
  
  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  _insertIsBlacklistedMerchant(data);
  
  for(var i = 0; i < data.count; i++) {
    if(data.phoneRank[i] == -1 && 
       data.mailRank[i] == -1 && 
       data.isBlacklisted[i] == false &&
       data.amount[i] < (1.0 / transactionAmountTol) * merchantAOV[data.merchantId[i]])
      continue;
  
    Object.keys(data).forEach(function(key){
      if(key == "count")      
        return;
      extData[key].push(data[key][i]);
    });
  }
  extData.count = extData.createdAt.length;
  
  return extData;
}

/**
 * @brief Inserts carrier fetched from numverify into data object
 * @param data object array containing transactions rows
 * @details 
  *data object is modified after the execution. 
 * data object must at least have the properties of rawCols object.
 *
 */
function _insertNumverifyData(data) {
  
  data.carrier = [];
  for(var i = 0; i < data.count; i++) {
    var url = 'http://apilayer.net/api/validate'
    + '?access_key=' + NUMVERIFY_API_TOKEN
    + '&number=' + data.phoneNumber[i]
    + "&format=1";
    
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    var json = response.getContentText();
    var numverifyData = JSON.parse(json);
    data.carrier.push(numverifyData.carrier);
  }
}

/**
 * @brief Inserts mail rank from mail blacklist sheet into data object
 * @param data object array containing transactions rows
 * @details 
 * data object is modified after the execution. 
 * data object must at least have the properties of rawCols object.
 * The mail blacklist sheet must be ordered by over-representation over datasets.
 *
 */
function _insertMailBlacklistRank(data) {
  var mailBlacklist = getMailBlacklist();

  data.mailRank = [];
  for(var k = 0; k < data.count; k++) {
    data.mailRank.push(-1);
    for(var i = 0; i < mailBlacklist.length; i++) {
      if(mailBlacklist[i] == data.email[k].split("@")[1]) {
        data.mailRank[k] = i + 1;
        break;
      }
    }
  }
}

/**
 * @brief Inserts phone rank from mail blacklist sheet into data object
 * @param data object array containing transactions rows
 * @details 
 * data object is modified after the execution. 
 * data object must at least have the properties of rawCols object.
 * The phone blacklist sheet must be ordered by over-representation over datasets.
 *
 */
function _insertPhoneBlacklistRank(data) {
  var phoneBlacklist = getPhoneBlacklist();

  data.phoneRank = [];
  var prefix2, prefix3, range;
  for(var k = 0; k < data.count; k++) {
    prefix2 = Number(data.phoneNumber[k].slice(0, 2)); // Indicatif de pays à 2 numéros. ex : 33
    prefix3 = Number(data.phoneNumber[k].slice(0, 3)); // Indicatif de pays à 3 numéros. ex : 212
    
    data.phoneRank.push(-1);
    for(var i = 0; i < phoneBlacklist.count; i++) {
      if(phoneBlacklist.countryPrefix[i] == prefix2 || phoneBlacklist.countryPrefix[i] == prefix3) {
        if(prefix2 != 33) {
          data.phoneRank[k] = i + 1;
          break;
        } else {
          range = Number(data.phoneNumber[k].slice(3));
          if(range >= phoneBlacklist.rangeStart[i] && range <= phoneBlacklist.rangeEnd[i]) {
            data.phoneRank[k] = i + 1;
            break;
          }
        }
      }
    }
  }
}

/**
 * @brief Inserts phone rank from merchant blacklist sheet into data object
 * @param data object array containing transactions rows
 * @details 
 * data object is modified after the execution. 
 * data object must at least have the properties of rawCols object.
 */
function _insertIsBlacklistedMerchant(data) {
  var merchantBlacklist = getMerchantBlacklist();
  
  data.isBlacklisted = []
  for(var k = 0; k < data.count; k++) {
    data.isBlacklisted.push(merchantBlacklist.indexOf(data.merchantId[k]) != -1);
  }
}

function _countTransactions() {
  var sheet = document.getSheetByName("Raw");
  var createdAt = sheet.getRange(_allCol(rawCols.createdAt)).getValues();
  var count = 1;
  while(createdAt[count] != "") {
    count++;
  }
  return count - 1;
}

function _extractValues(sheet, col, count, type) {
  var values = sheet.getRange(_allCol(col)).getValues().slice(1, count + 1);
  
  // Converts columns in array of js types
  for(var i = 0; i < count; i++) {
    if(type == "Date") {
      values[i] = new Date(values[i]);
    }
    else if(type == "Number") {
      values[i] = Number(values[i][0]);
    }
    else if(type == "Currency") {
      values[i] = Number(String(values[i]).split(" ")[0]);
    }
    else if(type == "String") {
      values[i] = String(values[i]);
    }
  }
  return values;
}

function _prepareForRange(array) {
  for(var i = 0; i < array.length; i++) {
    array[i] = [array[i]];
  }
  return array;
}

function _allCol(col) {
  return col + ":" + col;
}
