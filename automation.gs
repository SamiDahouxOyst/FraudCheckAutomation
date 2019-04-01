/**
 * @author Dahoux Sami sami.dahoux@oyst.com
 * @date 24 Mars 2019
 * @brief Automatisation de la vérificaiton de fraude
**/

NUMVERIFY_API_TOKEN = "72d8593e02b78207760b1f7a95fe542f";

var rawCols = {
  createdAt:"A",
  merchant:"C",
  amount:"F",
  phoneNumber:"I",
  email:"H",
  address:"M",
  fullName:"O",
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

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Process transactions', functionName: 'processTransactions_'},
  ];
  spreadsheet.addMenu('Actions', menuItems);
}

function runDemo() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  var data = _extractAll(document);

  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  // _insertNumverifyData(data);
  fillProcessedSheet(data);
    
  for(var i = 0; i < data.count; i++) {
    Logger.log(data.createdAt[i]);
    // Logger.log(data.carrier[i]);
    Logger.log(data.mailRank[i]);
    Logger.log(data.phoneRank[i]);
  }
}

function getMailBlacklist() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  document.setActiveSheet(document.getSheetByName("Blacklist Email"));
  
  return _extractValues(document, _allCol("A"), 4742, "String");
}

function getPhoneBlacklist() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  document.setActiveSheet(document.getSheetByName("Blacklist phone"));
  
  return {
    count: 70,
    countryPrefix:_extractValues(document, _allCol("A"), 70, "Number"),
    rangeStart:_extractValues(document, _allCol("C"), 70, "Number"),
    rangeEnd:_extractValues(document, _allCol("D"), 70, "Number")
  };
}

function fillProcessedSheet(data) {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = document.getSheetByName("Processed");
 
  document.setActiveSheet(sheet);
  Object.keys(processedCols).forEach(function(key){
    if(key == "count") 
      return;
    
    range = sheet.getRange(processedCols[key] + "2:" + processedCols[key] + (data.count + 1));
    range.setValues(_prepareForRange(data[key]));
  });
}

function processTransactions_() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  document.setActiveSheet(document.getSheetByName("Raw"));
  var data = _extractAll(document);
  
  _insertMailBlacklistRank(data);
  _insertPhoneBlacklistRank(data);
  
  fillProcessedSheet(data);
}

function _extractAll(document) {
  document.setActiveSheet(document.getSheetByName("Raw"));
  
  var count = _countTransactions(document);
  
  return {
    count:count,
    createdAt:_extractValues(document, _allCol(rawCols.createdAt), count, "Date"),
    fullName:_extractValues(document, _allCol(rawCols.fullName), count, "String"),
    merchant:_extractValues(document, _allCol(rawCols.merchant), count, "String"), 
    amount:_extractValues(document, _allCol(rawCols.amount), count, "Currency"), 
    phoneNumber:_extractValues(document,_allCol(rawCols.phoneNumber), count, "String"), 
    email:_extractValues(document, _allCol(rawCols.email), count, "String"), 
    address:_extractValues(document, _allCol(rawCols.address), count, "String"),
    orderId:_extractValues(document, _allCol(rawCols.orderId), count, "String")
  };
}

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

function _countTransactions(document) {
  
  // Dates et nombre de transactions
  var createdAt = document.getRangeByName(_allCol(rawCols.createdAt)).getValues();
  var count = 1;
  while(createdAt[count] != "") {
    count++;
  }
  return count - 1;
}

function _extractValues(document, range, count, type) {
  var values = document.getRangeByName(range).getValues().slice(1, count + 1);
  
  // Conversion des colonnes en types javascript
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
    else {
      Logger.log("_extractValues failed : unknown type ");
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
