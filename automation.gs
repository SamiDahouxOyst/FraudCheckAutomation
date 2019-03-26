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
  fullName:"O"
};

var processedCols = {
  merchant: "A",
  amount:"B",
  phoneNumber:"D",
  email:"C",
  fullName:"E",
  orderId:"F",
  mailBlacklistRank:"G",
  phoneBlacklistRank:"H"
};

var rawLabels = {
  createdAt:"Date de création",
  merchant:"Marchand",
  amount:"Montant €",
  phoneNumber:"Téléphone",
  email:"Email",
  address:"Adresse",
  fullName:"Nom"
};

function runDemo() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  var data = _extractAll(document);
  
  var mailBlacklist = getMailBlacklist();
  var phoneBlacklist = getPhoneBlacklist();
  
  _insertNumverifyData(data);
  for(var i = 0; i < data.count; i++) {
    Logger.log(data.createdAt[i]);
    Logger.log(data.carrier[i]);
  }
}

function getMailBlacklist() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = document.getSheetByName("Blacklist Email");
  document.setActiveSheet(sheet);
  return _extractValues(document, _allCol("A"), 4742, "String");
}

function getPhoneBlacklist() {
  var document = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = document.getSheetByName("Blacklist phone");
  document.setActiveSheet(sheet);
  return {
    countryPrefix:_extractValues(document, _allCol("A"), 70, "Number"),
    rangeStart:_extractValues(document, _allCol("C"), 70, "Number"),
    rangeEnd:_extractValues(document, _allCol("D"), 70, "Number")
  };
}


function _extractAll(document) {
  var count = _countTransactions(document);
  
  return {
    count:count,
    createdAt:_extractValues(document, _allCol(rawCols.createdAt), count, "Date"),
    merchant:_extractValues(document, _allCol(rawCols.merchant), count, "String"), 
    amount:_extractValues(document, _allCol(rawCols.amount), count, "Number"), 
    phoneNumber:_extractValues(document,_allCol(rawCols.phoneNumber), count, "String"), 
    email:_extractValues(document, _allCol(rawCols.email), count, "String"), 
    address:_extractValues(document, _allCol(rawCols.address), count, "String")
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

  data.rank = [];
  for(var k = 0; k < data.count; k++) {
    data.rank.push(-1);
    for(var i = 0; i < mailBlacklist.length; i++) {
      if(mailBlacklist[i] == data.email[k].split("@")[1]) {
        data.rank[k] = i;
        break;
      }
    }
  }
}

function _insertPhoneBlacklistRank(data) {
  var phoneBlacklist = getMailPhonelist();

  // TODO : Implémenter le ranking du numéro de téléphone
  data.rank = [];
  for(var k = 0; k < data.count; k++) {
    data.rank.push(-1);
    for(var i = 0; i < phoneBlacklist.length; i++) {

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
      
      // Convertion de devises
      if(values[i] == undefined) {
        values[i].split(" ")[0];
      }
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
