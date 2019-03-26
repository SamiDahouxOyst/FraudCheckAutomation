/**
 * @author Dahoux Sami sami.dahoux@oyst.com
 * @date 24 Mars 2019
 * @brief Automatisation de la vérificaiton de fraude
**/

NUMVERIFY_API_TOKEN = "72d8593e02b78207760b1f7a95fe542f";

var cols = {
  createdAt:"A",
  merchant:"C",
  amount:"F",
  phoneNumber:"I",
  email:"H",
  address:"M",
  fullName:"O"
};

var labels = {
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
  _insertNumverifyData(data);
  for(var i = 0; i < data.count; i++) {
    Logger.log(data.createdAt[i]);
    Logger.log(data.carrier[i]);
  }
}

function _extractAll(document) {
  var count = _countTransactions(document);
  
  return {
    count:count,
    createdAt:_extractValues(document, _allCol(cols.createdAt), count, "Date"),
    merchant:_extractValues(document, _allCol(cols.merchant), count, "String"), 
    amount:_extractValues(document, _allCol(cols.amount), count, "Number"), 
    phoneNumber:_extractValues(document,_allCol(cols.phoneNumber), count, "String"), 
    email:_extractValues(document, _allCol(cols.email), count, "String"), 
    address:_extractValues(document, _allCol(cols.address), count, "String")
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

function _countTransactions(document) {
  
  // Dates et nombre de transactions
  var createdAt = document.getRangeByName(_allCol(cols.createdAt)).getValues();
  var count = 0;
  while(createdAt[count + 1] != "") {
    count++;
  }
  return count;
}

function _extractValues(document, range, count, type) {
  var values = document.getRangeByName(range).getValues().slice(1, count + 1);
  
  // Conversion des colonnes en types javascript
  for(var i = 0; i < count; i++) {
    if(type == "Date") {
      values[i] = new Date(values[i]);
    }
    else if(type == "Number") {
      values[i] = Number(values[i][0].split(" ")[0]);
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
