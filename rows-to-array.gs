function rowsToArray(keyRangeText, valueRangeText, descriptionRangeText, outputAsJSON) {
  // you need to pass the input ranges as text defining three columns,
  // plus the outputAsJSON param, for example:
  // 'B2:B20', 'E2:E20', 'F2:F20', FALSE
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var keyRange = sheet.getRange(keyRangeText);
  var valueRange = sheet.getRange(valueRangeText);
  var descriptionRange = sheet.getRange(descriptionRangeText);

  // read
  var obj = {};
  for (var i = 1; i <= keyRange.getNumRows(); i++) {
    var currentKey = keyRange.getCell(i, 1).getValue();
    obj[currentKey] = {}
  }
  for (var i = 1; i <= keyRange.getNumRows(); i++) {
    var currentKey = keyRange.getCell(i, 1).getValue();
    var currentValue = valueRange.getCell(i, 1).getValue();
    var currentDescription = descriptionRange.getCell(i, 1).getValue();
    if (currentValue == '' || currentKey == '') {
      continue;
    }
    obj[currentKey][currentValue] = currentDescription
  }

  // write
  var outputArray = []
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      var text = JSON.stringify(obj[key]);
      if (outputAsJSON === TRUE) {
        // this turns the JSON string into 'Key1: Value1 \n Key2: Value2'
        text = text.replace(/(\{|\})/g, '')
        text = text.replace(/\"\,\"/g, '\n')
        text = text.replace(/\"\:\"/g, ': ')
        text = text.replace(/"/g, '')
      }
      var arr = [key, text]
      outputArray.push(arr)
    }
  }
  return outputArray;
}
