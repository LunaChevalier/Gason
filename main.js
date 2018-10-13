function getTestData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheetNumber = spreadsheet.getNumSheets()
  var allDatas = []

  for(var i = 0; i < sheetNumber; i++){
    var sheet = spreadsheet.getSheets()[i]
    var data = sheet.getDataRange().getValues()
    var dataClass = getDataClass(sheet, data)
    var keys = getJsonKeyRange(sheet, data)
    var endCol = getColumn(data, 'end', 1, 0)
    var jsonObj = new Object()
    var testDatas = []
    for(var j = 11; j < endCol; j++){
      testDatas.push(createJson(keys, dataClass, getTestDatas(data, j)))
    }
    jsonObj['fileName'] = data[0][4]
    jsonObj['testDatas'] = testDatas

    allDatas.push(jsonObj)
  }

  var json = JSON.stringify(allDatas)
  Logger.log(json)
  return json
}

function getJsonKeyRange(sheet, data){
  var lastIndexs = getLastIndex(sheet)
  var START_COLUMN_INDEX = 1
  var END_COLUMN_INDEX = 8
  var START_ROW_INDEX = 3
  var TAG_COUNT = lastIndexs[0] - START_ROW_INDEX
  var d = createDataInitialize(TAG_COUNT, END_COLUMN_INDEX)
  var i, j
  for(i = 0; i < TAG_COUNT; i++){
    for(j = 0; j < END_COLUMN_INDEX; j++){
      d[i][j] = data[i + START_ROW_INDEX][j + START_COLUMN_INDEX]
    }
  }
  return d
}

function getDataClass(sheet, data){
  var lastIndexs = getLastIndex(sheet)
  var DATA_CLASS_COLUMN_INDEX = 10
  var START_ROW_INDEX = 3
  var TAG_COUNT = lastIndexs[0] - START_ROW_INDEX
  var d = new Array()
  for(var i = START_ROW_INDEX; i < TAG_COUNT + START_ROW_INDEX; i++){
    d.push(data[i][DATA_CLASS_COLUMN_INDEX])
  }
  return d
}

function createDataInitialize(lowSize, colSize){
  var table = new Array(lowSize)
  for(var i = 0; i < lowSize; i++){
    table[i] = new Array(colSize)
  }

  return table
}


function getTestDatas(data, dataColum){
  var START_ROW_INDEX = 3
  var d = []
  for(var i = START_ROW_INDEX; i < data.length; i++){
    d.push(data[i][dataColum])
  }

  return d
}

function createJson(keys, dataClass, dataColum){
  var obj = new Object()
  var level = 0
  var keysJson = []
  // keyだけループする
  for(var i = 0; i < keys.length; i++){
    for(var j = 0; j < level - getLevel(keys[i]); j++){
      keysJson.pop()
    }
    // Key名を取得
    level = getLevel(keys[i])
    keysJson.push(keys[i][level])

    // クラスが設定されている場合，それを設定する
    if(dataClass[i] !== ''){
      var settingData = dataColum[i]
      if(settingData === '\"\"'){
        settingData = ''
      }
      if(settingData !== '' || dataColum[i]=== '\"\"'){
        setData(obj, keysJson, settingData)
      }
      keysJson.pop()
    }
  }

  return obj
}

function setData(settingObj, name, data){
  if(name.length == 1){
    settingObj[name[0]] = data
  } else {
    if(settingObj[name[0]] == undefined){
      settingObj[name[0]] = new Object()
    }
    var names = name.concat()
    names.shift()
    setData(settingObj[name[0]], names, data)
  }
}

function setObject(setObj, tagName){
  setObj[tagName] = new Object()
}

function getLevel(data){
  return getColumnSingle(data, '', 0)
}
