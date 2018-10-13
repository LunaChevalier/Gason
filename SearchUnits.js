function getRow(data, val, col, row) {
  for(var i = row; i < data.length; i++){
    if(data[i][col] === val){
      return i;
    }
  }
  return -1;
}

function getColumn(data, val, row, col) {
  var datas = data[row]
  for(var i = row; i < datas.length; i++){
    var v = datas[i]
    if(v === val){
      return i;
    }
  }
  return -1;
}

function getColumnSingle(data, val, row) {
  for(var i = row; i < data.length; i++){
    if(data[i] !== val){
      return i;
    }
  }
  return -1;
}

function getRowIndex(data, val) {
  return getRow(data, val, 0, 0)
}

function getLastIndex(sheet){
  return [sheet.getLastRow(), sheet.getLastColumn()]
}
