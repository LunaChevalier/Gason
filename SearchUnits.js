/**
 * 行番号取得
 * @param {Array} data 検索対象配列
 * @param {Object} val 検索対象オブジェクト
 * @param {number} col 検索対象列番号
 * @param {number} row 検索開始列
 */
function getRow(data, val, col, row) {
  for(var i = row; i < data.length; i++){
    if(data[i][col] === val){
      return i;
    }
  }
  return -1;
}

/**
 * 列番号取得
 * @param {Array} data 検索対象配列
 * @param {Object} val 検索対象オブジェクト
 * @param {number} row 検索対象行番号
 */
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

function getLastIndex(sheet){
  return [sheet.getLastRow(), sheet.getLastColumn()]
}
