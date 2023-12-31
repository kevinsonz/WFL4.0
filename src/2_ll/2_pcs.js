// Process LL個別処理

// シート初期化
function logSheetCreate(){
  setLog();
  setSheetInit();
}

// 書式・数式再設定（未完成？）
function formatReset() {
  setLogWidth(); // 幅
  setLogMidashi(); // 見出し行
  // setLogFormulaInit(); // 数式 // 干渉あり
  // setLogCheckBox(); // チェックボックス // 干渉あり
  setLogFont(); // フォント
  setLogAlignment(); // 寄せ
  setLogFilter(); // フィルタ
  setLogBorder(); // 罫線
  setLogGroup(); // グループ化
  // setLogList(); // リスト // 干渉あり
  // setLogColor(); // 色（背景・文字） // MBOエリアを黒くしちゃう。
  setLogHideCol(); // A列非表示化
}

// 全処理実行（土台づくり）
function setLog() {
  setLogRows(); // 行数調整（表示固定含む）
  setLogWidth(); // 幅
  setLogMidashi(); // 見出し行
  setLogFormulaInit(); // 数式
  setLogCheckBox(); // チェックボックス
  setLogFont(); // フォント
  setLogAlignment(); // 寄せ
  setLogFilter(); // フィルタ
  setLogBorder(); // 罫線
  setLogGroup(); // グループ化
  setLogList(); // リスト
  setLogColor(); // 色（背景・文字）
  setLogHideCol(); // A列非表示化
}

// MBOエリア作成セット（①チェックボックス潰し・②文字入力（「01週」・コード・「MBO」・「目標・予定」・③数式見直し）
function mboAreaCreate(start, unit) {
  // ①
  for(j=0; j < cbCol.length; j++){
      llSheet.getRange(start, cbCol[j], 4, 1).clearDataValidations();
      llSheet.getRange(start, cbCol[j], 4, 1).setValue('-');
  }

  // ②
  llSheet.getRange(start, logSheetCostCol, 4, 1).setValue(unit);
  const codeData = ['W', 'F', 'L', 'E'];
  const codeColor = ['blue', 'green', 'red', 'black'];
  for(j=0; j<4; j++){
    llSheet.getRange(start+j, logSheetCodeCol).clearDataValidations();
    llSheet.getRange(start+j, logSheetCodeCol).setValue(codeData[j]);
    llSheet.getRange(start+j, logSheetStartCol, 1, logSheetEndCol).setFontColor(codeColor[j]);
  }
  llSheet.getRange(start, logSheetTimeStartCol, 4, 1).setValue('');
  llSheet.getRange(start, logSheetTimeEndCol, 4, 1).setValue('');
  llSheet.getRange(start, logSheetKubunCol, 4, 1).setValue('MBO');
  llSheet.getRange(start, logSheetKoumokuCol, 4, 1).setValue('目標・予定');

  if(unit === '01日'){
    formulaStartTime(start+4, 1);
  }
}

// 日ヘッダーづくり（単体処理）
function dayHeaderCreate(row) {
  llSheet.getRange(row, logSheetStartCol).check();
  llSheet.getRange(row, logSheetCodeCol).clearDataValidations();
  headerList(row);
  for(j=1; j < cbCol.length; j++){
    llSheet.getRange(row, cbCol[j]).clearDataValidations();
    llSheet.getRange(row, cbCol[j]).setValue('-');
  }
}

// 中身（日ヘッダー・MBOエリア）初期化
function setSheetInit(){
  setDayHeaderFlag();
  setMboAreaInit();
}

// 日ヘッダーづくり
function setDayHeaderFlag() { // 日ヘッダーチェックOn
  for(i=0; i<maxDay; i++){
    dayHeaderCreate(logDayDataStartRow+(initRowsPerDay*i));
    formulaHeaderTime(logDayDataStartRow+(initRowsPerDay*i));
    colorHeader(logDayDataStartRow+(initRowsPerDay*i), 1);
  }
}

// MBOエリアづくり
function setMboAreaInit() { // MBOエリア作成
  // 01週
  mboAreaCreate(logDataStartRow, '01週');
  llSheet.getRange(logDataStartRow, logSheetDateCol, 4, 1).setFormula('=($B$2)-WEEKDAY($B$2,2)+1'); // 週エリア
  colorMboRowWeek(logDataStartRow, 4);
  fontWeightReset(logDataStartRow, 4, 'bold');
  // 01日
  for(i=0; i<maxDay; i++){
    const mboDayAreaStartRow = (initRowsPerDay * i);
    mboAreaCreate(logDayDataStartRow + 1 + mboDayAreaStartRow, '01日');
    colorMboRowDay(logDayDataStartRow + 1 + mboDayAreaStartRow, 4);
  }
}

// 非表示処理
function logSheetHideRows(){
  const todayDate = llSheet.getRange("A2").getValue();
  for(i=0; i<logDayDataNum; i++){
    const headerData = logData[i][logSheetStartCol - 1];
    const headerFlag = headerData === true;
    const headerDate = logData[i][logSheetDdCol - 1];
    const dateCheck = todayDate !== headerDate;
    const completeData = logData[i][logSheetCodeCol - 1];
    const completeFlag = completeData === '完了';
    const hideCheck = headerFlag && dateCheck && completeFlag;
    if(hideCheck){
      llSheet.hideRows(logData[i][logSheetAgeChiCol - 1] + 1, logData[i][logSheetSageChiCol - 1] - 1);
    }
  }
}

// 表示処理
function logSheetShowRows(){
  llSheet.showRows(logDayDataStartRow, logDayDataNum);
}

// コード色分け
function setFontCodeColor(){
  for(i=0; i<logDayDataNum; i++){
    const logHeaderCheck = logData[i][logSheetStartCol - 1];
    if(logHeaderCheck === false){
      fontCodeColor(logDayDataStartRow + i);
    }
  }
}

// バックアップ