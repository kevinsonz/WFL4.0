// LL共通処理

// データのみ削除（何のために作った？）
function setLogClear(){
  llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDataNum, logSheetEndCol).clearContent();
}

// 行数調整
function setLogRows() {
  setLogFrozenClear(); // 表示固定（解除）
  const initRows = initRowsPerDay * maxDay; // 行数＝何行？/日×何日？
  llSheet.deleteRows(logDataStartRow, logDataNum);
  llSheet.insertRowsAfter(logFilterRow, initRows+4);
  setLogFrozen(); // 表示固定
}

// 幅
function setLogWidth() {
  const widthPrm = [25,100,50,50,50,50,50,100,200,25,25,25,25,25,50,25,25,50];
  for(i=0; i<logSheetEndCol; i++){
    llSheet.setColumnWidth(1+i,widthPrm[i]);
  }
}

// 見出し
function setLogMidashi() {
  // 項目（1行目）
  for(i=0; i<logSheetEndCol; i++){
      llSheet.getRange(logSHeetStartRow,logSheetStartCol+i).setValue(midashi[i]);
  }
  // 項目（2行目）
  llSheet.getRange(logFilterRow,logSheetStartCol).setFormula('=if(month(today())=month(B2),text(day(today()),"00")&text(hour(now()),"00")&text(minute(now()),"00"),"-")');
  llSheet.getRange(logFilterRow,logSheetAgeCol).setFormula('=countifs(J3:J,true,A3:A,false)');
  llSheet.getRange(logFilterRow,logSheetSageCol).setFormula('=countifs(K3:K,true,A3:A,false)');
  llSheet.getRange(logFilterRow,logSheetAgeChiCol).setFormula('=sumifs(L3:L,A3:A,false)');
  llSheet.getRange(logFilterRow,logSheetSageChiCol).setFormula('=sumifs(M3:M,A3:A,false)');
  llSheet.getRange(logFilterRow,logSheetErrorCol).setFormula('=sum(N3:N)');
}

// 日付関数配布
function formulaDataDay(rowStart, rows){
  llSheet.getRange(rowStart, logSheetDateCol, rows,1).setFormula('=offset($B' + rowStart + ',-1,0)+if($A' + rowStart + '=TRUE,1,0)'); // 日エリア（2行目以降）
}

// 日ヘッダー行数式設定
function formulaHeaderTime(row){
  llSheet.getRange(row, logSheetTimeStartCol).setFormula('=$B' + row + '+$E' + row + '/24');
  llSheet.getRange(row, logSheetTimeEndCol).setFormula('=$C' + row);
  llSheet.getRange(row, logSheetAgeChiCol).setFormula('=row()');
  llSheet.getRange(row, logSheetSageChiCol).setFormula('=countifs(Q7:Q,Q' + row + ')');
}

// 開始時刻列数式設定
function formulaStartTime(rowStart, rows){
  llSheet.getRange(rowStart, logSheetTimeStartCol, rows, 1).setFormula('=$B' + rowStart + '+offset($C' + rowStart + ',if(offset(C' + rowStart + ',-1,2)="01日",-5,-1),1)');
}

// 終了時刻列数式設定
function formulaEndTime(rowStart, rows){
  llSheet.getRange(rowStart, logSheetTimeEndCol, rows, 1).setFormula('=$C' + rowStart + '+$E' + rowStart + '/24');
}

// アゲサゲ数値化数式設定
function formulaAgeSageDd(rowStart, rows) {
  llSheet.getRange(rowStart, logSheetAgeChiCol, rows, 1).setFormula('=if($J' + rowStart + '=TRUE,1,0)'); // ↑値列
  llSheet.getRange(rowStart, logSheetSageChiCol, rows, 1).setFormula('=if($K' + rowStart + '=TRUE,-1,0)'); // ↓値列
  llSheet.getRange(rowStart, logSheetDdCol, rows, 1).setFormula('=day($B' + rowStart + ')'); // dd列（データ行）
}

// Error列数式設定
function formulaRuleKizami(rowStart, rows) {
  llSheet.getRange(rowStart, logSheetErrorCol, rows, 1).setFormula('=if(AND($E' + rowStart + '<>"",$E' + rowStart + '<>"01日"),if(mod($E' + rowStart + ',0.25)<>0,1,0),0)'); // Error列
}

// 年・月・年月列数式設定
function formulaYyyyMm(rowStart, rows){
  llSheet.getRange(rowStart, logSheetYyyyCol, rows, 1).setFormula('=year($B' + rowStart + ')'); // yyyy列
  llSheet.getRange(rowStart, logSheetMmCol, rows, 1).setFormula('=month($B' + rowStart + ')'); // mm列
  llSheet.getRange(rowStart, logSheetYyyymmCol, rows, 1).setFormula('=$O' + rowStart + '*100+$P' + rowStart); // yyyymm列
}

// 数式（一括）
function setLogFormulaInit() {
  logLastRowGet();
  // 日付
  llSheet.getRange(logDayDataStartRow, logSheetDateCol).setFormula('=$B$2'); // 日エリア（1行目）
  formulaDataDay(logDayDataStartRow+1, logDayDataNum-1);
  // 非表示エリア
  formulaAgeSageDd(logDataStartRow, logDataNum);
  formulaYyyyMm(logFilterRow, logDataNum+1);
  llSheet.getRange(logFilterRow, logSheetDdCol).setFormula('=day(edate($B$2,1)-1)'); // dd列（フィルタ行）

  // 時刻（開始・終了）列
  // 日ヘッダー
  formulaHeaderTime(logDayDataStartRow);
  formulaRuleKizami(logDayDataStartRow, 1);

  // 日データ行
  llSheet.getRange(logDayDataStartRow, logSheetStartCol).check();
  formulaStartTime(logDayDataStartRow+1, logDayDataNum-1);
  formulaEndTime(logDayDataStartRow+1, logDayDataNum-1);
  formulaRuleKizami(logDayDataStartRow+1, logDayDataNum-1);
}

// 入力規則（チェックボックス）
function setLogCheckBox() {
  for(i=0; i < cbCol.length; i++){
    llSheet.getRange(logDataStartRow, cbCol[i], logDataNum, 1).insertCheckboxes();
    // llSheet.getRange(logDataStartRow, cbCol[i], logDataNum, 1).setFontColor('black'); // チェックボックスを灰色から黒くしたいが、上手く行かない。
  }
}

// フォントサイズ（小）
function fontSizeSmall(row, col, rows, cols){
  llSheet.getRange(row, col, rows, cols).setFontSize(8);
}

// フォント
function setLogFont() {
  llSheet.getRange(logSHeetStartRow, logSheetStartCol, logLastRow, logSheetEndCol).setFontSize(10);
  fontSizeSmall(logSHeetStartRow, logSheetTimeStartCol, 1, 3);
  fontSizeSmall(logDataStartRow, logSheetFontSizeSmallStartCol, logDataNum, logSheetFontSizeSmallNum);
  llSheet.getRange(logSHeetStartRow, logSheetStartCol, logLastRow, logSheetEndCol).setFontFamily("Roboto");
  fontWeightReset(logSHeetStartRow, logLastRow); // 太字を一旦全て解除。
  llSheet.getRange(logSHeetStartRow, logSheetStartCol, 1, logSheetEndCol).setFontWeight("bold");
  llSheet.getRange(logFilterRow, logSheetDateCol).setFontWeight("bold");
  llSheet.getRange(logFilterRow, logSheetDateCol).setFontColor("blue");
  llSheet.getRange(logFilterRow, logSheetDateCol).setNumberFormat("yyyy/mm");
  llSheet.getRange(logDataStartRow, logSheetDateCol, logDataNum, 1).setNumberFormat("yyyy/mm/dd(ddd)");
  llSheet.getRange(logDataStartRow, logSheetTimeStartCol, logDataNum, 2).setNumberFormat("hh:mm");
  llSheet.getRange(logDataStartRow, logSheetCostCol, logDataNum, 1).setNumberFormat("0.00");
}

// 寄せ
function setLogAlignment() {
  llSheet.getRange(logSHeetStartRow, logSheetStartCol, logLastRow, logSheetEndCol).setHorizontalAlignment("center");
  llSheet.getRange(logDataStartRow, logSheetFontSizeSmallStartCol, logDataNum, logSheetFontSizeSmallNum).setHorizontalAlignment("left");
  llSheet.getRange(logSHeetStartRow, logSheetStartCol, logLastRow, logSheetEndCol).setVerticalAlignment("middle");
  llSheet.getRange(logDataStartRow, logSheetMemoCol, logDataNum, 1).setVerticalAlignment("top");
}

// フィルタ
function setLogFilter() {
  let filterStatus = llSheet.getFilter(); // フィルタ設定状態を確認。
  if(filterStatus !== null){
    llSheet.getFilter().remove();
  }
    llSheet.getRange(logFilterRow, logSheetDateCol, (logDataNum + 1), 1).createFilter();
}

// 表示固定
function setLogFrozen() {
  llSheet.setFrozenRows(2);
  llSheet.setFrozenColumns(4);
}

// 表示固定（解除）
function setLogFrozenClear(){
  llSheet.setFrozenRows(0);
  llSheet.setFrozenColumns(0);
}

// 罫線
function setLogBorder() {
  llSheet.getRange(logSHeetStartRow,logSheetStartCol,logLastRow,logLastCol).setBorder(true,true,true,true,true,true);
}

// グループ化
function setLogGroup() {
  function grp(dpt) {
    llSheet.getRange(logSHeetStartRow, logSheetGrpStartCol, logLastRow, logSheetGrpNum).shiftColumnGroupDepth(dpt);
  }
  grp(-8);
  grp(1);
  llSheet.getRange(logSHeetStartRow, logSheetGrpStartCol, logLastRow, logSheetGrpNum).collapseGroups();
}

// 入力規則（リスト）※コード
function setLogList() {
  //プルダウンの選択肢を配列で指定
  const values = ['W', 'F', 'L', 'E', 'O', 'Z'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  //リストをセットするセル範囲を取得
  const cell = llSheet.getRange(logDataStartRow, logSheetCodeCol, logDataNum, 1);
  //セルに入力規則をセット
  cell.setDataValidation(rule);
}

// 入力規則（リスト）※日ヘッダー（状態）
function headerList(row) {
  //プルダウンの選択肢を配列で指定
  const values = ['未完', '完了'];
  //入力規則を作成
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values).build();
  //リストをセットするセル範囲を取得
  const cell = llSheet.getRange(row, logSheetCodeCol);
  //セルに入力規則をセット
  cell.setDataValidation(rule);
  fontSizeSmall(row, logSheetCodeCol, 1, 1);
  llSheet.getRange(row, logSheetCodeCol).setValue('未完');
}

// 日ヘッダー色付け
function colorHeader(startRow, rows){
      llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setBackground('Maroon');
      llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setFontColor('White');
}

// MBOエリア（01週）色付け
function colorMboRowWeek(startRow, rows){
      llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setBackground('midnightblue');
      llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setFontColor('White');
}

// MBOエリア（01日）色付け
function colorMboRowDay(startRow, rows){
      llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setBackground('LightSteelBlue');
      // llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setFontColor('Black'); // 色分けのため
}

// データエリア色付け
function colorDayData(startRow, rows){
  llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setFontColor('black');
  llSheet.getRange(startRow, logSheetStartCol, rows, logSheetEndCol).setBackground(null);
}

// 色
function setLogColor() {
  llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logLastCol).setBackground(null);
  llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logLastCol).setFontColor(null);
  colorMboRowWeek(logDataStartRow, 4);
  for(i=0; i<logDayDataNum; i++){
    const headerFlagData = llSheet.getRange(logDayDataStartRow+i, logSheetStartCol).getValue();
    const headerFlag = headerFlagData === true;
    if(headerFlag){
      colorHeader(logDayDataStartRow + i, 1);
      colorMboRowDay(logDayDataStartRow + i + 1, 4);
    }
  }
}

// A列隠し
function setLogHideCol(){
  llSheet.hideColumns(logSheetStartCol, 1);
}