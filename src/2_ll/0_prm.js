// Parameter LL共通変数

// ファイル・シート
const llSheet = wflFile.getSheetByName("Log"); // シート

// 行
const logSHeetStartRow = 1; // 開始行
const logFilterRow = 2; // フィルタ行
const logDataStartRow = 3; // データ開始行
const logDayDataStartRow = 7; // 日データ開始行
let logLastRow = '';
let logDataNum = '';
let logDayDataNum = '';
function logLastRowGet(){
  logLastRow = llSheet.getMaxRows(); // 最終行（シート全体）
  logDataNum = logLastRow - logFilterRow; // 最終行数（データ）
  logDayDataNum = logLastRow - logFilterRow - 4; // 週MBO*4行分を追加した数値。
}
logLastRowGet(); // 実行

// 列
const logSheetStartCol = 1; // 開始列
const logSheetDateCol = 2; // 日付列
const logSheetTimeStartCol = 3; // 開始時刻列
const logSheetTimeEndCol = 4; // 終了時刻列
const logSheetCostCol = 5; // 時間列
const logSheetCodeCol = 6; // コード列
const logSheetKubunCol = 7; // 区分列
const logSheetKoumokuCol = 8; // 項目列
const logSheetMemoCol = 9; // メモ列
const logSheetAgeCol = 10; // ↑列
const logSheetSageCol = 11; // ↓列
const logSheetAgeChiCol = 12; // ↑値列
const logSheetSageChiCol = 13; // ↓値列
const logSheetErrorCol = 14; // Error列
const logSheetYyyyCol = 15; // yyyy列
const logSheetMmCol = 16; // mm列
const logSheetDdCol = 17; // dd列
const logSheetYyyymmCol = 18; // yyyymm列
const logSheetEndCol = logSheetYyyymmCol; // シート最終列（設計上）
const logSheetGrpStartCol = logSheetAgeChiCol; // グループ化：開始列
const logSheetGrpNum = logSheetEndCol - logSheetGrpStartCol + 1; // グループ化：対象列数
const logSheetFontSizeSmallStartCol = logSheetKubunCol; // フォントサイズ小：開始列
const logSheetFontSizeSmallNum = logSheetMemoCol - logSheetKubunCol + 1; // フォントサイズ小：対象列数
const logLastCol = llSheet.getMaxColumns(); // シート最終列（状態）
const cbCol = [logSheetStartCol,logSheetAgeCol,logSheetSageCol]; // チェックボックス列
const logColCheck = logLastCol !== logSheetEndCol; // シートの列数が設計と異なるか？

// 行列
let logData = llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logSheetEndCol).getValues();
let logDataDsp = llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logSheetEndCol).getDisplayValues();
let logBgColor = llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logSheetEndCol).getBackgrounds();
let logFontColor = llSheet.getRange(logDayDataStartRow, logSheetStartCol, logDayDataNum, logSheetEndCol).getFontColorObjects();

// 他
const maxDay = llSheet.getRange(logFilterRow, logSheetDdCol).getValue(); // 何日？/月
const initRowsPerDay = 7; // 何行？/日
const midashi = ['＋','日付','開始(予定)','終了(振返)','時間(単位)','コード','区分','項目','メモ','↑','↓','↑値','↓値','Er','yyyy','mm','dd','yyyymm'];
let logSheetColCondition = llSheet.getRange(logSHeetStartRow, logSheetStartCol, 1, logLastCol).getValues(); // 項目状態取得

// 文字色設定
function fontColorReset(row, rows, color){
  llSheet.getRange(row, logSheetStartCol, rows, logLastCol).setFontColor(null);
  llSheet.getRange(row, logSheetStartCol, rows, logLastCol).setFontColor(color);
}

// 太字設定
function fontWeightReset(row, rows, weight){
  llSheet.getRange(row, logSheetStartCol, rows, logLastCol).setFontWeight('normal');
  if(weight === 'bold'){
    llSheet.getRange(row, logSheetStartCol, rows, logLastCol).setFontWeight('bold');
  }
}

// 今ココ
const todayYyyy = llSheet.getRange(logFilterRow, logSheetYyyyCol).getValue(); // 年
const todayMm = llSheet.getRange(logFilterRow, logSheetMmCol).getValue()-1; // 月
const nowDdHhMm = llSheet.getRange(logFilterRow, logSheetStartCol).getValue(); // 日・時・分
const todayDd = nowDdHhMm.substr(0,2);
const nowHh = nowDdHhMm.substr(2,2);
const nowMm = nowDdHhMm.substr(4,2);
const nowYMD = new Date(todayYyyy, todayMm, todayDd); // 年月日
const nowYMDHM = new Date(todayYyyy, todayMm, todayDd, nowHh, nowMm, 0); // 年月日時分