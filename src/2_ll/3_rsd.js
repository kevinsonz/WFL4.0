// Resident LL常駐処理

// 変数
const runFlagData = llSheet.getRange('H2').getValue();
const runFlagCols = logLastCol;
const runFlagOff = runFlagData === '' || runFlagCols !== 18;
const runFlagOn = runFlagData !== '' && runFlagCols === 18;

// 空白行：数式挿入
function setLogFormulaBlank() {
  logColCheckError(); // 列数チェック
  if(runFlagOn){
    for(i=0; i<logDayDataNum; i++){
      const colorData = logBgColor[i][logSheetDateCol - 1];
      const colorCheck = colorData === '#ffa500';
      if(colorCheck){
        colorDayData(logDayDataStartRow + i, 1);
      }
    }
    for(i=0; i<logDayDataNum; i++){
      const blankData = logData[i][logSheetDateCol - 1];
      const blankCheck = blankData === '';
      if(blankCheck){
        const rowColor = llSheet.getRange(logDayDataStartRow + i, logSheetDateCol).getBackground();
        const rowColorFlag = (rowColor === '#ffffff') || (rowColor === '#ffa500');
        if(rowColorFlag){
          formulaDataDay(logDayDataStartRow + i, 1);
          formulaStartTime(logDayDataStartRow + i, 1);
          formulaEndTime(logDayDataStartRow + i, 1);
          formulaAgeSageDd(logDayDataStartRow + i, 1);
          formulaRuleKizami(logDayDataStartRow + i, 1);
          formulaYyyyMm(logDayDataStartRow + i, 1);
        }
      }
    }
    imaKokoError();
  }
}

// 今ココ（日ヘッダー・データ行）
// データ行
function imaKokoError(){
  if(runFlagOn){
    for(i=0; i<logDayDataNum; i++){
      const imaKokoYyyy = logDataDsp[i][logSheetYyyyCol - 1];
      const imaKokoMm = logDataDsp[i][logSheetMmCol - 1] - 1;
      const imaKokoDd = logDataDsp[i][logSheetDdCol - 1];
      const imaKokoHourStart = Number(logDataDsp[i][logSheetTimeStartCol - 1].slice(0,2));
      const imaKokoMinuteStart = Number(logDataDsp[i][logSheetTimeStartCol - 1].slice(-2));
      const imaKokoHourEnd = Number(logDataDsp[i][logSheetTimeEndCol - 1].slice(0,2));
      const imaKokoMinuteEnd = Number(logDataDsp[i][logSheetTimeEndCol - 1].slice(-2));
      const imaKokoYMDHMStart = new Date(imaKokoYyyy, imaKokoMm, imaKokoDd, imaKokoHourStart, imaKokoMinuteStart, 0);
      const imaKokoYMDHMEnd = new Date(imaKokoYyyy, imaKokoMm, imaKokoDd, imaKokoHourEnd, imaKokoMinuteEnd, 0);
      const imakokoCheck = (imaKokoYMDHMStart < nowYMDHM) && (nowYMDHM < imaKokoYMDHMEnd);
      console.log('imaKokoYMDHMStart',imaKokoYMDHMStart);
      console.log('imaKokoYMDHMEnd',imaKokoYMDHMEnd);
      console.log('imakokoCheck',imakokoCheck);
      const logHeaderCheck = logData[i][logSheetStartCol - 1];
      const logErrorCheck = logData[i][logSheetErrorCol - 1];
      if(logHeaderCheck === false){
        const rowColor = logBgColor[i][logSheetStartCol - 1];
        console.log('rowColor', rowColor);
        const rowColorCheck = rowColor === '#ff0000' || rowColor === '#ffa500';
        if(logErrorCheck === 1){
          llSheet.getRange(logDayDataStartRow + i, logSheetStartCol, 1,logSheetEndCol).setBackground('red');
          llSheet.getRange(logDayDataStartRow + i, logSheetStartCol, 1,logSheetEndCol).setFontColor('white');
        }else{
          if(imakokoCheck){
            fontWeightReset(logDayDataStartRow + i, 1, 'bold'); // 太字
            fontCodeColor(logDayDataStartRow + i); // コード色分け
            llSheet.getRange(logDayDataStartRow + i, logSheetStartCol, 1,logSheetEndCol).setBackground('orange'); // 背景色
          }else if(rowColorCheck){
            fontWeightReset(logDayDataStartRow + i, 1); // 太字解除
            fontCodeColor(logDayDataStartRow + i); // コード色分け
            llSheet.getRange(logDayDataStartRow + i, logSheetStartCol, 1,logSheetEndCol).setBackground(null); // 背景色
          }
        }
      }
    }
  }
}

// コード色分け（単体）
function fontCodeColor(row){
  const logDayCode = llSheet.getRange(row, logSheetCodeCol).getValue();
  switch (logDayCode){
  case 'W':
    fontColorReset(row, 1, 'blue');
    break;
  case 'F':
    fontColorReset(row, 1, 'green');
    break;
  case 'L':
    fontColorReset(row, 1, 'red');
    break;
  case 'E':
    fontColorReset(row, 1, 'black');
    break;
  case 'O':
    fontColorReset(row, 1, 'gray');
    break;
  case 'Z':
    fontColorReset(row, 1, 'purple');
    break;
  case '':
    fontColorReset(row, 1, 'black');
    break;
  default:
    fontColorReset(row, 1, 'black');
  }
}

// 平日⇔休日：色つけ（日ヘッダー）　※MBO作ってから？

// MBO連携

// サポート機能 ※LL編 ※2023/08/26作成中
// 配列丸ごとだと上手く比較できない？cond〜は二次元、midachiは一次元？
function supportCol(){
  if(runFlagOn){
    let supportColCheck = logSheetColCondition[0].toString() !== midashi.toString();
    console.log('logSheetColCondition', logSheetColCondition[0]);
    console.log('midashi', midashi);
    if(supportColCheck){
      console.log('おかしい状態です。');
      for(i=0; i<logLastCol; i++){
        const blankData = logSheetColCondition[0][logSheetStartCol - 1 + i];
        const blankCheck = blankData === '';
        if(blankCheck){
          llSheet.getRange(logSHeetStartRow, logSheetStartCol + i, logLastRow, 1).setBackground('red');
        }
      }
    }else{
      console.log('おかしくないです。');
    }
  }
}

// コード色分け（全体）
function codeColorReset(){
  console.log('logData', logData);
  if(runFlagOn){
    for(i=0; i<logDayDataNum; i++){
      const logHeaderCheck = logData[i][logSheetStartCol - 1];
      if(logHeaderCheck === false){
        const colorData = logFontColor[i][7].asRgbColor().asHexString();
        const codeData = logData[i][logSheetCodeCol - 1];
        let codeColor16 = ''; // 16進数という意味。
        switch (codeData){
        case 'W':
          codeColor16 = '#0000ff';
          break;
        case 'F':
          codeColor16 = '#008000';
          break;
        case 'L':
          codeColor16 = '#ff0000';
          break;
        case 'E':
          codeColor16 = '#000000';
          break;
        case 'O':
          codeColor16 = '#808080';
          break;
        case 'Z':
          codeColor16 = '#800080';
          break;
        default:
          console.log('不明');
        }
        const codeColorCheck = colorData !== codeColor16;
        console.log('i', i);
        console.log('colorData', colorData);
        console.log('codeData', codeData);
        console.log('codeColor16', codeColor16);
        console.log('codeColorCheck', codeColorCheck);
        if(codeColorCheck){
          fontCodeColor(logDayDataStartRow + i);
        }
      }
    }
  }
}

// シート初期化？なぜ常駐処理に？
function logSheetCreateCall(){
  if(runFlagOn){
    logSheetCreate();
  }
}

// サポート機能 ※MBO編