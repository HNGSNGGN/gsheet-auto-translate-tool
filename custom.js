// flatten + 40000文字制限 + 多言語エラーチェック
function hasLoadingOrErrorStrings(values) {
  var flat = [].concat.apply([], values);
  var errorStrings = [
    '로드 중...', '#ERROR!', 'Loading...', '読み込み中...', '正在加载...', '正在加載...', '読み込み中', 'Chargement en cours...', 'Cargando...'
  ];
  for (var i = 0; i < errorStrings.length; i++) {
    if (flat.indexOf(errorStrings[i]) !== -1) {
      return true;
    }
  }
  return false;
}

// D列で最後のデータがある行番号を安全に見つける関数
function getLastDataRowInColumn(sheet, col, startRow) {
  var values = sheet.getRange(col + startRow + ':' + col + sheet.getMaxRows()).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      return startRow + i;
    }
  }
  return startRow - 1;
}

// エラー文字列判別関数
function isErrorCell(val) {
  // valが文字列で正確に"#REF!"で始まる場合のみtrue
  return (typeof val === 'string') && /^#REF!?/.test(val);
}

// 列番号 → アルファベット (A, B, ..., Z, AA, AB, ...)
function colToLetter(col) {
  var temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// 言語コード変換関数
function getLangCode(code, forFormula) {
  var map = {
    'KOR': 'ko',
    'JPN': 'ja',
    'ENG': 'en',
    'CHN': 'zh-CN',
    'TWN': 'zh-TW',
    'AUTO': forFormula ? 'auto' : ''
  };
  return map[code] || '';
}

// ステータスメッセージ出力（管理シートR16）
function printStatus(msg) {
  var mgmtSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('管理');
  mgmtSheet.getRange('R16').setValue(msg);
}

// 現在時刻（日本時間）
function nowString() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
}

// onOpenメニュー（日本語化）
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');

  var menuPairs = [
    {src: 'R11', tgt: 'S11'},
    {divider: true},
    {src: 'R17', tgt: 'S17'},
    {src: 'R18', tgt: 'S18'},
    {src: 'R19', tgt: 'S19'}
  ];

  var menu = ui.createMenu('現在タブ更新');
  var btnIdx = 1;
  menuPairs.forEach(function(pair) {
    if (pair.divider) {
      menu.addSeparator();
    } else {
      var srcCode = mgmtSheet.getRange(pair.src).getValue();
      var tgtCode = mgmtSheet.getRange(pair.tgt).getValue();
      var srcLabel = srcCode || 'SRC';
      var tgtLabel = tgtCode || 'TGT';
      var menuName = srcLabel + " → " + tgtLabel;
      menu.addItem(menuName, 'updateCurrentSheetOnly_' + btnIdx);
      btnIdx++;
    }
  });
  menu.addToUi();
}

// 各ボタン用関数
function updateCurrentSheetOnly_1() {
  updateCurrentSheetOnlyByLang('R11', 'S11');
}
function updateCurrentSheetOnly_2() {
  updateCurrentSheetOnlyByLang('R17', 'S17');
}
function updateCurrentSheetOnly_3() {
  updateCurrentSheetOnlyByLang('R18', 'S18');
}
function updateCurrentSheetOnly_4() {
  updateCurrentSheetOnlyByLang('R19', 'S19');
}

function updateCurrentSheetOnlyByLang(srcCell, tgtCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var originalSheet = ss.getActiveSheet();
  var mgmtSheet = ss.getSheetByName('管理');
  var currentTab = originalSheet.getName();

  // 安全にlastRowを計算
  var lastRow = getLastDataRowInColumn(mgmtSheet, 'D', 6);
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('管理シートD列にデータがありません。');
    return;
  }
  var tabPairs = mgmtSheet.getRange('C6:D' + lastRow).getValues();
  var sourceTab = null, targetTab = null;
  for (var i = 0; i < tabPairs.length; i++) {
    if (tabPairs[i][1] === currentTab) {
      sourceTab = tabPairs[i][0];
      targetTab = tabPairs[i][1];
      break;
    }
  }
  if (!sourceTab || !targetTab) {
    SpreadsheetApp.getUi().alert('管理シートD列に現在のタブ名がありません。');
    return;
  }

  var sourceUrl = mgmtSheet.getRange('R12').getValue();
  var match = sourceUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    SpreadsheetApp.getUi().alert('有効なGoogleスプレッドシートのリンクを入力してください。');
    return;
  }
  var sourceId = match[1];
  var sourceSs = SpreadsheetApp.openById(sourceId);

  var srcSheet = sourceSs.getSheetByName(sourceTab);
  if (!srcSheet) {
    SpreadsheetApp.getUi().alert('元のシートに該当するタブがありません。');
    return;
  }

  // ★ 現在のaタブ（コピーされるシート）の既存位置を記憶
  var sheets = ss.getSheets();
  var currIdx = -1;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === targetTab) {
      currIdx = i + 1; // moveActiveSheetは1-based index
      break;
    }
  }

  // 既存ターゲットシート削除後コピー（aタブ）
  var oldTarget = ss.getSheetByName(targetTab);
  if (oldTarget) ss.deleteSheet(oldTarget);
  var copied = srcSheet.copyTo(ss);
  copied.setName(targetTab);

  // ★ コピーされたaタブを既存位置へ移動
  if (currIdx > 0) {
    ss.setActiveSheet(copied);
    ss.moveActiveSheet(currIdx);
  }

  // 2. 元シートコピー → 一時シート（tmp2）生成（書式用）
  var tmp2Name = sourceTab + '_format_tmp';
  var oldTmp2 = ss.getSheetByName(tmp2Name);
  if (oldTmp2) ss.deleteSheet(oldTmp2);
  var tmp2 = srcSheet.copyTo(ss);
  tmp2.setName(tmp2Name);

  // 3. 翻訳数式入力およびflatten検査（一時シートtmp）
  var helperTab = sourceTab + '_tmp';
  var oldHelper = ss.getSheetByName(helperTab);
  if (oldHelper) ss.deleteSheet(oldHelper);

  var dataRange = srcSheet.getDataRange();
  var numRows = dataRange.getNumRows();
  var numCols = dataRange.getNumColumns();
  var endCell = colToLetter(numCols) + numRows;
  var rangeStr = "A1:" + endCell;

  var srcCode = mgmtSheet.getRange(srcCell).getValue();
  var tgtCode = mgmtSheet.getRange(tgtCell).getValue();
  var srcLangFormula = getLangCode(srcCode, true);
  var tgtLangFormula = getLangCode(tgtCode, true);

  var formula = '=MAP(IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + sourceId + '", "' + sourceTab + '!' + rangeStr + '"), LAMBDA(cell, IF(LEN(TRIM(CLEAN(cell)))=0, "", GOOGLETRANSLATE(cell, "' + srcLangFormula + '", "' + tgtLangFormula + '"))))';

  var newHelper = ss.insertSheet(helperTab);
  newHelper.getRange("A1").setFormula(formula);

  // flatten検査
  var maxWait = 120;
  var wait = 0;
  var isLoading = true;
  while (isLoading && wait < maxWait) {
    SpreadsheetApp.flush();
    Utilities.sleep(3000); // 3秒待機
    var vals = newHelper.getRange(rangeStr).getValues();
    isLoading = hasLoadingOrErrorStrings(vals);
    wait += 3;
  }

  if (!isLoading && copied) {
    var targetRange = copied.getRange(rangeStr);
    targetRange.clearDataValidations();

    var values = newHelper.getRange(rangeStr).getValues();

    for (var row = 0; row < numRows; row++) {
      for (var col = 0; col < numCols; col++) {
        var translatedVal = values[row][col];
        var srcRich = srcSheet.getRange(row + 1, col + 1).getRichTextValue();
        var targetCell = copied.getRange(row + 1, col + 1);

        if (isErrorCell(translatedVal)) continue;

        if (srcRich && srcRich.getLinkUrl()) {
          var link = srcRich.getLinkUrl();
          var richVal = SpreadsheetApp.newRichTextValue()
            .setText(translatedVal)
            .setLinkUrl(link)
            .build();
          targetCell.setRichTextValue(richVal);
        } else {
          targetCell.setValue(translatedVal);
        }
      }
    }


    // ✅ 書式コピー（tmp2 → 翻訳結果シート）
    var formatRange = tmp2.getRange(rangeStr);
    var tgtRange = copied.getRange(rangeStr);
    formatRange.copyTo(tgtRange, {formatOnly: true});
  } else {
    SpreadsheetApp.getUi().alert("公式計算が120秒以内に完了しませんでした。仮シート（" + helperTab + "）で直接ご確認ください。");
  }

  // 一時シート削除
  ss.deleteSheet(newHelper);
  ss.deleteSheet(tmp2);

  // ★ 最終的に該当タブをアクティブ化
  ss.setActiveSheet(copied);
}
