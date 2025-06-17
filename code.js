function startSequentialProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');

  // ✅ C/D列自動出力
  updateSheetList();

  // しばらく待機 (D列翻訳式の計算時間)
  Utilities.sleep(3000);

  // S21から開始タブ名を取得
  var startTab = mgmtSheet.getRange('S21').getValue();
  var ui = SpreadsheetApp.getUi();
  var msg = '';

  if (startTab && startTab.trim() !== '') {
    msg = '「' + startTab + '」から翻訳を開始しますか？\n\n(最初から実行したい場合はH23を削除してください)';
  } else {
    msg = '全てのタブを翻訳しますか？\n\n(特定のタブから実行する場合はH23でタブ名を選択してください)';
  }

  var result = ui.alert(msg, ui.ButtonSet.OK_CANCEL);
  if (result !== ui.Button.OK) {
    ui.alert('実行がキャンセルされました。');
    return;
  }


  // 初期化
  var prop = PropertiesService.getScriptProperties();
  prop.setProperty('current_index', '0');
  prop.setProperty('total_tabs', '0');

  // D列から処理するタブリストを取得
  var lastRow = getLastDataRowInColumn(mgmtSheet, 'D', 6);
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('管理シートD列にデータがありません。');
    return;
  }
  var tabPairs = mgmtSheet.getRange('C6:D' + lastRow).getValues();

  // S21が空でない場合、D列で開始タブの位置を検索
  var startIndex = 0;
  if (startTab && startTab.trim() !== '') {
    for (var i = 0; i < tabPairs.length; i++) {
      if (tabPairs[i][1] === startTab) { // D列(targetTab)で検索
        startIndex = i;
        break;
      }
    }
  }
  // S21が空の場合は startIndex = 0 (最初から)

  // 開始インデックスからタブリストをPropertiesに保存
  var actualTabs = tabPairs.slice(startIndex); // 開始タブから最後まで
  for (var i = 0; i < actualTabs.length; i++) {
    prop.setProperty('tab_' + i + '_source', actualTabs[i][0]);
    prop.setProperty('tab_' + i + '_target', actualTabs[i][1]);
  }
  prop.setProperty('total_tabs', actualTabs.length.toString());

  if (startTab && startTab.trim() !== '') {
    console.log('開始: ' + startTab + 'から ' + actualTabs.length + '個のタブを処理予定');
  } else {
    console.log('開始: 最初から ' + actualTabs.length + '個のタブを処理予定');
  }

  // 最初のタブ処理を開始
  processNextTab();
}



// 次のタブ処理関数（トリガーで呼び出される）
function processNextTab() {
  // ✅ まず現在のトリガーを削除
  deleteCurrentTriggers();
  
  var prop = PropertiesService.getScriptProperties();
  var currentIndex = parseInt(prop.getProperty('current_index') || '0');
  var totalTabs = parseInt(prop.getProperty('total_tabs') || '0');
  
  if (currentIndex >= totalTabs) {
    // 全てのタブ完了 → シートソート後終了
    sortSheetsByDColumnOrder();
    console.log('全てのタブ処理が完了しました！');
    // ✅ 全てのトリガーを整理
    deleteAllMyTriggers();
    return;
  }
  
  var sourceTab = prop.getProperty('tab_' + currentIndex + '_source');
  var targetTab = prop.getProperty('tab_' + currentIndex + '_target');
  
  if (!sourceTab || !targetTab) {
    // 空のタブはスキップして次へ
    scheduleNextTab(currentIndex + 1);
    return;
  }
  
  try {
    // 現在のタブを処理
    processOneTab(sourceTab, targetTab, 'R11', 'S11');
    
    // 進行状況ログ
    var progress = Math.round((currentIndex + 1) / totalTabs * 100);
    console.log(progress + '% 完了: ' + targetTab + ' 処理完了');
    
    // 次のタブのためのトリガーを予約
    scheduleNextTab(currentIndex + 1);
    
  } catch (e) {
    console.log('エラー発生: ' + targetTab + ' - ' + e);
    // エラーが発生しても次のタブへ
    scheduleNextTab(currentIndex + 1);
  }
}

// 次のタブ処理のためのトリガー予約
function scheduleNextTab(nextIndex) {
  var prop = PropertiesService.getScriptProperties();
  prop.setProperty('current_index', nextIndex.toString());
  
  // 1秒後に次のタブ処理トリガーを作成
  ScriptApp.newTrigger('processNextTab')
    .timeBased()
    .after(1000)
    .create();
}

// ====== 個別タブ処理関数 ======

function processOneTab(sourceTab, targetTab, srcCell, tgtCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  
  printStatus(nowString() + ' : ' + targetTab + 'タブ翻訳中...（5分以上停止時はこのタブから再実行）');
  
  var sourceUrl = mgmtSheet.getRange('R12').getValue();
  var match = sourceUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) throw new Error('有効なGoogleスプレッドシートのリンクを入力してください。');
  var sourceId = match[1];
  var sourceSs = SpreadsheetApp.openById(sourceId);

  var srcSheet = sourceSs.getSheetByName(sourceTab);
  if (!srcSheet) throw new Error(sourceTab + 'が元のシートに存在しません。');

  // 1. 既存のターゲットシートを削除後コピー（aタブ）
  var oldTarget = ss.getSheetByName(targetTab);
  if (oldTarget) ss.deleteSheet(oldTarget);
  var copied = srcSheet.copyTo(ss);
  copied.setName(targetTab);

  // 2. 元のシートをコピー → 一時シート(tmp2) (書式用)
  var tmp2Name = sourceTab + '_format_tmp';
  var oldTmp2 = ss.getSheetByName(tmp2Name);
  if (oldTmp2) ss.deleteSheet(oldTmp2);
  var tmp2 = srcSheet.copyTo(ss);
  tmp2.setName(tmp2Name);

  // 3. 翻訳式入力およびflatten検査（一時シート tmp）
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
    Utilities.sleep(3000);
    var vals = newHelper.getRange(rangeStr).getValues();
    isLoading = hasLoadingOrErrorStrings(vals);
    wait += 3;
  }

  if (!isLoading) {
    // 4. 値をコピー (tmp → aタブ)
    var values = newHelper.getRange(rangeStr).getValues();
    var targetRange = copied.getRange(rangeStr);
    targetRange.clearDataValidations();

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



    // 5. 書式をコピー (tmp2 → aタブ)
    var formatRange = tmp2.getRange(rangeStr);
    var tgtRange = copied.getRange(rangeStr);
    formatRange.copyTo(tgtRange, {formatOnly: true});
  }

  // 6. 一時シート(tmp, tmp2)を削除
  ss.deleteSheet(newHelper);
  ss.deleteSheet(tmp2);
}



// ====== C/D列自動出力関数 ======

function updateSheetList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  var sourceUrl = mgmtSheet.getRange('R12').getValue();
  var match = sourceUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    SpreadsheetApp.getUi().alert('有効なGoogleスプレッドシートのリンクを入力してください。');
    return;
  }
  var sourceId = match[1];
  var sourceSs = SpreadsheetApp.openById(sourceId);

  var sheets = sourceSs.getSheets();
  var cList = [];
  var dList = [];

  var srcCode = mgmtSheet.getRange('R11').getValue();
  var tgtCode = mgmtSheet.getRange('S11').getValue();
  var srcLang = getLangCode(srcCode, true);
  var tgtLang = getLangCode(tgtCode, true);

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (!sheet.isSheetHidden()) {
      var name = sheet.getName();
      cList.push([name]);
      if (srcLang && tgtLang && name.trim() !== '') {
        dList.push(['=GOOGLETRANSLATE("' + name + '","' + srcLang + '","' + tgtLang + '")']);
      } else {
        dList.push([name]);
      }
    }
  }

  if (cList.length > 0) {
    mgmtSheet.getRange('C6:C' + mgmtSheet.getMaxRows()).clearContent();
    mgmtSheet.getRange('D6:D' + mgmtSheet.getMaxRows()).clearContent();
    mgmtSheet.getRange(6, 3, cList.length, 1).setValues(cList);
    mgmtSheet.getRange(6, 4, dList.length, 1).setFormulas(dList);
    
    // ✅ 式計算待機後、重複処理
    Utilities.sleep(3000); // 翻訳式の計算待機
    removeDuplicatesInDColumn(cList.length);
  }
}

// 重複削除関数
function removeDuplicatesInDColumn(rowCount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  
  // D列の翻訳結果を取得
  var dRange = mgmtSheet.getRange(6, 4, rowCount, 1);
  var dValues = dRange.getValues();
  
  var nameCount = {};
  var newValues = [];
  
  for (var i = 0; i < dValues.length; i++) {
    var translatedName = dValues[i][0].toString().trim();
    
    // 重複チェックと番号付け
    if (nameCount[translatedName]) {
      nameCount[translatedName]++;
      translatedName = translatedName + '_(' + nameCount[translatedName] + ')';
    } else {
      nameCount[translatedName] = 1;
    }
    
    newValues.push([translatedName]);
  }
  
  // 重複削除された値でD列を更新
  dRange.setValues(newValues);
}



// ====== トリガー管理関数 ======

// processNextTabトリガーのみを削除する関数
function deleteProcessNextTabTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var deletedCount = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processNextTab') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  
  if (deletedCount > 0) {
    console.log('以前のprocessNextTabトリガー ' + deletedCount + '個を削除しました。');
  }
}

// 現在の関数のトリガーを削除
function deleteCurrentTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processNextTab') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// 全ての自分のトリガーを削除（完了時）
function deleteAllMyTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var funcName = triggers[i].getHandlerFunction();
    if (funcName === 'processNextTab') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// 既存のトリガーを削除 (sortSheetsByDColumnOrderから呼び出し)
function deleteMyScriptTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ====== ユーティリティ関数 ======

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

// 列番号→アルファベット変換
function colToLetter(col) {
  var temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// 現在時刻（日本時間）
function nowString() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
}

// ステータスメッセージ出力（管理シート R16）
function printStatus(msg) {
  var mgmtSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('管理');
  mgmtSheet.getRange('R16').setValue(msg);
}

// エラーメッセージ出力（管理シート R16）
function printError(msg) {
  var mgmtSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('管理');
  mgmtSheet.getRange('R16').setValue('エラー : ' + msg);
}

// #REF!のみ検出
function isErrorCell(val) {
  return (typeof val === 'string') && /^#REF!?/.test(val);
}

// flatten + 40000文字制限 + 多言語エラーチェック
function hasLoadingOrErrorStrings(values) {
  var flat = [].concat.apply([], values); // flatten
  var combined = flat.join('');
  combined = combined.slice(0, 40000); // 40000文字制限
  var errorStrings = [
    'ロード中...', '#ERROR!', 'Loading...', '読み込み中...', '正在加载...', '正在加載...', '読み込み中', 'Chargement en cours...', 'Cargando...'
  ]; // '로드중...' was changed to 'ロード中...'
  for (var i = 0; i < errorStrings.length; i++) {
    if (combined.indexOf(errorStrings[i]) !== -1) {
      return true;
    }
  }
  return false;
}

// D列で最後のデータがある行番号を安全に検索する関数
function getLastDataRowInColumn(sheet, col, startRow) {
  var values = sheet.getRange(col + startRow + ':' + col + sheet.getMaxRows()).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      return startRow + i;
    }
  }
  return startRow - 1;
}

// ====== シート並び替え関数 ======

// D列順に右側にソート、管理シートをアクティブにし、完了メッセージ・トリガー削除
// D列の順序で並び替え、D列にないシートは削除、管理シートを先頭に
function sortSheetsByDColumnOrder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  try {
    // D6:D105からシート名抽出（空の値を除去）
    var dNames = mgmtSheet.getRange('D6:D105').getValues()
      .map(function(row){ return row[0]; })
      .filter(function(name){ return !!name; });

    var sheets = ss.getSheets();
    var sheetNames = sheets.map(function(sheet){ return sheet.getName(); });

    // 管理シートを最も左（1番目）に固定
    if (ss.getSheets()[0].getName() !== '管理') {
      ss.setActiveSheet(mgmtSheet);
      ss.moveActiveSheet(1);
    }

    // D列の順序で管理シートの右側から順に移動（昇順）
    var idx = 2;
    for (var i = 0; i < dNames.length; i++) {
      var name = dNames[i];
      if (name !== '管理' && sheetNames.indexOf(name) !== -1) {
        var sheet = ss.getSheetByName(name);
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(idx);
        idx++;
      }
    }

    // ✅ D列にないシートは削除
    sheets = ss.getSheets(); // シート順序更新
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name !== '管理' && dNames.indexOf(name) === -1) {
        ss.deleteSheet(sheets[i]);
      }
    }

    // 最後に「管理」シートをアクティブにする
    ss.setActiveSheet(mgmtSheet);

    // ★ここでのみ「完了」メッセージ出力とトリガー削除
    printStatus(nowString() + ' : 完了');
    deleteMyScriptTriggers_();

  } catch (e) {
    printError('[右端から並び替え] ' + e); // This string was already in Japanese.
    deleteMyScriptTriggers_();
  }
}


// ====== ユーザーインターフェース関数 ======

// ボタン用：確認ダイアログ付き自動実行
function runAllStepsWithBlockParallelWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '全体の翻訳を開始しますか？（10タブで約12分かかります）',
    ui.ButtonSet.OK_CANCEL
  );
  if (result == ui.Button.OK) {
    startSequentialProcessing();
  }
}
