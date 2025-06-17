function startSequentialProcessingWithTimeLimit() {
  var startTime = new Date().getTime(); // 開始時間を記録
  var maxTime = 300 * 1000; // 300秒 = 300,000ms
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');

  // C/D列自動出力
  updateSheetList();
  Utilities.sleep(3000);

  // S21から開始タブ名を取得
  var startTab = mgmtSheet.getRange('S21').getValue();
  var ui = SpreadsheetApp.getUi();
  var msg = '';

  if (startTab && startTab.trim() !== '') {
    msg = '「' + startTab + '」から5分制限翻訳を開始しますか？\n\n(最初から実行したい場合はH23を削除してください)';
  } else {
    msg = '全てのタブを5分制限で翻訳しますか？\n\n(特定のタブから実行するにはH23でタブ名を選択してください)';
  }

  var result = ui.alert(msg, ui.ButtonSet.OK_CANCEL);
  if (result !== ui.Button.OK) {
    ui.alert('実行がキャンセルされました。');
    return;
  }

  // D列から処理するタブリストを取得
  var lastRow = getLastDataRowInColumn(mgmtSheet, 'D', 6);
  if (lastRow < 6) {
    ui.alert('管理シートD列にデータがありません。'); // This was already Japanese
    return;
  }
  var tabPairs = mgmtSheet.getRange('C6:D' + lastRow).getValues();

  // S21が空でない場合、D列で開始タブの位置を検索
  var startIndex = 0;
  if (startTab && startTab.trim() !== '') {
    for (var i = 0; i < tabPairs.length; i++) {
      if (tabPairs[i][1] === startTab) {
        startIndex = i;
        break;
      }
    }
  }

  var actualTabs = tabPairs.slice(startIndex);
  var completedTabs = [];
  var currentProcessingTab = null;

  // 各タブを順次処理
  for (var i = 0; i < actualTabs.length; i++) {
    var currentTime = new Date().getTime();
    var elapsedTime = currentTime - startTime;
    
    // ✅ 残り時間チェック (60秒の余裕を確保)
    if (elapsedTime > (maxTime - 60000)) {
      // 時間不足 - 安全に終了
      cleanupCurrentProcessing(currentProcessingTab);
      showPartialCompletionMessage(completedTabs, actualTabs[i][1]);
      return;
    }

    var sourceTab = actualTabs[i][0];
    var targetTab = actualTabs[i][1];
    currentProcessingTab = sourceTab;

    if (!sourceTab || !targetTab) continue;

    try {
      // 進行状況を表示
      printStatus(nowString() + ' : ' + targetTab + 'タブ翻訳中...（5分制限版）');
      
      // タブ処理
      processOneTabQuick(sourceTab, targetTab, 'R11', 'S11');
      completedTabs.push(targetTab);
      
      var progress = Math.round((i + 1) / actualTabs.length * 100);
      console.log(progress + '% 完了: ' + targetTab + ' 処理完了');
      
    } catch (e) {
      console.log('エラー発生: ' + targetTab + ' - ' + e);
      // エラーが発生しても続行
    }
  }

  // ✅ 全てのタブ完了
  printStatus(nowString() + ' : 完了（5分制限版）');

  // ★ 追加が必要: シートの並び替えと整理
  sortSheetsByDColumnOrder();

  ui.alert('全てのタブの翻訳が完了しました！');
}

// 現在処理中の一時シートを整理
function cleanupCurrentProcessing(sourceTab) {
  if (!sourceTab) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    var helperTab = sourceTab + '_tmp';
    var tmp2Name = sourceTab + '_format_tmp';
    
    var helper = ss.getSheetByName(helperTab);
    if (helper) ss.deleteSheet(helper);
    
    var tmp2 = ss.getSheetByName(tmp2Name);
    if (tmp2) ss.deleteSheet(tmp2);
    
    console.log('一時シートの整理完了: ' + sourceTab);
  } catch (e) {
    console.log('一時シートの整理中にエラー: ' + e);
  }
}

// 部分完了メッセージを表示
function showPartialCompletionMessage(completedTabs, nextTab) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  var ui = SpreadsheetApp.getUi();
  
  var lastCompletedTab = completedTabs.length > 0 ? completedTabs[completedTabs.length - 1] : 'なし';
  
  // S21に次の開始タブを設定
  if (nextTab) {
    mgmtSheet.getRange('H23').setValue(nextTab); // S21 -> H23 in comment
  }

    // ★ 追加が必要: 部分完了でも現在までに翻訳されたタブを並び替え
  sortSheetsByDColumnOrderPartial();
  
  var message = '時間制限のため、部分的に完了しました。\n\n';
  message += '完了したタブ数: ' + completedTabs.length + '個\n';
  message += '最後に完了したタブ: ' + lastCompletedTab + '\n\n';
  if (nextTab) {
    message += '次回実行時、「' + nextTab + '」から続行されます。\n';
    message += '(H23に自動的に設定されました)';
  }
  
  printStatus(nowString() + ' : ' + lastCompletedTab + 'タブまで処理完了');
  ui.alert(message);
}

// 高速処理用 processOneTab (既存と同じ、名前のみ変更)
function processOneTabQuick(sourceTab, targetTab, srcCell, tgtCell) {
  // 既存の processOneTab と同じロジック
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  
  var sourceUrl = mgmtSheet.getRange('R12').getValue();
  var match = sourceUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) throw new Error('有効なGoogleスプレッドシートのリンクを入力してください。'); // This was already Japanese
  var sourceId = match[1];
  var sourceSs = SpreadsheetApp.openById(sourceId);

  var srcSheet = sourceSs.getSheetByName(sourceTab);
  if (!srcSheet) throw new Error(sourceTab + 'が元のシートに存在しません。');

  // 既存のターゲットシートを削除後コピー
  var oldTarget = ss.getSheetByName(targetTab);
  if (oldTarget) ss.deleteSheet(oldTarget);
  var copied = srcSheet.copyTo(ss);
  copied.setName(targetTab);

  // 元のシートをコピー → 一時シート(tmp2) (書式用)
  var tmp2Name = sourceTab + '_format_tmp';
  var oldTmp2 = ss.getSheetByName(tmp2Name);
  if (oldTmp2) ss.deleteSheet(oldTmp2);
  var tmp2 = srcSheet.copyTo(ss);
  tmp2.setName(tmp2Name);

  // 翻訳式入力と処理 (既存と同じ)
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

  // flatten検査 (既存と同じ)
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
    // 値をコピー (既存と同じ)
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


    // 書式をコピー
    var formatRange = tmp2.getRange(rangeStr);
    var tgtRange = copied.getRange(rangeStr);
    formatRange.copyTo(tgtRange, {formatOnly: true});
  }

  // 一時シートを削除
  ss.deleteSheet(newHelper);
  ss.deleteSheet(tmp2);
}

// 部分完了用並び替え関数 (削除なしで並び替えのみ)
function sortSheetsByDColumnOrderPartial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  try {
    // D列の順序で並び替えのみ (削除はしない)
    var dNames = mgmtSheet.getRange('D6:D105').getValues()
      .map(function(row){ return row[0]; })
      .filter(function(name){ return !!name; });

    var sheets = ss.getSheets();
    var sheetNames = sheets.map(function(sheet){ return sheet.getName(); });

    // 管理シートを最も左に固定
    if (ss.getSheets()[0].getName() !== '管理') {
      ss.setActiveSheet(mgmtSheet);
      ss.moveActiveSheet(1);
    }

    // D列の順序で並び替え
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

    ss.setActiveSheet(mgmtSheet);
  } catch (e) {
    console.log('[部分並び替えエラー] ' + e);
  }
}
