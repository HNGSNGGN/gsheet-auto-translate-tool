// 管理シートR15の0～23（時間）に合わせて毎日自動実行トリガーを登録
function installDailyAutoRunTriggerByR15() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');
  var hour = mgmtSheet.getRange('R15').getValue();

  if (typeof hour !== "number" || isNaN(hour) || hour < 0 || hour > 23) {
    SpreadsheetApp.getActiveSpreadsheet().toast("R15セルに0～23のいずれかの時間を入力してください。", "エラー", 5);
    return;
  }

  // 既存のトリガー削除（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoRun') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // トリガー登録（毎日R15時～R15+1時の間にランダム実行）
  ScriptApp.newTrigger('dailyAutoRun')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("毎日 " + hour + "時～" + (hour+1) + "時の間に自動実行トリガーが登録されました。", "通知", 5);
}

// 自動実行トリガー解除
function removeDailyAutoRunTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = false;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyAutoRun') {
      ScriptApp.deleteTrigger(t);
      removed = true;
    }
  });
  if (removed) {
    SpreadsheetApp.getActiveSpreadsheet().toast("自動実行トリガーが解除されました。", "通知", 5);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("解除する自動実行トリガーがありません。", "通知", 5);
  }
}

function dailyAutoRun() {
  startSequentialProcessingAuto();
}


function startSequentialProcessingAuto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgmtSheet = ss.getSheetByName('管理');

  // 既存のトリガーを削除
  deleteProcessNextTabTriggers();

  // C/D列を自動出力
  updateSheetList();

  // 少し待機（D列の翻訳関数の計算時間）
  Utilities.sleep(3000);

  // 初期化
  var prop = PropertiesService.getScriptProperties();
  prop.setProperty('current_index', '0');
  prop.setProperty('total_tabs', '0');

  // D列から処理対象のタブ一覧を取得（全体）
  var lastRow = getLastDataRowInColumn(mgmtSheet, 'D', 6);
  if (lastRow < 6) {
    // トリガー環境ではUI使用禁止、そのままreturn
    return;
  }
  var tabPairs = mgmtSheet.getRange('C6:D' + lastRow).getValues();

  // S21は無視、最初から全体を保存
  for (var i = 0; i < tabPairs.length; i++) {
    prop.setProperty('tab_' + i + '_source', tabPairs[i][0]);
    prop.setProperty('tab_' + i + '_target', tabPairs[i][1]);
  }
  prop.setProperty('total_tabs', tabPairs.length.toString());

  // 最初のタブの処理をすぐに開始
  processNextTab();
}
