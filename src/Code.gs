/**
 * Amazon販売KPI管理ツール - メインファイル
 * 
 * このファイルはGoogle Apps Scriptプロジェクトのエントリーポイントです。
 * 各モジュールの初期化とメインメニューの制御を行います。
 */

// =============================================================================
// グローバル設定
// =============================================================================

/**
 * アプリケーション設定
 */
const APP_CONFIG = {
  name: 'Amazon販売KPI管理ツール',
  version: '1.0.0',
  author: 'KPI管理システム',
  description: 'Amazon販売データの統合管理とKPI可視化ツール'
};

/**
 * シート設定
 */
const SHEET_CONFIG = {
  KPI_MONTHLY: 'KPI月次管理',
  KPI_HISTORY: 'KPI履歴',
  SALES_HISTORY: '販売履歴',
  PURCHASE_HISTORY: '仕入履歴',
  INVENTORY: '在庫一覧',
  PRODUCT_MASTER: 'ASIN/SKUマスタ',
  SYNC_LOG: 'データ連携ログ',
  CONFIG: '設定',
  TEMP_DATA: '_一時データ'
};

// =============================================================================
// メインメニュー機能
// =============================================================================

/**
 * スプレッドシート開時のイベントハンドラー
 * カスタムメニューを作成します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu(APP_CONFIG.name)
      .addItem('📊 データ更新', 'manualDataSync')
      .addItem('🔄 KPI再計算', 'recalculateKPIs')
      .addSeparator()
      .addSubMenu(ui.createMenu('📥 インポート')
        .addItem('マカドCSV取り込み', 'importMakadoCSV')
        .addItem('過去データ取り込み', 'importHistoricalData')
        .addItem('SKUマッピング更新', 'updateSKUMapping'))
      .addSeparator()
      .addSubMenu(ui.createMenu('📈 レポート')
        .addItem('日次レポート生成', 'generateDailyReport')
        .addItem('週次レポート生成', 'generateWeeklyReport')
        .addItem('月次レポート生成', 'generateMonthlyReport')
        .addItem('過去実績表示', 'showHistoricalKPIs'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🔧 管理')
        .addItem('初期セットアップ', 'initialSetup')
        .addItem('設定', 'showSettingsDialog')
        .addItem('データクリーンアップ', 'cleanupData')
        .addItem('エラーログ確認', 'showErrorLog'))
      .addSeparator()
      .addItem('❓ ヘルプ', 'showHelp')
      .addToUi();
      
    // 初回起動時のチェック
    checkInitialSetup();
    
  } catch (error) {
    console.error('メニュー作成エラー:', error);
    ui.alert('エラー', 'メニューの作成中にエラーが発生しました。', ui.ButtonSet.OK);
  }
}

/**
 * インストール時のイベントハンドラー
 */
function onInstall() {
  onOpen();
}

// =============================================================================
// メインメニュー機能の実装
// =============================================================================

/**
 * 手動データ同期
 */
function manualDataSync() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'データ更新確認',
    'データを更新します。処理には数分かかる場合があります。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      // プログレス表示（簡易版）
      ui.alert('処理中', 'データ更新中です。しばらくお待ちください...', ui.ButtonSet.OK);
      
      // メインバッチ処理実行
      const processor = new BatchProcessor();
      const result = processor.runDailyBatch();
      
      // resultが正しく返されているか確認
      if (!result || typeof result !== 'object') {
        throw new Error('バッチ処理の結果が正しく返されませんでした');
      }
      
      // 更新件数を集計
      const totalRecords = (result.updateResults?.amazon?.recordCount || 0) + 
                         (result.updateResults?.makado?.recordCount || 0) + 
                         (result.updateResults?.inventory?.recordCount || 0);
      
      // 処理時間を安全に取得
      const duration = result.duration ? result.duration.toFixed(1) : '不明';
      
      // 結果表示
      ui.alert(
        '完了',
        `データ更新が完了しました。\n処理時間: ${duration}秒\n更新件数: ${totalRecords}件`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ErrorHandler.handleError(error, 'manualDataSync');
      ui.alert(
        'エラー',
        `データ更新中にエラーが発生しました:\n${error.message}`,
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * KPI再計算
 */
function recalculateKPIs() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('処理中', 'KPIを再計算しています。しばらくお待ちください...', ui.ButtonSet.OK);
    
    const calculator = new KPICalculator();
    calculator.recalculateAll();
    
    ui.alert('完了', 'KPIの再計算が完了しました。', ui.ButtonSet.OK);
    
  } catch (error) {
    ErrorHandler.handleError(error, 'recalculateKPIs');
    ui.alert('エラー', `KPI計算中にエラーが発生しました:\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * マカドCSV取り込み
 */
function importMakadoCSV() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'CSVファイル取り込み',
    'マカドからエクスポートしたCSVファイルを取り込みます。\n' +
    'ファイルをGoogleドライブにアップロードしてから、\n' +
    'ファイル名を入力してください。',
    ui.ButtonSet.OK
  );
  
  const response = ui.prompt(
    'ファイル名入力',
    'CSVファイル名を入力してください（例：20250731_2025_マカド_販売履歴.csv）',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const fileName = response.getResponseText();
    
    try {
      // 最適化版を使用
      const processor = new MakadoProcessorOptimized();
      const result = processor.processCSVFile(fileName);
      
      ui.alert(
        '完了',
        `マカドCSVの取り込みが完了しました。\n取り込み件数: ${result.recordCount}件\n処理時間: ${result.duration}秒`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ErrorHandler.handleError(error, 'importMakadoCSV');
      ui.alert('エラー', `CSV取り込み中にエラーが発生しました:\n${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * 初期セットアップ
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '初期セットアップ',
    'KPI管理ツールの初期セットアップを実行します。\n' +
    '・必要なシートの作成\n' +
    '・基本設定の初期化\n' +
    '・サンプルデータの投入\n\n' +
    '続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      const setupManager = new SetupManager();
      setupManager.runInitialSetup();
      
      ui.alert(
        '完了',
        '初期セットアップが完了しました。\n' +
        '設定シートでAPI認証情報を設定してください。',
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      ErrorHandler.handleError(error, 'initialSetup');
      ui.alert('エラー', `セットアップ中にエラーが発生しました:\n${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * ヘルプ表示
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  
  const helpText = `
${APP_CONFIG.name} v${APP_CONFIG.version}

【主要機能】
📊 データ更新: Amazon SP-APIとマカドデータを同期
🔄 KPI再計算: 全KPIを最新データで再計算
📥 インポート: CSV/過去データの取り込み
📈 レポート: 各種レポートの自動生成

【初回利用時】
1. 「管理」→「初期セットアップ」を実行
2. 「設定」シートでAPI認証情報を設定
3. 「データ更新」でデータ同期を開始

【サポート】
GitHub: https://github.com/taktok21/KPI-management-GoogleSpredSheet
問題が発生した場合は「管理」→「エラーログ確認」をご確認ください。
`;
  
  ui.alert('ヘルプ', helpText, ui.ButtonSet.OK);
}

// =============================================================================
// ユーティリティ関数
// =============================================================================

/**
 * プログレス表示ダイアログ
 */
function showProgressDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  
  // 注意: GASではモーダルダイアログは非同期で表示されるため、
  // 実際のプログレス表示は制限されます
  ui.alert(title, message + '\n\n処理中...', ui.ButtonSet.OK);
}

/**
 * 初期セットアップ状態をチェック
 */
function checkInitialSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_CONFIG.CONFIG);
    
    if (!configSheet) {
      // 初期セットアップが必要
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '初期セットアップが必要です',
        'KPI管理ツールを使用するには初期セットアップが必要です。\n今すぐ実行しますか？',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        initialSetup();
      }
    } else {
      // 必要なシートが存在するかチェック
      checkRequiredSheets();
    }
  } catch (error) {
    console.error('初期セットアップチェックエラー:', error);
  }
}

/**
 * 必要なシートの存在確認
 */
function checkRequiredSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = [
      SHEET_CONFIG.SALES_HISTORY,
      SHEET_CONFIG.PRODUCT_MASTER,
      SHEET_CONFIG.KPI_MONTHLY,
      SHEET_CONFIG.INVENTORY
    ];
    
    const missingSheets = [];
    requiredSheets.forEach(sheetName => {
      if (!ss.getSheetByName(sheetName)) {
        missingSheets.push(sheetName);
      }
    });
    
    if (missingSheets.length > 0) {
      console.log('不足しているシート:', missingSheets);
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '不足シートの検出',
        `以下のシートが不足しています：\n${missingSheets.join('\n')}\n\n不足しているシートを作成しますか？`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        createMissingSheets(missingSheets);
      }
    }
  } catch (error) {
    console.error('シート確認エラー:', error);
  }
}

/**
 * 不足しているシートを作成
 */
function createMissingSheets(missingSheets) {
  try {
    const setupManager = new SetupManager();
    
    missingSheets.forEach(sheetName => {
      console.log(`シート作成中: ${sheetName}`);
      switch (sheetName) {
        case SHEET_CONFIG.SALES_HISTORY:
          setupManager.createSalesHistorySheet();
          break;
        case SHEET_CONFIG.PRODUCT_MASTER:
          setupManager.createProductMasterSheet();
          break;
        case SHEET_CONFIG.KPI_MONTHLY:
          setupManager.createKPIMonthlySheet();
          break;
        case SHEET_CONFIG.INVENTORY:
          setupManager.createInventorySheet();
          break;
      }
    });
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('完了', `${missingSheets.length}個のシートを作成しました。`, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('シート作成エラー:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', `シート作成中にエラーが発生しました：\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 過去実績表示
 */
function showHistoricalKPIs() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const historyManager = new KPIHistoryManager();
    const historicalKPIs = historyManager.getHistoricalKPIs(12);
    
    if (historicalKPIs.length === 0) {
      ui.alert('データなし', '過去のKPIデータがありません。', ui.ButtonSet.OK);
      return;
    }
    
    // 過去実績レポートを生成
    let report = '📊 過去12ヶ月のKPI実績\n\n';
    
    historicalKPIs.forEach((kpi, index) => {
      if (kpi.hasData) {
        report += `【${kpi.yearMonth}】\n`;
        report += `売上高: ¥${NumberUtils.formatNumber(kpi.revenue)}\n`;
        report += `粗利益: ¥${NumberUtils.formatNumber(kpi.grossProfit)} (${kpi.profitMargin.toFixed(1)}%)\n`;
        report += `ROI: ${kpi.roi.toFixed(1)}%\n`;
        report += `達成率: ${kpi.profitGoalAchievement.toFixed(1)}%\n`;
        report += '\n';
      }
    });
    
    // 簡易レポート表示（将来的にはグラフ表示も検討）
    ui.alert('過去実績', report, ui.ButtonSet.OK);
    
  } catch (error) {
    ErrorHandler.handleError(error, 'showHistoricalKPIs');
    ui.alert('エラー', '過去実績の表示中にエラーが発生しました。', ui.ButtonSet.OK);
  }
}

// =============================================================================
// デバッグ用関数
// =============================================================================

/**
 * デバッグ情報表示（開発時のみ使用）
 */
function debugInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log('スプレッドシートID:', ss.getId());
  console.log('スプレッドシート名:', ss.getName());
  console.log('シート一覧:', ss.getSheets().map(sheet => sheet.getName()));
  
  // プロパティ確認
  const props = PropertiesService.getScriptProperties().getProperties();
  console.log('設定されているプロパティ:', Object.keys(props));
}

/**
 * テスト実行（開発時のみ使用）
 */
function runTests() {
  try {
    // テストランナー実行
    const testResults = TestRunner.runAllTests();
    console.log('テスト結果:', testResults);
    
    const ui = SpreadsheetApp.getUi();
    const passedTests = testResults.filter(r => r.status === 'PASS').length;
    const totalTests = testResults.length;
    
    ui.alert(
      'テスト結果',
      `実行: ${totalTests}件\n成功: ${passedTests}件\n失敗: ${totalTests - passedTests}件`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('テスト実行エラー:', error);
  }
}