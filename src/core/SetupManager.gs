/**
 * セットアップ管理クラス
 * 
 * 初期セットアップ、シート作成、データ構造の初期化を行います。
 */

class SetupManager {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  // =============================================================================
  // 初期セットアップ
  // =============================================================================

  /**
   * 初期セットアップ実行
   */
  runInitialSetup() {
    try {
      const setupSteps = [
        'スプレッドシート設定',
        '基本シート作成',
        'データ構造初期化',
        '設定シート作成',
        'サンプルデータ投入',
        '初期KPI計算'
      ];

      let currentStep = 0;

      // 1. スプレッドシート基本設定
      this.setupSpreadsheetProperties();
      currentStep++;

      // 2. 必要なシートを作成
      this.createAllSheets();
      currentStep++;

      // 3. データ構造を初期化
      this.initializeDataStructure();
      currentStep++;

      // 4. 設定シートを作成
      this.createConfigSheet();
      currentStep++;

      // 5. サンプルデータ投入（オプション）
      this.insertSampleData();
      currentStep++;

      // 6. 初期KPI計算
      this.runInitialKPICalculation();
      currentStep++;

      // セットアップ完了ログ
      this.logSetupCompletion();

      return {
        success: true,
        completedSteps: currentStep,
        totalSteps: setupSteps.length,
        message: '初期セットアップが完了しました',
        setupDate: new Date()
      };

    } catch (error) {
      ErrorHandler.handleError(error, 'SetupManager.runInitialSetup');
      throw error;
    }
  }

  // =============================================================================
  // スプレッドシート設定
  // =============================================================================

  /**
   * スプレッドシート基本設定
   */
  setupSpreadsheetProperties() {
    try {
      // スプレッドシート名設定
      this.spreadsheet.rename('Amazon販売KPI管理ツール');

      // タイムゾーン設定
      this.spreadsheet.setSpreadsheetTimeZone('Asia/Tokyo');

      // ロケール設定
      this.spreadsheet.setSpreadsheetLocale('ja_JP');

      // 繰り返し計算設定
      this.spreadsheet.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);

    } catch (error) {
      console.warn('スプレッドシート設定でエラー:', error);
    }
  }

  // =============================================================================
  // シート作成
  // =============================================================================

  /**
   * 全シート作成
   */
  createAllSheets() {
    const sheetsToCreate = [
      { name: SHEET_CONFIG.KPI_MONTHLY, method: 'createKPIMonthlySheet' },
      { name: SHEET_CONFIG.SALES_HISTORY, method: 'createSalesHistorySheet' },
      { name: SHEET_CONFIG.PURCHASE_HISTORY, method: 'createPurchaseHistorySheet' },
      { name: SHEET_CONFIG.INVENTORY, method: 'createInventorySheet' },
      { name: SHEET_CONFIG.PRODUCT_MASTER, method: 'createProductMasterSheet' },
      { name: SHEET_CONFIG.SYNC_LOG, method: 'createSyncLogSheet' },
      { name: SHEET_CONFIG.CONFIG, method: 'createConfigSheet' }
    ];

    sheetsToCreate.forEach(sheetInfo => {
      if (!this.spreadsheet.getSheetByName(sheetInfo.name)) {
        this[sheetInfo.method]();
      }
    });

    // デフォルトシートの削除
    this.removeDefaultSheets();
  }

  /**
   * KPI月次管理シート作成
   */
  createKPIMonthlySheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.KPI_MONTHLY);
    
    // シートの保護設定
    sheet.protect().setDescription('KPIダッシュボード（編集禁止）');

    // レイアウト設定
    this.setupKPIMonthlyLayout(sheet);
    
    return sheet;
  }

  /**
   * 販売履歴シート作成
   */
  createSalesHistorySheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.SALES_HISTORY);
    
    // ヘッダー行設定
    const headers = [
      '注文ID', 'ASIN', '注文日時', '統一SKU', 'マカドSKU', '商品名',
      '販売価格', '数量', '合計金額', '仕入原価', 'Amazon手数料', 'その他費用',
      '粗利益', '利益率', 'ステータス', '配送方法', 'データソース', '取込日時'
    ];
    
    this.setupSheetHeaders(sheet, headers);
    this.setupSalesHistoryFormatting(sheet);
    
    return sheet;
  }

  /**
   * 仕入履歴シート作成
   */
  createPurchaseHistorySheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.PURCHASE_HISTORY);
    
    const headers = [
      '統一SKU', 'ASIN', '仕入日', '仕入先', '数量', '単価',
      '合計金額', '送料', '備考', '登録日時'
    ];
    
    this.setupSheetHeaders(sheet, headers);
    this.setupPurchaseHistoryFormatting(sheet);
    
    return sheet;
  }

  /**
   * 在庫一覧シート作成
   */
  createInventorySheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.INVENTORY);
    
    const headers = [
      '統一SKU', 'ASIN', '商品名', '在庫数', '単位原価', '在庫金額',
      '保管場所', '最終入荷日', '最終販売日', '在庫日数', '回転日数',
      'アラート', '更新日時'
    ];
    
    this.setupSheetHeaders(sheet, headers);
    this.setupInventoryFormatting(sheet);
    
    return sheet;
  }

  /**
   * 商品マスターシート作成
   */
  createProductMasterSheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.PRODUCT_MASTER);
    
    const headers = [
      '統一SKU', 'ASIN', 'マカドSKU', 'AmazonSKU', '商品名',
      'カテゴリ', 'ブランド', 'JANコード', '仕入日', '作成日', '更新日'
    ];
    
    this.setupSheetHeaders(sheet, headers);
    this.setupProductMasterFormatting(sheet);
    
    return sheet;
  }

  /**
   * データ連携ログシート作成
   */
  createSyncLogSheet() {
    const sheet = this.spreadsheet.insertSheet(SHEET_CONFIG.SYNC_LOG);
    
    const headers = [
      '実行日時', '処理タイプ', 'ステータス', '処理時間', '取得件数',
      'エラー件数', 'メッセージ', '詳細'
    ];
    
    this.setupSheetHeaders(sheet, headers);
    this.setupSyncLogFormatting(sheet);
    
    return sheet;
  }

  /**
   * 設定シート作成
   */
  createConfigSheet() {
    const sheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.CONFIG) || 
                  this.spreadsheet.insertSheet(SHEET_CONFIG.CONFIG);
    
    this.setupConfigLayout(sheet);
    
    return sheet;
  }

  // =============================================================================
  // シートレイアウト設定
  // =============================================================================

  /**
   * 共通ヘッダー設定
   */
  setupSheetHeaders(sheet, headers) {
    // ヘッダー行設定
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダーフォーマット
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e3f2fd');
    headerRange.setHorizontalAlignment('center');
    
    // 行固定
    sheet.setFrozenRows(1);
    
    // 列幅自動調整
    sheet.autoResizeColumns(1, headers.length);
  }

  /**
   * KPI月次管理レイアウト設定
   */
  setupKPIMonthlyLayout(sheet) {
    // タイトル
    sheet.getRange('A1').setValue('Amazon販売KPI管理ダッシュボード');
    sheet.getRange('A1').setFontSize(18).setFontWeight('bold');
    sheet.getRange('A1:E1').merge();

    // 最終更新時刻
    sheet.getRange('F1').setValue('最終更新: ');
    sheet.getRange('G1').setFormula('=NOW()');
    sheet.getRange('G1').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // 月次KPIセクション
    sheet.getRange('A3').setValue('📊 月次KPI');
    sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
    sheet.getRange('A3').setBackground('#c8e6c9');

    const monthlyKPIs = [
      ['項目', '実績', '目標', '達成率', '前月比'],
      ['売上高', '', '3,200,000', '', ''],
      ['粗利益', '', '800,000', '', ''],
      ['利益率', '', '25%', '', ''],
      ['ROI', '', '30%', '', ''],
      ['販売数', '', '600', '', ''],
      ['在庫金額', '', '1,000,000', '', ''],
      ['在庫回転率', '', '1.0', '', ''],
      ['滞留在庫率', '', '10%', '', '']
    ];

    sheet.getRange(4, 1, monthlyKPIs.length, monthlyKPIs[0].length).setValues(monthlyKPIs);
    
    // 日次KPIセクション
    sheet.getRange('A14').setValue('📈 本日の実績');
    sheet.getRange('A14').setFontSize(14).setFontWeight('bold');
    sheet.getRange('A14').setBackground('#fff3e0');

    const dailyKPIs = [
      ['項目', '実績', '7日平均', '成長率'],
      ['本日売上', '', '', ''],
      ['本日利益', '', '', ''],
      ['本日販売数', '', '', ''],
      ['本日注文数', '', '', '']
    ];

    sheet.getRange(15, 1, dailyKPIs.length, dailyKPIs[0].length).setValues(dailyKPIs);

    // アラートセクション
    sheet.getRange('A21').setValue('⚠️ アラート');
    sheet.getRange('A21').setFontSize(14).setFontWeight('bold');
    sheet.getRange('A21').setBackground('#ffcdd2');

    // フォーマット調整
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 100);
  }

  /**
   * 設定シートレイアウト
   */
  setupConfigLayout(sheet) {
    sheet.clear();
    
    // タイトル
    sheet.getRange('A1').setValue('KPI管理ツール設定');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold');

    // Amazon SP-API設定セクション
    sheet.getRange('A3').setValue('Amazon SP-API設定');
    sheet.getRange('A3').setFontWeight('bold').setBackground('#e1f5fe');

    const apiSettings = [
      ['項目', '値', '説明'],
      ['Client ID', '', 'Amazon SP-API クライアントID'],
      ['Client Secret', '', 'Amazon SP-API クライアントシークレット'],
      ['Refresh Token', '', 'Amazon SP-API リフレッシュトークン'],
      ['Marketplace ID', 'A1VC38T7YXB528', '日本のマーケットプレイスID'],
      ['Endpoint', 'https://sellingpartnerapi-fe.amazon.com', 'API エンドポイント']
    ];

    sheet.getRange(4, 1, apiSettings.length, apiSettings[0].length).setValues(apiSettings);

    // 通知設定セクション
    sheet.getRange('A11').setValue('通知設定');
    sheet.getRange('A11').setFontWeight('bold').setBackground('#fff3e0');

    const notificationSettings = [
      ['項目', '値', '説明'],
      ['通知メールアドレス', '', 'エラー・アラート通知先'],
      ['Slack Webhook URL', '', 'Slack通知用URL（オプション）'],
      ['メール通知有効', 'TRUE', 'TRUE/FALSE'],
      ['Slack通知有効', 'FALSE', 'TRUE/FALSE']
    ];

    sheet.getRange(12, 1, notificationSettings.length, notificationSettings[0].length).setValues(notificationSettings);

    // KPI目標設定セクション
    sheet.getRange('A18').setValue('KPI目標設定');
    sheet.getRange('A18').setFontWeight('bold').setBackground('#c8e6c9');

    const kpiSettings = [
      ['項目', '値', '説明'],
      ['目標月利', '800000', '目標月間粗利益（円）'],
      ['目標利益率', '25', '目標利益率（%）'],
      ['目標ROI', '30', '目標投資利益率（%）'],
      ['最大在庫金額', '1000000', '在庫上限金額（円）'],
      ['滞留在庫日数閾値', '60', '滞留在庫判定日数'],
      ['低在庫アラート閾値', '5', '低在庫アラート数量']
    ];

    sheet.getRange(19, 1, kpiSettings.length, kpiSettings[0].length).setValues(kpiSettings);

    // データ更新設定セクション
    sheet.getRange('A27').setValue('データ更新設定');
    sheet.getRange('A27').setFontWeight('bold').setBackground('#f3e5f5');

    const updateSettings = [
      ['項目', '値', '説明'],
      ['自動更新有効', 'TRUE', '自動データ更新の有効/無効'],
      ['更新間隔（時間）', '24', '自動更新の間隔'],
      ['バッチサイズ', '100', '一度に処理するレコード数'],
      ['リトライ回数', '3', 'エラー時のリトライ回数'],
      ['タイムアウト（秒）', '300', '処理タイムアウト時間']
    ];

    sheet.getRange(28, 1, updateSettings.length, updateSettings[0].length).setValues(updateSettings);

    // 注意事項
    sheet.getRange('A35').setValue('⚠️ 注意事項');
    sheet.getRange('A35').setFontWeight('bold').setFontColor('#d32f2f');

    const notes = [
      '• Amazon SP-APIの認証情報は機密情報です。他人と共有しないでください。',
      '• 設定変更後は「データ更新」を実行して反映してください。',
      '• エラーが発生した場合は「管理」→「エラーログ確認」で詳細を確認してください。',
      '• 初回利用時は「管理」→「初期セットアップ」を実行してください。'
    ];

    notes.forEach((note, index) => {
      sheet.getRange(36 + index, 1).setValue(note);
    });

    // 列幅調整
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 300);

    // B列の値セルに薄い背景色を設定
    const valueRanges = [
      sheet.getRange('B5:B9'),   // API設定
      sheet.getRange('B13:B16'), // 通知設定
      sheet.getRange('B20:B25'), // KPI設定
      sheet.getRange('B29:B33')  // 更新設定
    ];

    valueRanges.forEach(range => {
      range.setBackground('#f5f5f5');
    });
  }

  // =============================================================================
  // フォーマット設定
  // =============================================================================

  /**
   * 販売履歴フォーマット設定
   */
  setupSalesHistoryFormatting(sheet) {
    // 日付列フォーマット
    const dateColumn = sheet.getRange('C:C');
    dateColumn.setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // 金額列フォーマット
    const priceColumns = ['G', 'I', 'J', 'K', 'L', 'M']; // 販売価格、合計金額、仕入原価、手数料、粗利益
    priceColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('¥#,##0');
    });

    // パーセント列フォーマット
    sheet.getRange('N:N').setNumberFormat('0.0%'); // 利益率

    // 条件付きフォーマット（利益率）
    const profitMarginRange = sheet.getRange('N:N');
    const profitRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.2)
      .setBackground('#ffcdd2')
      .setRanges([profitMarginRange])
      .build();
    
    sheet.setConditionalFormatRules([profitRule]);
  }

  /**
   * 仕入履歴フォーマット設定
   */
  setupPurchaseHistoryFormatting(sheet) {
    // 日付列フォーマット
    sheet.getRange('C:C').setNumberFormat('yyyy-mm-dd');
    sheet.getRange('J:J').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // 金額列フォーマット
    const priceColumns = ['F', 'G', 'H']; // 単価、合計金額、送料
    priceColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('¥#,##0');
    });

    // 数量列フォーマット
    sheet.getRange('E:E').setNumberFormat('#,##0');
  }

  /**
   * 在庫フォーマット設定
   */
  setupInventoryFormatting(sheet) {
    // 金額列フォーマット
    const priceColumns = ['E', 'F']; // 単位原価、在庫金額
    priceColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('¥#,##0');
    });

    // 数量列フォーマット
    sheet.getRange('D:D').setNumberFormat('#,##0');

    // 日付列フォーマット
    const dateColumns = ['H', 'I', 'M']; // 最終入荷日、最終販売日、更新日時
    dateColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('yyyy-mm-dd');
    });

    // 日数列フォーマット
    const dayColumns = ['J', 'K']; // 在庫日数、回転日数
    dayColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('#,##0');
    });

    // アラート列の条件付きフォーマット
    const alertRange = sheet.getRange('L:L');
    const alertRules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('滞留')
        .setBackground('#ffcdd2')
        .setFontColor('#d32f2f')
        .setRanges([alertRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('在庫切れ間近')
        .setBackground('#fff3e0')
        .setFontColor('#f57c00')
        .setRanges([alertRange])
        .build()
    ];
    
    sheet.setConditionalFormatRules(alertRules);
  }

  /**
   * 商品マスターフォーマット設定
   */
  setupProductMasterFormatting(sheet) {
    // 日付列フォーマット
    const dateColumns = ['I', 'J', 'K']; // 仕入日、作成日、更新日
    dateColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('yyyy-mm-dd');
    });
  }

  /**
   * データ連携ログフォーマット設定
   */
  setupSyncLogFormatting(sheet) {
    // 日時列フォーマット
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // 処理時間列フォーマット
    sheet.getRange('D:D').setNumberFormat('#,##0.0');

    // 件数列フォーマット
    const countColumns = ['E', 'F']; // 取得件数、エラー件数
    countColumns.forEach(col => {
      sheet.getRange(`${col}:${col}`).setNumberFormat('#,##0');
    });

    // ステータス列の条件付きフォーマット
    const statusRange = sheet.getRange('C:C');
    const statusRules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('SUCCESS')
        .setBackground('#c8e6c9')
        .setFontColor('#2e7d32')
        .setRanges([statusRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('ERROR')
        .setBackground('#ffcdd2')
        .setFontColor('#d32f2f')
        .setRanges([statusRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('WARNING')
        .setBackground('#fff3e0')
        .setFontColor('#f57c00')
        .setRanges([statusRange])
        .build()
    ];
    
    sheet.setConditionalFormatRules(statusRules);
  }

  // =============================================================================
  // データ構造初期化
  // =============================================================================

  /**
   * データ構造初期化
   */
  initializeDataStructure() {
    try {
      // トリガー設定
      this.setupTriggers();

      // 名前付き範囲設定
      this.setupNamedRanges();

      // データ検証設定
      this.setupDataValidation();

    } catch (error) {
      console.warn('データ構造初期化でエラー:', error);
    }
  }

  /**
   * トリガー設定
   */
  setupTriggers() {
    // 既存のトリガーをクリア
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runDailyBatch') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // 日次バッチトリガー設定（午前6時）
    ScriptApp.newTrigger('runDailyBatch')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
  }

  /**
   * 名前付き範囲設定
   */
  setupNamedRanges() {
    const namedRanges = [
      { name: 'SalesData', range: `${SHEET_CONFIG.SALES_HISTORY}!A:R` },
      { name: 'InventoryData', range: `${SHEET_CONFIG.INVENTORY}!A:M` },
      { name: 'ProductMaster', range: `${SHEET_CONFIG.PRODUCT_MASTER}!A:K` },
      { name: 'ConfigData', range: `${SHEET_CONFIG.CONFIG}!A:C` }
    ];

    namedRanges.forEach(rangeInfo => {
      try {
        this.spreadsheet.setNamedRange(rangeInfo.name, rangeInfo.range);
      } catch (error) {
        console.warn(`名前付き範囲設定エラー: ${rangeInfo.name}`, error);
      }
    });
  }

  /**
   * データ検証設定
   */
  setupDataValidation() {
    try {
      // 設定シートの検証ルール
      const configSheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.CONFIG);
      
      if (configSheet) {
        // Boolean値の検証
        const booleanRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['TRUE', 'FALSE'])
          .setAllowInvalid(false)
          .setHelpText('TRUE または FALSE を選択してください')
          .build();

        // Boolean値のセルに適用
        const booleanRanges = ['B15', 'B16', 'B29']; // メール通知、Slack通知、自動更新
        booleanRanges.forEach(cellAddress => {
          configSheet.getRange(cellAddress).setDataValidation(booleanRule);
        });

        // 数値の検証
        const numberRule = SpreadsheetApp.newDataValidation()
          .requireNumberGreaterThan(0)
          .setAllowInvalid(false)
          .setHelpText('0より大きい数値を入力してください')
          .build();

        // 数値のセルに適用
        const numberRanges = ['B20', 'B21', 'B22', 'B23', 'B24', 'B25', 'B30', 'B31', 'B32', 'B33'];
        numberRanges.forEach(cellAddress => {
          configSheet.getRange(cellAddress).setDataValidation(numberRule);
        });
      }

    } catch (error) {
      console.warn('データ検証設定でエラー:', error);
    }
  }

  // =============================================================================
  // サンプルデータ
  // =============================================================================

  /**
   * サンプルデータ投入
   */
  insertSampleData() {
    try {
      // サンプルデータ投入の確認
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'サンプルデータ',
        'テスト用のサンプルデータを投入しますか？\n（本番運用時は「いいえ」を選択してください）',
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        this.insertSampleSalesData();
        this.insertSampleInventoryData();
        this.insertSampleProductMaster();
      }

    } catch (error) {
      console.warn('サンプルデータ投入でエラー:', error);
    }
  }

  /**
   * サンプル販売データ投入
   */
  insertSampleSalesData() {
    const salesSheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.SALES_HISTORY);
    if (!salesSheet) return;

    const sampleData = [
      [
        '250-1234567-8901234', 'B001ABC123', new Date('2025-01-15 10:30:00'),
        'UNI-B001ABC123-250115', 'MAKAD-ABC123-250115', 'テスト商品A',
        1500, 1, 1500, 800, 150, 50, 500, 0.33, 'SHIPPED', 'FBA', 'SAMPLE', new Date()
      ],
      [
        '250-2345678-9012345', 'B002DEF456', new Date('2025-01-15 14:20:00'),
        'UNI-B002DEF456-250115', 'MAKAD-DEF456-250115', 'テスト商品B',
        2000, 2, 4000, 1200, 400, 100, 2300, 0.575, 'SHIPPED', 'FBA', 'SAMPLE', new Date()
      ]
    ];

    const range = salesSheet.getRange(2, 1, sampleData.length, sampleData[0].length);
    range.setValues(sampleData);
  }

  /**
   * サンプル在庫データ投入
   */
  insertSampleInventoryData() {
    const inventorySheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.INVENTORY);
    if (!inventorySheet) return;

    const sampleData = [
      [
        'UNI-B001ABC123-250115', 'B001ABC123', 'テスト商品A', 10, 800, 8000,
        'FBA', new Date('2025-01-10'), new Date('2025-01-15'), 5, 15, '', new Date()
      ],
      [
        'UNI-B002DEF456-250115', 'B002DEF456', 'テスト商品B', 25, 600, 15000,
        'FBA', new Date('2025-01-12'), new Date('2025-01-15'), 3, 10, '', new Date()
      ]
    ];

    const range = inventorySheet.getRange(2, 1, sampleData.length, sampleData[0].length);
    range.setValues(sampleData);
  }

  /**
   * サンプル商品マスター投入
   */
  insertSampleProductMaster() {
    const masterSheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.PRODUCT_MASTER);
    if (!masterSheet) return;

    const sampleData = [
      [
        'UNI-B001ABC123-250115', 'B001ABC123', 'MAKAD-ABC123-250115', '',
        'テスト商品A', 'テスト', 'テストブランド', '1234567890123',
        new Date('2025-01-15'), new Date(), new Date()
      ],
      [
        'UNI-B002DEF456-250115', 'B002DEF456', 'MAKAD-DEF456-250115', '',
        'テスト商品B', 'テスト', 'テストブランド', '1234567890124',
        new Date('2025-01-15'), new Date(), new Date()
      ]
    ];

    const range = masterSheet.getRange(2, 1, sampleData.length, sampleData[0].length);
    range.setValues(sampleData);
  }

  // =============================================================================
  // 初期計算・ログ
  // =============================================================================

  /**
   * 初期KPI計算実行
   */
  runInitialKPICalculation() {
    try {
      const calculator = new KPICalculator();
      calculator.recalculateAll();
    } catch (error) {
      console.warn('初期KPI計算でエラー:', error);
    }
  }

  /**
   * セットアップ完了ログ
   */
  logSetupCompletion() {
    try {
      const logSheet = this.spreadsheet.getSheetByName(SHEET_CONFIG.SYNC_LOG);
      if (logSheet) {
        logSheet.appendRow([
          new Date(),
          'INITIAL_SETUP',
          'SUCCESS',
          0,
          0,
          0,
          '初期セットアップ完了',
          'KPI管理ツールの初期セットアップが正常に完了しました'
        ]);
      }

      // Properties Serviceにもログ記録
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('SETUP_COMPLETED', new Date().toISOString());
      properties.setProperty('SETUP_VERSION', APP_CONFIG.version);

    } catch (error) {
      console.warn('セットアップログ記録でエラー:', error);
    }
  }

  // =============================================================================
  // ユーティリティ
  // =============================================================================

  /**
   * デフォルトシート削除
   */
  removeDefaultSheets() {
    try {
      const defaultSheetNames = ['シート1', 'Sheet1'];
      
      defaultSheetNames.forEach(name => {
        const sheet = this.spreadsheet.getSheetByName(name);
        if (sheet && this.spreadsheet.getSheets().length > 1) {
          this.spreadsheet.deleteSheet(sheet);
        }
      });

    } catch (error) {
      console.warn('デフォルトシート削除でエラー:', error);
    }
  }

  /**
   * セットアップ状態チェック
   */
  isSetupCompleted() {
    try {
      const properties = PropertiesService.getScriptProperties();
      return !!properties.getProperty('SETUP_COMPLETED');
    } catch (error) {
      return false;
    }
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * 初期セットアップ実行
 */
function runInitialSetup() {
  try {
    const setupManager = new SetupManager();
    return setupManager.runInitialSetup();
  } catch (error) {
    ErrorHandler.handleError(error, 'runInitialSetup');
    throw error;
  }
}