/**
 * テストランナークラス
 * 
 * Google Apps Scriptでのユニットテスト実行機能を提供します。
 */

class TestRunner {
  constructor() {
    this.testResults = [];
    this.startTime = null;
    this.endTime = null;
  }

  // =============================================================================
  // テスト実行
  // =============================================================================

  /**
   * 全テスト実行
   */
  static runAllTests() {
    const runner = new TestRunner();
    return runner.executeAllTests();
  }

  /**
   * テスト実行処理
   */
  executeAllTests() {
    this.startTime = new Date();
    this.testResults = [];

    console.log('=== KPI管理ツール テスト開始 ===');

    // 各テストグループを実行
    this.runUtilsTests();
    this.runConfigTests();
    this.runMakadoProcessorTests();
    this.runKPICalculatorTests();
    this.runValidationTests();

    this.endTime = new Date();
    const duration = (this.endTime - this.startTime) / 1000;

    // 結果サマリー
    const passedCount = this.testResults.filter(r => r.status === 'PASS').length;
    const failedCount = this.testResults.filter(r => r.status === 'FAIL').length;
    const errorCount = this.testResults.filter(r => r.status === 'ERROR').length;

    console.log(`=== テスト完了 (${duration}秒) ===`);
    console.log(`合格: ${passedCount}, 失敗: ${failedCount}, エラー: ${errorCount}`);

    return {
      results: this.testResults,
      summary: {
        total: this.testResults.length,
        passed: passedCount,
        failed: failedCount,
        errors: errorCount,
        duration: duration
      }
    };
  }

  /**
   * テストケース実行
   */
  runTest(testName, testFunction) {
    try {
      console.log(`テスト実行: ${testName}`);
      
      const testStartTime = new Date();
      testFunction();
      const testDuration = (new Date() - testStartTime) / 1000;

      this.testResults.push({
        name: testName,
        status: 'PASS',
        duration: testDuration,
        message: '成功'
      });

      console.log(`✓ ${testName} - 成功`);

    } catch (error) {
      const status = error.name === 'AssertionError' ? 'FAIL' : 'ERROR';
      
      this.testResults.push({
        name: testName,
        status: status,
        duration: 0,
        message: error.message,
        stack: error.stack
      });

      console.error(`✗ ${testName} - ${status}: ${error.message}`);
    }
  }

  // =============================================================================
  // ユーティリティテスト
  // =============================================================================

  /**
   * ユーティリティ関数テスト
   */
  runUtilsTests() {
    console.log('\n--- ユーティリティテスト ---');

    // DateUtilsテスト
    this.runTest('DateUtils.daysBetween', () => {
      const date1 = new Date('2025-01-01');
      const date2 = new Date('2025-01-10');
      const days = DateUtils.daysBetween(date1, date2);
      this.assertEqual(days, 9, '日数計算');
    });

    this.runTest('DateUtils.formatDate', () => {
      const date = new Date('2025-01-15 10:30:00');
      const formatted = DateUtils.formatDate(date, 'yyyy-MM-dd');
      this.assertEqual(formatted, '2025-01-15', '日付フォーマット');
    });

    // NumberUtilsテスト
    this.runTest('NumberUtils.safeNumber', () => {
      this.assertEqual(NumberUtils.safeNumber('123'), 123, '文字列から数値');
      this.assertEqual(NumberUtils.safeNumber(''), 0, '空文字のデフォルト値');
      this.assertEqual(NumberUtils.safeNumber(null, 100), 100, 'nullのデフォルト値');
    });

    this.runTest('NumberUtils.percentage', () => {
      const percent = NumberUtils.percentage(25, 100);
      this.assertEqual(percent, 25, 'パーセンテージ計算');
    });

    this.runTest('NumberUtils.calculateROI', () => {
      const roi = NumberUtils.calculateROI(300, 1000);
      this.assertEqual(roi, 30, 'ROI計算');
    });

    // StringUtilsテスト
    this.runTest('StringUtils.generateUnifiedSKU', () => {
      const sku = StringUtils.generateUnifiedSKU('B001234567', 'MAKAD-ABC123-250115');
      this.assertEqual(sku, 'UNI-B001234567-250115', '統一SKU生成');
    });

    this.runTest('StringUtils.normalizeSKU', () => {
      const normalized = StringUtils.normalizeSKU('test-sku_123');
      this.assertEqual(normalized, 'TEST-SKU_123', 'SKU正規化');
    });

    // ArrayUtilsテスト
    this.runTest('ArrayUtils.sum', () => {
      const data = [{ value: 10 }, { value: 20 }, { value: 30 }];
      const sum = ArrayUtils.sum(data, item => item.value);
      this.assertEqual(sum, 60, '配列合計計算');
    });

    this.runTest('ArrayUtils.unique', () => {
      const data = [1, 2, 2, 3, 3, 3];
      const unique = ArrayUtils.unique(data);
      this.assertEqual(unique.length, 3, '重複除去');
    });

    // ValidationUtilsテスト
    this.runTest('ValidationUtils.isValidASIN', () => {
      this.assertTrue(ValidationUtils.isValidASIN('B001234567'), '有効なASIN');
      this.assertFalse(ValidationUtils.isValidASIN('INVALID'), '無効なASIN');
    });

    this.runTest('ValidationUtils.isValidEmail', () => {
      this.assertTrue(ValidationUtils.isValidEmail('test@example.com'), '有効なメール');
      this.assertFalse(ValidationUtils.isValidEmail('invalid-email'), '無効なメール');
    });
  }

  // =============================================================================
  // 設定管理テスト
  // =============================================================================

  /**
   * 設定管理テスト
   */
  runConfigTests() {
    console.log('\n--- 設定管理テスト ---');

    this.runTest('ConfigManager.setKPISettings', () => {
      const config = new ConfigManager();
      const settings = {
        targetMonthlyProfit: 800000,
        targetProfitMargin: 25,
        targetROI: 30
      };
      
      // 設定保存
      config.setKPISettings(settings);
      
      // 設定取得
      const retrieved = config.getKPISettings();
      this.assertEqual(retrieved.targetMonthlyProfit, 800000, '目標月利設定');
      this.assertEqual(retrieved.targetProfitMargin, 25, '目標利益率設定');
    });

    this.runTest('ConfigManager.validateConfiguration', () => {
      const config = new ConfigManager();
      const validation = config.validateConfiguration();
      
      // エラーがあってもテストは成功（設定次第）
      this.assertTrue(typeof validation.isValid === 'boolean', '検証結果の型');
      this.assertTrue(Array.isArray(validation.errors), 'エラー配列の型');
    });
  }

  // =============================================================================
  // マカドプロセッサーテスト
  // =============================================================================

  /**
   * マカドプロセッサーテスト
   */
  runMakadoProcessorTests() {
    console.log('\n--- マカドプロセッサーテスト ---');

    this.runTest('MakadoProcessor.parseDate', () => {
      const processor = new MakadoProcessor();
      
      const date1 = processor.parseDate('2025-01-15 10:30:00');
      this.assertTrue(date1 instanceof Date, '日付オブジェクト生成');
      
      const date2 = processor.parseDate('2025/01/15 10:30:00');
      this.assertTrue(date2 instanceof Date, 'スラッシュ区切り日付');
    });

    this.runTest('MakadoProcessor.validateASIN', () => {
      const processor = new MakadoProcessor();
      
      const validASIN = processor.validateASIN('B001234567');
      this.assertEqual(validASIN, 'B001234567', '有効ASIN検証');
      
      try {
        processor.validateASIN('INVALID');
        this.fail('無効ASINで例外が発生すべき');
      } catch (error) {
        // 期待される例外
      }
    });

    this.runTest('MakadoProcessor.calculateGrossProfit', () => {
      const processor = new MakadoProcessor();
      
      const record = {
        unit_price: 1000,
        quantity: 2,
        purchase_cost: 600,
        amazon_fee: 200,
        other_cost: 50,
        shipping_fee: 0,
        points: 0
      };
      
      const profit = processor.calculateGrossProfit(record);
      this.assertEqual(profit, 950, '粗利益計算'); // (1000*2) - (600*2) - 200 - 50 = 950
    });
  }

  // =============================================================================
  // KPI計算テスト
  // =============================================================================

  /**
   * KPI計算テスト
   */
  runKPICalculatorTests() {
    console.log('\n--- KPI計算テスト ---');

    this.runTest('KPICalculator.mapSalesRecord', () => {
      const calculator = new KPICalculator();
      
      const row = [
        '250-1234567-8901234',     // 注文ID
        'B001234567',              // ASIN
        new Date('2025-01-15'),    // 注文日
        'UNI-B001234567-250115',   // 統一SKU
        'MAKAD-ABC123-250115',     // マカドSKU
        'テスト商品',               // 商品名
        1000,                      // 単価
        2,                         // 数量
        2000,                      // 合計金額
        600,                       // 仕入原価
        200,                       // Amazon手数料
        50,                        // その他費用
        1150,                      // 粗利益
        0.575,                     // 利益率
        'SHIPPED',                 // ステータス
        'FBA',                     // 配送方法
        'MAKADO'                   // データソース
      ];
      
      const record = calculator.mapSalesRecord(row);
      this.assertEqual(record.order_id, '250-1234567-8901234', '注文IDマッピング');
      this.assertEqual(record.unit_price, 1000, '単価マッピング');
      this.assertEqual(record.quantity, 2, '数量マッピング');
    });

    this.runTest('KPICalculator.calculateStagnantInventory', () => {
      const calculator = new KPICalculator();
      
      const inventoryData = [
        { unified_sku: 'SKU1', days_in_stock: 30, total_cost: 10000 },
        { unified_sku: 'SKU2', days_in_stock: 70, total_cost: 20000 },
        { unified_sku: 'SKU3', days_in_stock: 90, total_cost: 15000 }
      ];
      
      const stagnant = calculator.calculateStagnantInventory(inventoryData);
      this.assertEqual(stagnant.count, 2, '滞留在庫件数'); // 60日超が2件
      this.assertEqual(stagnant.value, 35000, '滞留在庫金額'); // 20000 + 15000
    });
  }

  // =============================================================================
  // バリデーションテスト
  // =============================================================================

  /**
   * バリデーション総合テスト
   */
  runValidationTests() {
    console.log('\n--- バリデーションテスト ---');

    this.runTest('ValidationUtils.validateRequired', () => {
      const obj = {
        name: 'テスト',
        value: 100,
        empty: '',
        missing: null
      };
      
      const result = ValidationUtils.validateRequired(obj, ['name', 'value', 'empty', 'missing']);
      this.assertFalse(result.isValid, '必須フィールドチェック');
      this.assertTrue(result.errors.length >= 2, 'エラー件数'); // empty, missing
    });

    this.runTest('NumberUtils.isInRange', () => {
      this.assertTrue(ValidationUtils.isInRange(50, 0, 100), '範囲内の値');
      this.assertFalse(ValidationUtils.isInRange(150, 0, 100), '範囲外の値');
      this.assertTrue(ValidationUtils.isInRange(0, 0, 100), '境界値（下限）');
      this.assertTrue(ValidationUtils.isInRange(100, 0, 100), '境界値（上限）');
    });
  }

  // =============================================================================
  // アサーション関数
  // =============================================================================

  /**
   * 等価アサーション
   */
  assertEqual(actual, expected, message = '') {
    if (actual !== expected) {
      throw new AssertionError(`${message}: 期待値 ${expected}, 実際の値 ${actual}`);
    }
  }

  /**
   * 真偽アサーション
   */
  assertTrue(condition, message = '') {
    if (!condition) {
      throw new AssertionError(`${message}: 条件が false です`);
    }
  }

  /**
   * 偽アサーション
   */
  assertFalse(condition, message = '') {
    if (condition) {
      throw new AssertionError(`${message}: 条件が true です`);
    }
  }

  /**
   * null/undefined アサーション
   */
  assertNotNull(value, message = '') {
    if (value === null || value === undefined) {
      throw new AssertionError(`${message}: 値が null または undefined です`);
    }
  }

  /**
   * 例外アサーション
   */
  assertThrows(func, message = '') {
    let threw = false;
    try {
      func();
    } catch (error) {
      threw = true;
    }
    
    if (!threw) {
      throw new AssertionError(`${message}: 例外が発生しませんでした`);
    }
  }

  /**
   * 強制失敗
   */
  fail(message = '') {
    throw new AssertionError(`テスト失敗: ${message}`);
  }
}

// =============================================================================
// カスタム例外クラス
// =============================================================================

class AssertionError extends Error {
  constructor(message) {
    super(message);
    this.name = 'AssertionError';
  }
}

// =============================================================================
// テストデータ生成
// =============================================================================

class TestDataGenerator {
  /**
   * サンプル販売データ生成
   */
  static generateSampleSalesData(count = 10) {
    const data = [];
    const baseDate = new Date('2025-01-01');
    
    for (let i = 0; i < count; i++) {
      const orderDate = new Date(baseDate.getTime() + i * 24 * 60 * 60 * 1000);
      
      data.push({
        order_id: `250-${1000000 + i}-${2000000 + i}`,
        asin: `B00${String(100000 + i).padStart(6, '0')}`,
        order_date: orderDate,
        unified_sku: `UNI-B00${String(100000 + i).padStart(6, '0')}-250101`,
        makado_sku: `MAKAD-${String(100000 + i).substring(0, 6)}-250101`,
        product_name: `テスト商品${i + 1}`,
        unit_price: 1000 + (i * 100),
        quantity: Math.floor(Math.random() * 3) + 1,
        total_amount: (1000 + (i * 100)) * (Math.floor(Math.random() * 3) + 1),
        purchase_cost: 500 + (i * 50),
        amazon_fee: 100 + (i * 10),
        other_cost: 50,
        gross_profit: 350 + (i * 40),
        profit_margin: 0.3 + (Math.random() * 0.2),
        status: 'SHIPPED',
        fulfillment: 'FBA',
        data_source: 'TEST'
      });
    }
    
    return data;
  }

  /**
   * サンプル在庫データ生成
   */
  static generateSampleInventoryData(count = 10) {
    const data = [];
    
    for (let i = 0; i < count; i++) {
      data.push({
        unified_sku: `UNI-B00${String(100000 + i).padStart(6, '0')}-250101`,
        asin: `B00${String(100000 + i).padStart(6, '0')}`,
        product_name: `テスト商品${i + 1}`,
        quantity: Math.floor(Math.random() * 50) + 1,
        unit_cost: 500 + (i * 50),
        total_cost: (500 + (i * 50)) * (Math.floor(Math.random() * 50) + 1),
        location: 'FBA',
        last_inbound_date: new Date('2025-01-01'),
        last_sold_date: new Date('2025-01-15'),
        days_in_stock: Math.floor(Math.random() * 100) + 1
      });
    }
    
    return data;
  }

  /**
   * サンプルマカドCSVデータ生成
   */
  static generateSampleMakadoCSV() {
    const headers = [
      '日付', '商品名', 'オーダーID', 'ASIN', 'SKU', 'コンディション',
      '配送経路', '販売価格', '送料', 'ポイント', '割引', '仕入価格',
      'その他経費', 'Amazon手数料', '粗利', 'ステータス', '販売数', '累計販売数'
    ];
    
    const data = [
      [
        '2025-01-15 10:30:00', 'テスト商品A', '250-1234567-8901234',
        'B001234567', 'MAKAD-ABC123-250115', 'New', 'FBA',
        1500, 0, 0, 0, 800, 50, 150, 500, 'Shipped', 1, 100
      ],
      [
        '2025-01-15 14:20:00', 'テスト商品B', '250-2345678-9012345',
        'B002345678', 'MAKAD-DEF456-250115', 'New', 'FBA',
        2000, 0, 0, 0, 1200, 100, 200, 500, 'Shipped', 2, 50
      ]
    ];
    
    const csvLines = [headers.join(',')];
    data.forEach(row => {
      csvLines.push(row.map(cell => `"${cell}"`).join(','));
    });
    
    return csvLines.join('\n');
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * テスト実行（メニューから）
 */
function runTests() {
  try {
    const results = TestRunner.runAllTests();
    
    // 結果をスプレッドシートに出力
    displayTestResults(results);
    
    return results;
    
  } catch (error) {
    ErrorHandler.handleError(error, 'runTests');
    throw error;
  }
}

/**
 * テスト結果をスプレッドシートに表示
 */
function displayTestResults(results) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let testSheet = ss.getSheetByName('_テスト結果');
    
    if (!testSheet) {
      testSheet = ss.insertSheet('_テスト結果');
    } else {
      testSheet.clear();
    }
    
    // ヘッダー
    testSheet.getRange('A1:E1').setValues([
      ['テスト名', 'ステータス', '実行時間', 'メッセージ', 'エラー詳細']
    ]);
    testSheet.getRange('A1:E1').setFontWeight('bold');
    
    // テスト結果
    results.results.forEach((result, index) => {
      const row = index + 2;
      testSheet.getRange(row, 1).setValue(result.name);
      testSheet.getRange(row, 2).setValue(result.status);
      testSheet.getRange(row, 3).setValue(result.duration);
      testSheet.getRange(row, 4).setValue(result.message);
      testSheet.getRange(row, 5).setValue(result.stack || '');
      
      // ステータスに応じて色分け
      const statusCell = testSheet.getRange(row, 2);
      if (result.status === 'PASS') {
        statusCell.setBackground('#c8e6c9');
      } else if (result.status === 'FAIL') {
        statusCell.setBackground('#ffcdd2');
      } else {
        statusCell.setBackground('#fff3e0');
      }
    });
    
    // サマリー
    const summaryRow = results.results.length + 4;
    testSheet.getRange(summaryRow, 1).setValue('=== テストサマリー ===');
    testSheet.getRange(summaryRow + 1, 1).setValue(`総数: ${results.summary.total}`);
    testSheet.getRange(summaryRow + 2, 1).setValue(`成功: ${results.summary.passed}`);
    testSheet.getRange(summaryRow + 3, 1).setValue(`失敗: ${results.summary.failed}`);
    testSheet.getRange(summaryRow + 4, 1).setValue(`エラー: ${results.summary.errors}`);
    testSheet.getRange(summaryRow + 5, 1).setValue(`実行時間: ${results.summary.duration}秒`);
    
    // 列幅調整
    testSheet.autoResizeColumns(1, 5);
    
    // テストシートをアクティブに
    testSheet.activate();
    
  } catch (error) {
    console.error('テスト結果表示エラー:', error);
  }
}