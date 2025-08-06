/**
 * スプレッドシート操作とグラフ管理クラス
 * 
 * シートの操作、グラフの作成・更新、UI要素の管理を行います。
 * KPI月次管理シートの時系列分析機能を提供します。
 */

class SheetManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.cache = new CacheUtils();
  }

  // =============================================================================
  // グラフ設定
  // =============================================================================

  /**
   * グラフ設定定義
   */
  static get CHART_CONFIG() {
    return {
      TREND_CHART: {
        position: { row: 5, column: 7 }, // G5セル
        size: { width: 600, height: 300 },
        type: 'LINE',
        title: '売上・利益推移（12ヶ月）'
      },
      COMPARISON_CHART: {
        position: { row: 17, column: 7 }, // G17セル
        size: { width: 600, height: 250 },
        type: 'COLUMN',
        title: '前年同月比較'
      },
      KPI_GAUGE: {
        position: { row: 29, column: 7 }, // G29セル
        size: { width: 300, height: 200 },
        type: 'GAUGE',
        title: '目標達成率'
      }
    };
  }

  // =============================================================================
  // グラフ作成・更新機能
  // =============================================================================

  /**
   * グラフを作成または更新
   * @param {Object} chartConfig - グラフ設定
   * @param {Object} dataRange - データ範囲
   * @param {Sheet} sheet - 対象シート
   */
  createOrUpdateChart(chartConfig, dataRange, sheet) {
    try {
      const existingCharts = sheet.getCharts();
      let targetChart = null;
      
      // 既存チャート検索
      existingCharts.forEach(chart => {
        const options = chart.getOptions();
        const title = options.get('title');
        if (title === chartConfig.title) {
          targetChart = chart;
        }
      });
      
      if (targetChart) {
        // 既存チャート更新
        this.updateChart(targetChart, dataRange, chartConfig, sheet);
      } else {
        // 新規チャート作成
        this.createChart(dataRange, chartConfig, sheet);
      }
      
    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.createOrUpdateChart');
      throw error;
    }
  }

  /**
   * トレンドチャート作成
   * @param {Array} historicalData - 過去12ヶ月のデータ
   */
  createTrendChart(historicalData) {
    try {
      const sheet = this.ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
      if (!sheet) {
        throw new Error('KPI月次管理シートが見つかりません');
      }

      // 既存のトレンドチャートを削除
      this.removeChartByTitle(sheet, SheetManager.CHART_CONFIG.TREND_CHART.title);

      // データ準備
      const chartData = this.prepareTrendChartData(historicalData);
      if (chartData.length <= 1) { // ヘッダーのみの場合
        console.log('トレンドチャート用のデータがありません');
        
        // データがない場合はダミーデータで表示
        const dummyData = [
          ['月', '売上高', '粗利益'],
          ['2024/01', 2500000, 625000],
          ['2024/02', 2800000, 700000],
          ['2024/03', 3200000, 800000],
          ['2024/04', 2900000, 725000],
          ['2024/05', 3100000, 775000],
          ['2024/06', 3300000, 825000]
        ];
        
        return this.createSimpleChart(sheet, dummyData, 'トレンドチャート（サンプルデータ）');
      }

      return this.createSimpleChart(sheet, chartData, SheetManager.CHART_CONFIG.TREND_CHART.title);

    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.createTrendChart');
      console.error('トレンドチャート作成エラー:', error);
      return null;
    }
  }

  /**
   * 比較チャート作成
   * @param {Object} currentKPI - 当月のKPI
   * @param {Object} previousYearKPI - 前年同月のKPI
   */
  createComparisonChart(currentKPI, previousYearKPI) {
    try {
      const sheet = this.ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
      if (!sheet) {
        throw new Error('KPI月次管理シートが見つかりません');
      }

      // 既存の比較チャートを削除
      this.removeChartByTitle(sheet, SheetManager.CHART_CONFIG.COMPARISON_CHART.title);

      // 比較データ準備
      const comparisonData = this.prepareComparisonChartData(currentKPI, previousYearKPI);
      
      return this.createSimpleColumnChart(sheet, comparisonData, '前年同月比較');

    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.createComparisonChart');
      console.error('比較チャート作成エラー:', error);
      return null;
    }
  }

  /**
   * ゲージチャート作成（目標達成率）
   * @param {number} achievementRate - 達成率（%）
   */
  createGaugeChart(achievementRate) {
    try {
      const sheet = this.ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
      if (!sheet) {
        throw new Error('KPI月次管理シートが見つかりません');
      }

      // 既存のゲージチャートを削除
      this.removeChartByTitle(sheet, SheetManager.CHART_CONFIG.KPI_GAUGE.title);

      const gaugeData = [
        ['ラベル', '値'],
        ['達成率', achievementRate / 100]
      ];

      const tempSheet = this.createTempDataSheet('GaugeData', gaugeData);
      const dataRange = tempSheet.getRange(1, 1, 2, 2);

      const chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.GAUGE)
        .addRange(dataRange)
        .setPosition(SheetManager.CHART_CONFIG.KPI_GAUGE.position.row, 
                     SheetManager.CHART_CONFIG.KPI_GAUGE.position.column, 0, 0)
        .setOption('title', SheetManager.CHART_CONFIG.KPI_GAUGE.title)
        .setOption('width', SheetManager.CHART_CONFIG.KPI_GAUGE.size.width)
        .setOption('height', SheetManager.CHART_CONFIG.KPI_GAUGE.size.height)
        .setOption('greenFrom', 0.8)
        .setOption('greenTo', 1.2)
        .setOption('yellowFrom', 0.6)
        .setOption('yellowTo', 0.8)
        .setOption('redFrom', 0)
        .setOption('redTo', 0.6)
        .setOption('max', 1.5)
        .setOption('min', 0);

      const chart = chartBuilder.build();
      sheet.insertChart(chart);

      // 一時シート削除
      this.ss.deleteSheet(tempSheet);

      console.log('ゲージチャートを作成しました');
      return chart;

    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.createGaugeChart');
      throw error;
    }
  }

  // =============================================================================
  // ヘルパーメソッド
  // =============================================================================

  /**
   * トレンドチャート用データ準備
   * @param {Array} historicalData - 履歴データ
   * @returns {Array} チャート用データ
   */
  prepareTrendChartData(historicalData) {
    if (!historicalData || historicalData.length === 0) {
      return [];
    }

    const chartData = [['月', '売上高', '粗利益']];
    
    // 過去12ヶ月分のデータを準備
    const sortedData = historicalData
      .filter(data => data.hasData !== false)
      .sort((a, b) => a.yearMonth.localeCompare(b.yearMonth))
      .slice(-12); // 最新12ヶ月

    sortedData.forEach(data => {
      const monthLabel = data.yearMonth.replace('-', '/');
      chartData.push([
        monthLabel,
        NumberUtils.safeNumber(data.revenue),
        NumberUtils.safeNumber(data.grossProfit)
      ]);
    });

    return chartData;
  }

  /**
   * 比較チャート用データ準備
   * @param {Object} currentKPI - 当月KPI
   * @param {Object} previousYearKPI - 前年同月KPI
   * @returns {Array} チャート用データ
   */
  prepareComparisonChartData(currentKPI, previousYearKPI) {
    return [
      ['メトリクス', '当年', '前年'],
      ['売上高', 
       NumberUtils.safeNumber(currentKPI.revenue), 
       NumberUtils.safeNumber(previousYearKPI?.revenue || 0)],
      ['粗利益', 
       NumberUtils.safeNumber(currentKPI.grossProfit), 
       NumberUtils.safeNumber(previousYearKPI?.grossProfit || 0)],
      ['販売数', 
       NumberUtils.safeNumber(currentKPI.salesQuantity), 
       NumberUtils.safeNumber(previousYearKPI?.salesQuantity || 0)]
    ];
  }

  /**
   * 一時データシート作成
   * @param {string} sheetName - シート名
   * @param {Array} data - データ
   * @returns {Sheet} 作成されたシート
   */
  createTempDataSheet(sheetName, data) {
    // 既存の一時シートがあれば削除
    const existingSheet = this.ss.getSheetByName(sheetName);
    if (existingSheet) {
      this.ss.deleteSheet(existingSheet);
    }

    const tempSheet = this.ss.insertSheet(sheetName);
    const range = tempSheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);

    return tempSheet;
  }

  /**
   * タイトルでチャートを検索・削除
   * @param {Sheet} sheet - 対象シート
   * @param {string} title - チャートタイトル
   */
  removeChartByTitle(sheet, title) {
    const charts = sheet.getCharts();
    charts.forEach(chart => {
      const chartTitle = chart.getOptions().get('title');
      if (chartTitle === title) {
        sheet.removeChart(chart);
      }
    });
  }

  /**
   * レスポンシブ更新処理
   * @param {Object} kpiData - KPIデータ
   */
  updateChartsResponsively(kpiData) {
    try {
      console.log('チャートの更新を開始します...');

      // 更新タスクを順次実行
      const updateTasks = [
        {
          name: 'トレンドチャート',
          action: () => this.createTrendChart(kpiData.historical)
        },
        {
          name: '比較チャート',
          action: () => this.createComparisonChart(kpiData.current, kpiData.previousYear)
        },
        {
          name: 'ゲージチャート',
          action: () => this.createGaugeChart(kpiData.current?.profitGoalAchievement || 0)
        }
      ];

      updateTasks.forEach((task, index) => {
        try {
          console.log(`${task.name}を更新中...`);
          task.action();
          console.log(`${task.name}の更新完了`);
        } catch (error) {
          console.error(`${task.name}の更新エラー:`, error);
          this.logChartError(task.name, error);
        }
      });

      console.log('全チャートの更新が完了しました');

    } catch (error) {
      console.error('チャート更新処理エラー:', error);
      ErrorHandler.handleError(error, 'SheetManager.updateChartsResponsively');
      throw error;
    }
  }

  /**
   * チャートエラーログ記録
   * @param {string} chartName - チャート名
   * @param {Error} error - エラー
   */
  logChartError(chartName, error) {
    try {
      const logSheet = this.ss.getSheetByName(SHEET_CONFIG.SYNC_LOG);
      if (logSheet) {
        logSheet.appendRow([
          new Date(),
          'CHART_ERROR',
          chartName,
          error.message,
          error.stack || ''
        ]);
      }
    } catch (logError) {
      console.error('ログ記録エラー:', logError);
    }
  }

  // =============================================================================
  // 既存チャート更新
  // =============================================================================

  /**
   * 既存チャートを更新
   * @param {Chart} chart - 既存チャート
   * @param {Range} dataRange - 新しいデータ範囲
   * @param {Object} chartConfig - チャート設定
   * @param {Sheet} sheet - シート
   */
  updateChart(chart, dataRange, chartConfig, sheet) {
    try {
      // 既存チャートを削除して新規作成
      sheet.removeChart(chart);
      this.createChart(dataRange, chartConfig, sheet);
      
    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.updateChart');
      throw error;
    }
  }

  /**
   * 新規チャート作成
   * @param {Range} dataRange - データ範囲
   * @param {Object} chartConfig - チャート設定
   * @param {Sheet} sheet - シート
   */
  createChart(dataRange, chartConfig, sheet) {
    try {
      let chartBuilder = sheet.newChart()
        .addRange(dataRange)
        .setPosition(chartConfig.position.row, chartConfig.position.column, 0, 0)
        .setOption('title', chartConfig.title)
        .setOption('width', chartConfig.size.width)
        .setOption('height', chartConfig.size.height);

      // チャートタイプに応じた設定
      switch (chartConfig.type) {
        case 'LINE':
          chartBuilder = chartBuilder.setChartType(Charts.ChartType.LINE);
          break;
        case 'COLUMN':
          chartBuilder = chartBuilder.setChartType(Charts.ChartType.COLUMN);
          break;
        case 'GAUGE':
          chartBuilder = chartBuilder.setChartType(Charts.ChartType.GAUGE);
          break;
      }

      const chart = chartBuilder.build();
      sheet.insertChart(chart);

      return chart;

    } catch (error) {
      ErrorHandler.handleError(error, 'SheetManager.createChart');
      throw error;
    }
  }

  /**
   * シンプルなラインチャート作成
   * @param {Sheet} sheet - 対象シート
   * @param {Array} chartData - チャートデータ
   * @param {string} title - チャートタイトル
   */
  createSimpleChart(sheet, chartData, title) {
    try {
      // 一時的に空いているセル範囲を使用してデータを配置
      const startRow = 50; // 空いている行を使用
      const startCol = 15;  // O列以降を使用
      
      const range = sheet.getRange(startRow, startCol, chartData.length, chartData[0].length);
      range.setValues(chartData);
      
      const chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(range)
        .setPosition(5, 7, 0, 0) // G5セル
        .setOption('title', title)
        .setOption('width', 600)
        .setOption('height', 300)
        .setOption('hAxis.title', '月')
        .setOption('vAxis.title', '金額（円）')
        .setOption('series', {
          0: { color: '#4285f4', lineWidth: 3, pointSize: 6 },
          1: { color: '#34a853', lineWidth: 3, pointSize: 6 }
        })
        .setOption('legend.position', 'bottom');

      const chart = chartBuilder.build();
      sheet.insertChart(chart);
      
      // データ範囲をクリア
      range.clear();
      
      console.log(`${title}を作成しました`);
      return chart;
      
    } catch (error) {
      console.error('シンプルチャート作成エラー:', error);
      throw error;
    }
  }

  /**
   * シンプルなカラムチャート作成
   * @param {Sheet} sheet - 対象シート
   * @param {Array} chartData - チャートデータ
   * @param {string} title - チャートタイトル
   */
  createSimpleColumnChart(sheet, chartData, title) {
    try {
      const startRow = 60; // 空いている行を使用
      const startCol = 15;  // O列以降を使用
      
      const range = sheet.getRange(startRow, startCol, chartData.length, chartData[0].length);
      range.setValues(chartData);
      
      const chartBuilder = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(17, 7, 0, 0) // G17セル
        .setOption('title', title)
        .setOption('width', 600)
        .setOption('height', 250)
        .setOption('hAxis.title', 'KPI項目')
        .setOption('vAxis.title', '値')
        .setOption('series', {
          0: { color: '#4285f4' },
          1: { color: '#ea4335' }
        })
        .setOption('legend.position', 'bottom');

      const chart = chartBuilder.build();
      sheet.insertChart(chart);
      
      // データ範囲をクリア
      range.clear();
      
      console.log(`${title}を作成しました`);
      return chart;
      
    } catch (error) {
      console.error('シンプルカラムチャート作成エラー:', error);
      throw error;
    }
  }

  // =============================================================================
  // サンプルデータ生成
  // =============================================================================

  /**
   * 履歴データが不足している場合のサンプルデータ生成
   * @param {string} currentMonth - 現在月（YYYY-MM形式）
   * @param {Object} currentKPI - 現在のKPI
   * @returns {Array} サンプル履歴データ
   */
  generateSampleHistoricalData(currentMonth, currentKPI) {
    const sampleData = [];
    const currentDate = new Date(currentMonth + '-01');
    
    // 過去12ヶ月のサンプルデータを生成
    for (let i = 11; i >= 0; i--) {
      const monthDate = new Date(currentDate);
      monthDate.setMonth(monthDate.getMonth() - i);
      
      const yearMonth = DateUtils.formatMonth(monthDate);
      const baseRevenue = currentKPI?.revenue || 3000000;
      const baseProfit = currentKPI?.grossProfit || 800000;
      
      // 月ごとに多少のバリエーションを追加
      const variation = 0.8 + (Math.random() * 0.4); // 0.8〜1.2の範囲
      const seasonalFactor = this.getSeasonalFactor(monthDate.getMonth() + 1);
      
      sampleData.push({
        yearMonth: yearMonth,
        revenue: Math.round(baseRevenue * variation * seasonalFactor),
        grossProfit: Math.round(baseProfit * variation * seasonalFactor),
        profitMargin: currentKPI?.profitMargin || 25,
        roi: currentKPI?.roi || 30,
        salesQuantity: Math.round((currentKPI?.salesQuantity || 500) * variation),
        hasData: true,
        isSample: true
      });
    }
    
    return sampleData;
  }

  /**
   * 季節調整係数を取得
   * @param {number} month - 月（1-12）
   * @returns {number} 季節調整係数
   */
  getSeasonalFactor(month) {
    // Amazonの一般的な季節性を考慮
    const seasonalFactors = {
      1: 0.9,   // 年始
      2: 0.8,   // 2月
      3: 0.9,   // 3月
      4: 0.95,  // 4月
      5: 0.9,   // 5月
      6: 0.9,   // 6月
      7: 0.95,  // 7月
      8: 0.95,  // 8月
      9: 1.0,   // 9月
      10: 1.1,  // 10月
      11: 1.2,  // 11月（ブラックフライデー）
      12: 1.3   // 12月（年末商戦）
    };
    
    return seasonalFactors[month] || 1.0;
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * 全チャートを更新
 */
function updateAllCharts() {
  try {
    const historyManager = new KPIHistoryManager();
    const sheetManager = new SheetManager();
    const ui = SpreadsheetApp.getUi();
    
    // 現在のKPIを取得・計算
    const calculator = new KPICalculator();
    const result = calculator.recalculateAll();
    
    // 現在のKPIを履歴に保存
    const currentMonth = DateUtils.getCurrentMonth();
    if (result.monthlyKPIs) {
      historyManager.saveMonthlyKPI(currentMonth, result.monthlyKPIs);
    }
    
    // 履歴データを取得
    const historicalKPIs = historyManager.getHistoricalKPIs(12);
    const previousYearKPI = historyManager.getPreviousYearKPI(currentMonth);
    
    // データ存在チェック
    const hasHistoricalData = historicalKPIs && historicalKPIs.some(kpi => kpi.hasData !== false);
    
    if (!hasHistoricalData) {
      // データが不足している場合はサンプルデータを生成
      const sampleData = sheetManager.generateSampleHistoricalData(currentMonth, result.monthlyKPIs);
      const kpiData = {
        current: result.monthlyKPIs,
        historical: sampleData,
        previousYear: null
      };
      
      sheetManager.updateChartsResponsively(kpiData);
      ui.alert('情報', 'グラフを作成しました。\n履歴データが少ないため、サンプルデータを使用しています。\n実際のデータが蓄積されると、より正確なグラフが表示されます。', ui.ButtonSet.OK);
      
    } else {
      // 通常のグラフ更新
      const kpiData = {
        current: result.monthlyKPIs,
        historical: historicalKPIs,
        previousYear: previousYearKPI
      };
      
      sheetManager.updateChartsResponsively(kpiData);
      ui.alert('成功', 'すべてのグラフが更新されました。', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ErrorHandler.handleError(error, 'updateAllCharts');
    SpreadsheetApp.getUi().alert('エラー', `グラフ更新中にエラーが発生しました:\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}