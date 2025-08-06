/**
 * テスト用グラフ作成関数
 * グラフ作成の問題を診断するための最小限の実装
 */

/**
 * 最もシンプルなグラフを作成
 */
function createTestChart() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('KPI月次管理');
    
    if (!sheet) {
      throw new Error('KPI月次管理シートが見つかりません');
    }
    
    // 既存のチャートをすべて削除
    const charts = sheet.getCharts();
    charts.forEach(chart => sheet.removeChart(chart));
    
    // シンプルなテストデータを直接シートに書き込む
    const testData = [
      ['月', '売上高', '粗利益'],
      ['2024/01', 2500000, 625000],
      ['2024/02', 2800000, 700000],
      ['2024/03', 3200000, 800000],
      ['2024/04', 2900000, 725000],
      ['2024/05', 3100000, 775000],
      ['2024/06', 3300000, 825000]
    ];
    
    // データを配置（AA列以降の見えない場所）
    const dataRange = sheet.getRange(1, 27, testData.length, testData[0].length);
    dataRange.setValues(testData);
    
    // 最もシンプルなチャートを作成
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(5, 7, 0, 0)
      .setOption('title', 'テストチャート')
      .setOption('width', 600)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(chart);
    
    // データをクリア
    dataRange.clear();
    
    SpreadsheetApp.getUi().alert('成功', 'テストチャートを作成しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'チャート作成エラー: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    console.error('チャート作成エラー:', error);
  }
}

/**
 * 実際のKPIデータでグラフを作成
 */
function createKPIChart() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('KPI月次管理');
    
    if (!sheet) {
      throw new Error('KPI月次管理シートが見つかりません');
    }
    
    // 既存のチャートをすべて削除
    const charts = sheet.getCharts();
    charts.forEach(chart => sheet.removeChart(chart));
    
    // 現在のKPIデータを取得
    const revenue = sheet.getRange('C4').getValue() || 3200000;
    const profit = sheet.getRange('C5').getValue() || 800000;
    
    // 過去6ヶ月のサンプルデータを生成
    const months = 6;
    const chartData = [['月', '売上高', '粗利益']];
    
    for (let i = months - 1; i >= 0; i--) {
      const date = new Date();
      date.setMonth(date.getMonth() - i);
      const monthStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM');
      
      // ランダムな変動を加える（±20%）
      const variation = 0.8 + (Math.random() * 0.4);
      chartData.push([
        monthStr,
        Math.round(revenue * variation),
        Math.round(profit * variation)
      ]);
    }
    
    // 現在月を追加
    const currentMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');
    chartData.push([currentMonth, revenue, profit]);
    
    // データを配置
    const dataRange = sheet.getRange(1, 27, chartData.length, chartData[0].length);
    dataRange.setValues(chartData);
    
    // チャートを作成
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(5, 7, 0, 0)
      .setOption('title', '売上・利益推移')
      .setOption('width', 600)
      .setOption('height', 300)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: '月'})
      .setOption('vAxis', {title: '金額（円）', format: '#,##0'})
      .setOption('series', {
        0: {color: '#4285f4', lineWidth: 3},
        1: {color: '#34a853', lineWidth: 3}
      })
      .build();
    
    sheet.insertChart(chart);
    
    // データをクリア
    dataRange.clear();
    
    SpreadsheetApp.getUi().alert('成功', 'KPIチャートを作成しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'チャート作成エラー: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    console.error('チャート作成エラー:', error);
  }
}

/**
 * 複数のグラフを作成
 */
function createAllTestCharts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('KPI月次管理');
    
    if (!sheet) {
      throw new Error('KPI月次管理シートが見つかりません');
    }
    
    // 既存のチャートをすべて削除
    const charts = sheet.getCharts();
    charts.forEach(chart => sheet.removeChart(chart));
    
    // 現在のKPIデータを取得
    const revenue = sheet.getRange('C4').getValue() || 3200000;
    const profit = sheet.getRange('C5').getValue() || 800000;
    const roi = sheet.getRange('C7').getValue() || 28;
    
    // 1. トレンドチャート
    const trendData = [['月', '売上高', '粗利益']];
    for (let i = 5; i >= 0; i--) {
      const date = new Date();
      date.setMonth(date.getMonth() - i);
      const monthStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM');
      const variation = 0.8 + (Math.random() * 0.4);
      trendData.push([
        monthStr,
        Math.round(revenue * variation),
        Math.round(profit * variation)
      ]);
    }
    
    // データ配置とチャート作成
    const trendRange = sheet.getRange(1, 27, trendData.length, trendData[0].length);
    trendRange.setValues(trendData);
    
    const trendChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(trendRange)
      .setPosition(5, 7, 0, 0)
      .setOption('title', '売上・利益推移（過去6ヶ月）')
      .setOption('width', 600)
      .setOption('height', 300)
      .build();
    
    sheet.insertChart(trendChart);
    trendRange.clear();
    
    // 2. 比較チャート
    const compareData = [
      ['項目', '実績', '目標'],
      ['売上高', revenue, 3200000],
      ['粗利益', profit, 800000],
      ['ROI(%)', roi, 30]
    ];
    
    const compareRange = sheet.getRange(10, 27, compareData.length, compareData[0].length);
    compareRange.setValues(compareData);
    
    const compareChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(compareRange)
      .setPosition(17, 7, 0, 0)
      .setOption('title', '実績 vs 目標')
      .setOption('width', 600)
      .setOption('height', 250)
      .build();
    
    sheet.insertChart(compareChart);
    compareRange.clear();
    
    SpreadsheetApp.getUi().alert('成功', 'すべてのテストチャートを作成しました。', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', 'チャート作成エラー: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    console.error('チャート作成エラー:', error);
  }
}