/**
 * KPI計算クラス
 * 
 * 各種KPI（売上、利益、在庫回転率、ROIなど）の計算機能を提供します。
 * 月利80万円達成を目指すためのKPI管理を行います。
 */

class KPICalculator {
  constructor() {
    this.cache = new CacheUtils();
    this.config = new ConfigManager();
  }

  // =============================================================================
  // メインKPI計算
  // =============================================================================

  /**
   * 全KPIを再計算
   */
  recalculateAll() {
    try {
      const startTime = new Date();
      
      console.log('KPI計算を開始します...');
      
      // 基本データ取得
      console.log('販売データを取得中...');
      const salesData = this.getSalesData();
      console.log(`販売データ取得完了: ${salesData.length}件`);
      
      console.log('在庫データを取得中...');
      const inventoryData = this.getInventoryData();
      console.log(`在庫データ取得完了: ${inventoryData.length}件`);
      
      console.log('仕入データを取得中...');
      const purchaseData = this.getPurchaseData();
      console.log(`仕入データ取得完了: ${purchaseData.length}件`);
      
      // 月次KPI計算
      console.log('月次KPI計算中...');
      const monthlyKPIs = this.calculateMonthlyKPIs(salesData, inventoryData, purchaseData);
      
      // 日次KPI計算
      console.log('日次KPI計算中...');
      const dailyKPIs = this.calculateDailyKPIs(salesData);
      
      // 商品別KPI計算
      console.log('商品別KPI計算中...');
      const productKPIs = this.calculateProductKPIs(salesData, inventoryData);
      
      // KPIダッシュボード更新
      console.log('KPIダッシュボード更新中...');
      this.updateKPIDashboard(monthlyKPIs, dailyKPIs);
      
      // 在庫分析更新
      console.log('在庫分析更新中...');
      this.updateInventoryAnalysis(inventoryData, productKPIs);
      
      // アラートチェック
      console.log('アラートチェック中...');
      const alerts = this.checkKPIAlerts(monthlyKPIs, dailyKPIs);
      
      const duration = (new Date() - startTime) / 1000;
      
      console.log(`KPI計算完了: ${duration}秒, アラート: ${alerts.length}件`);
      
      return {
        success: true,
        duration: duration,
        monthlyKPIs: monthlyKPIs,
        dailyKPIs: dailyKPIs,
        alerts: alerts.length,
        calculatedAt: new Date()
      };

    } catch (error) {
      console.error('KPI計算でエラーが発生:', error);
      ErrorHandler.handleError(error, 'KPICalculator.recalculateAll');
      throw error;
    }
  }

  // =============================================================================
  // データ取得
  // =============================================================================

  /**
   * 販売データ取得
   */
  getSalesData(startDate = null, endDate = null) {
    const cacheKey = `sales_data_${startDate}_${endDate}`;
    const cached = this.cache.get(cacheKey);
    
    if (cached) {
      return cached;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const salesSheet = ss.getSheetByName(SHEET_CONFIG.SALES_HISTORY);
      
      if (!salesSheet) {
        return [];
      }

      const lastRow = salesSheet.getLastRow();
      if (lastRow <= 1) {
        return [];
      }

      // データ取得（ヘッダー行を除く）
      const range = salesSheet.getRange(2, 1, lastRow - 1, salesSheet.getLastColumn());
      const rawData = range.getValues();
      
      // 日付フィルタリング
      const salesData = rawData
        .filter(row => row[0]) // 注文IDがある行のみ
        .map(row => this.mapSalesRecord(row))
        .filter(record => {
          if (!startDate && !endDate) return true;
          
          const orderDate = record.order_date;
          if (startDate && orderDate < startDate) return false;
          if (endDate && orderDate > endDate) return false;
          
          return true;
        });

      // キャッシュ保存
      this.cache.set(cacheKey, salesData, 300); // 5分間キャッシュ
      
      return salesData;

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.getSalesData');
      return [];
    }
  }

  /**
   * 在庫データ取得
   */
  getInventoryData() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = ss.getSheetByName(SHEET_CONFIG.INVENTORY);
      
      if (!inventorySheet) {
        return [];
      }

      const lastRow = inventorySheet.getLastRow();
      if (lastRow <= 1) {
        return [];
      }

      const range = inventorySheet.getRange(2, 1, lastRow - 1, inventorySheet.getLastColumn());
      const rawData = range.getValues();
      
      const inventoryData = rawData
        .filter(row => row[0]) // SKUがある行のみ
        .map(row => this.mapInventoryRecord(row));

      return inventoryData;

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.getInventoryData');
      return [];
    }
  }

  /**
   * 仕入データ取得
   */
  getPurchaseData(startDate = null, endDate = null) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const purchaseSheet = ss.getSheetByName(SHEET_CONFIG.PURCHASE_HISTORY);
      
      if (!purchaseSheet) {
        return [];
      }

      const lastRow = purchaseSheet.getLastRow();
      if (lastRow <= 1) {
        return [];
      }

      const range = purchaseSheet.getRange(2, 1, lastRow - 1, purchaseSheet.getLastColumn());
      const rawData = range.getValues();
      
      const purchaseData = rawData
        .filter(row => row[0]) // データがある行のみ
        .map(row => this.mapPurchaseRecord(row))
        .filter(record => {
          if (!startDate && !endDate) return true;
          
          const purchaseDate = record.purchase_date;
          if (startDate && purchaseDate < startDate) return false;
          if (endDate && purchaseDate > endDate) return false;
          
          return true;
        });

      return purchaseData;

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.getPurchaseData');
      return [];
    }
  }

  // =============================================================================
  // データマッピング
  // =============================================================================

  /**
   * 販売レコードをマッピング
   */
  mapSalesRecord(row) {
    const unit_price = NumberUtils.safeNumber(row[6]);
    const quantity = NumberUtils.safeInteger(row[7]);
    let total_amount = NumberUtils.safeNumber(row[8]);
    let gross_profit = NumberUtils.safeNumber(row[12]);
    
    // total_amountが0または空の場合は unit_price × quantity で計算
    if (total_amount === 0 && unit_price > 0 && quantity > 0) {
      total_amount = unit_price * quantity;
      
      // 粗利益も再計算（total_amount - purchase_cost - amazon_fee - other_cost）
      const purchase_cost = NumberUtils.safeNumber(row[9]);
      const amazon_fee = NumberUtils.safeNumber(row[10]);
      const other_cost = NumberUtils.safeNumber(row[11]);
      
      if (gross_profit === 0) {
        gross_profit = total_amount - purchase_cost - amazon_fee - other_cost;
      }
    }
    
    return {
      order_id: row[0],
      asin: row[1],
      order_date: new Date(row[2]),
      unified_sku: row[3],
      makado_sku: row[4],
      product_name: row[5],
      unit_price: unit_price,
      quantity: quantity,
      total_amount: total_amount,
      purchase_cost: NumberUtils.safeNumber(row[9]),
      amazon_fee: NumberUtils.safeNumber(row[10]),
      other_cost: NumberUtils.safeNumber(row[11]),
      gross_profit: gross_profit,
      profit_margin: NumberUtils.safeNumber(row[13]),
      status: row[14],
      fulfillment: row[15],
      data_source: row[16],
      import_timestamp: new Date(row[17])
    };
  }

  /**
   * 在庫レコードをマッピング
   */
  mapInventoryRecord(row) {
    return {
      unified_sku: row[0],
      asin: row[1],
      product_name: row[2],
      current_stock: NumberUtils.safeInteger(row[3]),
      purchase_cost: NumberUtils.safeNumber(row[4]),
      inventory_value: NumberUtils.safeNumber(row[5]),
      last_sold_date: row[6] ? new Date(row[6]) : null,
      days_in_stock: NumberUtils.safeInteger(row[7]),
      stock_status: row[8],
      reorder_point: NumberUtils.safeInteger(row[9]),
      updated_at: new Date(row[10])
    };
  }

  /**
   * 仕入レコードをマッピング
   */
  mapPurchaseRecord(row) {
    return {
      purchase_date: new Date(row[0]),
      asin: row[1],
      unified_sku: row[2],
      product_name: row[3],
      quantity: NumberUtils.safeInteger(row[4]),
      unit_cost: NumberUtils.safeNumber(row[5]),
      total_cost: NumberUtils.safeNumber(row[6]),
      supplier: row[7],
      purchase_order: row[8],
      status: row[9],
      notes: row[10]
    };
  }

  // =============================================================================
  // KPI計算
  // =============================================================================

  /**
   * 月次KPI計算
   */
  calculateMonthlyKPIs(salesData, inventoryData, purchaseData) {
    const currentMonth = DateUtils.getCurrentMonthRange();
    
    // 今月のデータでフィルタリング
    const monthlySales = salesData.filter(sale => 
      sale.order_date >= currentMonth.start && sale.order_date <= currentMonth.end
    );
    
    const monthlyPurchases = purchaseData.filter(purchase =>
      purchase.purchase_date >= currentMonth.start && purchase.purchase_date <= currentMonth.end
    );

    // 基本売上KPI
    const totalRevenue = ArrayUtils.sum(monthlySales, sale => sale.total_amount);
    const totalGrossProfit = ArrayUtils.sum(monthlySales, sale => sale.gross_profit);
    const totalQuantity = ArrayUtils.sum(monthlySales, sale => sale.quantity);
    const totalPurchaseAmount = ArrayUtils.sum(monthlyPurchases, purchase => purchase.total_cost);

    // 在庫KPI
    const totalInventoryValue = ArrayUtils.sum(inventoryData, inv => inv.inventory_value);
    const averageInventoryValue = this.calculateAverageInventoryValue(inventoryData);

    // 計算KPI
    const profitMargin = NumberUtils.percentage(totalGrossProfit, totalRevenue);
    const roi = NumberUtils.percentage(totalGrossProfit, totalPurchaseAmount);
    const inventoryTurnover = NumberUtils.calculateTurnoverRate(totalRevenue, averageInventoryValue);

    // 商品分析
    const uniqueASINs = ArrayUtils.unique(monthlySales, sale => sale.asin).length;
    const averageOrderValue = monthlySales.length > 0 ? totalRevenue / monthlySales.length : 0;

    // 在庫分析
    const stagnantInventory = this.calculateStagnantInventory(inventoryData);
    const lowStockItems = this.calculateLowStockItems(inventoryData);

    // 目標比較
    const kpiSettings = this.config.getKPISettings();
    const profitGoalAchievement = NumberUtils.percentage(totalGrossProfit, kpiSettings.targetMonthlyProfit);

    return {
      revenue: totalRevenue,
      grossProfit: totalGrossProfit,
      profitMargin: profitMargin,
      roi: roi,
      salesQuantity: totalQuantity,
      inventoryValue: totalInventoryValue,
      inventoryTurnover: inventoryTurnover,
      uniqueProducts: uniqueASINs,
      averageOrderValue: averageOrderValue,
      stagnantInventoryRate: stagnantInventory.rate,
      lowStockItemsCount: lowStockItems.length,
      profitGoalAchievement: profitGoalAchievement,
      month: DateUtils.formatDate(currentMonth.start, 'yyyy-MM'),
      calculatedAt: new Date()
    };
  }

  /**
   * 日次KPI計算
   */
  calculateDailyKPIs(salesData) {
    const today = DateUtils.getToday();
    const todaySales = salesData.filter(sale => 
      DateUtils.isSameDay(sale.order_date, today)
    );

    const last7Days = DateUtils.getLast7Days();
    const weekSales = salesData.filter(sale => 
      sale.order_date >= last7Days.start && sale.order_date <= last7Days.end
    );

    return {
      todayRevenue: ArrayUtils.sum(todaySales, sale => sale.total_amount),
      todayProfit: ArrayUtils.sum(todaySales, sale => sale.gross_profit),
      todayOrders: todaySales.length,
      todaySalesCount: ArrayUtils.sum(todaySales, sale => sale.quantity),
      weeklyAvgRevenue: ArrayUtils.average(weekSales, sale => sale.total_amount) * 7,
      weeklyAvgProfit: ArrayUtils.average(weekSales, sale => sale.gross_profit) * 7,
      growthRate: this.calculateGrowthRate(todaySales, weekSales),
      calculatedAt: new Date()
    };
  }

  /**
   * 商品別KPI計算
   */
  calculateProductKPIs(salesData, inventoryData) {
    const productSales = ArrayUtils.groupBy(salesData, sale => sale.asin);
    const productKPIs = [];

    Object.keys(productSales).forEach(asin => {
      const sales = productSales[asin];
      const inventory = inventoryData.find(inv => inv.asin === asin);
      
      const totalRevenue = ArrayUtils.sum(sales, sale => sale.total_amount);
      const totalProfit = ArrayUtils.sum(sales, sale => sale.gross_profit);
      const totalQuantity = ArrayUtils.sum(sales, sale => sale.quantity);

      productKPIs.push({
        asin: asin,
        productName: sales[0].product_name,
        revenue: totalRevenue,
        profit: totalProfit,
        quantity: totalQuantity,
        profitMargin: NumberUtils.percentage(totalProfit, totalRevenue),
        currentStock: inventory ? inventory.current_stock : 0,
        performance: this.calculateProductPerformance(totalRevenue, totalProfit, totalQuantity)
      });
    });

    return productKPIs.sort((a, b) => b.revenue - a.revenue);
  }

  // =============================================================================
  // KPIダッシュボード更新
  // =============================================================================

  /**
   * KPIダッシュボード更新
   */
  updateKPIDashboard(monthlyKPIs, dailyKPIs) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let kpiSheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
      
      if (!kpiSheet) {
        kpiSheet = this.createKPIDashboard(ss);
      }

      // 最終更新時刻
      kpiSheet.getRange('G1').setValue(DateUtils.formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss'));

      // 月次KPI更新
      this.updateMonthlyKPIValues(kpiSheet, monthlyKPIs);
      
      // 日次KPI更新
      this.updateDailyKPIValues(kpiSheet, dailyKPIs);

      console.log('KPIダッシュボード更新完了');

    } catch (error) {
      console.error('KPIダッシュボード更新エラー:', error);
      throw error;
    }
  }

  /**
   * 月次KPI値更新
   */
  updateMonthlyKPIValues(sheet, kpis) {
    const kpiSettings = this.config.getKPISettings();
    
    // 売上高
    sheet.getRange('B5').setValue(kpis.revenue);
    
    // 粗利益
    sheet.getRange('B6').setValue(kpis.grossProfit);
    sheet.getRange('D6').setValue(kpis.profitGoalAchievement / 100);
    
    // 利益率
    sheet.getRange('B7').setValue(kpis.profitMargin / 100);
    
    // ROI
    sheet.getRange('B8').setValue(kpis.roi / 100);
    
    // 販売数
    sheet.getRange('B9').setValue(kpis.salesQuantity);
    
    // 在庫金額
    sheet.getRange('B10').setValue(kpis.inventoryValue);
    
    // 在庫回転率
    sheet.getRange('B11').setValue(kpis.inventoryTurnover);
    
    // 滞留在庫率
    sheet.getRange('B12').setValue(kpis.stagnantInventoryRate / 100);
  }

  /**
   * 日次KPI値更新
   */
  updateDailyKPIValues(sheet, kpis) {
    // 本日売上
    sheet.getRange('B16').setValue(kpis.todayRevenue);
    sheet.getRange('C16').setValue(kpis.weeklyAvgRevenue / 7);
    sheet.getRange('D16').setValue(kpis.growthRate / 100);
    
    // 本日利益
    sheet.getRange('B17').setValue(kpis.todayProfit);
    sheet.getRange('C17').setValue(kpis.weeklyAvgProfit / 7);
    
    // 本日販売数
    sheet.getRange('B18').setValue(kpis.todaySalesCount);
    
    // 本日注文数
    sheet.getRange('B19').setValue(kpis.todayOrders);
  }

  /**
   * KPIダッシュボードシート作成
   */
  createKPIDashboard(spreadsheet) {
    const sheet = spreadsheet.insertSheet(SHEET_CONFIG.KPI_MONTHLY);
    
    // ヘッダー行の設定
    sheet.getRange('A1').setValue('Amazon販売KPI管理ダッシュボード');
    sheet.getRange('F1').setValue('最終更新: ');
    
    // 月次KPIセクション
    sheet.getRange('A3').setValue('📊 月次KPI');
    sheet.getRange('A4:E4').setValues([['項目', '実績', '目標', '達成率', '前月比']]);
    
    const monthlyItems = [
      ['売上高', '', '3,200,000', '', ''],
      ['粗利益', '', '800,000', '', ''],
      ['利益率', '', '25%', '', ''],
      ['ROI', '', '30%', '', ''],
      ['販売数', '', '600', '', ''],
      ['在庫金額', '', '1,000,000', '', ''],
      ['在庫回転率', '', '1', '', ''],
      ['滞留在庫率', '', '10%', '', '']
    ];
    
    sheet.getRange('A5:E12').setValues(monthlyItems);
    
    // 日次KPIセクション
    sheet.getRange('A14').setValue('📈 本日の実績');
    sheet.getRange('A15:D15').setValues([['項目', '実績', '7日平均', '成長率']]);
    
    const dailyItems = [
      ['本日売上', '', '', ''],
      ['本日利益', '', '', ''],
      ['本日販売数', '', '', ''],
      ['本日注文数', '', '', '']
    ];
    
    sheet.getRange('A16:D19').setValues(dailyItems);
    
    // アラートセクション
    sheet.getRange('A21').setValue('⚠️ アラート');
    
    // フォーマット設定
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    sheet.getRange('A3:A21').setFontWeight('bold');
    sheet.getRange('A4:E4').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.getRange('A15:D15').setFontWeight('bold').setBackground('#f0f0f0');
    
    // 金額列のフォーマット
    sheet.getRange('B5:B12').setNumberFormat('¥#,##0');
    sheet.getRange('C5:C12').setNumberFormat('¥#,##0');
    sheet.getRange('B16:C19').setNumberFormat('¥#,##0');
    
    // パーセント列のフォーマット
    sheet.getRange('D5:E12').setNumberFormat('0.0%');
    sheet.getRange('D16:D19').setNumberFormat('0.0%');
    
    return sheet;
  }

  // =============================================================================
  // ヘルパー関数
  // =============================================================================

  /**
   * 在庫分析更新
   */
  updateInventoryAnalysis(inventoryData, productKPIs) {
    console.log('在庫分析更新: 商品数', productKPIs.length);
  }

  /**
   * KPIアラートチェック
   */
  checkKPIAlerts(monthlyKPIs, dailyKPIs) {
    const alerts = [];

    // 利益目標未達アラート
    if (monthlyKPIs.profitGoalAchievement < 80) {
      alerts.push({
        type: 'profit_target',
        severity: 'high',
        message: `月間利益目標の達成率が${monthlyKPIs.profitGoalAchievement.toFixed(1)}%です`
      });
    }

    // 在庫回転率アラート
    if (monthlyKPIs.inventoryTurnover < 0.5) {
      alerts.push({
        type: 'inventory_turnover',
        severity: 'medium',
        message: `在庫回転率が低下しています: ${monthlyKPIs.inventoryTurnover.toFixed(2)}`
      });
    }

    // 滞留在庫アラート
    if (monthlyKPIs.stagnantInventoryRate > 15) {
      alerts.push({
        type: 'stagnant_inventory',
        severity: 'medium',
        message: `滞留在庫率が${monthlyKPIs.stagnantInventoryRate.toFixed(1)}%です`
      });
    }

    return alerts;
  }

  /**
   * 滞留在庫計算
   */
  calculateStagnantInventory(inventoryData) {
    const stagnantThreshold = 30; // 30日以上
    const stagnantItems = inventoryData.filter(inv => inv.days_in_stock > stagnantThreshold);
    
    return {
      count: stagnantItems.length,
      rate: NumberUtils.percentage(stagnantItems.length, inventoryData.length)
    };
  }

  /**
   * 低在庫商品計算
   */
  calculateLowStockItems(inventoryData) {
    return inventoryData.filter(inv => inv.current_stock < inv.reorder_point);
  }

  /**
   * 平均在庫価値計算
   */
  calculateAverageInventoryValue(inventoryData) {
    if (inventoryData.length === 0) return 0;
    const totalValue = ArrayUtils.sum(inventoryData, inv => inv.inventory_value);
    return totalValue / inventoryData.length;
  }

  /**
   * 成長率計算
   */
  calculateGrowthRate(todaySales, weekSales) {
    const todayRevenue = ArrayUtils.sum(todaySales, sale => sale.total_amount);
    const weeklyAvg = ArrayUtils.average(weekSales, sale => sale.total_amount);
    
    if (weeklyAvg === 0) return 0;
    return NumberUtils.percentage(todayRevenue - weeklyAvg, weeklyAvg);
  }

  /**
   * 商品パフォーマンス計算
   */
  calculateProductPerformance(revenue, profit, quantity) {
    if (revenue === 0) return 'poor';
    
    const profitMargin = NumberUtils.percentage(profit, revenue);
    
    if (profitMargin > 30 && quantity > 10) return 'excellent';
    if (profitMargin > 20 && quantity > 5) return 'good';
    if (profitMargin > 10) return 'average';
    return 'poor';
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * KPI再計算実行
 */
function runKPICalculation() {
  try {
    const calculator = new KPICalculator();
    return calculator.recalculateAll();
  } catch (error) {
    ErrorHandler.handleError(error, 'runKPICalculation');
    throw error;
  }
}