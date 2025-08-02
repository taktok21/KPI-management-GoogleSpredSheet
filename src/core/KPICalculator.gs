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
      
      // 基本データ取得
      const salesData = this.getSalesData();
      const inventoryData = this.getInventoryData();
      const purchaseData = this.getPurchaseData();
      
      // 月次KPI計算
      const monthlyKPIs = this.calculateMonthlyKPIs(salesData, inventoryData, purchaseData);
      
      // 日次KPI計算
      const dailyKPIs = this.calculateDailyKPIs(salesData);
      
      // 商品別KPI計算
      const productKPIs = this.calculateProductKPIs(salesData, inventoryData);
      
      // KPIダッシュボード更新
      this.updateKPIDashboard(monthlyKPIs, dailyKPIs);
      
      // 在庫分析更新
      this.updateInventoryAnalysis(inventoryData, productKPIs);
      
      // アラートチェック
      const alerts = this.checkKPIAlerts(monthlyKPIs, dailyKPIs);
      
      const duration = (new Date() - startTime) / 1000;
      
      return {
        success: true,
        duration: duration,
        monthlyKPIs: monthlyKPIs,
        dailyKPIs: dailyKPIs,
        alerts: alerts.length,
        calculatedAt: new Date()
      };

    } catch (error) {
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
    const cacheKey = 'inventory_data';
    const cached = this.cache.get(cacheKey);
    
    if (cached) {
      return cached;
    }

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

      this.cache.set(cacheKey, inventoryData, 600); // 10分間キャッシュ
      
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
      
      return rawData
        .filter(row => row[0]) // SKUがある行のみ
        .map(row => this.mapPurchaseRecord(row))
        .filter(record => {
          if (!startDate && !endDate) return true;
          
          const purchaseDate = record.purchase_date;
          if (startDate && purchaseDate < startDate) return false;
          if (endDate && purchaseDate > endDate) return false;
          
          return true;
        });

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.getPurchaseData');
      return [];
    }
  }

  // =============================================================================
  // データマッピング
  // =============================================================================

  /**
   * 販売レコードマッピング
   */
  mapSalesRecord(row) {
    return {
      order_id: row[0],
      asin: row[1],
      order_date: new Date(row[2]),
      unified_sku: row[3],
      makado_sku: row[4],
      product_name: row[5],
      unit_price: NumberUtils.safeNumber(row[6]),
      quantity: NumberUtils.safeInteger(row[7]),
      total_amount: NumberUtils.safeNumber(row[8]),
      purchase_cost: NumberUtils.safeNumber(row[9]),
      amazon_fee: NumberUtils.safeNumber(row[10]),
      other_cost: NumberUtils.safeNumber(row[11]),
      gross_profit: NumberUtils.safeNumber(row[12]),
      profit_margin: NumberUtils.safeNumber(row[13]),
      status: row[14],
      fulfillment: row[15],
      data_source: row[16]
    };
  }

  /**
   * 在庫レコードマッピング
   */
  mapInventoryRecord(row) {
    return {
      unified_sku: row[0],
      asin: row[1],
      product_name: row[2],
      quantity: NumberUtils.safeInteger(row[3]),
      unit_cost: NumberUtils.safeNumber(row[4]),
      total_cost: NumberUtils.safeNumber(row[5]),
      location: row[6],
      last_inbound_date: row[7] ? new Date(row[7]) : null,
      last_sold_date: row[8] ? new Date(row[8]) : null,
      days_in_stock: NumberUtils.safeInteger(row[9])
    };
  }

  /**
   * 仕入レコードマッピング
   */
  mapPurchaseRecord(row) {
    return {
      unified_sku: row[0],
      asin: row[1],
      purchase_date: new Date(row[2]),
      supplier: row[3],
      quantity: NumberUtils.safeInteger(row[4]),
      unit_cost: NumberUtils.safeNumber(row[5]),
      total_cost: NumberUtils.safeNumber(row[6]),
      shipping_cost: NumberUtils.safeNumber(row[7])
    };
  }

  // =============================================================================
  // 月次KPI計算
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
    const totalInventoryValue = ArrayUtils.sum(inventoryData, inv => inv.total_cost);
    const averageInventoryValue = this.calculateAverageInventoryValue(inventoryData);

    // 計算KPI
    const profitMargin = NumberUtils.percentage(totalGrossProfit, totalRevenue);
    const roi = NumberUtils.percentage(totalGrossProfit, totalPurchaseAmount);
    const inventoryTurnover = NumberUtils.calculateTurnoverRate(totalRevenue, averageInventoryValue);
    const turnoverDays = NumberUtils.calculateTurnoverDays(inventoryTurnover, 30);

    // 商品分析
    const uniqueASINs = ArrayUtils.unique(monthlySales, sale => sale.asin).length;
    const averageOrderValue = totalQuantity > 0 ? totalRevenue / monthlySales.length : 0;
    const averageSellingPrice = totalQuantity > 0 ? totalRevenue / totalQuantity : 0;

    // 在庫分析
    const stagnantInventory = this.calculateStagnantInventory(inventoryData);
    const lowStockItems = this.calculateLowStockItems(inventoryData);

    // 目標比較
    const kpiSettings = this.config.getKPISettings();
    const profitGoalAchievement = NumberUtils.percentage(totalGrossProfit, kpiSettings.targetMonthlyProfit);

    return {
      period: {
        start: currentMonth.start,
        end: currentMonth.end,
        year: currentMonth.start.getFullYear(),
        month: currentMonth.start.getMonth() + 1
      },
      
      // 売上・利益KPI
      totalRevenue: totalRevenue,
      totalGrossProfit: totalGrossProfit,
      profitMargin: profitMargin,
      roi: roi,
      profitGoalAchievement: profitGoalAchievement,
      
      // 販売KPI
      totalQuantity: totalQuantity,
      salesCount: monthlySales.length,
      uniqueASINs: uniqueASINs,
      averageOrderValue: averageOrderValue,
      averageSellingPrice: averageSellingPrice,
      
      // 在庫KPI
      totalInventoryValue: totalInventoryValue,
      averageInventoryValue: averageInventoryValue,
      inventoryTurnover: inventoryTurnover,
      turnoverDays: turnoverDays,
      stagnantInventoryValue: stagnantInventory.value,
      stagnantInventoryRate: stagnantInventory.rate,
      lowStockItemsCount: lowStockItems.count,
      
      // 仕入KPI
      totalPurchaseAmount: totalPurchaseAmount,
      purchaseCount: monthlyPurchases.length,
      averagePurchaseAmount: monthlyPurchases.length > 0 ? totalPurchaseAmount / monthlyPurchases.length : 0,
      
      // 効率性KPI
      grossProfitPerOrder: monthlySales.length > 0 ? totalGrossProfit / monthlySales.length : 0,
      revenuePerASIN: uniqueASINs > 0 ? totalRevenue / uniqueASINs : 0,
      
      calculatedAt: new Date()
    };
  }

  // =============================================================================
  // 日次KPI計算
  // =============================================================================

  /**
   * 日次KPI計算
   */
  calculateDailyKPIs(salesData) {
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayEnd = new Date(todayStart.getTime() + 24 * 60 * 60 * 1000 - 1);

    // 今日のデータ
    const todaySales = salesData.filter(sale => 
      sale.order_date >= todayStart && sale.order_date <= todayEnd
    );

    // 過去7日間の平均
    const weekAgo = new Date(todayStart.getTime() - 7 * 24 * 60 * 60 * 1000);
    const weekSales = salesData.filter(sale => 
      sale.order_date >= weekAgo && sale.order_date < todayStart
    );

    const todayRevenue = ArrayUtils.sum(todaySales, sale => sale.total_amount);
    const todayProfit = ArrayUtils.sum(todaySales, sale => sale.gross_profit);
    const todayQuantity = ArrayUtils.sum(todaySales, sale => sale.quantity);

    const weeklyAverageRevenue = weekSales.length > 0 ? ArrayUtils.sum(weekSales, sale => sale.total_amount) / 7 : 0;
    const weeklyAverageProfit = weekSales.length > 0 ? ArrayUtils.sum(weekSales, sale => sale.gross_profit) / 7 : 0;

    // 目標売上（月利80万円を30日で割った値）
    const dailyRevenueTarget = 800000 / 30 * (100 / 25); // 利益率25%前提

    return {
      date: today,
      
      // 今日の実績
      todayRevenue: todayRevenue,
      todayProfit: todayProfit,
      todayQuantity: todayQuantity,
      todayOrderCount: todaySales.length,
      
      // 平均比較
      weeklyAverageRevenue: weeklyAverageRevenue,
      weeklyAverageProfit: weeklyAverageProfit,
      revenueGrowthRate: weeklyAverageRevenue > 0 ? NumberUtils.percentage(todayRevenue - weeklyAverageRevenue, weeklyAverageRevenue) : 0,
      
      // 目標比較
      dailyRevenueTarget: dailyRevenueTarget,
      dailyRevenueAchievement: NumberUtils.percentage(todayRevenue, dailyRevenueTarget),
      
      calculatedAt: new Date()
    };
  }

  // =============================================================================
  // 商品別KPI計算
  // =============================================================================

  /**
   * 商品別KPI計算
   */
  calculateProductKPIs(salesData, inventoryData) {
    const productMetrics = {};
    
    // 販売データを商品別に集計
    salesData.forEach(sale => {
      const sku = sale.unified_sku;
      
      if (!productMetrics[sku]) {
        productMetrics[sku] = {
          unified_sku: sku,
          asin: sale.asin,
          product_name: sale.product_name,
          totalRevenue: 0,
          totalProfit: 0,
          totalQuantity: 0,
          salesCount: 0,
          firstSaleDate: sale.order_date,
          lastSaleDate: sale.order_date,
          averageSellingPrice: 0,
          profitMargin: 0
        };
      }
      
      const metrics = productMetrics[sku];
      metrics.totalRevenue += sale.total_amount;
      metrics.totalProfit += sale.gross_profit;
      metrics.totalQuantity += sale.quantity;
      metrics.salesCount += 1;
      
      if (sale.order_date < metrics.firstSaleDate) {
        metrics.firstSaleDate = sale.order_date;
      }
      if (sale.order_date > metrics.lastSaleDate) {
        metrics.lastSaleDate = sale.order_date;
      }
    });

    // 在庫情報をマージして完成
    Object.keys(productMetrics).forEach(sku => {
      const metrics = productMetrics[sku];
      const inventory = inventoryData.find(inv => inv.unified_sku === sku);
      
      // 計算KPI
      metrics.averageSellingPrice = metrics.totalQuantity > 0 ? metrics.totalRevenue / metrics.totalQuantity : 0;
      metrics.profitMargin = metrics.totalRevenue > 0 ? NumberUtils.percentage(metrics.totalProfit, metrics.totalRevenue) : 0;
      
      // 販売ベロシティ
      if (metrics.firstSaleDate && metrics.lastSaleDate) {
        const salesDays = DateUtils.daysBetween(metrics.firstSaleDate, metrics.lastSaleDate) || 1;
        metrics.dailySalesVelocity = metrics.totalQuantity / salesDays;
      } else {
        metrics.dailySalesVelocity = 0;
      }
      
      // 在庫情報
      if (inventory) {
        metrics.currentStock = inventory.quantity;
        metrics.stockValue = inventory.total_cost;
        metrics.daysInStock = inventory.days_in_stock;
        
        // 在庫切れ予測
        if (metrics.dailySalesVelocity > 0) {
          metrics.stockoutDays = Math.floor(inventory.quantity / metrics.dailySalesVelocity);
        } else {
          metrics.stockoutDays = Infinity;
        }
        
        // 在庫回転
        if (inventory.total_cost > 0) {
          metrics.inventoryTurnover = metrics.totalRevenue / inventory.total_cost;
        } else {
          metrics.inventoryTurnover = 0;
        }
      } else {
        metrics.currentStock = 0;
        metrics.stockValue = 0;
        metrics.daysInStock = 0;
        metrics.stockoutDays = 0;
        metrics.inventoryTurnover = 0;
      }
    });

    return Object.values(productMetrics);
  }

  // =============================================================================
  // 特殊計算
  // =============================================================================

  /**
   * 平均在庫価値計算
   */
  calculateAverageInventoryValue(inventoryData) {
    // 簡易版：現在の在庫価値を返す
    // 実際は過去数ヶ月の平均を計算すべき
    return ArrayUtils.sum(inventoryData, inv => inv.total_cost);
  }

  /**
   * 滞留在庫計算
   */
  calculateStagnantInventory(inventoryData) {
    const stagnantThreshold = this.config.getKPISettings().stagnantDaysThreshold;
    
    const stagnantItems = inventoryData.filter(inv => 
      inv.days_in_stock > stagnantThreshold
    );
    
    const stagnantValue = ArrayUtils.sum(stagnantItems, inv => inv.total_cost);
    const totalValue = ArrayUtils.sum(inventoryData, inv => inv.total_cost);
    
    return {
      count: stagnantItems.length,
      value: stagnantValue,
      rate: totalValue > 0 ? NumberUtils.percentage(stagnantValue, totalValue) : 0
    };
  }

  /**
   * 低在庫商品計算
   */
  calculateLowStockItems(inventoryData) {
    const lowStockThreshold = this.config.getKPISettings().lowStockThreshold;
    
    const lowStockItems = inventoryData.filter(inv => 
      inv.quantity > 0 && inv.quantity <= lowStockThreshold
    );
    
    return {
      count: lowStockItems.length,
      items: lowStockItems.map(inv => ({
        sku: inv.unified_sku,
        product_name: inv.product_name,
        quantity: inv.quantity
      }))
    };
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
      let dashboardSheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
      
      if (!dashboardSheet) {
        dashboardSheet = this.createKPIDashboard(ss);
      }

      // 月次KPIの更新
      this.updateMonthlySection(dashboardSheet, monthlyKPIs);
      
      // 日次KPIの更新
      this.updateDailySection(dashboardSheet, dailyKPIs);
      
      // 目標との比較グラフ更新
      this.updateTargetComparison(dashboardSheet, monthlyKPIs);
      
      // 最終更新時刻
      dashboardSheet.getRange('A1').setValue(`最終更新: ${new Date().toLocaleString('ja-JP')}`);

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.updateKPIDashboard');
    }
  }

  /**
   * KPIダッシュボードシート作成
   */
  createKPIDashboard(spreadsheet) {
    const sheet = spreadsheet.insertSheet(SHEET_CONFIG.KPI_MONTHLY);
    
    // 基本レイアウト設定
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 100);
    
    // ヘッダー設定
    sheet.getRange('A1').setValue('KPI管理ダッシュボード');
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    
    // 月次KPIセクション
    sheet.getRange('A3').setValue('月次KPI');
    sheet.getRange('A3').setFontWeight('bold').setBackground('#e1f5fe');
    
    const monthlyLabels = [
      '売上高', '粗利益', '利益率', 'ROI', '目標達成率',
      '販売数', '注文数', '取扱ASIN数', '平均注文額',
      '在庫金額', '在庫回転率', '回転日数', '滞留在庫率'
    ];
    
    monthlyLabels.forEach((label, index) => {
      sheet.getRange(4 + index, 1).setValue(label);
    });
    
    // 日次KPIセクション
    sheet.getRange('A18').setValue('本日の実績');
    sheet.getRange('A18').setFontWeight('bold').setBackground('#fff3e0');
    
    const dailyLabels = [
      '本日売上', '本日利益', '本日販売数', '本日注文数',
      '7日平均売上', '売上成長率', '目標達成率'
    ];
    
    dailyLabels.forEach((label, index) => {
      sheet.getRange(19 + index, 1).setValue(label);
    });
    
    return sheet;
  }

  /**
   * 月次セクション更新
   */
  updateMonthlySection(sheet, kpis) {
    const updates = [
      { cell: 'B4', value: kpis.totalRevenue, format: '¥#,##0' },
      { cell: 'B5', value: kpis.totalGrossProfit, format: '¥#,##0' },
      { cell: 'B6', value: kpis.profitMargin / 100, format: '0.0%' },
      { cell: 'B7', value: kpis.roi / 100, format: '0.0%' },
      { cell: 'B8', value: kpis.profitGoalAchievement / 100, format: '0.0%' },
      { cell: 'B9', value: kpis.totalQuantity, format: '#,##0' },
      { cell: 'B10', value: kpis.salesCount, format: '#,##0' },
      { cell: 'B11', value: kpis.uniqueASINs, format: '#,##0' },
      { cell: 'B12', value: kpis.averageOrderValue, format: '¥#,##0' },
      { cell: 'B13', value: kpis.totalInventoryValue, format: '¥#,##0' },
      { cell: 'B14', value: kpis.inventoryTurnover, format: '0.0' },
      { cell: 'B15', value: kpis.turnoverDays, format: '0' },
      { cell: 'B16', value: kpis.stagnantInventoryRate / 100, format: '0.0%' }
    ];
    
    updates.forEach(update => {
      const range = sheet.getRange(update.cell);
      range.setValue(update.value);
      range.setNumberFormat(update.format);
    });
  }

  /**
   * 日次セクション更新
   */
  updateDailySection(sheet, kpis) {
    const updates = [
      { cell: 'B19', value: kpis.todayRevenue, format: '¥#,##0' },
      { cell: 'B20', value: kpis.todayProfit, format: '¥#,##0' },
      { cell: 'B21', value: kpis.todayQuantity, format: '#,##0' },
      { cell: 'B22', value: kpis.todayOrderCount, format: '#,##0' },
      { cell: 'B23', value: kpis.weeklyAverageRevenue, format: '¥#,##0' },
      { cell: 'B24', value: kpis.revenueGrowthRate / 100, format: '0.0%' },
      { cell: 'B25', value: kpis.dailyRevenueAchievement / 100, format: '0.0%' }
    ];
    
    updates.forEach(update => {
      const range = sheet.getRange(update.cell);
      range.setValue(update.value);
      range.setNumberFormat(update.format);
    });
  }

  /**
   * 目標比較更新
   */
  updateTargetComparison(sheet, kpis) {
    // 目標値との比較表示用の条件付きフォーマット
    const targetRanges = [
      { range: 'B8', threshold: 1.0 }, // 利益目標達成率
      { range: 'B6', threshold: 0.25 }, // 利益率
      { range: 'B7', threshold: 0.30 }, // ROI
      { range: 'B16', threshold: 0.10 } // 滞留在庫率（逆）
    ];
    
    targetRanges.forEach(config => {
      const range = sheet.getRange(config.range);
      const value = range.getValue();
      
      if (config.range === 'B16') {
        // 滞留在庫率は低い方が良い
        if (value <= config.threshold) {
          range.setBackground('#c8e6c9'); // 緑
        } else {
          range.setBackground('#ffcdd2'); // 赤
        }
      } else {
        // その他は高い方が良い
        if (value >= config.threshold) {
          range.setBackground('#c8e6c9'); // 緑
        } else {
          range.setBackground('#ffcdd2'); // 赤
        }
      }
    });
  }

  // =============================================================================
  // 在庫分析更新
  // =============================================================================

  /**
   * 在庫分析シート更新
   */
  updateInventoryAnalysis(inventoryData, productKPIs) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let inventorySheet = ss.getSheetByName(SHEET_CONFIG.INVENTORY);
      
      if (!inventorySheet) {
        inventorySheet = this.createInventoryAnalysisSheet(ss);
      }

      // 商品別データマージ
      const analysisData = this.mergeInventoryWithKPIs(inventoryData, productKPIs);
      
      // 在庫アラート更新
      this.updateInventoryAlerts(inventorySheet, analysisData);
      
      // トップ/ワースト商品更新
      this.updateTopWorstProducts(inventorySheet, analysisData);

    } catch (error) {
      ErrorHandler.handleError(error, 'KPICalculator.updateInventoryAnalysis');
    }
  }

  // =============================================================================
  // アラートチェック
  // =============================================================================

  /**
   * KPIアラートチェック
   */
  checkKPIAlerts(monthlyKPIs, dailyKPIs) {
    const alerts = [];
    const kpiSettings = this.config.getKPISettings();

    // 利益率アラート
    if (monthlyKPIs.profitMargin < kpiSettings.targetProfitMargin) {
      alerts.push({
        type: 'LOW_PROFIT_MARGIN',
        severity: 'WARNING',
        message: `利益率が目標を下回っています: ${monthlyKPIs.profitMargin.toFixed(1)}% (目標: ${kpiSettings.targetProfitMargin}%)`,
        value: monthlyKPIs.profitMargin,
        target: kpiSettings.targetProfitMargin
      });
    }

    // ROIアラート
    if (monthlyKPIs.roi < kpiSettings.targetROI) {
      alerts.push({
        type: 'LOW_ROI',
        severity: 'WARNING',
        message: `ROIが目標を下回っています: ${monthlyKPIs.roi.toFixed(1)}% (目標: ${kpiSettings.targetROI}%)`,
        value: monthlyKPIs.roi,
        target: kpiSettings.targetROI
      });
    }

    // 在庫過多アラート
    if (monthlyKPIs.totalInventoryValue > kpiSettings.maxInventoryValue) {
      alerts.push({
        type: 'EXCESS_INVENTORY',
        severity: 'WARNING',
        message: `在庫金額が上限を超えています: ${NumberUtils.formatCurrency(monthlyKPIs.totalInventoryValue)}`,
        value: monthlyKPIs.totalInventoryValue,
        target: kpiSettings.maxInventoryValue
      });
    }

    // 滞留在庫アラート
    if (monthlyKPIs.stagnantInventoryRate > 15) {
      alerts.push({
        type: 'STAGNANT_INVENTORY',
        severity: 'CRITICAL',
        message: `滞留在庫率が高すぎます: ${monthlyKPIs.stagnantInventoryRate.toFixed(1)}%`,
        value: monthlyKPIs.stagnantInventoryRate,
        target: 10
      });
    }

    // 目標未達アラート
    if (monthlyKPIs.profitGoalAchievement < 80) {
      alerts.push({
        type: 'PROFIT_GOAL_BEHIND',
        severity: 'HIGH',
        message: `月利目標の進捗が遅れています: ${monthlyKPIs.profitGoalAchievement.toFixed(1)}%`,
        value: monthlyKPIs.profitGoalAchievement,
        target: 100
      });
    }

    return alerts;
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