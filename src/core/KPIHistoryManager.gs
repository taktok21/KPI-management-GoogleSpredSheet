/**
 * KPI履歴管理クラス
 * 
 * 過去のKPIデータを管理し、前月比・前年同月比などの比較分析を提供します。
 */

class KPIHistoryManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.cache = new CacheUtils();
  }

  // =============================================================================
  // KPI履歴の保存
  // =============================================================================

  /**
   * 月次KPIを履歴に保存
   * @param {string} yearMonth - 年月（YYYY-MM形式）
   * @param {Object} kpiData - KPIデータ
   */
  saveMonthlyKPI(yearMonth, kpiData) {
    try {
      const sheet = this.getOrCreateHistorySheet();
      const existingRow = this.findExistingRow(sheet, yearMonth);
      
      const recordData = [
        yearMonth,
        kpiData.revenue || 0,
        kpiData.grossProfit || 0,
        kpiData.profitMargin || 0,
        kpiData.roi || 0,
        kpiData.salesQuantity || 0,
        kpiData.inventoryValue || 0,
        kpiData.inventoryTurnover || 0,
        kpiData.stagnantInventoryRate || 0,
        kpiData.uniqueProducts || 0,
        kpiData.averageOrderValue || 0,
        kpiData.profitGoalAchievement || 0,
        kpiData.calculatedAt || new Date(),
        new Date() // created_at
      ];
      
      if (existingRow) {
        // 既存データを更新（created_at以外）
        const range = sheet.getRange(existingRow, 1, 1, recordData.length - 1);
        range.setValues([recordData.slice(0, -1)]);
      } else {
        // 新規データを追加
        sheet.appendRow(recordData);
      }
      
      // キャッシュをクリア
      this.cache.clear(`kpi_history_*`);
      
      console.log(`KPI履歴保存完了: ${yearMonth}`);
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.saveMonthlyKPI');
      throw error;
    }
  }

  // =============================================================================
  // KPI履歴の取得
  // =============================================================================

  /**
   * 過去のKPIデータを取得
   * @param {number} months - 取得する月数（デフォルト: 12ヶ月）
   * @returns {Array} KPIデータの配列
   */
  getHistoricalKPIs(months = 12) {
    const cacheKey = `kpi_history_${months}`;
    const cached = this.cache.get(cacheKey);
    
    if (cached) {
      return cached;
    }
    
    try {
      const sheet = this.getHistorySheet();
      if (!sheet) {
        return [];
      }
      
      const currentMonth = DateUtils.getCurrentMonth();
      const targetMonths = this.generateMonthList(currentMonth, months);
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return [];
      }
      
      // 全データを取得
      const allData = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
      const historicalData = [];
      
      // 対象月のデータを抽出
      targetMonths.forEach(month => {
        const monthData = allData.find(row => row[0] === month);
        if (monthData) {
          historicalData.push(this.mapHistoryRecord(monthData));
        } else {
          // データがない月はnullを追加
          historicalData.push({
            yearMonth: month,
            hasData: false
          });
        }
      });
      
      // キャッシュに保存（5分間）
      this.cache.set(cacheKey, historicalData, 300);
      
      return historicalData;
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.getHistoricalKPIs');
      return [];
    }
  }

  /**
   * 特定月のKPIデータを取得
   * @param {string} yearMonth - 年月（YYYY-MM形式）
   * @returns {Object|null} KPIデータ
   */
  getMonthKPI(yearMonth) {
    try {
      const sheet = this.getHistorySheet();
      if (!sheet) {
        return null;
      }
      
      const row = this.findExistingRow(sheet, yearMonth);
      if (!row) {
        return null;
      }
      
      const data = sheet.getRange(row, 1, 1, 14).getValues()[0];
      return this.mapHistoryRecord(data);
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.getMonthKPI');
      return null;
    }
  }

  // =============================================================================
  // 比較計算
  // =============================================================================

  /**
   * 前月比を計算
   * @param {Object} currentKPI - 当月のKPI
   * @param {Object} previousKPI - 前月のKPI
   * @returns {Object} 前月比データ
   */
  calculateMonthOverMonth(currentKPI, previousKPI) {
    if (!previousKPI || !previousKPI.hasData) {
      return null;
    }
    
    return {
      revenue: this.calculateGrowthRate(currentKPI.revenue, previousKPI.revenue),
      grossProfit: this.calculateGrowthRate(currentKPI.grossProfit, previousKPI.grossProfit),
      profitMargin: currentKPI.profitMargin - previousKPI.profitMargin,
      roi: currentKPI.roi - previousKPI.roi,
      salesQuantity: this.calculateGrowthRate(currentKPI.salesQuantity, previousKPI.salesQuantity),
      inventoryTurnover: currentKPI.inventoryTurnover - previousKPI.inventoryTurnover
    };
  }

  /**
   * 前年同月比を計算
   * @param {Object} currentKPI - 当月のKPI
   * @param {Object} yearAgoKPI - 前年同月のKPI
   * @returns {Object} 前年同月比データ
   */
  calculateYearOverYear(currentKPI, yearAgoKPI) {
    if (!yearAgoKPI || !yearAgoKPI.hasData) {
      return null;
    }
    
    return {
      revenue: this.calculateGrowthRate(currentKPI.revenue, yearAgoKPI.revenue),
      grossProfit: this.calculateGrowthRate(currentKPI.grossProfit, yearAgoKPI.grossProfit),
      profitMargin: currentKPI.profitMargin - yearAgoKPI.profitMargin,
      roi: currentKPI.roi - yearAgoKPI.roi,
      salesQuantity: this.calculateGrowthRate(currentKPI.salesQuantity, yearAgoKPI.salesQuantity)
    };
  }

  /**
   * 移動平均を計算
   * @param {Array} historicalKPIs - 過去のKPIデータ
   * @param {string} metric - 指標名
   * @param {number} period - 期間（デフォルト: 3ヶ月）
   * @returns {Array} 移動平均データ
   */
  calculateMovingAverage(historicalKPIs, metric, period = 3) {
    const movingAverages = [];
    
    for (let i = 0; i < historicalKPIs.length; i++) {
      if (i < period - 1) {
        movingAverages.push(null);
        continue;
      }
      
      let sum = 0;
      let count = 0;
      
      for (let j = 0; j < period; j++) {
        const kpi = historicalKPIs[i - j];
        if (kpi && kpi.hasData !== false && kpi[metric] !== null) {
          sum += kpi[metric];
          count++;
        }
      }
      
      movingAverages.push(count > 0 ? sum / count : null);
    }
    
    return movingAverages;
  }

  // =============================================================================
  // ヘルパーメソッド
  // =============================================================================

  /**
   * 履歴シートを取得または作成
   */
  getOrCreateHistorySheet() {
    let sheet = this.ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    
    if (!sheet) {
      sheet = this.ss.insertSheet(SHEET_CONFIG.KPI_HISTORY);
      this.setupHistorySheet(sheet);
    }
    
    return sheet;
  }

  /**
   * 履歴シートを取得
   */
  getHistorySheet() {
    return this.ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
  }

  /**
   * 履歴シートの初期設定
   */
  setupHistorySheet(sheet) {
    const headers = [
      '年月',
      '売上高',
      '粗利益',
      '利益率(%)',
      'ROI(%)',
      '販売数',
      '在庫金額',
      '在庫回転率',
      '滞留在庫率(%)',
      '取扱商品数',
      '平均注文額',
      '利益目標達成率(%)',
      '計算日時',
      '作成日時'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // 列幅の設定
    sheet.setColumnWidth(1, 80);  // 年月
    sheet.setColumnWidth(2, 100); // 売上高
    sheet.setColumnWidth(3, 100); // 粗利益
    
    // 数値フォーマット
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setNumberFormat('¥#,##0');      // 売上高
    sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setNumberFormat('¥#,##0');      // 粗利益
    sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setNumberFormat('0.0%');        // 利益率
    sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setNumberFormat('0.0%');        // ROI
    sheet.getRange(2, 7, sheet.getMaxRows() - 1, 1).setNumberFormat('¥#,##0');      // 在庫金額
    sheet.getRange(2, 8, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00');        // 在庫回転率
    sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1).setNumberFormat('0.0%');        // 滞留在庫率
    sheet.getRange(2, 11, sheet.getMaxRows() - 1, 1).setNumberFormat('¥#,##0');     // 平均注文額
    sheet.getRange(2, 12, sheet.getMaxRows() - 1, 1).setNumberFormat('0.0%');       // 利益目標達成率
  }

  /**
   * 既存行を検索
   */
  findExistingRow(sheet, yearMonth) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return null;
    }
    
    const yearMonthColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < yearMonthColumn.length; i++) {
      if (yearMonthColumn[i][0] === yearMonth) {
        return i + 2; // 行番号は1から始まり、ヘッダーが1行目
      }
    }
    
    return null;
  }

  /**
   * 対象月リストを生成
   */
  generateMonthList(currentMonth, months) {
    const monthList = [];
    const [year, month] = currentMonth.split('-').map(Number);
    
    for (let i = months - 1; i >= 0; i--) {
      const targetDate = new Date(year, month - 1 - i, 1);
      const yearMonth = DateUtils.formatDate(targetDate, 'yyyy-MM');
      monthList.push(yearMonth);
    }
    
    return monthList;
  }

  /**
   * 履歴レコードをマッピング
   */
  mapHistoryRecord(row) {
    return {
      yearMonth: row[0],
      revenue: row[1],
      grossProfit: row[2],
      profitMargin: row[3],
      roi: row[4],
      salesQuantity: row[5],
      inventoryValue: row[6],
      inventoryTurnover: row[7],
      stagnantInventoryRate: row[8],
      uniqueProducts: row[9],
      averageOrderValue: row[10],
      profitGoalAchievement: row[11],
      calculatedAt: row[12],
      createdAt: row[13],
      hasData: true
    };
  }

  /**
   * 成長率を計算
   */
  calculateGrowthRate(current, previous) {
    if (!previous || previous === 0) {
      return null;
    }
    
    return ((current - previous) / previous) * 100;
  }

  /**
   * 時系列データの効率的な取得
   * @param {string} startMonth - 開始月（YYYY-MM形式）
   * @param {string} endMonth - 終了月（YYYY-MM形式）
   * @param {Array} metrics - 取得するメトリクス配列
   * @returns {Array} 時系列データ
   */
  getTimeSeriesData(startMonth, endMonth, metrics = ['revenue', 'grossProfit']) {
    try {
      const sheet = this.getHistorySheet();
      if (!sheet) {
        return [];
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return [];
      }
      
      const data = sheet.getRange(1, 1, lastRow, 14).getValues();
      const headers = data[0];
      
      // インデックスマッピング作成
      const indexMap = {};
      headers.forEach((header, index) => {
        switch(header) {
          case '年月': indexMap['yearMonth'] = index; break;
          case '売上高': indexMap['revenue'] = index; break;
          case '粗利益': indexMap['grossProfit'] = index; break;
          case '利益率(%)': indexMap['profitMargin'] = index; break;
          case 'ROI(%)': indexMap['roi'] = index; break;
          case '販売数': indexMap['salesQuantity'] = index; break;
          case '在庫金額': indexMap['inventoryValue'] = index; break;
          case '在庫回転率': indexMap['inventoryTurnover'] = index; break;
          case '滞留在庫率(%)': indexMap['stagnantInventoryRate'] = index; break;
          case '取扱商品数': indexMap['uniqueProducts'] = index; break;
          case '平均注文額': indexMap['averageOrderValue'] = index; break;
          case '利益目標達成率(%)': indexMap['profitGoalAchievement'] = index; break;
          case '計算日時': indexMap['calculatedAt'] = index; break;
          case '作成日時': indexMap['createdAt'] = index; break;
        }
      });
      
      // 期間フィルタリング
      const filteredData = data.slice(1).filter(row => {
        const yearMonth = row[indexMap['yearMonth']];
        return yearMonth >= startMonth && yearMonth <= endMonth;
      });
      
      // メトリクス抽出
      return filteredData.map(row => {
        const result = { yearMonth: row[indexMap['yearMonth']] };
        metrics.forEach(metric => {
          if (indexMap[metric] !== undefined) {
            result[metric] = row[indexMap[metric]] || 0;
          }
        });
        return result;
      }).sort((a, b) => a.yearMonth.localeCompare(b.yearMonth));
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.getTimeSeriesData');
      return [];
    }
  }

  /**
   * キャッシュ機能付きデータ取得
   * @param {string} cacheKey - キャッシュキー
   * @param {Function} dataFunction - データ取得関数
   * @param {number} expireMinutes - キャッシュ有効期限（分）
   * @returns {*} データ
   */
  getCachedHistoricalData(cacheKey, dataFunction, expireMinutes = 30) {
    try {
      const cached = this.cache.get(cacheKey);
      if (cached) {
        return cached;
      }
      
      const data = dataFunction();
      this.cache.set(cacheKey, data, expireMinutes * 60);
      return data;
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.getCachedHistoricalData');
      return dataFunction();
    }
  }

  /**
   * 前年同月のKPIデータを取得
   * @param {string} currentMonth - 現在月（YYYY-MM形式）
   * @returns {Object|null} 前年同月のKPIデータ
   */
  getPreviousYearKPI(currentMonth) {
    try {
      const [year, month] = currentMonth.split('-');
      const previousYear = (parseInt(year) - 1).toString();
      const previousYearMonth = `${previousYear}-${month}`;
      
      return this.getMonthKPI(previousYearMonth);
      
    } catch (error) {
      ErrorHandler.handleError(error, 'KPIHistoryManager.getPreviousYearKPI');
      return null;
    }
  }

  /**
   * 指定月から過去N年分のデータを取得
   * @param {string} baseMonth - 基準月（YYYY-MM形式）
   * @param {number} years - 取得年数
   * @returns {Array} 過去N年分のデータ
   */
  getMultiYearData(baseMonth, years = 3) {
    const cacheKey = `multi_year_${baseMonth}_${years}`;
    
    return this.getCachedHistoricalData(cacheKey, () => {
      const data = [];
      const [baseYear, month] = baseMonth.split('-');
      
      for (let i = 0; i < years; i++) {
        const targetYear = (parseInt(baseYear) - i).toString();
        const targetMonth = `${targetYear}-${month}`;
        const kpiData = this.getMonthKPI(targetMonth);
        
        if (kpiData) {
          data.push({
            ...kpiData,
            yearOffset: i,
            label: i === 0 ? '当年' : `${i}年前`
          });
        }
      }
      
      return data;
    }, 60); // 1時間キャッシュ
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * KPI履歴を手動で保存
 */
function saveCurrentMonthKPI() {
  try {
    const calculator = new KPICalculator();
    const historyManager = new KPIHistoryManager();
    
    // 現在のKPIを計算
    const result = calculator.recalculateAll();
    const currentMonth = DateUtils.getCurrentMonth();
    
    // 履歴に保存
    historyManager.saveMonthlyKPI(currentMonth, result.monthlyKPIs);
    
    SpreadsheetApp.getUi().alert('完了', `${currentMonth}のKPIを履歴に保存しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    ErrorHandler.handleError(error, 'saveCurrentMonthKPI');
    SpreadsheetApp.getUi().alert('エラー', 'KPI履歴の保存中にエラーが発生しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}