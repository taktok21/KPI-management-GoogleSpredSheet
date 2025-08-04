/**
 * バッチ処理クラス
 * 
 * 日次バッチ処理、データ同期、KPI計算の統合処理を行います。
 */

class BatchProcessor {
  constructor() {
    this.config = new ConfigManager();
    this.cache = new CacheUtils();
    this.startTime = new Date();
    this.processLog = [];
  }

  // =============================================================================
  // メインバッチ処理
  // =============================================================================

  /**
   * 日次バッチ処理実行
   */
  runDailyBatch() {
    return ErrorHandler.executeWithRetry(() => {
      return this.executeBatchProcess();
    }, 'BatchProcessor.runDailyBatch', {
      maxRetries: 2,
      retryDelay: 5000
    });
  }

  /**
   * バッチ処理の実行
   */
  executeBatchProcess() {
    try {
      this.startTime = new Date();
      this.processLog = [];
      
      this.log('INFO', 'バッチ処理を開始しました');
      
      // 1. 前処理
      this.preProcess();
      
      // 2. データ取得・更新
      const updateResults = this.updateAllData();
      
      // 3. KPI計算
      const kpiResults = this.calculateKPIs();
      
      // 4. アラートチェック
      const alerts = this.checkAlerts(kpiResults);
      
      // 5. 後処理
      this.postProcess(updateResults, kpiResults, alerts);
      
      const duration = (new Date() - this.startTime) / 1000;
      this.log('INFO', `バッチ処理が完了しました (処理時間: ${duration}秒)`);
      
      return {
        success: true,
        duration: duration,
        updateResults: updateResults,
        kpiResults: kpiResults,
        alerts: alerts.length,
        processLog: this.processLog,
        executedAt: this.startTime
      };

    } catch (error) {
      this.log('ERROR', `バッチ処理でエラーが発生しました: ${error.message}`);
      this.saveProcessLog('ERROR');
      throw error;
    }
  }

  // =============================================================================
  // 前処理・後処理
  // =============================================================================

  /**
   * バッチ処理前処理
   */
  preProcess() {
    try {
      this.log('INFO', '前処理を開始');
      
      // キャッシュクリア
      this.cache.clear();
      this.log('INFO', 'キャッシュをクリアしました');
      
      // 設定検証
      const validation = this.config.validateConfiguration();
      if (!validation.isValid) {
        throw new Error(`設定エラー: ${validation.errors.join(', ')}`);
      }
      this.log('INFO', '設定検証が完了しました');
      
      // 前回処理状況確認
      this.checkPreviousExecution();
      
      this.log('INFO', '前処理が完了しました');

    } catch (error) {
      this.log('ERROR', `前処理でエラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * バッチ処理後処理
   */
  postProcess(updateResults, kpiResults, alerts) {
    try {
      this.log('INFO', '後処理を開始');
      
      // 処理ログ保存
      this.saveProcessLog('SUCCESS');
      
      // アラート通知
      if (alerts.length > 0) {
        this.sendAlertNotifications(alerts);
        this.log('INFO', `${alerts.length}件のアラートを送信しました`);
      }
      
      // 統計情報更新
      this.updateProcessingStats(updateResults, kpiResults);
      
      // 次回実行時刻設定
      this.scheduleNextExecution();
      
      this.log('INFO', '後処理が完了しました');

    } catch (error) {
      this.log('WARN', `後処理でエラー: ${error.message}`);
    }
  }

  // =============================================================================
  // データ更新処理
  // =============================================================================

  /**
   * 全データ更新
   */
  updateAllData() {
    const results = {
      amazon: { success: false, recordCount: 0 },
      makado: { success: false, recordCount: 0 },
      inventory: { success: false, recordCount: 0 }
    };

    try {
      this.log('INFO', 'データ更新を開始');

      // Amazon SP-APIデータ更新
      try {
        this.log('INFO', 'Amazon SP-APIデータ取得中...');
        results.amazon = this.updateAmazonData();
        this.log('INFO', `Amazonデータ: ${results.amazon.recordCount}件取得`);
      } catch (error) {
        this.log('ERROR', `Amazonデータ取得エラー: ${error.message}`);
        results.amazon.error = error.message;
      }

      // マカドCSVデータ確認・取得
      try {
        this.log('INFO', 'マカドデータ確認中...');
        results.makado = this.updateMakadoData();
        if (results.makado.recordCount > 0) {
          this.log('INFO', `マカドデータ: ${results.makado.recordCount}件取得`);
        } else {
          this.log('INFO', 'マカドデータ: 新しいファイルはありません');
        }
      } catch (error) {
        this.log('ERROR', `マカドデータ取得エラー: ${error.message}`);
        results.makado.error = error.message;
      }

      // 在庫データ更新
      try {
        this.log('INFO', '在庫データ更新中...');
        results.inventory = this.updateInventoryData();
        this.log('INFO', `在庫データ: ${results.inventory.recordCount}件更新`);
      } catch (error) {
        this.log('ERROR', `在庫データ更新エラー: ${error.message}`);
        results.inventory.error = error.message;
      }

      this.log('INFO', 'データ更新が完了しました');
      return results;

    } catch (error) {
      this.log('ERROR', `データ更新で致命的エラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * Amazonデータ更新
   */
  updateAmazonData() {
    // Amazon SP-API連携は後で実装
    // 現在は仮実装
    this.log('INFO', 'Amazon SP-API連携は開発中です');
    
    return {
      success: true,
      recordCount: 0,
      source: 'Amazon SP-API',
      message: '開発中のため実際のデータ取得はスキップしました'
    };
  }

  /**
   * マカドデータ更新
   */
  updateMakadoData() {
    try {
      // 最新のマカドCSVファイルを検索
      const latestFile = this.findLatestMakadoCSV();
      
      if (!latestFile) {
        return {
          success: true,
          recordCount: 0,
          source: 'Makado CSV',
          message: '新しいCSVファイルが見つかりませんでした'
        };
      }

      // ファイルが既に処理済みかチェック
      if (this.isFileAlreadyProcessed(latestFile)) {
        return {
          success: true,
          recordCount: 0,
          source: 'Makado CSV',
          message: 'ファイルは既に処理済みです'
        };
      }

      // CSVファイル処理
      const processor = new MakadoProcessor();
      const result = processor.importCSV(latestFile.getName());
      
      // 処理済みファイルとしてマーク
      this.markFileAsProcessed(latestFile);

      return {
        success: result.success,
        recordCount: result.recordCount,
        source: 'Makado CSV',
        fileName: latestFile.getName(),
        message: `CSVファイルを正常に処理しました`
      };

    } catch (error) {
      return {
        success: false,
        recordCount: 0,
        source: 'Makado CSV',
        error: error.message
      };
    }
  }

  /**
   * 在庫データ更新
   */
  updateInventoryData() {
    try {
      // 販売データから在庫情報を更新
      const salesData = this.getSalesDataForInventoryUpdate();
      const inventoryUpdates = this.calculateInventoryUpdates(salesData);
      
      // 在庫シート更新
      this.applyInventoryUpdates(inventoryUpdates);
      
      return {
        success: true,
        recordCount: inventoryUpdates.length,
        source: 'Inventory Calculation',
        message: '在庫データを更新しました'
      };

    } catch (error) {
      return {
        success: false,
        recordCount: 0,
        source: 'Inventory Calculation',
        error: error.message
      };
    }
  }

  // =============================================================================
  // KPI計算処理
  // =============================================================================

  /**
   * KPI計算実行
   */
  calculateKPIs() {
    try {
      this.log('INFO', 'KPI計算を開始');

      const calculator = new KPICalculator();
      const result = calculator.recalculateAll();

      this.log('INFO', `KPI計算が完了しました (処理時間: ${result.duration}秒)`);
      
      return result;

    } catch (error) {
      this.log('ERROR', `KPI計算でエラー: ${error.message}`);
      throw error;
    }
  }

  // =============================================================================
  // アラート処理
  // =============================================================================

  /**
   * アラートチェック
   */
  checkAlerts(kpiResults) {
    try {
      this.log('INFO', 'アラートチェックを開始');

      const calculator = new KPICalculator();
      const alerts = calculator.checkKPIAlerts(kpiResults.monthlyKPIs, kpiResults.dailyKPIs);

      // 在庫アラートも追加
      const inventoryAlerts = this.checkInventoryAlerts();
      alerts.push(...inventoryAlerts);

      this.log('INFO', `アラートチェックが完了しました (${alerts.length}件のアラート)`);
      
      return alerts;

    } catch (error) {
      this.log('ERROR', `アラートチェックでエラー: ${error.message}`);
      return [];
    }
  }

  /**
   * 在庫アラートチェック
   */
  checkInventoryAlerts() {
    const alerts = [];

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = ss.getSheetByName(SHEET_CONFIG.INVENTORY);
      
      if (!inventorySheet) return alerts;

      const lastRow = inventorySheet.getLastRow();
      if (lastRow <= 1) return alerts;

      const data = inventorySheet.getRange(2, 1, lastRow - 1, 13).getValues();
      const kpiSettings = this.config.getKPISettings();

      data.forEach(row => {
        const sku = row[0];
        const productName = row[2];
        const quantity = NumberUtils.safeInteger(row[3]);
        const daysInStock = NumberUtils.safeInteger(row[9]);

        // 低在庫アラート
        if (quantity > 0 && quantity <= kpiSettings.lowStockThreshold) {
          alerts.push({
            type: 'LOW_STOCK',
            severity: 'WARNING',
            sku: sku,
            productName: productName,
            message: `在庫が少なくなっています: ${productName} (残り${quantity}個)`,
            value: quantity,
            target: kpiSettings.lowStockThreshold
          });
        }

        // 在庫切れアラート
        if (quantity === 0) {
          alerts.push({
            type: 'OUT_OF_STOCK',
            severity: 'HIGH',
            sku: sku,
            productName: productName,
            message: `在庫切れです: ${productName}`,
            value: 0,
            target: 1
          });
        }

        // 滞留在庫アラート
        if (daysInStock > kpiSettings.stagnantDaysThreshold) {
          alerts.push({
            type: 'STAGNANT_STOCK',
            severity: 'WARNING',
            sku: sku,
            productName: productName,
            message: `滞留在庫です: ${productName} (${daysInStock}日間)`,
            value: daysInStock,
            target: kpiSettings.stagnantDaysThreshold
          });
        }
      });

    } catch (error) {
      this.log('ERROR', `在庫アラートチェックでエラー: ${error.message}`);
    }

    return alerts;
  }

  /**
   * アラート通知送信
   */
  sendAlertNotifications(alerts) {
    try {
      const notificationService = new NotificationService();
      notificationService.sendAlert(alerts);
    } catch (error) {
      this.log('ERROR', `アラート通知送信でエラー: ${error.message}`);
    }
  }

  // =============================================================================
  // ヘルパー関数
  // =============================================================================

  /**
   * 最新のマカドCSVファイル検索
   */
  findLatestMakadoCSV() {
    try {
      const searchQuery = 'title contains "マカド" and title contains ".csv"';
      const files = DriveApp.searchFiles(searchQuery);
      
      let latestFile = null;
      let latestDate = new Date(0);

      while (files.hasNext()) {
        const file = files.next();
        const lastUpdated = file.getLastUpdated();
        
        if (lastUpdated > latestDate) {
          latestDate = lastUpdated;
          latestFile = file;
        }
      }

      return latestFile;

    } catch (error) {
      this.log('ERROR', `マカドCSVファイル検索でエラー: ${error.message}`);
      return null;
    }
  }

  /**
   * ファイル処理済みチェック
   */
  isFileAlreadyProcessed(file) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const processedFiles = properties.getProperty('PROCESSED_FILES');
      
      if (!processedFiles) return false;
      
      const fileList = JSON.parse(processedFiles);
      const fileId = file.getId();
      
      return fileList.includes(fileId);

    } catch (error) {
      return false;
    }
  }

  /**
   * ファイルを処理済みとしてマーク
   */
  markFileAsProcessed(file) {
    try {
      const properties = PropertiesService.getScriptProperties();
      let processedFiles = properties.getProperty('PROCESSED_FILES');
      
      let fileList = [];
      if (processedFiles) {
        fileList = JSON.parse(processedFiles);
      }
      
      fileList.push(file.getId());
      
      // 最新50ファイルのみ保持
      if (fileList.length > 50) {
        fileList = fileList.slice(-50);
      }
      
      properties.setProperty('PROCESSED_FILES', JSON.stringify(fileList));

    } catch (error) {
      this.log('ERROR', `ファイル処理済みマークでエラー: ${error.message}`);
    }
  }

  /**
   * 前回実行状況確認
   */
  checkPreviousExecution() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const lastExecution = properties.getProperty('LAST_BATCH_EXECUTION');
      
      if (lastExecution) {
        const lastDate = new Date(lastExecution);
        const hoursSinceLastRun = (this.startTime - lastDate) / (1000 * 60 * 60);
        
        if (hoursSinceLastRun < 1) {
          this.log('WARN', `前回実行から${hoursSinceLastRun.toFixed(1)}時間しか経過していません`);
        }
        
        this.log('INFO', `前回実行: ${lastDate.toLocaleString('ja-JP')}`);
      }

    } catch (error) {
      this.log('WARN', `前回実行状況確認でエラー: ${error.message}`);
    }
  }

  /**
   * 次回実行時刻設定
   */
  scheduleNextExecution() {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('LAST_BATCH_EXECUTION', this.startTime.toISOString());
      
      const tomorrow = new Date(this.startTime.getTime() + 24 * 60 * 60 * 1000);
      properties.setProperty('NEXT_BATCH_EXECUTION', tomorrow.toISOString());

    } catch (error) {
      this.log('WARN', `次回実行時刻設定でエラー: ${error.message}`);
    }
  }

  /**
   * 処理統計更新
   */
  updateProcessingStats(updateResults, kpiResults) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const today = DateUtils.formatDate(new Date(), 'yyyy-MM-dd');
      
      // 日次統計更新
      const statsKey = `BATCH_STATS_${today}`;
      const stats = {
        executionCount: (parseInt(properties.getProperty(`${statsKey}_COUNT`)) || 0) + 1,
        totalDuration: (parseFloat(properties.getProperty(`${statsKey}_DURATION`)) || 0) + kpiResults.duration,
        totalRecords: (parseInt(properties.getProperty(`${statsKey}_RECORDS`)) || 0) + 
                     updateResults.amazon.recordCount + updateResults.makado.recordCount,
        lastExecution: this.startTime.toISOString()
      };
      
      Object.keys(stats).forEach(key => {
        properties.setProperty(`${statsKey}_${key.toUpperCase()}`, stats[key].toString());
      });

    } catch (error) {
      this.log('WARN', `処理統計更新でエラー: ${error.message}`);
    }
  }

  /**
   * 在庫データ取得（更新用）
   */
  getSalesDataForInventoryUpdate() {
    try {
      const calculator = new KPICalculator();
      
      // 過去7日間の販売データを取得
      const endDate = new Date();
      const startDate = new Date(endDate.getTime() - 7 * 24 * 60 * 60 * 1000);
      
      return calculator.getSalesData(startDate, endDate);

    } catch (error) {
      this.log('ERROR', `在庫更新用販売データ取得でエラー: ${error.message}`);
      return [];
    }
  }

  /**
   * 在庫更新計算
   */
  calculateInventoryUpdates(salesData) {
    const updates = [];

    try {
      // SKU別の販売数集計
      const skuSales = ArrayUtils.groupBy(salesData, sale => sale.unified_sku);
      
      Object.keys(skuSales).forEach(sku => {
        const sales = skuSales[sku];
        const totalQuantity = ArrayUtils.sum(sales, sale => sale.quantity);
        const lastSaleDate = ArrayUtils.max(sales, sale => sale.order_date.getTime());
        
        updates.push({
          sku: sku,
          lastSaleDate: new Date(lastSaleDate),
          recentSalesQuantity: totalQuantity
        });
      });

    } catch (error) {
      this.log('ERROR', `在庫更新計算でエラー: ${error.message}`);
    }

    return updates;
  }

  /**
   * 在庫更新適用
   */
  applyInventoryUpdates(updates) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = ss.getSheetByName(SHEET_CONFIG.INVENTORY);
      
      if (!inventorySheet || updates.length === 0) return;

      const lastRow = inventorySheet.getLastRow();
      if (lastRow <= 1) return;

      const data = inventorySheet.getRange(2, 1, lastRow - 1, 13).getValues();
      
      updates.forEach(update => {
        const rowIndex = data.findIndex(row => row[0] === update.sku);
        if (rowIndex >= 0) {
          // 最終販売日更新
          inventorySheet.getRange(rowIndex + 2, 9).setValue(update.lastSaleDate);
          
          // 在庫日数再計算
          const lastInboundDate = data[rowIndex][7];
          if (lastInboundDate) {
            const daysInStock = DateUtils.daysBetween(lastInboundDate, new Date());
            inventorySheet.getRange(rowIndex + 2, 10).setValue(daysInStock);
          }
          
          // 更新日時
          inventorySheet.getRange(rowIndex + 2, 13).setValue(new Date());
        }
      });

    } catch (error) {
      this.log('ERROR', `在庫更新適用でエラー: ${error.message}`);
    }
  }

  // =============================================================================
  // ログ管理
  // =============================================================================

  /**
   * ログ記録
   */
  log(level, message) {
    const logEntry = {
      timestamp: new Date(),
      level: level,
      message: message
    };
    
    this.processLog.push(logEntry);
    console.log(`[${level}] ${message}`);
  }

  /**
   * 処理ログ保存
   */
  saveProcessLog(status) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName(SHEET_CONFIG.SYNC_LOG);
      
      if (!logSheet) return;

      const duration = (new Date() - this.startTime) / 1000;
      const errorCount = this.processLog.filter(log => log.level === 'ERROR').length;
      const warningCount = this.processLog.filter(log => log.level === 'WARN').length;
      
      const logSummary = this.processLog
        .filter(log => log.level === 'INFO' || log.level === 'ERROR')
        .map(log => `[${log.level}] ${log.message}`)
        .join('\n');

      logSheet.appendRow([
        this.startTime,
        'DAILY_BATCH',
        status,
        duration,
        this.processLog.length,
        errorCount + warningCount,
        `バッチ処理${status === 'SUCCESS' ? '成功' : '失敗'}`,
        logSummary
      ]);

    } catch (error) {
      console.error('処理ログ保存でエラー:', error);
    }
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * 日次バッチ処理実行（グローバル関数）
 */
function runDailyBatch() {
  const processor = new BatchProcessor();
  return processor.runDailyBatch();
}

// =============================================================================
// 通知サービス（簡易版）
// =============================================================================

class NotificationService {
  constructor() {
    this.config = new ConfigManager();
  }

  /**
   * アラート通知送信
   */
  sendAlert(alerts) {
    if (!alerts || alerts.length === 0) return;

    const notificationSettings = this.config.getNotificationSettings();
    
    // メール通知
    if (notificationSettings.emailEnabled && notificationSettings.email) {
      this.sendEmailAlert(alerts, notificationSettings.email);
    }
    
    // Slack通知
    if (notificationSettings.slackEnabled && notificationSettings.slackWebhook) {
      this.sendSlackAlert(alerts, notificationSettings.slackWebhook);
    }
  }

  /**
   * メールアラート送信
   */
  sendEmailAlert(alerts, email) {
    try {
      const subject = `[KPI管理] ${alerts.length}件のアラート - ${new Date().toLocaleDateString('ja-JP')}`;
      
      let body = 'KPI管理システムからのアラート通知です。\n\n';
      
      alerts.forEach((alert, index) => {
        body += `【アラート ${index + 1}】\n`;
        body += `重要度: ${alert.severity}\n`;
        body += `内容: ${alert.message}\n`;
        if (alert.productName) {
          body += `商品: ${alert.productName}\n`;
        }
        body += '\n';
      });
      
      body += `\n詳細確認: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body
      });

    } catch (error) {
      console.error('メールアラート送信エラー:', error);
    }
  }

  /**
   * Slackアラート送信
   */
  sendSlackAlert(alerts, webhookUrl) {
    try {
      const payload = {
        text: `🚨 KPI管理システム アラート通知 (${alerts.length}件)`,
        attachments: alerts.map(alert => ({
          color: this.getSeverityColor(alert.severity),
          title: alert.message,
          fields: [
            {
              title: '重要度',
              value: alert.severity,
              short: true
            },
            {
              title: '商品',
              value: alert.productName || 'N/A',
              short: true
            }
          ]
        }))
      };
      
      UrlFetchApp.fetch(webhookUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });

    } catch (error) {
      console.error('Slackアラート送信エラー:', error);
    }
  }

  /**
   * 重要度カラー取得
   */
  getSeverityColor(severity) {
    const colors = {
      'CRITICAL': '#ff0000',
      'HIGH': '#ff9900',
      'WARNING': '#ffcc00',
      'INFO': '#0099ff'
    };
    
    return colors[severity] || '#cccccc';
  }
}

// =============================================================================
// グローバル関数（トリガーから呼び出し用）
// =============================================================================

/**
 * 日次バッチ処理実行（トリガーから呼び出し）
 */
function runDailyBatch() {
  try {
    const processor = new BatchProcessor();
    return processor.runDailyBatch();
  } catch (error) {
    ErrorHandler.handleError(error, 'runDailyBatch');
    throw error;
  }
}