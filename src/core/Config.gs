/**
 * 設定管理クラス
 * 
 * API認証情報、アプリケーション設定の管理を行います。
 * セキュリティを考慮して、機密情報はProperties Serviceに保存します。
 */

class ConfigManager {
  constructor() {
    this.scriptProperties = PropertiesService.getScriptProperties();
    this.userProperties = PropertiesService.getUserProperties();
  }

  // =============================================================================
  // API認証情報管理
  // =============================================================================

  /**
   * Amazon SP-API認証情報を設定
   */
  setAmazonCredentials(credentials) {
    const requiredFields = ['clientId', 'clientSecret', 'refreshToken'];
    
    for (const field of requiredFields) {
      if (!credentials[field]) {
        throw new Error(`必須フィールドが不足しています: ${field}`);
      }
    }

    this.scriptProperties.setProperties({
      'SP_API_CLIENT_ID': credentials.clientId,
      'SP_API_CLIENT_SECRET': credentials.clientSecret,
      'SP_API_REFRESH_TOKEN': credentials.refreshToken,
      'SP_API_MARKETPLACE_ID': credentials.marketplaceId || 'A1VC38T7YXB528', // Japan
      'SP_API_ENDPOINT': credentials.endpoint || 'https://sellingpartnerapi-fe.amazon.com'
    });

    this.logConfigChange('Amazon SP-API認証情報を更新しました');
  }

  /**
   * Amazon SP-API認証情報を取得
   */
  getAmazonCredentials() {
    return {
      clientId: this.scriptProperties.getProperty('SP_API_CLIENT_ID'),
      clientSecret: this.scriptProperties.getProperty('SP_API_CLIENT_SECRET'),
      refreshToken: this.scriptProperties.getProperty('SP_API_REFRESH_TOKEN'),
      marketplaceId: this.scriptProperties.getProperty('SP_API_MARKETPLACE_ID') || 'A1VC38T7YXB528',
      endpoint: this.scriptProperties.getProperty('SP_API_ENDPOINT') || 'https://sellingpartnerapi-fe.amazon.com'
    };
  }

  /**
   * Keepa API認証情報を設定
   */
  setKeepaCredentials(apiKey) {
    if (!apiKey) {
      throw new Error('Keepa APIキーが指定されていません');
    }

    this.scriptProperties.setProperty('KEEPA_API_KEY', apiKey);
    this.logConfigChange('Keepa API認証情報を更新しました');
  }

  /**
   * Keepa API認証情報を取得
   */
  getKeepaCredentials() {
    return {
      apiKey: this.scriptProperties.getProperty('KEEPA_API_KEY')
    };
  }

  // =============================================================================
  // 通知設定管理
  // =============================================================================

  /**
   * 通知設定を設定
   */
  setNotificationSettings(settings) {
    const properties = {};
    
    if (settings.email) {
      properties.NOTIFICATION_EMAIL = settings.email;
      properties.EMAIL_ENABLED = settings.emailEnabled ? 'true' : 'false';
    }
    
    if (settings.slackWebhook) {
      properties.SLACK_WEBHOOK_URL = settings.slackWebhook;
      properties.SLACK_ENABLED = settings.slackEnabled ? 'true' : 'false';
    }

    this.scriptProperties.setProperties(properties);
    this.logConfigChange('通知設定を更新しました');
  }

  /**
   * 通知設定を取得
   */
  getNotificationSettings() {
    return {
      email: this.scriptProperties.getProperty('NOTIFICATION_EMAIL'),
      emailEnabled: this.scriptProperties.getProperty('EMAIL_ENABLED') === 'true',
      slackWebhook: this.scriptProperties.getProperty('SLACK_WEBHOOK_URL'),
      slackEnabled: this.scriptProperties.getProperty('SLACK_ENABLED') === 'true'
    };
  }

  // =============================================================================
  // アプリケーション設定管理
  // =============================================================================

  /**
   * データ更新設定を設定
   */
  setDataUpdateSettings(settings) {
    const properties = {
      'AUTO_UPDATE_ENABLED': settings.autoUpdateEnabled ? 'true' : 'false',
      'UPDATE_INTERVAL_HOURS': (settings.updateIntervalHours || 24).toString(),
      'BATCH_SIZE': (settings.batchSize || 100).toString(),
      'RETRY_COUNT': (settings.retryCount || 3).toString(),
      'TIMEOUT_SECONDS': (settings.timeoutSeconds || 300).toString()
    };

    this.scriptProperties.setProperties(properties);
    this.logConfigChange('データ更新設定を更新しました');
  }

  /**
   * データ更新設定を取得
   */
  getDataUpdateSettings() {
    return {
      autoUpdateEnabled: this.scriptProperties.getProperty('AUTO_UPDATE_ENABLED') === 'true',
      updateIntervalHours: parseInt(this.scriptProperties.getProperty('UPDATE_INTERVAL_HOURS')) || 24,
      batchSize: parseInt(this.scriptProperties.getProperty('BATCH_SIZE')) || 100,
      retryCount: parseInt(this.scriptProperties.getProperty('RETRY_COUNT')) || 3,
      timeoutSeconds: parseInt(this.scriptProperties.getProperty('TIMEOUT_SECONDS')) || 300
    };
  }

  /**
   * KPI設定を設定
   */
  setKPISettings(settings) {
    const properties = {
      'TARGET_MONTHLY_PROFIT': (settings.targetMonthlyProfit || 800000).toString(),
      'TARGET_PROFIT_MARGIN': (settings.targetProfitMargin || 25).toString(),
      'TARGET_ROI': (settings.targetROI || 30).toString(),
      'MAX_INVENTORY_VALUE': (settings.maxInventoryValue || 1000000).toString(),
      'STAGNANT_DAYS_THRESHOLD': (settings.stagnantDaysThreshold || 60).toString(),
      'LOW_STOCK_THRESHOLD': (settings.lowStockThreshold || 5).toString()
    };

    this.scriptProperties.setProperties(properties);
    this.logConfigChange('KPI設定を更新しました');
  }

  /**
   * KPI設定を取得
   */
  getKPISettings() {
    return {
      targetMonthlyProfit: parseInt(this.scriptProperties.getProperty('TARGET_MONTHLY_PROFIT')) || 800000,
      targetProfitMargin: parseFloat(this.scriptProperties.getProperty('TARGET_PROFIT_MARGIN')) || 25,
      targetROI: parseFloat(this.scriptProperties.getProperty('TARGET_ROI')) || 30,
      maxInventoryValue: parseInt(this.scriptProperties.getProperty('MAX_INVENTORY_VALUE')) || 1000000,
      stagnantDaysThreshold: parseInt(this.scriptProperties.getProperty('STAGNANT_DAYS_THRESHOLD')) || 60,
      lowStockThreshold: parseInt(this.scriptProperties.getProperty('LOW_STOCK_THRESHOLD')) || 5
    };
  }

  // =============================================================================
  // ユーザー固有設定
  // =============================================================================

  /**
   * ユーザープリファレンスを設定
   */
  setUserPreferences(prefs) {
    const properties = {
      'TIMEZONE': prefs.timezone || 'Asia/Tokyo',
      'DATE_FORMAT': prefs.dateFormat || 'yyyy-MM-dd',
      'CURRENCY_FORMAT': prefs.currencyFormat || '¥#,##0',
      'DASHBOARD_REFRESH_INTERVAL': (prefs.dashboardRefreshInterval || 60).toString()
    };

    this.userProperties.setProperties(properties);
  }

  /**
   * ユーザープリファレンスを取得
   */
  getUserPreferences() {
    return {
      timezone: this.userProperties.getProperty('TIMEZONE') || 'Asia/Tokyo',
      dateFormat: this.userProperties.getProperty('DATE_FORMAT') || 'yyyy-MM-dd',
      currencyFormat: this.userProperties.getProperty('CURRENCY_FORMAT') || '¥#,##0',
      dashboardRefreshInterval: parseInt(this.userProperties.getProperty('DASHBOARD_REFRESH_INTERVAL')) || 60
    };
  }

  // =============================================================================
  // 設定検証
  // =============================================================================

  /**
   * 設定の整合性をチェック
   */
  validateConfiguration() {
    const errors = [];
    
    // Amazon認証情報チェック
    const amazonCreds = this.getAmazonCredentials();
    if (!amazonCreds.clientId || !amazonCreds.clientSecret || !amazonCreds.refreshToken) {
      errors.push('Amazon SP-API認証情報が不完全です');
    }

    // 通知設定チェック
    const notificationSettings = this.getNotificationSettings();
    if (notificationSettings.emailEnabled && !notificationSettings.email) {
      errors.push('メール通知が有効ですが、メールアドレスが設定されていません');
    }

    // KPI設定チェック
    const kpiSettings = this.getKPISettings();
    if (kpiSettings.targetProfitMargin < 0 || kpiSettings.targetProfitMargin > 100) {
      errors.push('目標利益率は0-100の範囲で設定してください');
    }

    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * 必要最小限の設定がされているかチェック
   */
  isConfigurationComplete() {
    const amazonCreds = this.getAmazonCredentials();
    return !!(amazonCreds.clientId && amazonCreds.clientSecret && amazonCreds.refreshToken);
  }

  // =============================================================================
  // 設定のバックアップ・復元
  // =============================================================================

  /**
   * 設定をエクスポート（機密情報は除く）
   */
  exportSettings() {
    return {
      dataUpdate: this.getDataUpdateSettings(),
      kpi: this.getKPISettings(),
      userPreferences: this.getUserPreferences(),
      notification: {
        emailEnabled: this.getNotificationSettings().emailEnabled,
        slackEnabled: this.getNotificationSettings().slackEnabled
      },
      exportedAt: new Date().toISOString(),
      version: APP_CONFIG.version
    };
  }

  /**
   * 設定をインポート
   */
  importSettings(settingsData) {
    try {
      if (settingsData.dataUpdate) {
        this.setDataUpdateSettings(settingsData.dataUpdate);
      }
      
      if (settingsData.kpi) {
        this.setKPISettings(settingsData.kpi);
      }
      
      if (settingsData.userPreferences) {
        this.setUserPreferences(settingsData.userPreferences);
      }

      this.logConfigChange('設定をインポートしました');
      return { success: true };
      
    } catch (error) {
      this.logConfigChange(`設定インポートエラー: ${error.message}`);
      return { success: false, error: error.message };
    }
  }

  // =============================================================================
  // 設定変更ログ
  // =============================================================================

  /**
   * 設定変更をログに記録
   */
  logConfigChange(message) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName('_設定変更ログ');
      
      if (!logSheet) {
        logSheet = ss.insertSheet('_設定変更ログ');
        logSheet.getRange(1, 1, 1, 4).setValues([
          ['日時', 'ユーザー', '変更内容', 'IPアドレス']
        ]);
      }

      const now = new Date();
      const user = Session.getActiveUser().getEmail();
      const ip = ''; // GASでは取得制限あり
      
      logSheet.appendRow([now, user, message, ip]);
      
    } catch (error) {
      console.error('設定変更ログ記録エラー:', error);
    }
  }

  /**
   * 最近の設定変更を取得
   */
  getRecentConfigChanges(limit = 10) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('_設定変更ログ');
      
      if (!logSheet) return [];

      const lastRow = logSheet.getLastRow();
      if (lastRow <= 1) return [];

      const startRow = Math.max(2, lastRow - limit + 1);
      const numRows = lastRow - startRow + 1;
      
      const data = logSheet.getRange(startRow, 1, numRows, 4).getValues();
      
      return data.map(row => ({
        timestamp: row[0],
        user: row[1],
        message: row[2],
        ip: row[3]
      })).reverse();
      
    } catch (error) {
      console.error('設定変更ログ取得エラー:', error);
      return [];
    }
  }

  // =============================================================================
  // 設定リセット
  // =============================================================================

  /**
   * 全設定をリセット（注意：復元不可）
   */
  resetAllSettings() {
    const confirmation = SpreadsheetApp.getUi().alert(
      '設定リセット確認',
      '全ての設定がリセットされます。この操作は元に戻せません。\n続行しますか？',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (confirmation === SpreadsheetApp.getUi().Button.YES) {
      // スクリプトプロパティを削除
      this.scriptProperties.deleteAllProperties();
      
      // ユーザープロパティを削除
      this.userProperties.deleteAllProperties();
      
      this.logConfigChange('全設定をリセットしました');
      
      return { success: true };
    }
    
    return { success: false, message: 'ユーザーによりキャンセルされました' };
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * 設定ダイアログを表示
 */
function showSettingsDialog() {
  const ui = SpreadsheetApp.getUi();
  const config = new ConfigManager();
  
  try {
    // 現在の設定を確認
    const validation = config.validateConfiguration();
    
    let message = '現在の設定状況:\n\n';
    
    if (config.isConfigurationComplete()) {
      message += '✅ 基本設定: 完了\n';
    } else {
      message += '❌ 基本設定: 未完了（Amazon API認証情報が必要）\n';
    }
    
    const kpiSettings = config.getKPISettings();
    message += `📊 目標月利: ${kpiSettings.targetMonthlyProfit.toLocaleString()}円\n`;
    message += `📈 目標利益率: ${kpiSettings.targetProfitMargin}%\n`;
    message += `🔄 目標ROI: ${kpiSettings.targetROI}%\n\n`;
    
    if (!validation.isValid) {
      message += '⚠️ 設定上の問題:\n';
      validation.errors.forEach(error => {
        message += `• ${error}\n`;
      });
      message += '\n';
    }
    
    message += '詳細な設定は「設定」シートで行ってください。\n';
    message += '「設定」シートを開きますか？';
    
    const response = ui.alert('設定確認', message, ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      openConfigSheet();
    }
    
  } catch (error) {
    ErrorHandler.handleError(error, 'showSettingsDialog');
    ui.alert('エラー', `設定確認中にエラーが発生しました:\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 設定シートを開く
 */
function openConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_CONFIG.CONFIG);
  
  if (!configSheet) {
    // 設定シートを作成
    const setupManager = new SetupManager();
    setupManager.createConfigSheet();
    configSheet = ss.getSheetByName(SHEET_CONFIG.CONFIG);
  }
  
  configSheet.activate();
}