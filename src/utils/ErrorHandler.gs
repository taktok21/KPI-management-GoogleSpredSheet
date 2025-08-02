/**
 * エラーハンドリングクラス
 * 
 * アプリケーション全体のエラー処理、ログ記録、通知機能を提供します。
 */

class ErrorHandler {
  constructor() {
    this.maxRetries = 3;
    this.retryDelay = 1000; // ミリ秒
    this.notificationThreshold = 5; // エラー回数の閾値
  }

  // =============================================================================
  // メインエラーハンドリング
  // =============================================================================

  /**
   * エラーを処理（ログ記録・通知）
   */
  static handleError(error, context = '', additionalInfo = {}) {
    const handler = new ErrorHandler();
    return handler.processError(error, context, additionalInfo);
  }

  /**
   * エラー処理の実行
   */
  processError(error, context, additionalInfo) {
    const errorInfo = this.createErrorInfo(error, context, additionalInfo);
    
    // ログ記録
    this.logError(errorInfo);
    
    // 重要度判定
    const severity = this.determineSeverity(error, context);
    errorInfo.severity = severity;
    
    // 通知判定
    if (this.shouldNotify(errorInfo)) {
      this.sendNotification(errorInfo);
    }
    
    // 統計更新
    this.updateErrorStats(errorInfo);
    
    return errorInfo;
  }

  /**
   * エラー情報オブジェクトの作成
   */
  createErrorInfo(error, context, additionalInfo) {
    return {
      timestamp: new Date(),
      context: context,
      errorType: error.constructor.name,
      message: error.message,
      stack: error.stack,
      userEmail: this.getCurrentUser(),
      spreadsheetId: this.getCurrentSpreadsheetId(),
      additionalInfo: additionalInfo,
      sessionId: this.getSessionId()
    };
  }

  // =============================================================================
  // リトライ機能付き実行
  // =============================================================================

  /**
   * リトライ付きで関数を実行
   */
  static executeWithRetry(func, context = '', options = {}) {
    const handler = new ErrorHandler();
    return handler.retryExecution(func, context, options);
  }

  /**
   * リトライ実行の処理
   */
  async retryExecution(func, context, options = {}) {
    const maxRetries = options.maxRetries || this.maxRetries;
    const retryDelay = options.retryDelay || this.retryDelay;
    const exponentialBackoff = options.exponentialBackoff !== false;
    
    let lastError;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await func();
      } catch (error) {
        lastError = error;
        
        // エラーログ記録
        this.logRetryAttempt(error, context, attempt, maxRetries);
        
        // リトライ可能かチェック
        if (!this.isRetryableError(error) || attempt === maxRetries) {
          // 最終的にエラーハンドリング
          this.handleError(error, `${context} (${attempt}回目の試行後に失敗)`);
          throw error;
        }
        
        // 遅延（指数バックオフ）
        const delay = exponentialBackoff 
          ? retryDelay * Math.pow(2, attempt - 1)
          : retryDelay;
        
        if (delay > 0) {
          Utilities.sleep(delay);
        }
      }
    }
    
    throw lastError;
  }

  /**
   * リトライ可能なエラーかチェック
   */
  isRetryableError(error) {
    // HTTP関連のリトライ可能エラー
    const retryableHttpCodes = [408, 429, 500, 502, 503, 504];
    
    if (error.response && retryableHttpCodes.includes(error.response.status)) {
      return true;
    }
    
    // ネットワークエラー
    if (error.message && (
      error.message.includes('fetch') ||
      error.message.includes('timeout') ||
      error.message.includes('network') ||
      error.message.includes('connection')
    )) {
      return true;
    }
    
    // Google API関連の一時的エラー
    if (error.message && (
      error.message.includes('Service invoked too many times') ||
      error.message.includes('Rate limit exceeded') ||
      error.message.includes('Backend Error')
    )) {
      return true;
    }
    
    return false;
  }

  // =============================================================================
  // エラーログ記録
  // =============================================================================

  /**
   * エラーログを記録
   */
  logError(errorInfo) {
    try {
      // コンソールログ
      console.error('エラー発生:', {
        context: errorInfo.context,
        message: errorInfo.message,
        timestamp: errorInfo.timestamp
      });
      
      // スプレッドシートログ
      this.saveErrorToSheet(errorInfo);
      
      // Properties Serviceに最新エラー保存
      this.saveLatestError(errorInfo);
      
    } catch (logError) {
      console.error('エラーログ記録に失敗:', logError);
    }
  }

  /**
   * エラーをスプレッドシートに保存
   */
  saveErrorToSheet(errorInfo) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let errorSheet = ss.getSheetByName('_エラーログ');
      
      if (!errorSheet) {
        errorSheet = ss.insertSheet('_エラーログ');
        
        // ヘッダー行設定
        const headers = [
          '発生時刻', 'コンテキスト', 'エラー種別', 'メッセージ',
          '重要度', 'ユーザー', 'セッションID', '追加情報'
        ];
        
        errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        errorSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        errorSheet.setFrozenRows(1);
      }
      
      // エラー情報を追加
      const row = [
        errorInfo.timestamp,
        errorInfo.context,
        errorInfo.errorType,
        errorInfo.message,
        errorInfo.severity || 'UNKNOWN',
        errorInfo.userEmail,
        errorInfo.sessionId,
        JSON.stringify(errorInfo.additionalInfo)
      ];
      
      errorSheet.appendRow(row);
      
      // 古いログの削除（1000行を超えたら古いものから削除）
      this.cleanupOldLogs(errorSheet, 1000);
      
    } catch (error) {
      console.error('スプレッドシートへのエラーログ保存に失敗:', error);
    }
  }

  /**
   * 最新エラーをProperties Serviceに保存
   */
  saveLatestError(errorInfo) {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('LATEST_ERROR', JSON.stringify({
        timestamp: errorInfo.timestamp.toISOString(),
        context: errorInfo.context,
        message: errorInfo.message,
        severity: errorInfo.severity
      }));
    } catch (error) {
      console.error('最新エラー保存に失敗:', error);
    }
  }

  /**
   * リトライ試行をログ記録
   */
  logRetryAttempt(error, context, attempt, maxRetries) {
    console.warn(`リトライ ${attempt}/${maxRetries}: ${context} - ${error.message}`);
  }

  // =============================================================================
  // エラー重要度判定
  // =============================================================================

  /**
   * エラーの重要度を判定
   */
  determineSeverity(error, context) {
    // CRITICAL: システムの基本機能に影響
    if (this.isCriticalError(error, context)) {
      return 'CRITICAL';
    }
    
    // HIGH: 主要機能に影響
    if (this.isHighSeverityError(error, context)) {
      return 'HIGH';
    }
    
    // MEDIUM: 部分的な機能に影響
    if (this.isMediumSeverityError(error, context)) {
      return 'MEDIUM';
    }
    
    // LOW: 軽微な影響
    return 'LOW';
  }

  /**
   * 重大エラーの判定
   */
  isCriticalError(error, context) {
    const criticalPatterns = [
      'Authentication',
      'Authorization',
      'Invalid API key',
      'Permission denied',
      'Script timeout',
      'Memory limit exceeded'
    ];
    
    const criticalContexts = [
      'initialSetup',
      'runDailyBatch',
      'fetchAmazonData'
    ];
    
    return criticalPatterns.some(pattern => 
      error.message && error.message.includes(pattern)
    ) || criticalContexts.includes(context);
  }

  /**
   * 高重要度エラーの判定
   */
  isHighSeverityError(error, context) {
    const highPatterns = [
      'Rate limit exceeded',
      'Service unavailable',
      'Network error',
      'Data validation failed'
    ];
    
    return highPatterns.some(pattern => 
      error.message && error.message.includes(pattern)
    );
  }

  /**
   * 中重要度エラーの判定
   */
  isMediumSeverityError(error, context) {
    const mediumPatterns = [
      'CSV parse error',
      'Format error',
      'Calculation error'
    ];
    
    return mediumPatterns.some(pattern => 
      error.message && error.message.includes(pattern)
    );
  }

  // =============================================================================
  // エラー通知
  // =============================================================================

  /**
   * 通知が必要かチェック
   */
  shouldNotify(errorInfo) {
    // 重大エラーは常に通知
    if (errorInfo.severity === 'CRITICAL') {
      return true;
    }
    
    // 高重要度エラーで一定回数発生したら通知
    if (errorInfo.severity === 'HIGH') {
      const recentErrors = this.getRecentErrorCount(errorInfo.context, 60); // 1時間以内
      return recentErrors >= 3;
    }
    
    return false;
  }

  /**
   * エラー通知送信
   */
  sendNotification(errorInfo) {
    try {
      const config = new ConfigManager();
      const notificationSettings = config.getNotificationSettings();
      
      // メール通知
      if (notificationSettings.emailEnabled && notificationSettings.email) {
        this.sendEmailNotification(errorInfo, notificationSettings.email);
      }
      
      // Slack通知
      if (notificationSettings.slackEnabled && notificationSettings.slackWebhook) {
        this.sendSlackNotification(errorInfo, notificationSettings.slackWebhook);
      }
      
    } catch (error) {
      console.error('エラー通知送信に失敗:', error);
    }
  }

  /**
   * メール通知送信
   */
  sendEmailNotification(errorInfo, email) {
    try {
      const subject = `[KPI管理] ${errorInfo.severity}エラー発生 - ${errorInfo.context}`;
      
      const body = `
KPI管理システムでエラーが発生しました。

【エラー情報】
発生時刻: ${errorInfo.timestamp.toLocaleString('ja-JP')}
重要度: ${errorInfo.severity}
コンテキスト: ${errorInfo.context}
エラー種別: ${errorInfo.errorType}
メッセージ: ${errorInfo.message}
ユーザー: ${errorInfo.userEmail}

【対応】
スプレッドシートの「_エラーログ」シートで詳細を確認してください。
重大なエラーの場合は、システム管理者にお問い合わせください。

スプレッドシートURL: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
`;
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body
      });
      
    } catch (error) {
      console.error('メール通知送信に失敗:', error);
    }
  }

  /**
   * Slack通知送信
   */
  sendSlackNotification(errorInfo, webhookUrl) {
    try {
      const color = this.getSeverityColor(errorInfo.severity);
      
      const payload = {
        text: `KPI管理システムでエラーが発生しました`,
        attachments: [{
          color: color,
          title: `${errorInfo.severity}エラー - ${errorInfo.context}`,
          fields: [
            {
              title: 'エラーメッセージ',
              value: errorInfo.message,
              short: false
            },
            {
              title: '発生時刻',
              value: errorInfo.timestamp.toLocaleString('ja-JP'),
              short: true
            },
            {
              title: 'ユーザー',
              value: errorInfo.userEmail,
              short: true
            }
          ],
          footer: 'KPI管理システム',
          ts: Math.floor(errorInfo.timestamp.getTime() / 1000)
        }]
      };
      
      UrlFetchApp.fetch(webhookUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
      
    } catch (error) {
      console.error('Slack通知送信に失敗:', error);
    }
  }

  /**
   * 重要度に応じた色を取得
   */
  getSeverityColor(severity) {
    const colors = {
      'CRITICAL': '#ff0000',
      'HIGH': '#ff9900',
      'MEDIUM': '#ffcc00',
      'LOW': '#0099ff'
    };
    
    return colors[severity] || '#cccccc';
  }

  // =============================================================================
  // エラー統計・分析
  // =============================================================================

  /**
   * エラー統計を更新
   */
  updateErrorStats(errorInfo) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const today = DateUtils.formatDate(new Date(), 'yyyy-MM-dd');
      
      // 日次エラーカウント更新
      const dailyKey = `ERROR_COUNT_${today}`;
      const dailyCount = parseInt(properties.getProperty(dailyKey)) || 0;
      properties.setProperty(dailyKey, (dailyCount + 1).toString());
      
      // コンテキスト別エラーカウント更新
      const contextKey = `ERROR_COUNT_${errorInfo.context}_${today}`;
      const contextCount = parseInt(properties.getProperty(contextKey)) || 0;
      properties.setProperty(contextKey, (contextCount + 1).toString());
      
      // 重要度別エラーカウント更新
      const severityKey = `ERROR_COUNT_${errorInfo.severity}_${today}`;
      const severityCount = parseInt(properties.getProperty(severityKey)) || 0;
      properties.setProperty(severityKey, (severityCount + 1).toString());
      
    } catch (error) {
      console.error('エラー統計更新に失敗:', error);
    }
  }

  /**
   * 最近のエラー回数を取得
   */
  getRecentErrorCount(context, minutesBack = 60) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = ss.getSheetByName('_エラーログ');
      
      if (!errorSheet) return 0;
      
      const lastRow = errorSheet.getLastRow();
      if (lastRow <= 1) return 0;
      
      const cutoffTime = new Date(Date.now() - (minutesBack * 60 * 1000));
      const data = errorSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      let count = 0;
      for (let i = data.length - 1; i >= 0; i--) {
        const timestamp = data[i][0];
        const errorContext = data[i][1];
        
        if (timestamp < cutoffTime) break;
        if (errorContext === context) count++;
      }
      
      return count;
      
    } catch (error) {
      console.error('最近のエラー回数取得に失敗:', error);
      return 0;
    }
  }

  // =============================================================================
  // ユーティリティ
  // =============================================================================

  /**
   * 現在のユーザーを取得
   */
  getCurrentUser() {
    try {
      return Session.getActiveUser().getEmail();
    } catch (error) {
      return 'unknown';
    }
  }

  /**
   * 現在のスプレッドシートIDを取得
   */
  getCurrentSpreadsheetId() {
    try {
      return SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch (error) {
      return 'unknown';
    }
  }

  /**
   * セッションIDを取得（簡易版）
   */
  getSessionId() {
    const sessionKey = 'CURRENT_SESSION_ID';
    let sessionId = PropertiesService.getUserProperties().getProperty(sessionKey);
    
    if (!sessionId) {
      sessionId = Utilities.getUuid();
      PropertiesService.getUserProperties().setProperty(sessionKey, sessionId);
    }
    
    return sessionId;
  }

  /**
   * 古いログをクリーンアップ
   */
  cleanupOldLogs(sheet, maxRows) {
    try {
      const lastRow = sheet.getLastRow();
      
      if (lastRow > maxRows) {
        const rowsToDelete = lastRow - maxRows;
        sheet.deleteRows(2, rowsToDelete); // ヘッダー行は保持
      }
    } catch (error) {
      console.error('ログクリーンアップに失敗:', error);
    }
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * エラーログ表示
 */
function showErrorLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('_エラーログ');
    
    if (!errorSheet) {
      SpreadsheetApp.getUi().alert(
        'エラーログ',
        'エラーログはまだ記録されていません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    errorSheet.activate();
    
    // 最新のエラーを最上部に表示するためソート
    const lastRow = errorSheet.getLastRow();
    if (lastRow > 1) {
      const range = errorSheet.getRange(2, 1, lastRow - 1, errorSheet.getLastColumn());
      range.sort({column: 1, ascending: false});
    }
    
  } catch (error) {
    ErrorHandler.handleError(error, 'showErrorLog');
  }
}

/**
 * エラー統計表示
 */
function showErrorStats() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const today = DateUtils.formatDate(new Date(), 'yyyy-MM-dd');
    
    const dailyCount = parseInt(properties.getProperty(`ERROR_COUNT_${today}`)) || 0;
    const criticalCount = parseInt(properties.getProperty(`ERROR_COUNT_CRITICAL_${today}`)) || 0;
    const highCount = parseInt(properties.getProperty(`ERROR_COUNT_HIGH_${today}`)) || 0;
    
    const message = `
本日のエラー統計:

総エラー数: ${dailyCount}件
重大エラー: ${criticalCount}件
高重要度エラー: ${highCount}件

詳細は「_エラーログ」シートをご確認ください。
`;
    
    SpreadsheetApp.getUi().alert('エラー統計', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    ErrorHandler.handleError(error, 'showErrorStats');
  }
}