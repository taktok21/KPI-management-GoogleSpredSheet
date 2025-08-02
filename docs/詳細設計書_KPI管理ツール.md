# 詳細設計書：Amazon販売KPI管理ツール

## 1. システム構成詳細

### 1.1 Google Apps Script プロジェクト構成
```
KPI管理ツール/
├── Code.gs           # メイン処理
├── Config.gs         # 設定管理
├── DataFetcher.gs    # データ取得処理
├── DataProcessor.gs  # データ処理
├── KPICalculator.gs  # KPI計算
├── UIController.gs   # UI制御
├── ErrorHandler.gs   # エラー処理
└── Utils.gs          # ユーティリティ
```

### 1.2 スプレッドシート構成詳細

#### シート定義
```javascript
const SHEET_CONFIG = {
  KPI_MONTHLY: 'KPI月次管理',
  SALES_HISTORY: '販売履歴',
  PURCHASE_HISTORY: '仕入履歴',
  INVENTORY: '在庫一覧',
  PRODUCT_MASTER: 'ASIN/SKUマスタ',
  SYNC_LOG: 'データ連携ログ',
  CONFIG: '設定',
  TEMP_DATA: '_一時データ'
};
```

## 2. データモデル詳細

### 2.1 販売履歴テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| order_id | STRING | Amazon注文ID | PRIMARY KEY |
| order_date | DATETIME | 注文日時（JST） | NOT NULL |
| asin | STRING | Amazon商品ID | NOT NULL |
| sku | STRING | 統一SKU | NOT NULL |
| product_name | STRING | 商品名 | NOT NULL |
| quantity | INTEGER | 販売数量 | NOT NULL, > 0 |
| unit_price | DECIMAL | 単価 | NOT NULL, >= 0 |
| total_price | DECIMAL | 合計金額 | NOT NULL |
| purchase_cost | DECIMAL | 仕入原価 | NOT NULL, >= 0 |
| amazon_fee | DECIMAL | Amazon手数料 | NOT NULL, >= 0 |
| fba_fee | DECIMAL | FBA手数料 | >= 0 |
| other_cost | DECIMAL | その他費用 | >= 0 |
| gross_profit | DECIMAL | 粗利益 | CALCULATED |
| profit_rate | DECIMAL | 利益率 | CALCULATED |
| status | STRING | 注文ステータス | IN ('Shipped', 'Pending', 'Cancelled') |
| fulfillment | STRING | 配送方法 | IN ('FBA', 'FBM') |
| marketplace | STRING | マーケットプレイス | DEFAULT 'JP' |

### 2.2 在庫一覧テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| sku | STRING | 統一SKU | PRIMARY KEY |
| asin | STRING | Amazon商品ID | NOT NULL |
| product_name | STRING | 商品名 | NOT NULL |
| condition | STRING | 商品状態 | IN ('New', 'Used') |
| quantity | INTEGER | 在庫数 | NOT NULL, >= 0 |
| unit_cost | DECIMAL | 単位原価 | NOT NULL, >= 0 |
| total_cost | DECIMAL | 在庫金額 | CALCULATED |
| location | STRING | 在庫場所 | IN ('FBA', 'Self') |
| last_inbound_date | DATE | 最終入荷日 | |
| last_sold_date | DATE | 最終販売日 | |
| days_in_stock | INTEGER | 在庫日数 | CALCULATED |
| turnover_days | INTEGER | 回転日数 | CALCULATED |
| monthly_sales | INTEGER | 月間販売数 | CALCULATED |
| alert_flag | STRING | アラート | IN ('滞留', '在庫切れ間近', NULL) |

### 2.3 SKUマッピングテーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| unified_sku | STRING | 統一SKU | PRIMARY KEY |
| amazon_sku | STRING | Amazon SKU | UNIQUE |
| makado_sku | STRING | マカドSKU | UNIQUE |
| asin | STRING | ASIN | NOT NULL |
| jan_code | STRING | JANコード | |
| product_name | STRING | 正式商品名 | NOT NULL |
| category | STRING | カテゴリ | |
| brand | STRING | ブランド | |
| purchase_date | DATE | 仕入日 | |
| created_at | DATETIME | 作成日時 | DEFAULT NOW |
| updated_at | DATETIME | 更新日時 | ON UPDATE NOW |

## 3. API連携詳細

### 3.1 Amazon SP-API連携

#### 認証設定
```javascript
const SP_API_CONFIG = {
  endpoint: 'https://sellingpartnerapi-fe.amazon.com',
  marketplaceId: 'A1VC38T7YXB528', // Japan
  credentials: {
    clientId: PropertiesService.getScriptProperties().getProperty('SP_API_CLIENT_ID'),
    clientSecret: PropertiesService.getScriptProperties().getProperty('SP_API_CLIENT_SECRET'),
    refreshToken: PropertiesService.getScriptProperties().getProperty('SP_API_REFRESH_TOKEN')
  }
};
```

#### データ取得関数
```javascript
// 注文データ取得
function fetchAmazonOrders(startDate, endDate) {
  const accessToken = getAccessToken();
  const url = `${SP_API_CONFIG.endpoint}/orders/v0/orders`;
  
  const params = {
    MarketplaceIds: SP_API_CONFIG.marketplaceId,
    CreatedAfter: startDate.toISOString(),
    CreatedBefore: endDate.toISOString(),
    OrderStatuses: 'Shipped,Pending',
    MaxResultsPerPage: 100
  };
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'x-amz-access-token': accessToken
    },
    muteHttpExceptions: true
  };
  
  return fetchWithPagination(url, params, options);
}

// 在庫データ取得
function fetchInventorySummaries() {
  const accessToken = getAccessToken();
  const url = `${SP_API_CONFIG.endpoint}/fba/inventory/v1/summaries`;
  
  const params = {
    granularityType: 'Marketplace',
    granularityId: SP_API_CONFIG.marketplaceId,
    marketplaceIds: SP_API_CONFIG.marketplaceId
  };
  
  // 実装継続...
}
```

### 3.2 マカドCSV処理

#### CSVパーサー
```javascript
function parseMakadoCSV(csvContent) {
  // 文字コード変換（Shift-JIS → UTF-8）
  const decodedContent = Utilities.newBlob(csvContent)
    .getDataAsString('Shift_JIS');
  
  const lines = decodedContent.split('\n');
  const headers = parseCSVLine(lines[0]);
  
  const columnMapping = {
    '日付': 'order_date',
    '商品名': 'product_name',
    'オーダーID': 'order_id',
    'ASIN': 'asin',
    'SKU': 'makado_sku',
    'コンディション': 'condition',
    '配送経路': 'fulfillment',
    '販売価格': 'unit_price',
    '送料': 'shipping',
    'ポイント': 'points',
    '割引': 'discount',
    '仕入価格': 'purchase_cost',
    'その他経費': 'other_cost',
    'Amazon手数料': 'amazon_fee',
    '粗利': 'gross_profit',
    'ステータス': 'status',
    '販売数': 'quantity'
  };
  
  const data = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    
    const values = parseCSVLine(lines[i]);
    const row = {};
    
    headers.forEach((header, index) => {
      const mappedKey = columnMapping[header];
      if (mappedKey && values[index] !== undefined) {
        row[mappedKey] = convertValue(values[index], mappedKey);
      }
    });
    
    // SKU統一処理
    row.unified_sku = generateUnifiedSKU(row.asin, row.makado_sku);
    
    data.push(row);
  }
  
  return data;
}

// CSVライン解析（ダブルクォート対応）
function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current);
  return result;
}
```

## 4. KPI計算ロジック詳細

### 4.1 基本KPI計算
```javascript
class KPICalculator {
  // 粗利益計算
  calculateGrossProfit(salesData) {
    return salesData.reduce((total, sale) => {
      const revenue = sale.unit_price * sale.quantity;
      const cost = sale.purchase_cost * sale.quantity;
      const fees = sale.amazon_fee + (sale.fba_fee || 0) + (sale.other_cost || 0);
      return total + (revenue - cost - fees);
    }, 0);
  }
  
  // ROI計算
  calculateROI(grossProfit, totalInvestment) {
    if (totalInvestment === 0) return 0;
    return (grossProfit / totalInvestment) * 100;
  }
  
  // 在庫回転率計算
  calculateInventoryTurnover(salesAmount, averageInventory) {
    if (averageInventory === 0) return 0;
    return salesAmount / averageInventory;
  }
  
  // 回転日数計算
  calculateTurnoverDays(inventoryTurnover, daysInPeriod = 30) {
    if (inventoryTurnover === 0) return Infinity;
    return daysInPeriod / inventoryTurnover;
  }
  
  // カート取得率計算（推定）
  estimateBuyBoxPercentage(salesData) {
    // 販売価格と市場価格の比較から推定
    // 実装は簡略化
    return 85; // デフォルト値
  }
}
```

### 4.2 高度なKPI計算
```javascript
// 商品別パフォーマンス分析
function analyzeProductPerformance(salesData, inventoryData) {
  const productMetrics = {};
  
  // 商品ごとに集計
  salesData.forEach(sale => {
    const sku = sale.unified_sku;
    
    if (!productMetrics[sku]) {
      productMetrics[sku] = {
        totalRevenue: 0,
        totalProfit: 0,
        totalQuantity: 0,
        salesCount: 0,
        firstSaleDate: sale.order_date,
        lastSaleDate: sale.order_date
      };
    }
    
    const metrics = productMetrics[sku];
    metrics.totalRevenue += sale.unit_price * sale.quantity;
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
  
  // 在庫情報とマージ
  Object.keys(productMetrics).forEach(sku => {
    const inventory = inventoryData.find(inv => inv.sku === sku);
    const metrics = productMetrics[sku];
    
    if (inventory) {
      metrics.currentStock = inventory.quantity;
      metrics.stockValue = inventory.total_cost;
      metrics.daysInStock = inventory.days_in_stock;
      
      // 販売速度計算（日販）
      const salesDays = daysBetween(metrics.firstSaleDate, metrics.lastSaleDate) || 1;
      metrics.dailySalesVelocity = metrics.totalQuantity / salesDays;
      
      // 在庫切れまでの予測日数
      if (metrics.dailySalesVelocity > 0) {
        metrics.stockoutDays = Math.floor(inventory.quantity / metrics.dailySalesVelocity);
      }
    }
    
    // 利益率計算
    metrics.profitMargin = metrics.totalRevenue > 0 
      ? (metrics.totalProfit / metrics.totalRevenue) * 100 
      : 0;
    
    // 平均販売単価
    metrics.averageSellingPrice = metrics.totalQuantity > 0
      ? metrics.totalRevenue / metrics.totalQuantity
      : 0;
  });
  
  return productMetrics;
}
```

## 5. データ同期処理

### 5.1 メインバッチ処理
```javascript
function runDailyBatch() {
  const startTime = new Date();
  const log = [];
  
  try {
    // 1. 設定読み込み
    log.push('バッチ処理開始');
    const config = loadConfiguration();
    
    // 2. Amazon注文データ取得
    log.push('Amazon注文データ取得中...');
    const orders = fetchAmazonOrdersForToday();
    log.push(`注文データ: ${orders.length}件取得`);
    
    // 3. Amazon在庫データ取得
    log.push('Amazon在庫データ取得中...');
    const inventory = fetchCurrentInventory();
    log.push(`在庫データ: ${inventory.length}件取得`);
    
    // 4. マカドCSV確認・取り込み
    log.push('マカドデータ確認中...');
    const makadoData = checkAndImportMakadoCSV();
    if (makadoData.length > 0) {
      log.push(`マカドデータ: ${makadoData.length}件取得`);
    }
    
    // 5. データ統合
    log.push('データ統合処理中...');
    const integratedData = integrateAllData(orders, inventory, makadoData);
    
    // 6. データ保存
    log.push('データ保存中...');
    saveToSpreadsheet(integratedData);
    
    // 7. KPI計算
    log.push('KPI計算中...');
    calculateAndUpdateKPIs();
    
    // 8. アラートチェック
    log.push('アラートチェック中...');
    const alerts = checkAlerts();
    if (alerts.length > 0) {
      sendAlertNotifications(alerts);
      log.push(`アラート: ${alerts.length}件検出`);
    }
    
    // 9. 完了処理
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    log.push(`バッチ処理完了（処理時間: ${duration}秒）`);
    
    // ログ保存
    saveProcessingLog(log, 'SUCCESS');
    
  } catch (error) {
    log.push(`エラー発生: ${error.toString()}`);
    saveProcessingLog(log, 'ERROR');
    sendErrorNotification(error);
    throw error;
  }
}
```

### 5.2 データ統合処理
```javascript
function integrateAllData(amazonOrders, inventory, makadoData) {
  // SKUマッピングテーブル読み込み
  const skuMapping = loadSKUMapping();
  
  // 統合データ配列
  const integratedData = {
    sales: [],
    inventory: [],
    unmappedItems: []
  };
  
  // Amazon注文データの処理
  amazonOrders.forEach(order => {
    order.orderItems.forEach(item => {
      const mappedSku = findUnifiedSKU(item.sku, item.asin, skuMapping);
      
      if (mappedSku) {
        integratedData.sales.push({
          order_id: order.amazonOrderId,
          order_date: new Date(order.purchaseDate),
          asin: item.asin,
          sku: mappedSku,
          product_name: item.title,
          quantity: item.quantityOrdered,
          unit_price: parseFloat(item.itemPrice.amount),
          amazon_fee: calculateAmazonFee(item),
          status: order.orderStatus,
          fulfillment: order.fulfillmentChannel,
          source: 'Amazon'
        });
      } else {
        integratedData.unmappedItems.push({
          type: 'order',
          sku: item.sku,
          asin: item.asin,
          title: item.title
        });
      }
    });
  });
  
  // マカドデータの統合
  makadoData.forEach(row => {
    const mappedSku = findUnifiedSKU(row.makado_sku, row.asin, skuMapping);
    
    if (mappedSku) {
      // 既存注文との重複チェック
      const exists = integratedData.sales.some(
        sale => sale.order_id === row.order_id && sale.asin === row.asin
      );
      
      if (!exists) {
        integratedData.sales.push({
          ...row,
          sku: mappedSku,
          source: 'Makado'
        });
      }
    }
  });
  
  // 在庫データの処理
  inventory.forEach(inv => {
    const mappedSku = findUnifiedSKU(inv.sellerSku, inv.asin, skuMapping);
    
    if (mappedSku) {
      integratedData.inventory.push({
        sku: mappedSku,
        asin: inv.asin,
        product_name: inv.productName,
        quantity: inv.totalQuantity,
        condition: inv.condition,
        location: 'FBA'
      });
    }
  });
  
  return integratedData;
}
```

## 6. エラーハンドリング

### 6.1 エラーハンドラー実装
```javascript
class ErrorHandler {
  constructor() {
    this.maxRetries = 3;
    this.retryDelay = 1000; // ミリ秒
  }
  
  // リトライ付き実行
  async executeWithRetry(func, context = '') {
    let lastError;
    
    for (let attempt = 1; attempt <= this.maxRetries; attempt++) {
      try {
        return await func();
      } catch (error) {
        lastError = error;
        
        // エラーログ記録
        this.logError(error, context, attempt);
        
        // リトライ可能なエラーかチェック
        if (!this.isRetryableError(error) || attempt === this.maxRetries) {
          throw error;
        }
        
        // 指数バックオフ
        const delay = this.retryDelay * Math.pow(2, attempt - 1);
        Utilities.sleep(delay);
      }
    }
    
    throw lastError;
  }
  
  // エラー種別判定
  isRetryableError(error) {
    const retryableStatuses = [429, 500, 502, 503, 504];
    
    if (error.response && retryableStatuses.includes(error.response.status)) {
      return true;
    }
    
    // ネットワークエラー
    if (error.message && error.message.includes('fetch')) {
      return true;
    }
    
    return false;
  }
  
  // エラーログ記録
  logError(error, context, attempt = 1) {
    const errorLog = {
      timestamp: new Date().toISOString(),
      context: context,
      attempt: attempt,
      errorType: error.constructor.name,
      message: error.message,
      stack: error.stack,
      details: error.response ? {
        status: error.response.status,
        statusText: error.response.statusText,
        data: error.response.data
      } : null
    };
    
    // スプレッドシートに記録
    this.saveErrorLog(errorLog);
    
    // 重大エラーの場合は即座に通知
    if (this.isCriticalError(error)) {
      this.sendImmediateAlert(errorLog);
    }
  }
  
  // 重大エラー判定
  isCriticalError(error) {
    const criticalPatterns = [
      'Authentication',
      'Authorization',
      'Rate limit exceeded',
      'Invalid API key'
    ];
    
    return criticalPatterns.some(pattern => 
      error.message && error.message.includes(pattern)
    );
  }
}
```

## 7. UI制御

### 7.1 カスタムメニュー実装
```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('KPI管理ツール')
    .addItem('データ更新', 'manualDataSync')
    .addItem('KPI再計算', 'recalculateKPIs')
    .addSeparator()
    .addSubMenu(ui.createMenu('インポート')
      .addItem('マカドCSV取り込み', 'importMakadoCSV')
      .addItem('過去データ取り込み', 'importHistoricalData'))
    .addSeparator()
    .addSubMenu(ui.createMenu('レポート')
      .addItem('日次レポート生成', 'generateDailyReport')
      .addItem('週次レポート生成', 'generateWeeklyReport')
      .addItem('月次レポート生成', 'generateMonthlyReport'))
    .addSeparator()
    .addItem('設定', 'showSettingsDialog')
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

// 手動データ同期
function manualDataSync() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'データ更新',
    'データを更新します。処理には数分かかる場合があります。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      ui.alert('処理中...', '処理が完了するまでお待ちください。', ui.ButtonSet.OK);
      runDailyBatch();
      ui.alert('完了', 'データ更新が完了しました。', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('エラー', `データ更新中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}
```

### 7.2 ダッシュボード更新
```javascript
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
  
  // 現在の月次データ取得
  const currentMonth = new Date();
  const monthlyData = getMonthlyData(currentMonth);
  
  // KPI計算
  const kpis = calculateMonthlyKPIs(monthlyData);
  
  // ダッシュボード更新
  const updates = [
    // 基本KPI
    { cell: 'B2', value: kpis.grossProfit, format: '¥#,##0' },
    { cell: 'B3', value: kpis.revenue, format: '¥#,##0' },
    { cell: 'B4', value: kpis.profitMargin, format: '0.0%' },
    { cell: 'B5', value: kpis.roi, format: '0.0%' },
    
    // 在庫KPI
    { cell: 'E2', value: kpis.inventoryValue, format: '¥#,##0' },
    { cell: 'E3', value: kpis.inventoryTurnover, format: '0.0' },
    { cell: 'E4', value: kpis.turnoverDays, format: '0' },
    { cell: 'E5', value: kpis.stagnantRate, format: '0.0%' },
    
    // 販売KPI
    { cell: 'H2', value: kpis.totalQuantity, format: '#,##0' },
    { cell: 'H3', value: kpis.uniqueASINs, format: '#,##0' },
    { cell: 'H4', value: kpis.averagePrice, format: '¥#,##0' },
    { cell: 'H5', value: kpis.repeatRate, format: '0.0%' }
  ];
  
  // バッチ更新
  updates.forEach(update => {
    const range = dashboardSheet.getRange(update.cell);
    range.setValue(update.value);
    range.setNumberFormat(update.format);
  });
  
  // 条件付き書式設定
  applyConditionalFormatting(dashboardSheet, kpis);
  
  // グラフ更新
  updateCharts(dashboardSheet, monthlyData);
  
  // 最終更新日時
  dashboardSheet.getRange('K1').setValue(new Date());
}
```

## 8. セキュリティ実装

### 8.1 API認証情報管理
```javascript
class CredentialManager {
  constructor() {
    this.scriptProperties = PropertiesService.getScriptProperties();
  }
  
  // 認証情報の安全な保存
  setCredentials(type, credentials) {
    const encrypted = this.encrypt(JSON.stringify(credentials));
    this.scriptProperties.setProperty(`${type}_CREDENTIALS`, encrypted);
  }
  
  // 認証情報の取得
  getCredentials(type) {
    const encrypted = this.scriptProperties.getProperty(`${type}_CREDENTIALS`);
    if (!encrypted) return null;
    
    try {
      const decrypted = this.decrypt(encrypted);
      return JSON.parse(decrypted);
    } catch (error) {
      console.error('Failed to decrypt credentials:', error);
      return null;
    }
  }
  
  // 簡易暗号化（本番環境では適切な暗号化ライブラリを使用）
  encrypt(text) {
    const key = this.getEncryptionKey();
    return Utilities.base64Encode(text);
  }
  
  decrypt(encrypted) {
    return Utilities.newBlob(
      Utilities.base64Decode(encrypted)
    ).getDataAsString();
  }
  
  // 暗号化キー取得
  getEncryptionKey() {
    let key = this.scriptProperties.getProperty('ENCRYPTION_KEY');
    if (!key) {
      key = Utilities.getUuid();
      this.scriptProperties.setProperty('ENCRYPTION_KEY', key);
    }
    return key;
  }
}
```

## 9. パフォーマンス最適化

### 9.1 バッチ処理最適化
```javascript
// スプレッドシート一括更新
function batchUpdateSpreadsheet(sheetName, data, startRow = 2) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(sheetName);
  
  if (!sheet || !data || data.length === 0) return;
  
  // ヘッダー行を除いたデータ範囲
  const numRows = data.length;
  const numCols = data[0].length;
  
  // 既存データクリア
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols)
      .clearContent();
  }
  
  // 一括書き込み
  const range = sheet.getRange(startRow, 1, numRows, numCols);
  range.setValues(data);
  
  // フォーマット設定も一括で
  applyBatchFormatting(sheet, startRow, numRows);
}

// キャッシュ活用
class DataCache {
  constructor(cacheExpiry = 300) { // 5分
    this.cache = CacheService.getScriptCache();
    this.expiry = cacheExpiry;
  }
  
  get(key) {
    const cached = this.cache.get(key);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        return cached;
      }
    }
    return null;
  }
  
  set(key, value) {
    const toCache = typeof value === 'object' 
      ? JSON.stringify(value) 
      : value;
    
    this.cache.put(key, toCache, this.expiry);
  }
  
  clear(pattern = null) {
    if (pattern) {
      // パターンマッチでキャッシュクリア
      const keys = this.cache.getAll();
      Object.keys(keys).forEach(key => {
        if (key.includes(pattern)) {
          this.cache.remove(key);
        }
      });
    } else {
      // 全キャッシュクリア
      this.cache.removeAll([]);
    }
  }
}
```

## 10. 監視・アラート

### 10.1 アラート定義
```javascript
const ALERT_RULES = {
  // 利益率低下
  LOW_PROFIT_MARGIN: {
    condition: (kpi) => kpi.profitMargin < 20,
    message: (kpi) => `利益率が${kpi.profitMargin.toFixed(1)}%に低下しています（目標: 25%以上）`,
    severity: 'WARNING'
  },
  
  // 在庫過多
  EXCESS_INVENTORY: {
    condition: (kpi) => kpi.inventoryValue > 1000000,
    message: (kpi) => `在庫金額が${(kpi.inventoryValue / 10000).toFixed(0)}万円を超えています`,
    severity: 'WARNING'
  },
  
  // 滞留在庫
  STAGNANT_INVENTORY: {
    condition: (kpi) => kpi.stagnantRate > 15,
    message: (kpi) => `60日以上の滞留在庫が${kpi.stagnantRate.toFixed(1)}%あります`,
    severity: 'CRITICAL'
  },
  
  // 日次売上低下
  LOW_DAILY_SALES: {
    condition: (kpi) => kpi.dailyRevenue < 50000,
    message: (kpi) => `本日の売上が${kpi.dailyRevenue}円です（目標: 9万円/日）`,
    severity: 'INFO'
  }
};

// アラートチェック実行
function checkAlerts() {
  const currentKPIs = getCurrentKPIs();
  const alerts = [];
  
  Object.entries(ALERT_RULES).forEach(([ruleName, rule]) => {
    if (rule.condition(currentKPIs)) {
      alerts.push({
        rule: ruleName,
        message: rule.message(currentKPIs),
        severity: rule.severity,
        timestamp: new Date(),
        kpiSnapshot: currentKPIs
      });
    }
  });
  
  return alerts;
}
```

### 10.2 通知実装
```javascript
// 統合通知システム
class NotificationService {
  constructor() {
    this.emailEnabled = true;
    this.slackEnabled = this.checkSlackConfig();
  }
  
  // アラート通知送信
  sendAlert(alerts) {
    if (!alerts || alerts.length === 0) return;
    
    // 重要度でグループ化
    const groupedAlerts = this.groupAlertsBySeverity(alerts);
    
    // メール通知
    if (this.emailEnabled) {
      this.sendEmailNotification(groupedAlerts);
    }
    
    // Slack通知
    if (this.slackEnabled) {
      this.sendSlackNotification(groupedAlerts);
    }
    
    // ログ記録
    this.logAlerts(alerts);
  }
  
  // メール通知
  sendEmailNotification(groupedAlerts) {
    const recipient = PropertiesService.getScriptProperties()
      .getProperty('ALERT_EMAIL');
    
    if (!recipient) return;
    
    const subject = `[KPI管理] アラート通知 - ${new Date().toLocaleDateString('ja-JP')}`;
    
    let body = 'KPI管理システムからのアラート通知です。\n\n';
    
    ['CRITICAL', 'WARNING', 'INFO'].forEach(severity => {
      if (groupedAlerts[severity] && groupedAlerts[severity].length > 0) {
        body += `【${severity}】\n`;
        groupedAlerts[severity].forEach(alert => {
          body += `・${alert.message}\n`;
        });
        body += '\n';
      }
    });
    
    body += '詳細はスプレッドシートをご確認ください。\n';
    body += SpreadsheetApp.getActiveSpreadsheet().getUrl();
    
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });
  }
  
  // Slack通知
  sendSlackNotification(groupedAlerts) {
    const webhookUrl = PropertiesService.getScriptProperties()
      .getProperty('SLACK_WEBHOOK_URL');
    
    if (!webhookUrl) return;
    
    const attachments = [];
    
    ['CRITICAL', 'WARNING', 'INFO'].forEach(severity => {
      if (groupedAlerts[severity] && groupedAlerts[severity].length > 0) {
        const color = {
          CRITICAL: '#ff0000',
          WARNING: '#ff9900',
          INFO: '#0099ff'
        }[severity];
        
        attachments.push({
          color: color,
          title: `${severity} アラート`,
          fields: groupedAlerts[severity].map(alert => ({
            value: alert.message,
            short: false
          }))
        });
      }
    });
    
    const payload = {
      text: 'KPI管理システムからのアラート',
      attachments: attachments
    };
    
    UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  }
}
```

## 11. テスト仕様

### 11.1 単体テスト
```javascript
// テストランナー
function runAllTests() {
  const tests = [
    testSKUMapping,
    testKPICalculation,
    testDataIntegration,
    testErrorHandling
  ];
  
  const results = [];
  
  tests.forEach(test => {
    try {
      test();
      results.push({ name: test.name, status: 'PASS' });
    } catch (error) {
      results.push({ 
        name: test.name, 
        status: 'FAIL', 
        error: error.message 
      });
    }
  });
  
  // 結果出力
  console.log('Test Results:');
  results.forEach(result => {
    console.log(`${result.name}: ${result.status}`);
    if (result.error) {
      console.log(`  Error: ${result.error}`);
    }
  });
}

// SKUマッピングテスト
function testSKUMapping() {
  const testCases = [
    {
      input: { makadoSku: 'MAKAD-ABC123-240101', asin: 'B001234567' },
      expected: 'UNI-B001234567-240101'
    },
    {
      input: { amazonSku: 'AMZ-001', asin: 'B001234567' },
      expected: 'UNI-B001234567-DEFAULT'
    }
  ];
  
  testCases.forEach(testCase => {
    const result = generateUnifiedSKU(
      testCase.input.asin, 
      testCase.input.makadoSku || testCase.input.amazonSku
    );
    
    if (result !== testCase.expected) {
      throw new Error(
        `Expected ${testCase.expected}, got ${result}`
      );
    }
  });
}
```

## 12. 実装スケジュール

### Phase 1: 基本機能（Week 1-2）
- [ ] プロジェクト設定
- [ ] 基本的なデータモデル実装
- [ ] Amazon SP-API連携（基本）
- [ ] マカドCSV読み込み
- [ ] 基本的なKPI計算

### Phase 2: 自動化（Week 3）
- [ ] バッチ処理実装
- [ ] エラーハンドリング
- [ ] 通知機能
- [ ] ログ機能

### Phase 3: 高度な機能（Week 4-5）
- [ ] 詳細なKPI分析
- [ ] ダッシュボード最適化
- [ ] パフォーマンスチューニング
- [ ] テスト実装

### Phase 4: 運用開始（Week 6）
- [ ] 本番データ移行
- [ ] 運用マニュアル作成
- [ ] トレーニング実施
- [ ] 本番運用開始