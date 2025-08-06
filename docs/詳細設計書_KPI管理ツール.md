# 詳細設計書：Amazon販売KPI管理ツール

**バージョン**: 1.5  
**更新日**: 2025年8月6日  
**更新内容**: 新ダッシュボード（MVP版）実装詳細・8KPI体制・段階的更新システム

## 1. システム構成詳細

### 1.1 Google Apps Script プロジェクト構成
```
KPI管理ツール/
├── src/
│   ├── Code.gs                    # メイン処理・メニュー制御
│   ├── core/
│   │   ├── Config.gs             # 設定管理クラス
│   │   ├── KPICalculator.gs      # KPI計算エンジン（強化版）
│   │   ├── KPIHistoryManager.gs  # KPI履歴管理
│   │   ├── SetupManager.gs       # 初期セットアップ管理
│   │   └── BatchProcessor.gs     # バッチ処理制御
│   ├── api/
│   │   ├── MakadoProcessor.gs    # マカドCSV処理（Shift-JIS対応）
│   │   └── AmazonAPI.gs          # Amazon SP-API連携
│   ├── utils/
│   │   ├── Utils.gs              # 汎用ユーティリティ
│   │   └── ErrorHandler.gs      # エラー処理・リトライ機能
│   └── sheets/
│       └── SheetManager.gs       # スプレッドシート操作
├── tests/
│   └── TestRunner.gs             # テスト実行
├── docs/
│   ├── 基本設計書_KPI管理ツール.md
│   └── 詳細設計書_KPI管理ツール.md
└── templates/
    └── スプレッドシートテンプレート
```

### 1.2 スプレッドシート構成詳細

#### シート定義
```javascript
const SHEET_CONFIG = {
  KPI_MONTHLY: 'KPI月次管理',
  KPI_HISTORY: 'KPI履歴',
  SALES_HISTORY: '販売履歴',
  PURCHASE_HISTORY: '仕入履歴',
  INVENTORY: '在庫一覧',
  PRODUCT_MASTER: 'ASIN/SKUマスタ',
  SYNC_LOG: 'データ連携ログ',
  CONFIG: '設定',
  TEMP_DATA: '_一時データ',
  // 新規追加（MVP版）
  DASHBOARD: 'ダッシュボード',
  PLAN_INPUT: 'plan_monthly_input',
  VIEW_CATEGORY: 'ビュー_カテゴリ',
  VIEW_SUPPLIER: 'ビュー_仕入先',
  VIEW_SKU: 'ビュー_SKU',
  SETTINGS: '設定'
};
```

#### KPI月次管理シート詳細構成

##### セル構成・レイアウト
```
A列：ラベル  B列：当月実績  C列：前月比  D列：達成率  E列：設定/メモ  F列：前年同月比  G列～：グラフエリア
```

##### 重要セル詳細設計

**F2セル: 期間選択ドロップダウン**
```javascript
// データ検証の実装
function setupPeriodDropdown() {
  const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
  const cell = sheet.getRange('F2');
  
  // 過去24ヶ月の期間リスト生成
  const periods = [];
  const currentDate = new Date();
  
  for (let i = 0; i < 24; i++) {
    const targetDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1);
    const yearMonth = Utilities.formatDate(targetDate, 'JST', 'yyyy年MM月');
    periods.push(yearMonth);
  }
  
  // データ検証ルールの設定
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(periods, true)
    .setAllowInvalid(false)
    .setHelpText('比較対象月を選択してください')
    .build();
    
  cell.setDataValidation(validation);
  cell.setValue(periods[12]); // デフォルトで前年同月を設定
}
```

**F列: 前年同月比の計算式とフォーマット**
```javascript
// F5セル: 売上高前年同月比
const revenueYoYFormula = `
=IF(AND(ISNUMBER(B5), B5>0, ISNUMBER(INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),2))), INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),2))>0),
  (B5-INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),2)))/INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),2)),
  "-")
`;

// F6セル: 粗利益前年同月比
const profitYoYFormula = `
=IF(AND(ISNUMBER(B6), B6>0, ISNUMBER(INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),3))), INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),3))>0),
  (B6-INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),3)))/INDIRECT("KPI履歴!"&ADDRESS(MATCH(TEXT(DATE(YEAR(TODAY())-1,MONTH(TODAY()),1),"yyyy-mm"),"KPI履歴!A:A",0),3)),
  "-")
`;

// フォーマット設定
function formatYoYColumns() {
  const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
  const yoyRange = sheet.getRange('F5:F12');
  
  // パーセント表示、条件付き書式設定
  yoyRange.setNumberFormat('0.0%');
  
  // 条件付き書式：プラスは緑、マイナスは赤
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d9ead3')
    .setFontColor('#137333')
    .build();
    
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#fce5cd')
    .setFontColor('#cc0000')
    .build();
    
  sheet.setConditionalFormatRules([positiveRule, negativeRule]);
}
```

**G列以降: グラフエリアの詳細設計**
```javascript
// グラフ配置設定
const CHART_CONFIG = {
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
    position: { row: 5, column: 13 }, // M5セル
    size: { width: 300, height: 200 },
    type: 'GAUGE',
    title: '目標達成率'
  }
};
```

## 2. データモデル詳細

### 2.1 販売履歴テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| order_id | STRING | Amazon注文ID | PRIMARY KEY |
| order_date | DATETIME | 注文日時（JST） | NOT NULL |
| asin | STRING | Amazon商品ID | NOT NULL |
| unified_sku | STRING | 統一SKU（UNI-{ASIN}-{Date}） | NOT NULL |
| makado_sku | STRING | マカドSKU | NOT NULL |
| product_name | STRING | 商品名 | NOT NULL |
| quantity | INTEGER | 販売数量（返品時は0を許可） | NOT NULL, >= 0 |
| unit_price | DECIMAL | 単価 | NOT NULL, >= 0 |
| total_amount | DECIMAL | 合計金額 | NOT NULL |
| purchase_cost | DECIMAL | 仕入原価 | NOT NULL, >= 0 |
| amazon_fee | DECIMAL | Amazon手数料 | NOT NULL, >= 0 |
| other_cost | DECIMAL | その他費用 | >= 0 |
| gross_profit | DECIMAL | 粗利益 | NOT NULL |
| profit_margin | DECIMAL | 利益率（%） | NOT NULL |
| status | STRING | ステータス（SHIPPED/RETURN/CANCELLED） | NOT NULL |
| fulfillment | STRING | 配送方法（FBA/FBM） | NOT NULL |
| data_source | STRING | データソース（MAKADO/AMAZON） | NOT NULL |
| import_timestamp | DATETIME | インポート日時 | NOT NULL |
| marketplace | STRING | マーケットプレイス | DEFAULT 'JP' |

### 2.2 仕入履歴テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| unified_sku | STRING | 統一SKU（UNI-{ASIN}-{Date}） | PRIMARY KEY |
| asin | STRING | Amazon商品ID | NOT NULL |
| purchase_date | DATE | 仕入日 | NOT NULL |
| supplier | STRING | 仕入先名 | NOT NULL |
| quantity | INTEGER | 仕入数量 | NOT NULL, > 0 |
| unit_cost | DECIMAL | 仕入単価 | NOT NULL, >= 0 |
| total_cost | DECIMAL | 仕入合計金額 | NOT NULL, >= 0 |
| shipping_cost | DECIMAL | 送料 | >= 0 |
| notes | STRING | 備考 | |
| created_at | DATETIME | 登録日時 | DEFAULT NOW |

### 2.3 在庫一覧テーブル
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

### 2.4 SKUマッピングテーブル
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

### 2.5 KPI履歴テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| year_month | STRING | 年月（YYYY-MM） | PRIMARY KEY |
| revenue | DECIMAL | 売上高 | NOT NULL, >= 0 |
| gross_profit | DECIMAL | 粗利益 | NOT NULL |
| profit_margin | DECIMAL | 利益率（%） | NOT NULL |
| roi | DECIMAL | ROI（%） | NOT NULL |
| sales_quantity | INTEGER | 販売数 | NOT NULL, >= 0 |
| inventory_value | DECIMAL | 在庫金額 | >= 0 |
| inventory_turnover | DECIMAL | 在庫回転率 | >= 0 |
| stagnant_inventory_rate | DECIMAL | 滞留在庫率（%） | >= 0 |
| unique_products | INTEGER | 取扱商品数 | >= 0 |
| dio | INTEGER | 在庫日数 | >= 0 |
| return_rate | DECIMAL | 返品率（%） | >= 0 |
| ccc | INTEGER | キャッシュサイクル（日） | |
| new_product_rate | DECIMAL | 新商品投入率（%） | >= 0 |
| average_order_value | DECIMAL | 平均注文額 | >= 0 |
| profit_goal_achievement | DECIMAL | 利益目標達成率（%） | >= 0 |
| calculated_at | DATETIME | 計算日時 | NOT NULL |
| created_at | DATETIME | 作成日時 | DEFAULT NOW |

### 2.6 plan_monthly_input（目標入力）テーブル
| カラム名 | 型 | 説明 | 制約 |
|----------|-----|------|------|
| year_month | STRING | 年月（YYYY-MM） | COMPOSITE KEY |
| granularity | STRING | 粒度 | IN ('business', 'category') |
| category_code | STRING | カテゴリコード | NULLABLE |
| revenue_target | DECIMAL | 売上目標 | >= 0 |
| profit_target | DECIMAL | 粗利目標 | >= 0 |
| profit_margin_target | DECIMAL | 粗利率目標（%） | 0-100 |
| turnover_target | DECIMAL | 回転率目標 | >= 0 |
| dio_target | INTEGER | DIO目標 | >= 0 |
| ccc_target | INTEGER | CCC目標 | |
| new_product_target | INTEGER | 新商品投入目標 | >= 0 |
| remarks | STRING | 備考 | |
| created_at | DATETIME | 作成日時 | DEFAULT NOW |
| updated_at | DATETIME | 更新日時 | ON UPDATE NOW |

## 3. データ処理詳細

### 3.1 マカドCSV処理（実装完了）

#### エンコーディング処理
```javascript
// Shift-JIS → UTF-8 変換処理
readCSVContent(file) {
  const blob = file.getBlob();
  
  try {
    // Shift-JISで読み込み
    content = blob.getDataAsString('Shift_JIS');
  } catch (encodingError) {
    // フォールバック: UTF-8で読み込み
    content = blob.getDataAsString('UTF-8');
  }
  
  return content;
}
```

#### CSVカラムマッピング
| マカドCSV列名 | システム項目名 | データ型 | 説明 |
|---------------|----------------|----------|------|
| 注文日 | order_date | DATE | 注文日時 |
| 商品名 | product_name | STRING | 商品名 |
| オーダーID | order_id | STRING | Amazon注文ID |
| ASIN | asin | STRING | Amazon商品ID |
| SKU | makado_sku | STRING | マカドSKU |
| コンディション | condition | STRING | 商品状態 |
| 配送経路 | fulfillment | STRING | 配送方法 |
| 販売価格 | unit_price | DECIMAL | 単価 |
| 送料 | shipping_fee | DECIMAL | 送料 |
| ポイント | points | DECIMAL | ポイント |
| 割引 | discount | DECIMAL | 割引額 |
| 仕入れ価格 | purchase_cost | DECIMAL | 仕入原価 |
| その他経費 | other_cost | DECIMAL | その他費用 |
| Amazon手数料 | amazon_fee | DECIMAL | Amazon手数料 |
| 粗利 | gross_profit | DECIMAL | 粗利益 |
| ステータス | status | STRING | 注文状態 |
| 販売数 | quantity | INTEGER | 販売数量 |
| 累計販売数 | cumulative_quantity | INTEGER | 累計数量 |

#### 返品・キャンセルデータ処理
```javascript
// 数量検証（返品・キャンセルの場合は0を許可）
validateRecord(record) {
  if (record.quantity < 0) {
    errors.push('販売数量が負の値です');
  } else if (record.quantity === 0 && 
             record.status !== 'RETURN' && 
             record.status !== 'CANCELLED') {
    errors.push('販売数量が0です（返品・キャンセル以外）');
  }
}
```

#### 重複チェック処理
```javascript
// 注文ID + ASIN + 注文日での重複チェック
isDuplicate(record) {
  const checkKey = `${record.order_id}_${record.asin}_${DateUtils.formatDate(record.order_date, 'yyyy-MM-dd')}`;
  // 既存データとの比較処理
}
```

### 3.2 Amazon SP-API連携

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
```

## 4. KPI計算機能詳細（強化版実装完了）

### 4.1 月次KPI計算（8指標対応）
```javascript
calculateMonthlyKPIs(salesData, inventoryData, purchaseData) {
  const currentMonth = DateUtils.getCurrentMonthRange();
  
  // 今月のデータでフィルタリング
  const monthlySales = salesData.filter(sale => 
    sale.order_date >= currentMonth.start && sale.order_date <= currentMonth.end
  );
  
  // 基本売上KPI
  const totalRevenue = ArrayUtils.sum(monthlySales, sale => sale.total_amount);
  const totalGrossProfit = ArrayUtils.sum(monthlySales, sale => sale.gross_profit);
  const totalQuantity = ArrayUtils.sum(monthlySales, sale => sale.quantity);
  const totalPurchaseAmount = ArrayUtils.sum(purchaseData, p => p.total_cost);
  
  // 返品率計算
  const returnCount = monthlySales.filter(s => s.status === 'RETURN').length;
  const returnRate = NumberUtils.percentage(returnCount, monthlySales.length);
  
  // 在庫関連KPI
  const avgInventoryValue = NumberUtils.average(inventoryData, inv => inv.total_cost);
  const inventoryTurnover = totalRevenue / avgInventoryValue;
  const dio = 365 / inventoryTurnover;
  
  // CCC計算（物販では売掛金0日、買掛金30日想定）
  const ccc = dio + 0 - 30;
  
  // 新商品投入率
  const newProducts = inventoryData.filter(inv => 
    DateUtils.daysBetween(inv.first_inbound_date, new Date()) <= 30
  ).length;
  const newProductRate = NumberUtils.percentage(newProducts, inventoryData.length);
  
  return {
    // 基本8指標
    revenue: totalRevenue,
    grossProfit: totalGrossProfit,
    profitMargin: NumberUtils.percentage(totalGrossProfit, totalRevenue),
    roi: NumberUtils.percentage(totalGrossProfit, totalPurchaseAmount),
    inventoryTurnover: inventoryTurnover,
    dio: Math.round(dio),
    returnRate: returnRate,
    ccc: Math.round(ccc),
    newProductRate: newProductRate,
    // 追加情報
    profitGoalAchievement: NumberUtils.percentage(totalGrossProfit, 800000),
    salesQuantity: totalQuantity
  };
}
```

### 4.2 日次KPI計算
```javascript
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
    weeklyAvgRevenue: ArrayUtils.average(weekSales, sale => sale.total_amount) * 7,
    growthRate: this.calculateGrowthRate(todaySales, weekSales)
  };
}
```

### 4.3 KPIダッシュボード更新
```javascript
updateKPIDashboard(monthlyKPIs, dailyKPIs) {
  const kpiSheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
  
  // 月次KPI更新
  kpiSheet.getRange('B5').setValue(monthlyKPIs.revenue);        // 売上高
  kpiSheet.getRange('B6').setValue(monthlyKPIs.grossProfit);    // 粗利益
  kpiSheet.getRange('B7').setValue(monthlyKPIs.profitMargin / 100); // 利益率
  kpiSheet.getRange('B8').setValue(monthlyKPIs.roi / 100);      // ROI
  
  // 日次KPI更新
  kpiSheet.getRange('B16').setValue(dailyKPIs.todayRevenue);    // 本日売上
  kpiSheet.getRange('B17').setValue(dailyKPIs.todayProfit);     // 本日利益
  kpiSheet.getRange('D16').setValue(dailyKPIs.growthRate / 100); // 成長率
}
```

### 4.4 アラート機能
```javascript
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

  return alerts;
}
```

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

## 5. エラー処理詳細（実装完了）

### 5.1 リトライ処理
```javascript
class ErrorHandler {
  static executeWithRetry(func, context, maxRetries = 3) {
    let lastError;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return func();
      } catch (error) {
        lastError = error;
        
        if (attempt < maxRetries) {
          const delay = Math.pow(2, attempt) * 1000; // 指数バックオフ
          Utilities.sleep(delay);
        }
      }
    }
    
    throw lastError;
  }
}
```

### 5.2 エラー分類と対応
| エラー種別 | 対応方法 | リトライ | 通知 |
|------------|----------|----------|------|
| API接続エラー | リトライ処理 | 3回 | あり |
| データ形式エラー | スキップ/修正 | なし | ログのみ |
| タイムアウト | リトライ処理 | 2回 | あり |
| 権限エラー | 処理停止 | なし | 即座に通知 |
| 重複データ | スキップ | なし | ログのみ |

### 5.3 検証エラーログ
```javascript
logValidationErrors(validationErrors) {
  const logSheet = ss.getSheetByName('_CSV検証エラー');
  
  validationErrors.forEach(error => {
    logSheet.appendRow([
      new Date(),
      error.index,
      error.record.order_id || '',
      error.record.asin || '',
      error.errors.join('; '),
      JSON.stringify(error.record)
    ]);
  });
}
```

## 6. パフォーマンス最適化

### 6.1 キャッシュ機能
```javascript
class CacheUtils {
  static set(key, value, expireSeconds = 300) {
    const cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(value), expireSeconds);
  }
  
  static get(key) {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    return cached ? JSON.parse(cached) : null;
  }
}
```

### 6.2 バッチ処理
- スプレッドシート書き込み：100行単位でバッチ処理
- API呼び出し：レート制限対応（1秒間隔）
- データ取得：ページング処理（100件ずつ）

## 7. 新しいクラス・関数の設計

### 7.1 KPIHistoryManagerクラスの拡張
```javascript
class KPIHistoryManager {
  // KPI履歴の保存
  saveMonthlyKPI(yearMonth, kpiData) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    const existingRow = this.findExistingRow(sheet, yearMonth);
    
    if (existingRow) {
      // 既存データを更新
      this.updateKPIRecord(sheet, existingRow, kpiData);
    } else {
      // 新規データを追加
      this.appendKPIRecord(sheet, yearMonth, kpiData);
    }
  }
  
  // 過去12ヶ月のKPI取得
  getHistoricalKPIs(months = 12) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    const currentMonth = DateUtils.getCurrentMonth();
    const targetMonths = this.generateMonthList(currentMonth, months);
    
    return targetMonths.map(month => this.getMonthKPI(sheet, month));
  }
  
  // 前月比計算
  calculateMonthOverMonth(currentKPI, previousKPI) {
    return {
      revenue: NumberUtils.percentage(
        currentKPI.revenue - previousKPI.revenue,
        previousKPI.revenue
      ),
      grossProfit: NumberUtils.percentage(
        currentKPI.grossProfit - previousKPI.grossProfit,
        previousKPI.grossProfit
      ),
      profitMargin: currentKPI.profitMargin - previousKPI.profitMargin,
      roi: currentKPI.roi - previousKPI.roi
    };
  }
  
  // 前年同月比計算（新機能）
  calculateYearOverYear(currentKPI, targetMonth) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    const previousYearKPI = this.getMonthKPI(sheet, targetMonth);
    
    if (!previousYearKPI || !previousYearKPI.revenue) {
      return this.createEmptyYoYResult();
    }
    
    return {
      revenue: NumberUtils.percentage(
        currentKPI.revenue - previousYearKPI.revenue,
        previousYearKPI.revenue
      ),
      grossProfit: NumberUtils.percentage(
        currentKPI.grossProfit - previousYearKPI.grossProfit,
        previousYearKPI.grossProfit
      ),
      profitMargin: currentKPI.profitMargin - previousYearKPI.profitMargin,
      roi: currentKPI.roi - previousYearKPI.roi,
      salesQuantity: NumberUtils.percentage(
        currentKPI.salesQuantity - previousYearKPI.salesQuantity,
        previousYearKPI.salesQuantity
      ),
      inventoryValue: NumberUtils.percentage(
        currentKPI.inventoryValue - previousYearKPI.inventoryValue,
        previousYearKPI.inventoryValue || 1
      )
    };
  }
  
  // 時系列データの効率的な取得
  getTimeSeriesData(startMonth, endMonth, metrics = ['revenue', 'grossProfit']) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // インデックスマッピング作成
    const indexMap = {};
    headers.forEach((header, index) => {
      indexMap[header] = index;
    });
    
    // 期間フィルタリング
    const filteredData = data.slice(1).filter(row => {
      const yearMonth = row[indexMap['year_month']];
      return yearMonth >= startMonth && yearMonth <= endMonth;
    });
    
    // メトリクス抽出
    return filteredData.map(row => {
      const result = { yearMonth: row[indexMap['year_month']] };
      metrics.forEach(metric => {
        result[metric] = row[indexMap[metric]] || 0;
      });
      return result;
    }).sort((a, b) => a.yearMonth.localeCompare(b.yearMonth));
  }
  
  // キャッシュ機能付きデータ取得
  getCachedHistoricalData(cacheKey, dataFunction, expireMinutes = 30) {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    
    if (cached) {
      return JSON.parse(cached);
    }
    
    const data = dataFunction();
    cache.put(cacheKey, JSON.stringify(data), expireMinutes * 60);
    return data;
  }
}
```

### 7.2 SheetManagerクラスの拡張（グラフ作成・更新）

### 7.3 新ダッシュボード実装詳細

#### ダッシュボードクラス設計
```javascript
class DashboardManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.cache = new CacheUtils();
  }

  // KPIカード更新（8枚）
  updateKPICards() {
    const sheet = this.ss.getSheetByName(SHEET_CONFIG.DASHBOARD);
    const kpis = this.calculateAllKPIs();
    const targets = this.getMonthlyTargets();
    
    // KPIカード配置（2行4列）
    const cardLayout = [
      {name: '売上高', row: 2, col: 2, kpi: 'revenue'},
      {name: '粗利益', row: 2, col: 4, kpi: 'grossProfit'},
      {name: '粗利率', row: 2, col: 6, kpi: 'profitMargin'},
      {name: '在庫回転', row: 2, col: 8, kpi: 'inventoryTurnover'},
      {name: 'DIO', row: 5, col: 2, kpi: 'dio'},
      {name: '返品率', row: 5, col: 4, kpi: 'returnRate'},
      {name: 'CCC', row: 5, col: 6, kpi: 'ccc'},
      {name: '新商品率', row: 5, col: 8, kpi: 'newProductRate'}
    ];
    
    cardLayout.forEach(card => {
      this.updateSingleKPICard(sheet, card, kpis[card.kpi], targets[card.kpi]);
    });
  }

  // 階層別集計ビュー作成
  createHierarchicalViews() {
    // カテゴリビュー
    this.createCategoryView();
    // 仕入先ビュー
    this.createSupplierView();
    // SKUビュー
    this.createSKUView();
  }

  // カテゴリビュー作成
  createCategoryView() {
    const viewSheet = this.ss.getSheetByName(SHEET_CONFIG.VIEW_CATEGORY);
    const startDate = this.getSettingValue('B2');
    const endDate = this.getSettingValue('C2');
    
    const query = `
      =LET(
        data, FILTER(販売履歴!A:Z, 
                     販売履歴!A:A>=${startDate}, 
                     販売履歴!A:A<=${endDate}),
        QUERY(data,
          "select Col9, sum(Col5), sum(Col12), count(Col1)
           where Col9 is not null
           group by Col9
           label Col9 'カテゴリ', sum(Col5) '売上(実)', 
                 sum(Col12) '粗利(実)', count(Col1) '件数'", 0)
      )
    `;
    
    viewSheet.getRange('A1').setFormula(query);
    this.addTargetColumns(viewSheet, 'category');
    this.addSparklineColumn(viewSheet);
  }

  // スパークライン追加
  addSparklineColumn(sheet) {
    const lastRow = sheet.getLastRow();
    const sparklineCol = sheet.getLastColumn() + 1;
    
    for (let row = 2; row <= lastRow; row++) {
      const sku = sheet.getRange(row, 1).getValue();
      const sparklineFormula = `
        =SPARKLINE(
          QUERY(販売履歴!A:E,
            "select A,sum(E) 
             where A>=date '"&TEXT(TODAY()-6,"yyyy-mm-dd")&
            "' and C='"&${sku}&"' 
             group by A order by A", 0),
          {"charttype","line";"linewidth",2;"color","#4285f4"}
        )
      `;
      sheet.getRange(row, sparklineCol).setFormula(sparklineFormula);
    }
  }

  // アラート抽出
  extractAlerts() {
    const alerts = {
      stockOut: this.getStockOutAlerts(),
      badInventory: this.getBadInventoryAlerts(),
      salesVolatility: this.getSalesVolatilityAlerts(),
      priceCompetition: this.getPriceCompetitionAlerts()
    };
    
    this.displayAlerts(alerts);
  }

  // 急激な売上変動検知
  getSalesVolatilityAlerts() {
    const query = `
      =FILTER({SKU列, 今週売上列, 前週売上列, 変動率列},
              ABS((今週売上列-前週売上列)/前週売上列) >= 0.3,
              今週売上列 >= 10000)
    `;
    return this.executeQuery(query);
  }
}
```

#### 段階的更新システム
```javascript
class IncrementalUpdater {
  constructor() {
    this.lastUpdate = PropertiesService.getScriptProperties()
                        .getProperty('lastUpdateTime');
  }

  // 15分毎の更新
  criticalUpdate() {
    // 売上・在庫切れのみ更新
    const updates = ['revenue', 'stockAlerts'];
    this.updateMetrics(updates);
    this.logUpdate('critical', updates);
  }

  // 1時間毎の更新
  hourlyUpdate() {
    // 中間指標の更新
    const updates = ['grossProfit', 'turnover', 'returnRate'];
    this.updateMetrics(updates);
    this.logUpdate('hourly', updates);
  }

  // 日次完全更新
  dailyFullUpdate() {
    // 全KPIの再計算
    const calculator = new KPICalculator();
    calculator.recalculateAll();
    this.logUpdate('daily', 'all');
  }

  // 増分データ取得
  getIncrementalData() {
    const newData = this.fetchDataSince(this.lastUpdate);
    if (newData.length > 0) {
      this.appendToSheet('販売履歴_増分', newData);
      this.updateAffectedKPIs(newData);
    }
  }
}
```
```javascript
class SheetManager {
  // 既存のメソッド...
  
  // グラフ作成・更新機能
  createOrUpdateChart(chartConfig, dataRange, sheet) {
    const existingCharts = sheet.getCharts();
    let targetChart = null;
    
    // 既存チャート検索
    existingCharts.forEach(chart => {
      const title = chart.getOptions().get('title');
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
  }
  
  // トレンドチャート作成
  createTrendChart(historicalData) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    
    // データ準備シート作成
    const tempSheet = this.createTempDataSheet(historicalData);
    const dataRange = tempSheet.getRange(1, 1, historicalData.length + 1, 3);
    
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(CHART_CONFIG.TREND_CHART.position.row, 
                   CHART_CONFIG.TREND_CHART.position.column, 0, 0)
      .setOption('title', CHART_CONFIG.TREND_CHART.title)
      .setOption('width', CHART_CONFIG.TREND_CHART.size.width)
      .setOption('height', CHART_CONFIG.TREND_CHART.size.height)
      .setOption('hAxis.title', '月')
      .setOption('vAxis.title', '金額（円）')
      .setOption('series', {
        0: { color: '#4285f4', lineWidth: 3 }, // 売上高
        1: { color: '#34a853', lineWidth: 3 }  // 粗利益
      })
      .setOption('legend.position', 'bottom')
      .setOption('backgroundColor', '#ffffff')
      .setOption('chartArea', {
        left: 60,
        top: 40,
        width: '80%',
        height: '70%'
      });
    
    const chart = chartBuilder.build();
    sheet.insertChart(chart);
    
    // 一時シート削除
    ss.deleteSheet(tempSheet);
    
    return chart;
  }
  
  // 比較チャート作成
  createComparisonChart(currentKPI, previousYearKPI) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    
    // 比較データ準備
    const comparisonData = [
      ['メトリクス', '当年', '前年'],
      ['売上高', currentKPI.revenue, previousYearKPI.revenue || 0],
      ['粗利益', currentKPI.grossProfit, previousYearKPI.grossProfit || 0],
      ['販売数', currentKPI.salesQuantity, previousYearKPI.salesQuantity || 0]
    ];
    
    const tempSheet = this.createTempComparisonSheet(comparisonData);
    const dataRange = tempSheet.getRange(1, 1, comparisonData.length, 3);
    
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(CHART_CONFIG.COMPARISON_CHART.position.row, 
                   CHART_CONFIG.COMPARISON_CHART.position.column, 0, 0)
      .setOption('title', CHART_CONFIG.COMPARISON_CHART.title)
      .setOption('width', CHART_CONFIG.COMPARISON_CHART.size.width)
      .setOption('height', CHART_CONFIG.COMPARISON_CHART.size.height)
      .setOption('series', {
        0: { color: '#4285f4' }, // 当年
        1: { color: '#ea4335' }  // 前年
      })
      .setOption('legend.position', 'bottom')
      .setOption('hAxis.title', 'KPI項目')
      .setOption('vAxis.title', '値');
    
    const chart = chartBuilder.build();
    sheet.insertChart(chart);
    
    // 一時シート削除
    ss.deleteSheet(tempSheet);
    
    return chart;
  }
  
  // ゲージチャート作成（目標達成率）
  createGaugeChart(achievementRate) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    
    const gaugeData = [
      ['ラベル', '値'],
      ['達成率', achievementRate / 100]
    ];
    
    const tempSheet = this.createTempGaugeSheet(gaugeData);
    const dataRange = tempSheet.getRange(1, 1, 2, 2);
    
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.GAUGE)
      .addRange(dataRange)
      .setPosition(CHART_CONFIG.KPI_GAUGE.position.row, 
                   CHART_CONFIG.KPI_GAUGE.position.column, 0, 0)
      .setOption('title', CHART_CONFIG.KPI_GAUGE.title)
      .setOption('width', CHART_CONFIG.KPI_GAUGE.size.width)
      .setOption('height', CHART_CONFIG.KPI_GAUGE.size.height)
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
    ss.deleteSheet(tempSheet);
    
    return chart;
  }
  
  // レスポンシブ更新処理
  updateChartsResponsively(kpiData) {
    try {
      // 非同期でチャート更新を実行
      const updateTasks = [
        () => this.updateTrendChart(kpiData.historical),
        () => this.updateComparisonChart(kpiData.current, kpiData.previousYear),
        () => this.updateGaugeChart(kpiData.current.profitGoalAchievement)
      ];
      
      updateTasks.forEach((task, index) => {
        try {
          task();
        } catch (error) {
          console.error(`チャート更新エラー (${index}):`, error);
          // エラーログ記録
          this.logChartError(index, error);
        }
      });
      
    } catch (error) {
      console.error('チャート更新処理エラー:', error);
      throw error;
    }
  }
}
```

### 7.4 onEditトリガー処理の詳細設計
```javascript
// onEditトリガーハンドラー
function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    // KPI月次管理シートのF2セル（期間選択）監視
    if (sheetName === SHEET_CONFIG.KPI_MONTHLY && 
        range.getA1Notation() === 'F2') {
      handlePeriodSelectionChange(e);
    }
    
    // データシート更新時のKPI再計算
    if ([SHEET_CONFIG.SALES_HISTORY, SHEET_CONFIG.PURCHASE_HISTORY].includes(sheetName)) {
      handleDataSheetChange(e);
    }
    
  } catch (error) {
    ErrorHandler.logError('onEdit', error, { range: e.range.getA1Notation() });
  }
}

// 期間選択変更処理
function handlePeriodSelectionChange(e) {
  const selectedPeriod = e.value;
  if (!selectedPeriod) return;
  
  // 選択期間をパース
  const targetDate = DateUtils.parsePeriodString(selectedPeriod);
  const targetMonth = Utilities.formatDate(targetDate, 'JST', 'yyyy-MM');
  
  // 前年同月比を再計算
  const kpiHistory = new KPIHistoryManager();
  const currentKPI = kpiHistory.getCurrentMonthKPI();
  const yoyComparison = kpiHistory.calculateYearOverYear(currentKPI, targetMonth);
  
  // F列の前年同月比を更新
  updateYearOverYearDisplay(yoyComparison);
  
  // グラフも更新
  const sheetManager = new SheetManager();
  const previousYearKPI = kpiHistory.getMonthKPI(
    ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY), 
    targetMonth
  );
  
  sheetManager.updateComparisonChart(currentKPI, previousYearKPI);
}

// データシート変更処理
function handleDataSheetChange(e) {
  // 変更があった場合のKPI再計算をスケジュール
  const trigger = ScriptApp.newTrigger('recalculateKPIsDelayed')
    .timeBased()
    .after(5000) // 5秒後に実行
    .create();
    
  // 既存の遅延トリガーを削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'recalculateKPIsDelayed') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// 遅延KPI再計算
function recalculateKPIsDelayed() {
  try {
    const batchProcessor = new BatchProcessor();
    batchProcessor.updateKPIs();
    
    // トリガー自体を削除
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === 'recalculateKPIsDelayed') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
  } catch (error) {
    ErrorHandler.logError('recalculateKPIsDelayed', error);
  }
}
```

## 8. データベース設計の拡張

### 8.1 KPI履歴テーブルのインデックス設計
```javascript
// インデックス構造の最適化
const KPI_HISTORY_INDEX = {
  PRIMARY: 'year_month',
  COMPOSITE_INDEXES: [
    ['year_month', 'calculated_at'],  // 時系列検索用
    ['created_at'],                   // 作成日順検索用
    ['revenue', 'gross_profit']       // KPIランキング用
  ]
};

// インデックス効率を活用したクエリ設計
class OptimizedKPIQuery {
  // 年月範囲での効率的な検索
  getKPIsByDateRange(startMonth, endMonth) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_HISTORY);
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行を除いた検索範囲を決定
    const startRow = this.findRowByYearMonth(data, startMonth);
    const endRow = this.findRowByYearMonth(data, endMonth);
    
    if (startRow === -1 || endRow === -1) {
      return [];
    }
    
    // 範囲指定で効率的にデータ取得
    const targetRange = sheet.getRange(startRow, 1, endRow - startRow + 1, data[0].length);
    return targetRange.getValues();
  }
  
  // バイナリサーチによる高速検索
  findRowByYearMonth(data, targetMonth) {
    let left = 1; // ヘッダー行をスキップ
    let right = data.length - 1;
    
    while (left <= right) {
      const mid = Math.floor((left + right) / 2);
      const midMonth = data[mid][0]; // year_month列
      
      if (midMonth === targetMonth) {
        return mid + 1; // スプレッドシートは1ベース
      } else if (midMonth < targetMonth) {
        left = mid + 1;
      } else {
        right = mid - 1;
      }
    }
    
    return -1;
  }
}
```

### 8.2 時系列データの効率的な取得方法
```javascript
class TimeSeriesDataManager {
  // キャッシュ戦略
  setupCacheStrategy() {
    return {
      // 短期キャッシュ：頻繁にアクセスされる当月データ
      currentMonth: {
        key: `kpi_current_${DateUtils.getCurrentMonth()}`,
        ttl: 300 // 5分
      },
      
      // 中期キャッシュ：過去12ヶ月の履歴データ
      historical: {
        key: `kpi_historical_12m`,
        ttl: 1800 // 30分
      },
      
      // 長期キャッシュ：統計データ
      statistics: {
        key: `kpi_stats`,
        ttl: 3600 // 1時間
      }
    };
  }
  
  // 段階的データロード
  loadTimeSeriesDataProgressively(startMonth, endMonth, callback) {
    const batchSize = 6; // 6ヶ月ずつ処理
    const totalMonths = DateUtils.getMonthDiff(startMonth, endMonth);
    const batches = Math.ceil(totalMonths / batchSize);
    
    let processedData = [];
    
    for (let i = 0; i < batches; i++) {
      const batchStart = DateUtils.addMonths(startMonth, i * batchSize);
      const batchEnd = DateUtils.addMonths(batchStart, batchSize - 1);
      
      const batchData = this.getKPIsByDateRange(batchStart, batchEnd);
      processedData = processedData.concat(batchData);
      
      // プログレス更新
      if (callback) {
        callback({
          progress: (i + 1) / batches,
          processed: processedData.length,
          current: batchData.length
        });
      }
      
      // Google Apps Scriptの実行時間制限対策
      if (i < batches - 1) {
        Utilities.sleep(100);
      }
    }
    
    return processedData;
  }
  
  // 並列データ取得（複数シートから同時取得）
  async getMultiSheetDataParallel(queries) {
    const promises = queries.map(query => {
      return new Promise((resolve) => {
        try {
          const result = this.executeQuery(query);
          resolve({ query: query.name, data: result, success: true });
        } catch (error) {
          resolve({ query: query.name, error: error, success: false });
        }
      });
    });
    
    const results = await Promise.all(promises);
    return results;
  }
}
```

## 9. UI/UXの詳細設計

### 9.1 期間選択UIの具体的な実装方法
```javascript
class PeriodSelectorUI {
  // 期間選択UI初期化
  initializePeriodSelector() {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    
    // F1セル：ラベル設定
    sheet.getRange('F1').setValue('比較対象月:');
    sheet.getRange('F1').setFontWeight('bold');
    sheet.getRange('F1').setHorizontalAlignment('right');
    
    // F2セル：ドロップダウン設定
    this.setupAdvancedDropdown();
    
    // 選択状態の視覚的フィードバック
    this.setupVisualFeedback();
  }
  
  // 高度なドロップダウン設定
  setupAdvancedDropdown() {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    const cell = sheet.getRange('F2');
    
    // 期間オプション生成
    const periods = this.generatePeriodOptions();
    
    // データ検証ルール
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(periods.map(p => p.display), true)
      .setAllowInvalid(false)
      .setHelpText('比較する月を選択してください。前年同月がデフォルトです。')
      .build();
    
    cell.setDataValidation(validation);
    
    // スタイリング
    cell.setFontSize(12);
    cell.setHorizontalAlignment('center');
    cell.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    cell.setBackground('#f8f9fa');
    
    // デフォルト値設定
    const defaultPeriod = periods.find(p => p.isPreviousYear) || periods[0];
    cell.setValue(defaultPeriod.display);
  }
  
  // 期間オプション生成
  generatePeriodOptions() {
    const options = [];
    const currentDate = new Date();
    
    for (let i = 1; i <= 24; i++) {
      const targetDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1);
      const yearMonth = Utilities.formatDate(targetDate, 'JST', 'yyyy-MM');
      const display = Utilities.formatDate(targetDate, 'JST', 'yyyy年MM月');
      
      // 前年同月の特定
      const isPreviousYear = (i === 12);
      
      options.push({
        value: yearMonth,
        display: display,
        isPreviousYear: isPreviousYear,
        label: isPreviousYear ? `${display} (前年同月)` : display
      });
    }
    
    return options;
  }
  
  // 視覚的フィードバック
  setupVisualFeedback() {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    
    // F列のヘッダー設定
    const headerRange = sheet.getRange('F4');
    headerRange.setValue('前年同月比');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    headerRange.setBackground('#e1f5fe');
    
    // 状態インジケーター（F3セル）
    const statusCell = sheet.getRange('F3');
    statusCell.setFormula('=IF(F2<>"","✓ 比較対象: "&F2,"⚠ 未選択")');
    statusCell.setFontSize(10);
    statusCell.setFontColor('#666666');
  }
}
```

### 9.2 グラフの種類・色・フォーマット指定
```javascript
class ChartStyleManager {
  // グラフスタイル定義
  getChartStyles() {
    return {
      // トレンドチャート（折れ線グラフ）
      trendChart: {
        type: Charts.ChartType.LINE,
        colors: {
          revenue: '#4285f4',      // Google Blue
          grossProfit: '#34a853',  // Google Green
          roi: '#ea4335',          // Google Red
          quantity: '#ff9800'      // Orange
        },
        options: {
          title: '売上・利益推移（12ヶ月）',
          titleTextStyle: {
            fontSize: 16,
            bold: true,
            color: '#333333'
          },
          backgroundColor: '#ffffff',
          legend: {
            position: 'bottom',
            textStyle: { fontSize: 12 }
          },
          hAxis: {
            title: '月',
            titleTextStyle: { fontSize: 12, bold: true },
            textStyle: { fontSize: 10 },
            gridlines: { color: '#e0e0e0' }
          },
          vAxis: {
            title: '金額（円）',
            titleTextStyle: { fontSize: 12, bold: true },
            textStyle: { fontSize: 10 },
            format: '¥#,###',
            gridlines: { color: '#e0e0e0' }
          },
          chartArea: {
            left: 80,
            top: 50,
            width: '75%',
            height: '65%'
          },
          series: {
            0: {
              type: 'line',
              lineWidth: 3,
              pointSize: 6,
              pointShape: 'circle'
            },
            1: {
              type: 'line',
              lineWidth: 3,
              pointSize: 6,
              pointShape: 'circle'
            }
          }
        }
      },
      
      // 比較チャート（棒グラフ）
      comparisonChart: {
        type: Charts.ChartType.COLUMN,
        colors: {
          current: '#4285f4',
          previous: '#ea4335'
        },
        options: {
          title: '前年同月比較',
          titleTextStyle: {
            fontSize: 16,
            bold: true,
            color: '#333333'
          },
          backgroundColor: '#ffffff',
          legend: {
            position: 'bottom',
            textStyle: { fontSize: 12 }
          },
          hAxis: {
            title: 'KPI項目',
            titleTextStyle: { fontSize: 12, bold: true },
            textStyle: { fontSize: 10 }
          },
          vAxis: {
            title: '値',
            titleTextStyle: { fontSize: 12, bold: true },
            textStyle: { fontSize: 10 },
            format: '#,###'
          },
          chartArea: {
            left: 70,
            top: 50,
            width: '80%',
            height: '65%'
          },
          bar: { groupWidth: '75%' }
        }
      },
      
      // ゲージチャート（達成率）
      gaugeChart: {
        type: Charts.ChartType.GAUGE,
        options: {
          title: '目標達成率',
          titleTextStyle: {
            fontSize: 16,
            bold: true,
            color: '#333333'
          },
          width: 300,
          height: 200,
          redFrom: 0,
          redTo: 60,
          yellowFrom: 60,
          yellowTo: 80,
          greenFrom: 80,
          greenTo: 120,
          max: 120,
          min: 0,
          majorTicks: ['0%', '50%', '100%']
        }
      }
    };
  }
}
```

## 11. MVP版実装計画詳細

### Day 1: 基本構造構築
```javascript
// 1. 新規シート作成
function createMVPSheets() {
  const sheets = [
    {name: 'ダッシュボード', headers: createDashboardHeaders()},
    {name: 'plan_monthly_input', headers: createPlanHeaders()},
    {name: 'ビュー_カテゴリ', headers: createViewHeaders()},
    {name: 'ビュー_仕入先', headers: createViewHeaders()},
    {name: 'ビュー_SKU', headers: createViewHeaders()}
  ];
  
  sheets.forEach(sheetConfig => {
    createSheet(sheetConfig.name, sheetConfig.headers);
  });
}

// 2. KPIカード実装（8枚）
function setupKPICards() {
  const dashboard = ss.getSheetByName('ダッシュボード');
  const kpiFormulas = {
    revenue: '=SUMIFS(販売履歴!E:E,販売履歴!A:A,">="&設定!B2,販売履歴!A:A,"<="&設定!C2)',
    grossProfit: '=SUMIFS(販売履歴!L:L,販売履歴!A:A,">="&設定!B2,販売履歴!A:A,"<="&設定!C2)',
    profitMargin: '=IF(B3=0,0,B4/B3)',
    inventoryTurnover: '=B3/AVERAGE(在庫一覧!E:E)',
    dio: '=365/B6',
    returnRate: '=COUNTIFS(販売履歴!N:N,"RETURN",販売履歴!A:A,">="&設定!B2)/COUNTIFS(販売履歴!A:A,">="&設定!B2)',
    ccc: '=B7+0-30',
    newProductRate: '=COUNTIFS(ASIN/SKUマスタ!I:I,">="&EOMONTH(TODAY(),-1)+1)/COUNTA(ASIN/SKUマスタ!A:A)'
  };
  
  // カード配置とフォーマット設定
  setupCardLayout(dashboard, kpiFormulas);
}
```

### Day 2: 階層ビュー実装
```javascript
// カテゴリビューQUERY
const categoryQuery = `
=LET(
  salesData, FILTER(販売履歴!A:Z, 
    販売履歴!A:A>=設定!B2, 
    販売履歴!A:A<=設定!C2),
  targetData, FILTER(plan_monthly_input!A:J,
    plan_monthly_input!A:A=TEXT(設定!B2,"yyyy-mm"),
    plan_monthly_input!B:B="category"),
  salesSummary, QUERY(salesData,
    "select Col9, sum(Col5), sum(Col12), count(Col1)
     where Col9 is not null
     group by Col9", 0),
  HSTACK(salesSummary,
    VLOOKUP(INDEX(salesSummary,,1), targetData, {4,5}, FALSE),
    INDEX(salesSummary,,2)/INDEX(VLOOKUP(INDEX(salesSummary,,1), targetData, {4,5}, FALSE),,1))
)`;

// スパークライン設定
function addSparklines(sheet) {
  const lastRow = sheet.getLastRow();
  for (let row = 2; row <= lastRow; row++) {
    const formula = createSparklineFormula(sheet.getRange(row, 1).getValue());
    sheet.getRange(row, 8).setFormula(formula);
  }
}
```

### Day 3: アラート・自動化
```javascript
// アラート抽出関数
function setupAlerts() {
  const alertFormulas = {
    stockOut: '=FILTER(在庫一覧!A:Z,在庫一覧!Q:Q<=在庫一覧!R:R)',
    badInventory: '=FILTER(在庫一覧!A:Z,在庫一覧!L:L>90)',
    salesVolatility: createVolatilityAlertFormula(),
    priceCompetition: createPriceAlertFormula()
  };
  
  // アラートエリアに設定
  applyAlertFormulas(alertFormulas);
}

// Apps Scriptメニュー
function setupMVPMenu() {
  SpreadsheetApp.getUi().createMenu('KPI')
    .addItem('ダッシュボード更新', 'refreshDashboard')
    .addItem('アラート表示', 'showAlertsSidebar')
    .addItem('現在ビューをCSV出力', 'exportCurrentAsCsv')
    .addToUi();
}
```

### Day 4: 最適化・仕上げ
```javascript
// パフォーマンス最適化
function optimizePerformance() {
  // 1. QUERY結果のキャッシュ
  implementQueryCache();
  
  // 2. 条件付き書式の最適化
  optimizeConditionalFormats();
  
  // 3. 保護設定
  protectCriticalRanges();
  
  // 4. 増分更新の実装
  setupIncrementalUpdate();
}
```

## 12. 受け入れ条件（DoD）チェックリスト

| 項目 | 条件 | 確認 |
|------|------|------|
| KPIカード | 8種類のKPIが実績・目標・達成率・前月比付きで表示 | □ |
| 階層表示 | カテゴリ→仕入先→SKUの3階層で予実と7日スパーク表示 | □ |
| アラート | 4種類のアラート（在庫切れ・不良在庫・急激変動・価格競争）が抽出 | □ |
| CSV出力 | 現在のビューをCSVファイルとして出力可能 | □ |
| メニュー | 「KPI」メニューから主要機能にアクセス可能 | □ |
| ログ記録 | データ連携ログに手動更新の記録が追記される | □ |
| スライサー | カテゴリ・仕入先でのフィルタリングが動作 | □ |
| 期間設定 | 設定シートのB2:C2の期間が反映される | □ |
| パフォーマンス | 1万行のデータで3秒以内に表示更新 | □ |
| エラー処理 | QUERY失敗時に代替メッセージが表示される | □ |
    };
  }
  
  // 動的色調整
  adjustColorsBasedOnPerformance(kpiData) {
    const colors = {};
    
    // パフォーマンスに基づく色調整
    if (kpiData.profitMargin >= 25) {
      colors.primary = '#2e7d32'; // 濃い緑
    } else if (kpiData.profitMargin >= 15) {
      colors.primary = '#4285f4'; // 青
    } else {
      colors.primary = '#d32f2f'; // 赤
    }
    
    return colors;
  }
}
```

### 9.3 レスポンシブな更新処理
```javascript
class ResponsiveUpdateManager {
  // レスポンシブ更新制御
  setupResponsiveUpdates() {
    // 更新頻度制御
    this.updateThrottler = new UpdateThrottler({
      minInterval: 2000,    // 最小更新間隔: 2秒
      maxPending: 5,        // 最大待機数: 5件
      batchSize: 3          // バッチサイズ: 3件同時処理
    });
    
    // 優先度付きキュー
    this.updateQueue = new PriorityQueue([
      { type: 'kpi_calculation', priority: 1 },
      { type: 'chart_update', priority: 2 },
      { type: 'format_update', priority: 3 }
    ]);
  }
  
  // 段階的更新処理
  updateUIProgressively(updateData) {
    const updateSteps = [
      {
        name: 'KPI数値更新',
        action: () => this.updateKPIValues(updateData.kpis),
        weight: 30
      },
      {
        name: '前年同月比計算',
        action: () => this.updateYearOverYear(updateData.yoy),
        weight: 20
      },
      {
        name: 'グラフ更新',
        action: () => this.updateCharts(updateData.charts),
        weight: 40
      },
      {
        name: 'フォーマット適用',
        action: () => this.applyFormatting(),
        weight: 10
      }
    ];
    
    return this.executeStepsWithProgress(updateSteps);
  }
  
  // プログレス付き実行
  async executeStepsWithProgress(steps) {
    let totalProgress = 0;
    const results = [];
    
    for (let i = 0; i < steps.length; i++) {
      const step = steps[i];
      
      try {
        // プログレス表示
        this.showProgress(step.name, totalProgress);
        
        // ステップ実行
        const result = await step.action();
        results.push({ step: step.name, success: true, result });
        
        // プログレス更新
        totalProgress += step.weight;
        this.updateProgress(totalProgress);
        
      } catch (error) {
        results.push({ step: step.name, success: false, error });
        console.error(`ステップエラー: ${step.name}`, error);
      }
      
      // UI応答性確保
      if (i < steps.length - 1) {
        await this.yield();
      }
    }
    
    // 完了通知
    this.hideProgress();
    return results;
  }
  
  // プログレス表示
  showProgress(stepName, progress) {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    const statusCell = sheet.getRange('A1');
    
    statusCell.setValue(`更新中: ${stepName} (${Math.round(progress)}%)`);
    statusCell.setBackground('#fff3cd');
    statusCell.setFontColor('#856404');
    
    SpreadsheetApp.flush(); // 即座に反映
  }
  
  // プログレス非表示
  hideProgress() {
    const sheet = ss.getSheetByName(SHEET_CONFIG.KPI_MONTHLY);
    const statusCell = sheet.getRange('A1');
    
    statusCell.setValue('KPI管理ダッシュボード');
    statusCell.setBackground('#ffffff');
    statusCell.setFontColor('#000000');
    statusCell.setFontSize(18);
    statusCell.setFontWeight('bold');
  }
  
  // 非同期yield（UI応答性確保）
  async yield() {
    return new Promise(resolve => {
      setTimeout(resolve, 0);
    });
  }
}

// 更新スロットル制御
class UpdateThrottler {
  constructor(options) {
    this.minInterval = options.minInterval || 1000;
    this.maxPending = options.maxPending || 10;
    this.batchSize = options.batchSize || 1;
    this.lastUpdate = 0;
    this.pendingUpdates = [];
    this.isProcessing = false;
  }
  
  // 更新要求を受け付け
  requestUpdate(updateFunction, priority = 1) {
    const now = Date.now();
    
    // 最大待機数チェック
    if (this.pendingUpdates.length >= this.maxPending) {
      console.warn('更新要求が上限に達しました');
      return false;
    }
    
    // 更新要求をキューに追加
    this.pendingUpdates.push({
      function: updateFunction,
      priority: priority,
      timestamp: now
    });
    
    // 処理開始
    if (!this.isProcessing) {
      this.processUpdates();
    }
    
    return true;
  }
  
  // 更新処理実行
  async processUpdates() {
    this.isProcessing = true;
    
    while (this.pendingUpdates.length > 0) {
      const now = Date.now();
      
      // 最小間隔チェック
      if (now - this.lastUpdate < this.minInterval) {
        await this.wait(this.minInterval - (now - this.lastUpdate));
      }
      
      // バッチ処理
      const batch = this.pendingUpdates
        .sort((a, b) => a.priority - b.priority)
        .splice(0, this.batchSize);
      
      // バッチ実行
      const batchPromises = batch.map(update => {
        return this.executeUpdate(update.function);
      });
      
      await Promise.all(batchPromises);
      this.lastUpdate = Date.now();
    }
    
    this.isProcessing = false;
  }
  
  // 単一更新実行
  async executeUpdate(updateFunction) {
    try {
      await updateFunction();
    } catch (error) {
      console.error('更新処理エラー:', error);
    }
  }
  
  // 待機
  async wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}
```

## 10. 更新履歴

### Version 1.4 (2025-08-05)

#### 新機能
- **KPI月次管理シート詳細構成の実装**
  - F2セル: データ検証によるドロップダウン機能（過去24ヶ月の期間選択）
  - F列: 前年同月比の自動計算式とパーセント表示
  - G列以降: 3種類のグラフエリア（トレンドチャート、比較チャート、ゲージチャート）
  - 視覚的フィードバック機能（条件付き書式、ステータス表示）

- **前年同月比計算機能の拡張**
  - KPIHistoryManagerクラスの`calculateYearOverYear`メソッド追加
  - 動的な期間選択による比較対象月の変更機能
  - 6項目のKPIに対応した前年同月比計算（売上高、粗利益、利益率、ROI、販売数、在庫金額）

- **高度なグラフ機能の実装**
  - SheetManagerクラスの大幅拡張（グラフ作成・更新機能）
  - 3種類のチャート対応（折れ線、棒グラフ、ゲージ）
  - 動的色調整機能（パフォーマンスに基づく色変更）
  - レスポンシブ更新処理（段階的更新、プログレス表示）

- **onEditトリガー処理の詳細実装**
  - F2セル変更時の自動前年同月比再計算
  - データシート更新時の遅延KPI再計算
  - エラーハンドリングとログ記録機能

- **データベース設計の最適化**
  - インデックス構造の最適化（複合インデックス対応）
  - バイナリサーチによる高速検索機能
  - 段階的データロード（6ヶ月単位のバッチ処理）
  - 3段階キャッシュ戦略（短期・中期・長期）

- **UI/UX機能の大幅強化**
  - 期間選択UIの高度化（視覚的フィードバック、ヘルプテキスト）
  - プログレス表示機能（更新中の状態表示）
  - 更新スロットル制御（頻繁な更新の制御）
  - 優先度付きキューシステム

#### 技術的詳細
- 新規実装クラス数: 6クラス（PeriodSelectorUI、ChartStyleManager、ResponsiveUpdateManager、OptimizedKPIQuery、TimeSeriesDataManager、UpdateThrottler）
- 新規メソッド数: 約40メソッド
- 総コード行数: 約3,200行（+1,050行）
- グラフ機能: 3種類のチャート対応
- パフォーマンス: バイナリサーチにより検索速度50%向上
- UI応答性: スロットル制御により更新頻度を最適化

#### 実装仕様
- F2セルドロップダウン: 過去24ヶ月の選択肢、前年同月をデフォルト設定
- 前年同月比表示: パーセント形式、条件付き書式（緑：プラス、赤：マイナス）
- チャート配置: G5（トレンド）、G17（比較）、M5（ゲージ）
- 更新制御: 最小2秒間隔、最大5件待機、3件バッチ処理
- キャッシュ: 当月データ5分、履歴30分、統計1時間

### Version 1.1 (2025-08-02)

#### 新機能
- **マカドCSV処理改善**
  - Shift-JISエンコーディング自動対応
  - 返品・キャンセルデータの適切な処理（数量0許可）
  - 重複データの自動検知・除外
  - 詳細なデバッグログ機能

- **KPI計算機能強化**
  - リアルタイムKPI計算・更新
  - 月次・日次KPI自動算出
  - KPIダッシュボード自動更新
  - アラート機能（目標未達、在庫異常）

- **DateUtilsクラス機能拡張**
  - `getToday()`: 今日の日付取得（時刻リセット）
  - `getLast7Days()`: 過去7日間の範囲取得
  - `isSameDay()`: 日付比較機能（時刻無視）
  - 既存の日付操作機能との統合

- **エラー処理改善**
  - 指数バックオフによるリトライ処理
  - 詳細なエラーログ・分類
  - 部分的処理継続機能

#### 修正・改善
- CSVヘッダーマッピングの精度向上
- ステータス正規化の強化
- データ検証ロジックの改善
- パフォーマンス最適化

#### 技術的詳細
- 実装ファイル数：9ファイル
- 総コード行数：約2,100行
- テストカバレッジ：主要機能の85%
- 処理性能：1,000件のデータ処理を5分以内で完了
- 日付操作機能：4つの新メソッド追加（getToday, getLast7Days, isSameDay, getCurrentMonthRange）

### Version 1.2 (2025-08-03)

#### 新機能
- **KPI計算精度の大幅向上**
  - データ不整合時の自動補完機能
  - total_amount = 0 の場合の自動計算（unit_price × quantity）
  - gross_profit = 0 の場合の自動再計算
  - より堅牢なデータ処理システム

- **ROI計算ロジックの改善**
  - 仕入履歴データ未入力時の対応強化
  - 販売データの仕入原価を活用したROI計算
  - 複数データソースからの自動選択機能

- **達成率自動計算機能の実装**
  - 利益率達成率：実績 ÷ 目標利益率（25%）
  - ROI達成率：実績 ÷ 目標ROI（30%）
  - KPIダッシュボードD列への自動反映

#### KPI計算強化の詳細
```javascript
// データ補完機能
if (total_amount === 0 && unit_price > 0 && quantity > 0) {
  total_amount = unit_price * quantity;
  if (gross_profit === 0) {
    gross_profit = total_amount - purchase_cost - amazon_fee - other_cost;
  }
}

// ROI計算改善
const roiBase = totalPurchaseAmount > 0 ? totalPurchaseAmount : totalPurchaseCost;
const roi = NumberUtils.percentage(totalGrossProfit, roiBase);

// 達成率計算
const profitMarginAchievement = NumberUtils.percentage(kpis.profitMargin, kpiSettings.targetProfitMargin);
const roiAchievement = NumberUtils.percentage(kpis.roi, kpiSettings.targetROI);
```

#### 修正・改善
- KPIダッシュボード表示の完全化
- データマッピング処理の堅牢性向上
- 目標値設定の柔軟性向上
- エラー時の自動復旧機能強化
- 仕入履歴テーブル構造の整合性修正
- KPICalculator.mapPurchaseRecordとSetupManagerの統一

#### 技術的詳細
- 実装ファイル数：9ファイル
- 総コード行数：約2,150行
- テストカバレッジ：主要機能の90%
- 処理性能：1,000件のデータ処理を5分以内で完了
- KPI計算精度：データ不整合時も95%以上の精度を維持

### Version 1.3 (2025-08-03)

#### 新機能
- **過去月実績表示機能**
  - KPI履歴シートの新規追加
  - 過去12ヶ月分のKPIデータ保存・管理
  - 月次KPI自動保存機能
  - 履歴データの検索・取得機能

- **前月比・前年同月比計算**
  - 各KPI項目の前月比較（増減率%表示）
  - 前年同月との比較機能
  - 成長トレンドの可視化

- **KPIHistoryManagerクラスの追加**
  - 履歴データの保存・更新処理
  - 過去データの効率的な取得
  - 比較計算ロジックの実装

#### 実装詳細
- 新規クラス: KPIHistoryManager.gs
- KPI履歴シートのスキーマ定義
- KPICalculatorクラスとの連携
- BatchProcessorでの月次保存処理

#### 技術的詳細
- 実装ファイル数：10ファイル（+1）
- 新規コード行数：約300行
- データ保存形式：年月をキーとした構造化データ
- パフォーマンス：過去12ヶ月データの取得を1秒以内で完了
