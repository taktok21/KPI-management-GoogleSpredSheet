/**
 * マカドCSV処理クラス
 * 
 * マカドから出力されたCSVファイルの読み込み、解析、データ変換を行います。
 * Shift-JISエンコーディングの対応、SKU統一処理も含みます。
 */

class MakadoProcessor {
  constructor() {
    this.cache = new CacheUtils();
    this.errorHandler = new ErrorHandler();
  }

  // =============================================================================
  // CSVファイル処理
  // =============================================================================

  /**
   * マカドCSVファイルをインポート
   */
  importCSV(fileName) {
    return ErrorHandler.executeWithRetry(() => {
      return this.processCSVFile(fileName);
    }, 'MakadoProcessor.importCSV');
  }

  /**
   * CSVファイル処理の実行
   */
  processCSVFile(fileName) {
    try {
      // ファイル取得
      const file = this.findCSVFile(fileName);
      if (!file) {
        throw new Error(`CSVファイルが見つかりません: ${fileName}`);
      }

      // ファイル内容読み込み
      const content = this.readCSVContent(file);
      
      // データ解析
      const parsedData = this.parseCSVData(content);
      
      // データ変換・検証
      const processedData = this.processData(parsedData);
      
      // スプレッドシートに保存
      const result = this.saveToSpreadsheet(processedData);
      
      // SKUマッピング更新
      this.updateSKUMapping(processedData);
      
      console.log(`ファイル処理完了: ${fileName}, 処理件数: ${processedData.length}`);
      
      return {
        success: true,
        fileName: fileName,
        recordCount: processedData.length,
        importedAt: new Date(),
        ...result
      };

    } catch (error) {
      ErrorHandler.handleError(error, 'MakadoProcessor.processCSVFile', { fileName });
      throw error;
    }
  }

  /**
   * CSVファイルを検索
   */
  findCSVFile(fileName) {
    try {
      // Googleドライブから検索
      const files = DriveApp.getFilesByName(fileName);
      
      if (files.hasNext()) {
        return files.next();
      }
      
      // 部分一致で検索
      const searchQuery = `title contains "${fileName.split('.')[0]}"`;
      const searchFiles = DriveApp.searchFiles(searchQuery);
      
      while (searchFiles.hasNext()) {
        const file = searchFiles.next();
        if (file.getName().includes('マカド') && file.getName().endsWith('.csv')) {
          return file;
        }
      }
      
      return null;

    } catch (error) {
      throw new Error(`ファイル検索に失敗しました: ${error.message}`);
    }
  }

  /**
   * CSVファイル内容を読み込み
   */
  readCSVContent(file) {
    try {
      // ファイルサイズチェック
      const maxSize = 10 * 1024 * 1024; // 10MB
      if (file.getSize() > maxSize) {
        throw new Error('ファイルサイズが大きすぎます（10MB以下にしてください）');
      }

      console.log(`ファイル読み込み開始: ${file.getName()}, サイズ: ${file.getSize()}bytes`);

      // Blobとして取得
      const blob = file.getBlob();
      
      // Shift-JISからUTF-8に変換
      let content;
      try {
        content = blob.getDataAsString('Shift_JIS');
        console.log('Shift-JISエンコーディングで読み込み成功');
      } catch (encodingError) {
        // フォールバック: UTF-8として読み込み
        console.warn('Shift-JIS読み込みに失敗、UTF-8で試行:', encodingError);
        content = blob.getDataAsString('UTF-8');
        console.log('UTF-8エンコーディングで読み込み成功');
      }

      // 最初の100文字を表示（デバッグ用）
      console.log('読み込み内容サンプル:', content.substring(0, 100));

      return content;

    } catch (error) {
      throw new Error(`ファイル読み込みに失敗しました: ${error.message}`);
    }
  }

  // =============================================================================
  // データ解析・変換
  // =============================================================================

  /**
   * CSVデータを解析
   */
  parseCSVData(content) {
    try {
      const lines = content.split('\n').filter(line => line.trim());
      
      console.log(`CSV解析開始: 全${lines.length}行`);
      
      if (lines.length < 2) {
        throw new Error('CSVファイルにデータがありません');
      }

      // ヘッダー行取得
      const headerLine = lines[0];
      const headers = CSVUtils.parseCSVLine(headerLine);
      
      console.log('CSVヘッダー:', headers);
      
      // カラムマッピング定義
      const columnMapping = this.getColumnMapping();
      
      console.log('カラムマッピング:', columnMapping);
      
      // データ行解析
      const records = [];
      const errors = [];
      
      for (let i = 1; i < lines.length; i++) {
        try {
          const line = lines[i];
          if (!line.trim()) continue;
          
          const values = CSVUtils.parseCSVLine(line);
          const record = this.mapCSVRecord(headers, values, columnMapping);
          
          if (record) {
            records.push(record);
          }
        } catch (error) {
          errors.push({
            line: i + 1,
            content: lines[i],
            error: error.message
          });
        }
      }

      console.log(`CSV解析完了: 成功${records.length}件、エラー${errors.length}件`);

      // エラーがある場合は警告
      if (errors.length > 0) {
        console.warn(`CSV解析で${errors.length}行のエラーがありました:`, errors.slice(0, 5));
      }

      return {
        records: records,
        totalLines: lines.length - 1,
        successLines: records.length,
        errorLines: errors.length,
        errors: errors
      };

    } catch (error) {
      throw new Error(`CSV解析に失敗しました: ${error.message}`);
    }
  }

  /**
   * カラムマッピング定義
   */
  getColumnMapping() {
    return {
      '注文日': 'order_date',
      '商品名': 'product_name',
      'オーダーID': 'order_id',
      'ASIN': 'asin',
      'SKU': 'makado_sku',
      'コンディション': 'condition',
      '配送経路': 'fulfillment',
      '販売価格': 'unit_price',
      '送料': 'shipping_fee',
      'ポイント': 'points',
      '割引': 'discount',
      '仕入れ価格': 'purchase_cost',
      'その他経費': 'other_cost',
      'Amazon手数料': 'amazon_fee',
      '粗利': 'gross_profit',
      'ステータス': 'status',
      '販売数': 'quantity',
      '累計販売数': 'cumulative_quantity'
    };
  }

  /**
   * CSVレコードをマッピング
   */
  mapCSVRecord(headers, values, columnMapping) {
    const record = {};
    
    console.log(`レコードマッピング開始: ヘッダー数=${headers.length}, 値数=${values.length}`);
    
    // ヘッダーとバリューの対応付け
    headers.forEach((header, index) => {
      const mappedKey = columnMapping[header];
      if (mappedKey && values[index] !== undefined) {
        const value = this.convertValue(values[index], mappedKey);
        record[mappedKey] = value;
        
        if (index < 5) { // 最初の5列のみログ出力
          console.log(`マッピング: ${header} (${index}) -> ${mappedKey} = ${value}`);
        }
      } else if (index < 5) {
        console.log(`マッピング対象外: ${header} (${index}) = ${values[index]}`);
      }
    });

    console.log('マッピング後のレコード構造:', Object.keys(record));

    // 必須フィールドチェック
    const requiredFields = ['order_date', 'order_id', 'asin', 'makado_sku'];
    const missingFields = [];
    
    for (const field of requiredFields) {
      if (!record[field]) {
        missingFields.push(field);
      }
    }
    
    if (missingFields.length > 0) {
      throw new Error(`必須フィールドが不足しています: ${missingFields.join(', ')}`);
    }

    // 統一SKU生成
    record.unified_sku = this.generateUnifiedSKU(record.asin, record.makado_sku);
    
    // データソース識別
    record.data_source = 'MAKADO';
    record.import_timestamp = new Date();

    return record;
  }

  /**
   * データ型変換
   */
  convertValue(value, fieldName) {
    if (!value || value === '') return null;

    // 文字列のトリミング
    if (typeof value === 'string') {
      value = value.trim();
    }

    switch (fieldName) {
      case 'order_date':
        return this.parseDate(value);
        
      case 'unit_price':
      case 'shipping_fee':
      case 'points':
      case 'discount':
      case 'purchase_cost':
      case 'other_cost':
      case 'amazon_fee':
      case 'gross_profit':
        return NumberUtils.safeNumber(value);
        
      case 'quantity':
      case 'cumulative_quantity':
        return NumberUtils.safeInteger(value);
        
      case 'asin':
        return this.validateASIN(value);
        
      case 'makado_sku':
        return StringUtils.normalizeSKU(value);
        
      case 'order_id':
        return this.validateOrderID(value);
        
      case 'fulfillment':
        return this.normalizeFulfillment(value);
        
      case 'status':
        return this.normalizeStatus(value);
        
      default:
        return value;
    }
  }

  // =============================================================================
  // データ処理・検証
  // =============================================================================

  /**
   * データ処理
   */
  processData(parsedData) {
    const processedRecords = [];
    const validationErrors = [];

    parsedData.records.forEach((record, index) => {
      try {
        // データ検証
        const validation = this.validateRecord(record);
        if (!validation.isValid) {
          validationErrors.push({
            index: index,
            record: record,
            errors: validation.errors
          });
          return;
        }

        // データ補正
        const correctedRecord = this.correctRecord(record);
        
        // 重複チェック
        if (!this.isDuplicate(correctedRecord)) {
          processedRecords.push(correctedRecord);
        }

      } catch (error) {
        validationErrors.push({
          index: index,
          record: record,
          errors: [error.message]
        });
      }
    });

    // 検証エラーがある場合は警告
    if (validationErrors.length > 0) {
      console.warn(`データ検証で${validationErrors.length}件のエラーがありました`);
      this.logValidationErrors(validationErrors);
    }

    return processedRecords;
  }

  /**
   * レコード検証
   */
  validateRecord(record) {
    const errors = [];

    // 日付検証
    if (!record.order_date || !(record.order_date instanceof Date)) {
      errors.push('注文日時が不正です');
    }

    // 金額検証
    if (record.unit_price < 0) {
      errors.push('販売価格が負の値です');
    }

    if (record.purchase_cost < 0) {
      errors.push('仕入価格が負の値です');
    }

    // 数量検証（返品・キャンセルの場合は0を許可）
    if (record.quantity < 0) {
      errors.push('販売数量が負の値です');
    } else if (record.quantity === 0 && record.status !== 'RETURN' && record.status !== 'CANCELLED') {
      errors.push('販売数量が0です（返品・キャンセル以外）');
    }

    // ASIN検証
    if (!ValidationUtils.isValidASIN(record.asin)) {
      errors.push('ASINの形式が不正です');
    }

    // 利益率チェック（警告レベル）
    if (record.unit_price > 0) {
      const profitMargin = NumberUtils.calculateProfitMargin(record.gross_profit, record.unit_price);
      if (profitMargin < -50 || profitMargin > 90) {
        errors.push(`利益率が異常です: ${profitMargin.toFixed(1)}%`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * レコード補正
   */
  correctRecord(record) {
    const corrected = {...record};

    // 日時をJSTに変換
    corrected.order_date = DateUtils.toJST(corrected.order_date);

    // 粗利益の再計算（整合性チェック）
    const calculatedProfit = this.calculateGrossProfit(corrected);
    if (Math.abs(calculatedProfit - corrected.gross_profit) > 1) {
      console.warn(`粗利益に差異があります。計算値: ${calculatedProfit}, CSV値: ${corrected.gross_profit}`);
      corrected.gross_profit_calculated = calculatedProfit;
    }

    // 配送方法正規化
    corrected.fulfillment = this.normalizeFulfillment(corrected.fulfillment);

    return corrected;
  }

  /**
   * 粗利益計算
   */
  calculateGrossProfit(record) {
    const revenue = NumberUtils.safeNumber(record.unit_price) * NumberUtils.safeNumber(record.quantity);
    const cost = NumberUtils.safeNumber(record.purchase_cost) * NumberUtils.safeNumber(record.quantity);
    const fees = NumberUtils.safeNumber(record.amazon_fee) + 
                NumberUtils.safeNumber(record.other_cost) - 
                NumberUtils.safeNumber(record.shipping_fee) - 
                NumberUtils.safeNumber(record.points);

    return revenue - cost - fees;
  }

  /**
   * 重複チェック
   */
  isDuplicate(record) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const salesSheet = ss.getSheetByName(SHEET_CONFIG.SALES_HISTORY);
      
      if (!salesSheet) return false;

      // 既存データの最後の100行をチェック（パフォーマンス考慮）
      const lastRow = salesSheet.getLastRow();
      if (lastRow <= 1) return false;

      const startRow = Math.max(2, lastRow - 99);
      const checkRange = salesSheet.getRange(startRow, 1, lastRow - startRow + 1, 4);
      const existingData = checkRange.getValues();

      // 注文ID + ASIN + 注文日で重複チェック
      const checkKey = `${record.order_id}_${record.asin}_${DateUtils.formatDate(record.order_date, 'yyyy-MM-dd')}`;

      for (const row of existingData) {
        const existingKey = `${row[0]}_${row[1]}_${DateUtils.formatDate(row[2], 'yyyy-MM-dd')}`;
        if (existingKey === checkKey) {
          return true;
        }
      }

      return false;

    } catch (error) {
      console.warn('重複チェックでエラー:', error);
      return false;
    }
  }

  // =============================================================================
  // ヘルパー関数
  // =============================================================================

  /**
   * 日付解析
   */
  parseDate(dateString) {
    if (!dateString) return null;

    // 複数の日付フォーマットに対応
    const formats = [
      /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/, // 2025-01-01 10:30:00
      /^\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}:\d{2}$/, // 2025/01/01 10:30:00
      /^\d{4}-\d{2}-\d{2}$/, // 2025-01-01
      /^\d{4}\/\d{2}\/\d{2}$/ // 2025/01/01
    ];

    for (const format of formats) {
      if (format.test(dateString)) {
        const date = new Date(dateString);
        if (!isNaN(date.getTime())) {
          return date;
        }
      }
    }

    throw new Error(`日付形式が不正です: ${dateString}`);
  }

  /**
   * ASIN検証
   */
  validateASIN(asin) {
    if (!asin) throw new Error('ASINが空です');
    
    const normalized = asin.toString().toUpperCase().trim();
    
    if (!ValidationUtils.isValidASIN(normalized)) {
      throw new Error(`ASIN形式が不正です: ${asin}`);
    }
    
    return normalized;
  }

  /**
   * 注文ID検証
   */
  validateOrderID(orderId) {
    if (!orderId) throw new Error('注文IDが空です');
    
    const normalized = orderId.toString().trim();
    
    // Amazon注文IDの基本パターンチェック
    if (!/^\d{3}-\d{7}-\d{7}$/.test(normalized)) {
      console.warn(`注文ID形式が異常です: ${orderId}`);
    }
    
    return normalized;
  }

  /**
   * 配送方法正規化
   */
  normalizeFulfillment(fulfillment) {
    if (!fulfillment) return 'UNKNOWN';
    
    const normalized = fulfillment.toString().toUpperCase().trim();
    
    if (normalized === 'FBA' || normalized.includes('FBA')) {
      return 'FBA';
    } else if (normalized === 'FBM' || normalized.includes('自己発送')) {
      return 'FBM';
    }
    
    return normalized;
  }

  /**
   * ステータス正規化
   */
  normalizeStatus(status) {
    if (!status) return 'UNKNOWN';
    
    const normalized = status.toString().trim().toUpperCase();
    
    const statusMap = {
      'SHIPPED': 'SHIPPED',
      'PENDING': 'PENDING',
      'CANCELLED': 'CANCELLED',
      'CANCELED': 'CANCELLED',
      'REFUNDED': 'REFUNDED',
      'RETURN': 'RETURN',
      'RETURNED': 'RETURN'
    };
    
    return statusMap[normalized] || normalized;
  }

  /**
   * 統一SKU生成
   */
  generateUnifiedSKU(asin, makadoSku) {
    return StringUtils.generateUnifiedSKU(asin, makadoSku);
  }

  // =============================================================================
  // データ保存
  // =============================================================================

  /**
   * スプレッドシートに保存
   */
  saveToSpreadsheet(processedData) {
    try {
      if (!processedData || processedData.length === 0) {
        return { savedCount: 0 };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let salesSheet = ss.getSheetByName(SHEET_CONFIG.SALES_HISTORY);
      
      if (!salesSheet) {
        salesSheet = this.createSalesHistorySheet(ss);
      }

      // データを配列形式に変換
      const dataArray = this.convertToSheetData(processedData);
      
      // バッチで追加
      const startRow = salesSheet.getLastRow() + 1;
      const range = salesSheet.getRange(startRow, 1, dataArray.length, dataArray[0].length);
      range.setValues(dataArray);

      // フォーマット適用
      this.applyDataFormatting(salesSheet, startRow, dataArray.length);

      return {
        savedCount: processedData.length,
        startRow: startRow,
        endRow: startRow + dataArray.length - 1
      };

    } catch (error) {
      throw new Error(`スプレッドシート保存に失敗しました: ${error.message}`);
    }
  }

  /**
   * 販売履歴シート作成
   */
  createSalesHistorySheet(spreadsheet) {
    const sheet = spreadsheet.insertSheet(SHEET_CONFIG.SALES_HISTORY);
    
    // ヘッダー行設定
    const headers = [
      '注文ID', 'ASIN', '注文日時', '統一SKU', 'マカドSKU', '商品名',
      '販売価格', '数量', '合計金額', '仕入原価', 'Amazon手数料', 'その他費用',
      '粗利益', '利益率', 'ステータス', '配送方法', 'データソース', '取込日時'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // 列幅調整
    sheet.autoResizeColumns(1, headers.length);
    
    return sheet;
  }

  /**
   * シートデータ形式に変換
   */
  convertToSheetData(records) {
    return records.map(record => [
      record.order_id,
      record.asin,
      record.order_date,
      record.unified_sku,
      record.makado_sku,
      record.product_name,
      record.unit_price,
      record.quantity,
      record.unit_price * record.quantity,
      record.purchase_cost,
      record.amazon_fee,
      record.other_cost,
      record.gross_profit,
      record.unit_price > 0 ? NumberUtils.calculateProfitMargin(record.gross_profit, record.unit_price * record.quantity) : 0,
      record.status,
      record.fulfillment,
      record.data_source,
      record.import_timestamp
    ]);
  }

  /**
   * データフォーマット適用
   */
  applyDataFormatting(sheet, startRow, rowCount) {
    try {
      // 日付列のフォーマット
      const dateRange = sheet.getRange(startRow, 3, rowCount, 1);
      dateRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
      
      // 金額列のフォーマット
      const priceColumns = [7, 9, 10, 11, 12, 13]; // 販売価格、合計金額、仕入原価、手数料、粗利益
      priceColumns.forEach(col => {
        const range = sheet.getRange(startRow, col, rowCount, 1);
        range.setNumberFormat('¥#,##0');
      });
      
      // パーセント列のフォーマット
      const percentRange = sheet.getRange(startRow, 14, rowCount, 1); // 利益率
      percentRange.setNumberFormat('0.0%');
      
    } catch (error) {
      console.warn('フォーマット適用でエラー:', error);
    }
  }

  // =============================================================================
  // SKUマッピング更新
  // =============================================================================

  /**
   * SKUマッピングを更新
   */
  updateSKUMapping(processedData) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let masterSheet = ss.getSheetByName(SHEET_CONFIG.PRODUCT_MASTER);
      
      if (!masterSheet) {
        masterSheet = this.createProductMasterSheet(ss);
      }

      const mappingUpdates = [];
      
      processedData.forEach(record => {
        if (!this.skuMappingExists(record.unified_sku)) {
          mappingUpdates.push([
            record.unified_sku,
            record.asin,
            record.makado_sku,
            '', // Amazon SKU（後で更新）
            record.product_name,
            '', // カテゴリ
            '', // ブランド
            this.extractPurchaseDate(record.makado_sku),
            new Date(), // 作成日
            new Date()  // 更新日
          ]);
        }
      });

      if (mappingUpdates.length > 0) {
        const startRow = masterSheet.getLastRow() + 1;
        const range = masterSheet.getRange(startRow, 1, mappingUpdates.length, 10);
        range.setValues(mappingUpdates);
      }

    } catch (error) {
      console.warn('SKUマッピング更新でエラー:', error);
    }
  }

  /**
   * 商品マスターシート作成
   */
  createProductMasterSheet(spreadsheet) {
    const sheet = spreadsheet.insertSheet(SHEET_CONFIG.PRODUCT_MASTER);
    
    const headers = [
      '統一SKU', 'ASIN', 'マカドSKU', 'AmazonSKU', '商品名',
      'カテゴリ', 'ブランド', '仕入日', '作成日', '更新日'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    return sheet;
  }

  /**
   * SKUマッピング存在チェック
   */
  skuMappingExists(unifiedSku) {
    // 簡易実装（実際はより効率的な検索を実装）
    return false;
  }

  /**
   * マカドSKUから仕入日を抽出
   */
  extractPurchaseDate(makadoSku) {
    if (!makadoSku) return null;
    
    const match = makadoSku.match(/(\d{6})$/);
    if (match) {
      const dateStr = match[1];
      const year = 2000 + parseInt(dateStr.substr(0, 2));
      const month = parseInt(dateStr.substr(2, 2)) - 1; // 0ベース
      const day = parseInt(dateStr.substr(4, 2));
      
      return new Date(year, month, day);
    }
    
    return null;
  }

  // =============================================================================
  // ログ・レポート
  // =============================================================================

  /**
   * 検証エラーをログ記録
   */
  logValidationErrors(validationErrors) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName('_CSV検証エラー');
      
      if (!logSheet) {
        logSheet = ss.insertSheet('_CSV検証エラー');
        logSheet.getRange(1, 1, 1, 6).setValues([
          ['日時', 'インデックス', '注文ID', 'ASIN', 'エラー内容', 'レコード']
        ]);
      }

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

    } catch (error) {
      console.error('検証エラーログ記録に失敗:', error);
    }
  }
}

// =============================================================================
// グローバル関数（メニューから呼び出し用）
// =============================================================================

/**
 * マカドCSV取り込み処理
 */
function processMakadoCSV(fileName) {
  try {
    const processor = new MakadoProcessor();
    return processor.importCSV(fileName);
  } catch (error) {
    ErrorHandler.handleError(error, 'processMakadoCSV');
    throw error;
  }
}