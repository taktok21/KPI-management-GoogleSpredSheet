/**
 * マカドCSV処理クラス（最適化版）
 * 
 * 処理時間を短縮するための最適化を実装
 */

class MakadoProcessorOptimized {
  constructor() {
    this.cache = new CacheUtils();
    this.errorHandler = new ErrorHandler();
    // バッチサイズを設定
    this.BATCH_SIZE = 100;
  }

  /**
   * CSVファイル処理の実行（最適化版）
   */
  processCSVFile(fileName) {
    try {
      console.log(`CSVファイル処理開始: ${fileName}`);
      const startTime = new Date();
      
      // ファイル取得
      const file = this.findCSVFile(fileName);
      if (!file) {
        throw new Error(`CSVファイルが見つかりません: ${fileName}`);
      }

      // ファイル内容読み込み（デバッグログを削減）
      const content = this.readCSVContent(file);
      
      // データ解析（ログを最小限に）
      const parsedData = this.parseCSVDataOptimized(content);
      
      // バッチ処理で保存
      const result = this.saveToSpreadsheetBatch(parsedData.records);
      
      // SKUマッピング更新
      this.updateSKUMapping(parsedData.records);
      
      const duration = (new Date() - startTime) / 1000;
      console.log(`処理完了: ${duration}秒, ${result.savedCount}件`);
      
      return {
        success: true,
        fileName: fileName,
        recordCount: result.savedCount,
        duration: duration,
        importedAt: new Date()
      };

    } catch (error) {
      ErrorHandler.handleError(error, 'MakadoProcessorOptimized.processCSVFile', { fileName });
      throw error;
    }
  }

  /**
   * CSVファイル内容を読み込み（ログ削減版）
   */
  readCSVContent(file) {
    try {
      const blob = file.getBlob();
      
      try {
        return blob.getDataAsString('Shift_JIS');
      } catch (encodingError) {
        return blob.getDataAsString('UTF-8');
      }

    } catch (error) {
      throw new Error(`ファイル読み込みに失敗しました: ${error.message}`);
    }
  }

  /**
   * CSVデータを解析（最適化版）
   */
  parseCSVDataOptimized(content) {
    const lines = content.split('\n').filter(line => line.trim());
    
    if (lines.length < 2) {
      throw new Error('CSVファイルにデータがありません');
    }

    // ヘッダー行取得
    const headerLine = lines[0];
    const headers = CSVUtils.parseCSVLine(headerLine);
    
    // カラムマッピング定義
    const columnMapping = this.getColumnMapping();
    
    // データ行解析（ログを5件ごとに）
    const records = [];
    const errors = [];
    
    for (let i = 1; i < lines.length; i++) {
      try {
        const line = lines[i];
        if (!line.trim()) continue;
        
        const values = CSVUtils.parseCSVLine(line);
        const record = this.mapCSVRecordOptimized(headers, values, columnMapping);
        
        if (record) {
          records.push(record);
          
          // 100件ごとに進捗ログ
          if (records.length % 100 === 0) {
            console.log(`処理中: ${records.length}件完了`);
          }
        }
      } catch (error) {
        errors.push({
          line: i + 1,
          error: error.message
        });
      }
    }

    console.log(`解析完了: 成功${records.length}件、エラー${errors.length}件`);

    return {
      records: records,
      totalLines: lines.length - 1,
      successLines: records.length,
      errorLines: errors.length,
      errors: errors
    };
  }

  /**
   * CSVレコードをマッピング（ログ削減版）
   */
  mapCSVRecordOptimized(headers, values, columnMapping) {
    const record = {};
    
    // ヘッダーとバリューの対応付け（ログなし）
    headers.forEach((header, index) => {
      const mappedKey = columnMapping[header];
      if (mappedKey && values[index] !== undefined) {
        record[mappedKey] = this.convertValue(values[index], mappedKey);
      }
    });

    // 必須フィールドチェック
    const requiredFields = ['order_date', 'order_id', 'asin', 'makado_sku'];
    for (const field of requiredFields) {
      if (!record[field]) {
        throw new Error(`必須フィールドが不足しています: ${field}`);
      }
    }

    // 統一SKU生成
    record.unified_sku = this.generateUnifiedSKU(record.asin, record.makado_sku);
    record.data_source = 'MAKADO';
    record.import_timestamp = new Date();

    return record;
  }

  /**
   * スプレッドシートにバッチ保存（高速化版）
   */
  saveToSpreadsheetBatch(processedData) {
    try {
      if (!processedData || processedData.length === 0) {
        return { savedCount: 0 };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let salesSheet = ss.getSheetByName(SHEET_CONFIG.SALES_HISTORY);
      
      if (!salesSheet) {
        salesSheet = this.createSalesHistorySheet(ss);
      }

      // 既存データの重複チェック（最適化版）
      const existingDataCache = this.buildExistingDataCache(salesSheet);
      
      // 重複を除外
      const uniqueData = processedData.filter(record => {
        const checkKey = `${record.order_id}_${record.asin}_${DateUtils.formatDate(record.order_date, 'yyyy-MM-dd')}`;
        return !existingDataCache.has(checkKey);
      });

      if (uniqueData.length === 0) {
        console.log('新規データはありません（すべて重複）');
        return { savedCount: 0 };
      }

      // データを配列形式に変換
      const dataArray = this.convertToSheetData(uniqueData);
      
      // バッチサイズごとに分割して保存
      let savedCount = 0;
      const startRow = salesSheet.getLastRow() + 1;
      
      for (let i = 0; i < dataArray.length; i += this.BATCH_SIZE) {
        const batch = dataArray.slice(i, i + this.BATCH_SIZE);
        const currentRow = startRow + i;
        const range = salesSheet.getRange(currentRow, 1, batch.length, batch[0].length);
        range.setValues(batch);
        savedCount += batch.length;
        
        // 進捗表示
        console.log(`保存中: ${savedCount}/${dataArray.length}件`);
        
        // 処理の一時停止（レート制限対策）
        if (i + this.BATCH_SIZE < dataArray.length) {
          Utilities.sleep(100); // 0.1秒待機
        }
      }

      // フォーマット適用（最後に一括で）
      this.applyDataFormattingBatch(salesSheet, startRow, savedCount);

      return {
        savedCount: savedCount,
        startRow: startRow,
        endRow: startRow + savedCount - 1
      };

    } catch (error) {
      throw new Error(`スプレッドシート保存に失敗しました: ${error.message}`);
    }
  }

  /**
   * 既存データのキャッシュを構築（高速化）
   */
  buildExistingDataCache(salesSheet) {
    const cache = new Set();
    const lastRow = salesSheet.getLastRow();
    
    if (lastRow <= 1) return cache;

    // 最新の1000件のみチェック（パフォーマンス対策）
    const checkRows = Math.min(1000, lastRow - 1);
    const startRow = Math.max(2, lastRow - checkRows + 1);
    
    const range = salesSheet.getRange(startRow, 1, checkRows, 4);
    const existingData = range.getValues();

    existingData.forEach(row => {
      if (row[0]) { // order_idが存在する場合
        const checkKey = `${row[0]}_${row[1]}_${DateUtils.formatDate(new Date(row[2]), 'yyyy-MM-dd')}`;
        cache.add(checkKey);
      }
    });

    console.log(`重複チェック用キャッシュ構築完了: ${cache.size}件`);
    return cache;
  }

  /**
   * フォーマット適用（バッチ版）
   */
  applyDataFormattingBatch(sheet, startRow, rowCount) {
    try {
      // すべてのフォーマットを一括で適用
      const batchFormat = [];
      
      // 日付列
      batchFormat.push({
        range: sheet.getRange(startRow, 3, rowCount, 1),
        format: 'yyyy-mm-dd hh:mm:ss'
      });
      
      // 金額列
      const priceColumns = [7, 9, 10, 11, 12, 13];
      priceColumns.forEach(col => {
        batchFormat.push({
          range: sheet.getRange(startRow, col, rowCount, 1),
          format: '¥#,##0'
        });
      });
      
      // パーセント列
      batchFormat.push({
        range: sheet.getRange(startRow, 14, rowCount, 1),
        format: '0.0%'
      });
      
      // 一括適用
      batchFormat.forEach(item => {
        item.range.setNumberFormat(item.format);
      });
      
    } catch (error) {
      console.warn('フォーマット適用でエラー:', error);
    }
  }

  // =============================================================================
  // 独立実装メソッド（MakadoProcessorへの依存を削除）
  // =============================================================================

  findCSVFile(fileName) {
    try {
      const files = DriveApp.getFilesByName(fileName);
      
      if (files.hasNext()) {
        return files.next();
      }
      
      // フォルダ検索も試行
      const folders = DriveApp.getFolders();
      while (folders.hasNext()) {
        const folder = folders.next();
        const folderFiles = folder.getFilesByName(fileName);
        if (folderFiles.hasNext()) {
          return folderFiles.next();
        }
      }
      
      return null;
    } catch (error) {
      console.error('ファイル検索エラー:', error);
      return null;
    }
  }

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

  convertValue(value, fieldName) {
    if (!value || value === '') return null;
    
    const stringValue = value.toString().trim();
    
    // 数値フィールド
    const numericFields = ['unit_price', 'quantity', 'purchase_cost', 'amazon_fee', 'other_cost', 'gross_profit'];
    if (numericFields.includes(fieldName)) {
      const numValue = parseFloat(stringValue.replace(/[¥,]/g, ''));
      return isNaN(numValue) ? 0 : numValue;
    }
    
    // 日付フィールド
    if (fieldName === 'order_date') {
      try {
        return new Date(stringValue);
      } catch (error) {
        return new Date();
      }
    }
    
    return stringValue;
  }

  generateUnifiedSKU(asin, makadoSku) {
    return StringUtils.generateUnifiedSKU(asin, makadoSku);
  }

  convertToSheetData(records) {
    return records.map(record => [
      record.order_id || '',
      record.asin || '',
      record.order_date || new Date(),
      record.unified_sku || '',
      record.makado_sku || '',
      record.product_name || '',
      record.unit_price || 0,
      record.quantity || 0,
      record.total_amount || 0,
      record.purchase_cost || 0,
      record.amazon_fee || 0,
      record.other_cost || 0,
      record.gross_profit || 0,
      record.profit_margin || 0,
      record.status || '',
      record.fulfillment || '',
      record.data_source || 'MAKADO',
      record.import_timestamp || new Date()
    ]);
  }

  createSalesHistorySheet(spreadsheet) {
    const sheet = spreadsheet.insertSheet(SHEET_CONFIG.SALES_HISTORY);
    
    const headers = [
      '注文ID', 'ASIN', '注文日時', '統一SKU', 'マカドSKU', '商品名',
      '販売価格', '数量', '合計金額', '仕入原価', 'Amazon手数料', 'その他費用',
      '粗利益', '利益率', 'ステータス', '配送方法', 'データソース', '取込日時'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    return sheet;
  }

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

  skuMappingExists(unifiedSku) {
    // 簡易実装（実際はより効率的な検索を実装）
    return false;
  }

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
}

/**
 * 最適化版のマカドCSV取り込み処理
 */
function processMakadoCSVOptimized(fileName) {
  try {
    const processor = new MakadoProcessorOptimized();
    return processor.processCSVFile(fileName);
  } catch (error) {
    ErrorHandler.handleError(error, 'processMakadoCSVOptimized');
    throw error;
  }
}