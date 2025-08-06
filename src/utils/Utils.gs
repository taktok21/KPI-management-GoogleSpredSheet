/**
 * ユーティリティ関数集
 * 
 * 各種共通処理、日付操作、文字列操作、データ変換などの
 * 汎用的な機能を提供します。
 */

// =============================================================================
// 日付・時刻操作
// =============================================================================

class DateUtils {
  /**
   * 日付をJST（日本標準時）に変換
   */
  static toJST(date) {
    if (!date) return null;
    
    if (typeof date === 'string') {
      date = new Date(date);
    }
    
    // UTCからJSTに変換（+9時間）
    const jstOffset = 9 * 60 * 60 * 1000; // 9時間をミリ秒で
    return new Date(date.getTime() + jstOffset);
  }

  /**
   * 日付を指定フォーマットで文字列に変換
   */
  static formatDate(date, format = 'yyyy-MM-dd') {
    if (!date) return '';
    
    if (typeof date === 'string') {
      date = new Date(date);
    }

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');

    return format
      .replace('yyyy', year)
      .replace('MM', month)
      .replace('dd', day)
      .replace('HH', hours)
      .replace('mm', minutes)
      .replace('ss', seconds);
  }

  /**
   * 日付の差分を日数で取得
   */
  static daysBetween(date1, date2) {
    if (!date1 || !date2) return 0;
    
    const msPerDay = 24 * 60 * 60 * 1000;
    const utc1 = Date.UTC(date1.getFullYear(), date1.getMonth(), date1.getDate());
    const utc2 = Date.UTC(date2.getFullYear(), date2.getMonth(), date2.getDate());
    
    return Math.floor((utc2 - utc1) / msPerDay);
  }

  /**
   * 月の開始日と終了日を取得
   */
  static getMonthRange(year, month) {
    const startDate = new Date(year, month - 1, 1); // 月は0ベース
    const endDate = new Date(year, month, 0); // 次の月の0日 = 今月の最終日
    
    return {
      start: startDate,
      end: endDate
    };
  }

  /**
   * 今月の範囲を取得
   */
  static getCurrentMonthRange() {
    const now = new Date();
    return this.getMonthRange(now.getFullYear(), now.getMonth() + 1);
  }

  /**
   * 今日の日付を取得（時刻を00:00:00にリセット）
   */
  static getToday() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return today;
  }

  /**
   * 過去7日間の範囲を取得
   */
  static getLast7Days() {
    const end = this.getToday();
    const start = new Date(end);
    start.setDate(start.getDate() - 7);
    
    return {
      start: start,
      end: end
    };
  }

  /**
   * 同じ日付かどうかを判定（時刻は無視）
   */
  static isSameDay(date1, date2) {
    if (!date1 || !date2) return false;
    
    const d1 = new Date(date1);
    const d2 = new Date(date2);
    
    return d1.getFullYear() === d2.getFullYear() &&
           d1.getMonth() === d2.getMonth() &&
           d1.getDate() === d2.getDate();
  }

  /**
   * ISO週番号を取得
   */
  static getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  /**
   * 現在の年月を取得（YYYY-MM形式）
   */
  static getCurrentMonth() {
    const now = new Date();
    return this.formatDate(now, 'yyyy-MM');
  }

  /**
   * 日付を年月形式（YYYY-MM）で取得
   */
  static formatMonth(date) {
    if (!date) return '';
    
    if (typeof date === 'string') {
      date = new Date(date);
    }
    
    return this.formatDate(date, 'yyyy-MM');
  }
}

// =============================================================================
// 文字列操作
// =============================================================================

class StringUtils {
  /**
   * 文字列をタイトルケースに変換
   */
  static toTitleCase(str) {
    if (!str) return '';
    
    return str.replace(/\w\S*/g, (txt) => {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
  }

  /**
   * 文字列を切り詰め
   */
  static truncate(str, maxLength, suffix = '...') {
    if (!str || str.length <= maxLength) return str;
    
    return str.substring(0, maxLength - suffix.length) + suffix;
  }

  /**
   * HTML/XMLエスケープ
   */
  static escapeHtml(str) {
    if (!str) return '';
    
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#x27;'
    };
    
    return str.replace(/[&<>"']/g, (m) => map[m]);
  }

  /**
   * 日本語文字エンコーディング変換
   */
  static convertEncoding(str, fromEncoding = 'Shift_JIS', toEncoding = 'UTF-8') {
    try {
      // GASのUtilitiesを使用してエンコーディング変換
      const blob = Utilities.newBlob(str, 'text/plain', fromEncoding);
      return blob.getDataAsString(toEncoding);
    } catch (error) {
      console.warn('エンコーディング変換に失敗:', error);
      return str;
    }
  }

  /**
   * SKU正規化
   */
  static normalizeSKU(sku) {
    if (!sku) return '';
    
    // 大文字に統一、特殊文字除去
    return sku.toString().toUpperCase().replace(/[^A-Z0-9\-_]/g, '');
  }

  /**
   * 統一SKU生成
   */
  static generateUnifiedSKU(asin, originalSku, purchaseDate = null) {
    if (!asin) return '';
    
    let suffix = 'DEFAULT';
    
    if (originalSku && originalSku.includes('MAKAD-')) {
      // マカドSKUから日付抽出
      const match = originalSku.match(/(\d{6})$/);
      if (match) {
        suffix = match[1];
      }
    } else if (purchaseDate) {
      // 購入日から生成
      const date = new Date(purchaseDate);
      const year = date.getFullYear().toString().substr(2);
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      suffix = year + month + day;
    }
    
    return `UNI-${asin}-${suffix}`;
  }
}

// =============================================================================
// 数値・計算操作
// =============================================================================

class NumberUtils {
  /**
   * 安全な数値変換
   */
  static safeNumber(value, defaultValue = 0) {
    if (value === null || value === undefined || value === '') {
      return defaultValue;
    }
    
    const num = parseFloat(value);
    return isNaN(num) ? defaultValue : num;
  }

  /**
   * 安全な整数変換
   */
  static safeInteger(value, defaultValue = 0) {
    if (value === null || value === undefined || value === '') {
      return defaultValue;
    }
    
    const num = parseInt(value);
    return isNaN(num) ? defaultValue : num;
  }

  /**
   * パーセンテージ計算
   */
  static percentage(part, whole, decimalPlaces = 2) {
    if (!whole || whole === 0) return 0;
    
    const percent = (part / whole) * 100;
    return Math.round(percent * Math.pow(10, decimalPlaces)) / Math.pow(10, decimalPlaces);
  }

  /**
   * 数値をカンマ区切りでフォーマット
   */
  static formatNumber(num) {
    const safeNum = this.safeNumber(num);
    return safeNum.toLocaleString('ja-JP');
  }

  /**
   * 通貨フォーマット
   */
  static formatCurrency(amount, currency = 'JPY') {
    const num = this.safeNumber(amount);
    
    if (currency === 'JPY') {
      return `¥${num.toLocaleString('ja-JP')}`;
    }
    
    return num.toLocaleString('ja-JP', {
      style: 'currency',
      currency: currency
    });
  }

  /**
   * ROI計算
   */
  static calculateROI(profit, investment) {
    const profitNum = this.safeNumber(profit);
    const investmentNum = this.safeNumber(investment);
    
    if (investmentNum === 0) return 0;
    
    return this.percentage(profitNum, investmentNum);
  }

  /**
   * 利益率計算
   */
  static calculateProfitMargin(profit, revenue) {
    const profitNum = this.safeNumber(profit);
    const revenueNum = this.safeNumber(revenue);
    
    if (revenueNum === 0) return 0;
    
    return this.percentage(profitNum, revenueNum);
  }

  /**
   * 在庫回転率計算
   */
  static calculateTurnoverRate(salesAmount, averageInventory) {
    const salesNum = this.safeNumber(salesAmount);
    const inventoryNum = this.safeNumber(averageInventory);
    
    if (inventoryNum === 0) return 0;
    
    return salesNum / inventoryNum;
  }

  /**
   * 回転日数計算
   */
  static calculateTurnoverDays(turnoverRate, periodDays = 30) {
    const rate = this.safeNumber(turnoverRate);
    
    if (rate === 0) return Infinity;
    
    return periodDays / rate;
  }
}

// =============================================================================
// 配列・オブジェクト操作
// =============================================================================

class ArrayUtils {
  /**
   * 配列をチャンクに分割
   */
  static chunk(array, size) {
    if (!Array.isArray(array) || size <= 0) return [];
    
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }

  /**
   * 配列から重複を除去
   */
  static unique(array, keyFunc = null) {
    if (!Array.isArray(array)) return [];
    
    if (keyFunc) {
      const seen = new Set();
      return array.filter(item => {
        const key = keyFunc(item);
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      });
    }
    
    return [...new Set(array)];
  }

  /**
   * 配列をキーでグループ化
   */
  static groupBy(array, keyFunc) {
    if (!Array.isArray(array)) return {};
    
    return array.reduce((groups, item) => {
      const key = keyFunc(item);
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {});
  }

  /**
   * 配列の合計値計算
   */
  static sum(array, keyFunc = null) {
    if (!Array.isArray(array)) return 0;
    
    return array.reduce((sum, item) => {
      const value = keyFunc ? keyFunc(item) : item;
      return sum + NumberUtils.safeNumber(value);
    }, 0);
  }

  /**
   * 配列の平均値計算
   */
  static average(array, keyFunc = null) {
    if (!Array.isArray(array) || array.length === 0) return 0;
    
    const total = this.sum(array, keyFunc);
    return total / array.length;
  }

  /**
   * 配列の最大値取得
   */
  static max(array, keyFunc = null) {
    if (!Array.isArray(array) || array.length === 0) return null;
    
    return array.reduce((max, item) => {
      const value = keyFunc ? keyFunc(item) : item;
      const numValue = NumberUtils.safeNumber(value);
      return numValue > NumberUtils.safeNumber(max) ? numValue : NumberUtils.safeNumber(max);
    }, NumberUtils.safeNumber(keyFunc ? keyFunc(array[0]) : array[0]));
  }

  /**
   * 配列の最小値取得
   */
  static min(array, keyFunc = null) {
    if (!Array.isArray(array) || array.length === 0) return null;
    
    return array.reduce((min, item) => {
      const value = keyFunc ? keyFunc(item) : item;
      const numValue = NumberUtils.safeNumber(value);
      return numValue < NumberUtils.safeNumber(min) ? numValue : NumberUtils.safeNumber(min);
    }, NumberUtils.safeNumber(keyFunc ? keyFunc(array[0]) : array[0]));
  }
}

// =============================================================================
// CSV操作
// =============================================================================

class CSVUtils {
  /**
   * CSVライン解析（ダブルクォート対応）
   */
  static parseCSVLine(line, delimiter = ',') {
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
      } else if (char === delimiter && !inQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    
    result.push(current.trim());
    return result;
  }

  /**
   * CSV文字列を2次元配列に変換
   */
  static parseCSV(csvString, hasHeader = true, delimiter = ',') {
    if (!csvString) return [];
    
    const lines = csvString.split('\n').filter(line => line.trim());
    const data = lines.map(line => this.parseCSVLine(line, delimiter));
    
    if (hasHeader) {
      const headers = data[0];
      const rows = data.slice(1);
      
      return {
        headers: headers,
        data: rows.map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] || '';
          });
          return obj;
        })
      };
    }
    
    return { headers: [], data: data };
  }

  /**
   * 2次元配列をCSV文字列に変換
   */
  static arrayToCSV(data, includeHeaders = true) {
    if (!Array.isArray(data) || data.length === 0) return '';
    
    const csvLines = [];
    
    // ヘッダー行
    if (includeHeaders && typeof data[0] === 'object') {
      const headers = Object.keys(data[0]);
      csvLines.push(headers.map(h => this.escapeCSVField(h)).join(','));
      
      // データ行
      data.forEach(row => {
        const values = headers.map(h => this.escapeCSVField(row[h]));
        csvLines.push(values.join(','));
      });
    } else {
      // 2次元配列の場合
      data.forEach(row => {
        if (Array.isArray(row)) {
          csvLines.push(row.map(cell => this.escapeCSVField(cell)).join(','));
        }
      });
    }
    
    return csvLines.join('\n');
  }

  /**
   * CSVフィールドのエスケープ
   */
  static escapeCSVField(field) {
    if (field === null || field === undefined) return '';
    
    const str = field.toString();
    
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    
    return str;
  }
}

// =============================================================================
// バリデーション
// =============================================================================

class ValidationUtils {
  /**
   * メールアドレス形式チェック
   */
  static isValidEmail(email) {
    if (!email) return false;
    
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  /**
   * URL形式チェック
   */
  static isValidURL(url) {
    if (!url) return false;
    
    try {
      new URL(url);
      return true;
    } catch {
      return false;
    }
  }

  /**
   * ASIN形式チェック
   */
  static isValidASIN(asin) {
    if (!asin) return false;
    
    const asinRegex = /^B[0-9A-Z]{9}$/;
    return asinRegex.test(asin);
  }

  /**
   * 日付形式チェック
   */
  static isValidDate(dateString) {
    if (!dateString) return false;
    
    const date = new Date(dateString);
    return date instanceof Date && !isNaN(date);
  }

  /**
   * 数値範囲チェック
   */
  static isInRange(value, min, max) {
    const num = NumberUtils.safeNumber(value);
    return num >= min && num <= max;
  }

  /**
   * 必須フィールドチェック
   */
  static validateRequired(obj, requiredFields) {
    const errors = [];
    
    requiredFields.forEach(field => {
      if (!obj[field] || obj[field] === '') {
        errors.push(`必須フィールド「${field}」が入力されていません`);
      }
    });
    
    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }
}

// =============================================================================
// キャッシュ管理
// =============================================================================

class CacheUtils {
  constructor(cacheService = CacheService.getScriptCache(), defaultExpiry = 300) {
    this.cache = cacheService;
    this.defaultExpiry = defaultExpiry; // 秒
  }

  /**
   * キャッシュ取得
   */
  get(key) {
    try {
      const cached = this.cache.get(key);
      if (cached) {
        return JSON.parse(cached);
      }
    } catch (error) {
      console.warn('キャッシュ取得エラー:', error);
    }
    return null;
  }

  /**
   * キャッシュ設定
   */
  set(key, value, expiry = this.defaultExpiry) {
    try {
      const toCache = JSON.stringify(value);
      this.cache.put(key, toCache, expiry);
      return true;
    } catch (error) {
      console.warn('キャッシュ設定エラー:', error);
      return false;
    }
  }

  /**
   * キャッシュ削除
   */
  remove(key) {
    try {
      this.cache.remove(key);
      return true;
    } catch (error) {
      console.warn('キャッシュ削除エラー:', error);
      return false;
    }
  }

  /**
   * キャッシュクリア
   */
  clear(pattern = null) {
    try {
      if (pattern) {
        // パターンマッチでの削除は現在のGASでは制限あり
        console.warn('パターンマッチでのキャッシュクリアは制限があります');
      } else {
        // 全キャッシュクリア
        this.cache.removeAll([]);
      }
      return true;
    } catch (error) {
      console.warn('キャッシュクリアエラー:', error);
      return false;
    }
  }

  /**
   * キャッシュかヒットしない場合の取得・設定
   */
  getOrSet(key, valueFunc, expiry = this.defaultExpiry) {
    let value = this.get(key);
    
    if (value === null) {
      value = valueFunc();
      this.set(key, value, expiry);
    }
    
    return value;
  }
}