/**
 * Amazon SP-API接続テスト
 * 
 * 認証情報が正しく設定されているか確認します
 */

/**
 * 簡易接続テスト
 */
function testAmazonConnection() {
  console.log('Amazon SP-API接続テストを開始します...');
  
  try {
    // 認証情報の確認
    const credentials = checkAmazonCredentials();
    if (!credentials.isValid) {
      throw new Error('認証情報が不完全です');
    }
    
    // アクセストークンの取得テスト
    const accessToken = getAmazonAccessToken();
    if (accessToken) {
      console.log('✅ アクセストークン取得成功');
      console.log('トークンの最初の20文字:', accessToken.substring(0, 20) + '...');
      return true;
    } else {
      console.error('❌ アクセストークン取得失敗');
      return false;
    }
    
  } catch (error) {
    console.error('❌ テスト失敗:', error.message);
    return false;
  }
}

/**
 * 認証情報の存在確認
 */
function checkAmazonCredentials() {
  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('SP_API_CLIENT_ID');
  const clientSecret = properties.getProperty('SP_API_CLIENT_SECRET');
  const refreshToken = properties.getProperty('SP_API_REFRESH_TOKEN');
  
  console.log('=== 認証情報チェック ===');
  console.log('Client ID:', clientId ? '✅ 設定済み' : '❌ 未設定');
  console.log('Client Secret:', clientSecret ? '✅ 設定済み' : '❌ 未設定');
  console.log('Refresh Token:', refreshToken ? '✅ 設定済み' : '❌ 未設定');
  
  return {
    isValid: !!(clientId && clientSecret && refreshToken),
    clientId: clientId,
    clientSecret: clientSecret,
    refreshToken: refreshToken
  };
}

/**
 * アクセストークンの取得
 */
function getAmazonAccessToken() {
  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('SP_API_CLIENT_ID');
  const clientSecret = properties.getProperty('SP_API_CLIENT_SECRET');
  const refreshToken = properties.getProperty('SP_API_REFRESH_TOKEN');
  
  const url = 'https://api.amazon.com/auth/o2/token';
  
  const payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
    client_secret: clientSecret
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: Object.entries(payload).map(([key, value]) => `${key}=${encodeURIComponent(value)}`).join('&'),
    muteHttpExceptions: true
  };
  
  try {
    console.log('アクセストークンを取得中...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      return data.access_token;
    } else {
      console.error('エラーレスポンス:', responseText);
      return null;
    }
    
  } catch (error) {
    console.error('リクエストエラー:', error);
    return null;
  }
}

/**
 * 実際のAPI呼び出しテスト（注文データ取得）
 */
function testGetOrders() {
  console.log('=== 注文データ取得テスト ===');
  
  try {
    // アクセストークンを取得
    const accessToken = getAmazonAccessToken();
    if (!accessToken) {
      throw new Error('アクセストークンの取得に失敗しました');
    }
    
    // 注文データを取得（過去7日間）
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 7);
    
    const url = 'https://sellingpartnerapi-fe.amazon.com/orders/v0/orders';
    const queryParams = {
      MarketplaceIds: 'A1VC38T7YXB528', // 日本
      CreatedAfter: startDate.toISOString(),
      CreatedBefore: endDate.toISOString(),
      MaxResultsPerPage: 10
    };
    
    const queryString = Object.entries(queryParams)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&');
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'x-amz-access-token': accessToken,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    console.log('注文データを取得中...');
    const response = UrlFetchApp.fetch(`${url}?${queryString}`, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      console.log('✅ 注文データ取得成功');
      console.log('注文数:', data.Orders ? data.Orders.length : 0);
      
      if (data.Orders && data.Orders.length > 0) {
        console.log('最初の注文:', {
          OrderId: data.Orders[0].AmazonOrderId,
          OrderDate: data.Orders[0].PurchaseDate,
          OrderStatus: data.Orders[0].OrderStatus
        });
      }
      
      return true;
    } else {
      console.error('❌ エラーレスポンス:', responseText);
      return false;
    }
    
  } catch (error) {
    console.error('❌ テスト失敗:', error.message);
    return false;
  }
}

/**
 * 在庫データ取得テスト
 */
function testGetInventory() {
  console.log('=== 在庫データ取得テスト ===');
  
  try {
    const accessToken = getAmazonAccessToken();
    if (!accessToken) {
      throw new Error('アクセストークンの取得に失敗しました');
    }
    
    const url = 'https://sellingpartnerapi-fe.amazon.com/fba/inventory/v1/summaries';
    const queryParams = {
      marketplaceIds: 'A1VC38T7YXB528',
      granularityType: 'Marketplace',
      granularityId: 'A1VC38T7YXB528'
    };
    
    const queryString = Object.entries(queryParams)
      .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
      .join('&');
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'x-amz-access-token': accessToken,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    console.log('在庫データを取得中...');
    const response = UrlFetchApp.fetch(`${url}?${queryString}`, options);
    const responseCode = response.getResponseCode();
    
    console.log('レスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      console.log('✅ 在庫データ取得成功');
      console.log('在庫アイテム数:', data.inventorySummaries ? data.inventorySummaries.length : 0);
      return true;
    } else {
      console.error('❌ エラーレスポンス:', response.getContentText());
      return false;
    }
    
  } catch (error) {
    console.error('❌ テスト失敗:', error.message);
    return false;
  }
}

/**
 * 総合テスト実行
 */
function runAllAmazonTests() {
  console.log('========================================');
  console.log('Amazon SP-API 総合テスト開始');
  console.log('========================================\n');
  
  const results = {
    credentials: false,
    accessToken: false,
    orders: false,
    inventory: false
  };
  
  // 1. 認証情報チェック
  console.log('【テスト1】認証情報チェック');
  const credentials = checkAmazonCredentials();
  results.credentials = credentials.isValid;
  console.log('結果:', results.credentials ? '✅ 成功' : '❌ 失敗');
  console.log('');
  
  if (!results.credentials) {
    console.error('認証情報が設定されていません。テストを中止します。');
    return results;
  }
  
  // 2. アクセストークン取得
  console.log('【テスト2】アクセストークン取得');
  results.accessToken = !!getAmazonAccessToken();
  console.log('結果:', results.accessToken ? '✅ 成功' : '❌ 失敗');
  console.log('');
  
  if (!results.accessToken) {
    console.error('アクセストークンが取得できません。認証情報を確認してください。');
    return results;
  }
  
  // 3. 注文データ取得
  console.log('【テスト3】注文データ取得');
  results.orders = testGetOrders();
  console.log('結果:', results.orders ? '✅ 成功' : '❌ 失敗');
  console.log('');
  
  // 4. 在庫データ取得
  console.log('【テスト4】在庫データ取得');
  results.inventory = testGetInventory();
  console.log('結果:', results.inventory ? '✅ 成功' : '❌ 失敗');
  console.log('');
  
  // 総合結果
  console.log('========================================');
  console.log('テスト結果サマリー');
  console.log('========================================');
  console.log('認証情報チェック:', results.credentials ? '✅' : '❌');
  console.log('アクセストークン:', results.accessToken ? '✅' : '❌');
  console.log('注文データ取得:', results.orders ? '✅' : '❌');
  console.log('在庫データ取得:', results.inventory ? '✅' : '❌');
  
  const allPassed = Object.values(results).every(result => result);
  console.log('\n総合結果:', allPassed ? '✅ すべて成功！' : '❌ 一部失敗');
  
  return results;
}