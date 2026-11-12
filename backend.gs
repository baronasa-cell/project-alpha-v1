/**
 * 統合管理システム - バックエンドAPI (Google Apps Script)
 * VERSION: 0.63
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const props = PropertiesService.getScriptProperties();
const AUTH_KEY = (props.getProperty('AUTH_KEY') || props.getProperty('inventory_auth_key') || "").trim();

/**
 * 値を SHA-256 でハッシュ化する (認証用)
 */
function hashValue(value) {
  if (!value) return "";
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  // Array.from を使って確実に JavaScript の配列として扱い、各バイトを 0-255 に変換してから 16進数文字列にする
  return Array.from(digest).map(b => {
    const unsignedByte = b < 0 ? b + 256 : b;
    return unsignedByte.toString(16).padStart(2, '0');
  }).join('');
}

/**
 * 1回のリクエスト内でのデータ再利用用キャッシュ
 */
const DataCache = {
  values: {},
  getValues: function(sheetName) {
    if (!sheetName) return null;
    if (!this.values[sheetName]) {
      console.time('fetchSheet:' + sheetName);
      const sheet = SS.getSheetByName(sheetName);
      if (!sheet) return null;
      
      const lastRow = sheet.getLastRow();
      if (lastRow === 0) return [];

      const rawValues = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      
      // 末尾の空行を完全に除去
      let valuesLength = rawValues.length;
      while (valuesLength > 0) {
        const lastRowData = rawValues[valuesLength - 1];
        const isRowEmpty = lastRowData.every(cell => cell === "" || cell === null || cell === undefined);
        if (!isRowEmpty) break;
        valuesLength--;
      }
      
      const values = (valuesLength === rawValues.length) ? rawValues : rawValues.slice(0, valuesLength);
      this.values[sheetName] = values;

      console.timeEnd('fetchSheet:' + sheetName);
      console.log(sheetName + ': Rows=' + values.length + ' (LastRow=' + lastRow + ')');
      
      if (sheetName === 'T_在庫管理') {
        console.log('T_在庫管理: Final Rows=' + values.length);
      }
    }
    return this.values[sheetName];
  },
  clear: function(sheetName) {
    if (sheetName) delete this.values[sheetName];
    else this.values = {};
  }
};

/**
 * google.script.run (本番環境) 用のエントリーポイント
 */
function apiEntryPoint(payload) {
  if (!payload) return { status: 'error', message: 'Payload is empty' };
  console.log('API Request (google.script.run):', payload.action);
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('Spreadsheet not found or not accessible');
    
    const result = handleRequest(payload);
    if (!result) {
      console.error('handleRequest returned null for action:', payload.action);
      return { status: 'error', message: 'Handler returned no data' };
    }
    
    // GASのシリアライズ制限に配慮し、確実にプレーンなオブジェクトであることを保証
    return JSON.parse(JSON.stringify(result));
  } catch (e) {
    console.error('API Entry Error:', e.toString(), e.stack);
    return { status: 'error', message: 'Server Exception: ' + e.toString() };
  }
}

/**
 * POSTリクエストの処理 (ローカル開発/外部連携用)
 */
function doPost(e) {
  console.log('doPost: Received request');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const p = JSON.parse(e.postData.contents);
    const inputHash = hashValue(p.key);
    
    // セキュリティ強化: ハッシュ一致のみを許可 (生パスワードでの一致は認めない)
    if (inputHash !== AUTH_KEY) {
      const logMsg = `Auth Failure. Expected: [${AUTH_KEY}] (len:${AUTH_KEY.length}), Received: [${inputHash}] (len:${inputHash.length})`;
      console.warn(logMsg);
      Logger.log(logMsg); 
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unauthorized' })).setMimeType(ContentService.MimeType.JSON);
    }
    const result = handleRequest(p);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * リクエスト処理の共通ロジック
 */
function handleRequest(p) {
  const action = p.action;
  const payload = p;
    
    let result = { status: 'error', message: 'Unknown action' };

    switch (action) {
      case 'getMasters':
        result = { status: 'success', data: fetchMasterData(p.includeInactive === true) };
        break;
      case 'getInitData':
        var scope = p.scope || 'all';
        var masters = (p.skipMasters === true) ? undefined : fetchMasterData(false);
        result = { status: 'success', data: { masters: masters, historyData: fetchHistoryData(scope) } };
        break;
      case 'verifyAndAddMaster':
        result = verifyAndAddMaster(p.sheet, p.value);
        break;
      case 'registerTransaction':
        result = registerTransaction(p.sheet, p.data, p.scope);
        break;
      case 'registerBulk':
        result = registerBulkTransactions(p.transactions, p.scope);
        break;
      case 'updateTransaction':
        result = updateTransaction(p.id, p.updates, p.scope);
        break;
      case 'getHistory':
        result = { status: 'success', data: fetchHistoryData() };
        break;
      case 'updateStockThreshold':
        result = updateStockThreshold(p.itemName, p.threshold);
        break;
      case 'updateStockItemStatus':
        result = updateStockItemStatus(p.itemName, p.status);
        break;
      case 'updateStockBulk':
        result = updateStockBulk(p.thresholdUpdates, p.statusUpdates);
        break;
      case 'exportLedgerReport':
        result = exportLedgerReport(p.period, p.includePersonal === true);
        break;
      case 'uploadProductImage':
        result = uploadProductImage(p.itemName, p.base64);
        break;
      case 'registerStocktake':
        result = registerStocktake(p.stocktakeData, p.note);
        break;
      case 'updateMasterRecord':
        result = updateMasterRecord(p.masterName, p.id, p.updates);
        break;
      case 'addMaterialToManufacturing':
        result = addMaterialToManufacturing(p.manufacturingId, p.item, p.quantity, p.reason);
        break;
       case 'rebuildInventorySummary':
        result = rebuildInventorySummary();
        break;
      case 'searchHistory':
        result = { status: 'success', data: searchHistory(p.target, p.keyword, p.dateFrom, p.dateTo) };
        break;
    }

    return result;
}

function doGet(e) {
  // GETでもマスタ取得を可能にしておく
  if (e.parameter && e.parameter.action === 'getMasters') {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: fetchMasterData()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // WebAppの表示
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('在庫・販売・利益 統合管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * トランザクション登録処理
 */
/**
 * 各種取引の登録処理
 * @param {string} scope 最新履歴を取得するスコープ（新規追加）
 */
function registerTransaction(sheetName, data, scope = 'all') {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  const headers = sheet.getDataRange().getValues()[0];
  
  // 1. 自動採番
  const prefixes = { 'T_仕入': 'P', 'T_経費': 'E', 'T_製造': 'M', 'T_販売': 'S' };
  const prefix = prefixes[sheetName];
  const idValue = generateNextId(sheet, prefix);
  let masterAdded = false;

  // 2. ヘッダーに基づいたマッピング定義
  const keyMap = {
    'T_仕入': { '仕入ID': idValue, 'ステータス': 'status', '仕入日': 'date', '区分': 'category', '仕入先': 'vendor', '品名': 'item', '数量': 'quantity', '価格': 'price', '支払方法': 'payment', '備考': 'note' },
    'T_経費': { '経費ID': idValue, 'ステータス': 'status', '注文日': 'date', '仕訳': 'account', '購入先': 'vendor', '品名': 'item', '数量': 'quantity', '合計金額': 'price', '支払方法': 'payment', 'レシート': 'receipt', '管理対象': 'isStock', '備考': 'note' },
    'T_製造': { '製造ID': idValue, '品名': 'item', '数量': 'quantity', '製造開始日': 'date', '備考': 'note', 'ステータス': 'status' },
    'T_販売': { 
      '販売ID': idValue, '品名': 'item', '売先': 'buyer', '数量': 'quantity',
      '発送方法': 'shipping', 'ステータス': 'status', '追跡番号': 'trackingNumber', '備考': 'note'
    }
  };

  // 3. マスタ自動登録の実施とフラグ管理
  if (sheetName === 'T_仕入') {
    if (data.vendor) if (verifyAndAddMaster('M_仕入先', data.vendor, 1).added) masterAdded = true;
    if (data.payment) if (verifyAndAddMaster('M_支払', data.payment).added) masterAdded = true;
    if (data.item) {
        if (verifyAndAddMaster('M_商品', data.item, data.category || 'パーツ').added) masterAdded = true;
    }
  } else if (sheetName === 'T_経費') {
    if (data.vendor) if (verifyAndAddMaster('M_仕入先', data.vendor, 2).added) masterAdded = true;
    if (data.payment) if (verifyAndAddMaster('M_支払', data.payment).added) masterAdded = true;
    if (data.account) if (verifyAndAddMaster('M_仕訳', data.account).added) masterAdded = true;
    if (data.item) {
        if (data.isStock == 1) {
            if (verifyAndAddMaster('M_商品', data.item, '経費').added) masterAdded = true;
        } else {
            if (verifyAndAddMaster('M_経費品名', data.item).added) masterAdded = true;
        }
    }
  }

  // 4. 販売時の追加計算（手数料等）
  if (sheetName === 'T_販売') {
    const sellers = getSheetDataAsObjects('M_売先');
    const buyerRow = sellers.find(r => r['売先'] === data.buyer);
    const rate = buyerRow ? (parseFloat(buyerRow['手数料率']) || 0) : 0;
    data.commission = Math.round(data.price * rate);
  }
  data.updatedAt = new Date();

  // 5. 在庫バリデーション
  if (sheetName === 'T_販売' && (data.status === '取引開始' || data.status === '受注' || data.status === '開始' || data.status === '入金待ち')) {
      const isPersonal = data.type === 'personal';
      const prodMaster = getSheetDataAsObjects('M_商品');
      const prodExists = prodMaster.some(r => r['品名'] === data.item);
      
      if (!(isPersonal && !prodExists)) {
          const availability = checkStockAvailability(data.item, data.quantity);
          if (!availability.ok) {
              throw new Error(`在庫不足: ${data.item} が ${availability.shortage} 不足しています。`);
          }
      }
  }
  if (sheetName === 'T_製造' && (data.status === '製造開始' || data.status === '開始')) {
      const availability = checkBOMAvailability(data.item, data.quantity);
      if (!availability.ok) {
          throw new Error(`部品不足: ${availability.message}`);
      }
  }

  // 6. 元々の詳細なマッピングロジックでrowDataを生成
  const mapping = keyMap[sheetName] || {};
  const rowData = headers.map(raw_h => {
    const h = raw_h ? raw_h.toString().trim() : "";
    if (h === headers[0]) return idValue;
    
    if (sheetName === 'T_販売') {
      if (h === '価格' || h === '販売価格') return data.price !== undefined ? data.price : "";
      if (h === '送料') return data.shippingCost !== undefined ? data.shippingCost : 0;
      if (h === '送料負担区分') {
        const payer = data.shippingPayer || "出品者";
        return payer === "落札者" ? 1 : 0;
      }
      if (h === '手数料') return data.commission !== undefined ? data.commission : 0;
      if (h === '管理区分' || h === '管理対象外') return (data.type === 'personal') ? 1 : 0;
      if (h === '販売開始日') {
        if (data.date) {
            return data.date.replace(/-/g, '/');
        } else {
            return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
        }
      }
      if (h === '合計単価') return ""; 
    }
    
    if (sheetName === 'T_経費' && h === '管理対象') return data.isStock == 1 ? 1 : 0;

    if (sheetName === 'T_仕入' && h === '入庫日' && (data.status === '入庫済み' || data.status === '入庫済')) {
      return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    }
    if (sheetName === 'T_経費' && h === '完了日' && data.status === '完了') {
      return data.date ? data.date.replace(/-/g, '/') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    }

    const key = mapping[h];
    if (typeof key === 'string' && data[key] !== undefined) {
        let val = data[key];
        if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
            val = val.replace(/-/g, '/');
        }
        return val;
    }
    
    if (h === '単価' && data.price !== undefined && data.quantity) {
      const qty = parseFloat(data.quantity) || 0;
      return qty > 0 ? roundTo2dp(data.price / qty) : 0;
    }
    if (h === '最終更新日') return data.updatedAt || new Date();
    
    return "";
  });

  sheet.appendRow(rowData);
  SpreadsheetApp.flush();
  DataCache.clear(sheetName); 
  
  const newRowIdx = sheet.getLastRow();
  if (sheetName === 'T_仕入' || sheetName === 'T_経費') formatColumn(sheet, newRowIdx, '単価', '0.00');

  handleStatusLogic(sheetName, idValue, data.status, getRowAsObject(sheet, newRowIdx));

  return { 
    status: 'success', 
    id: idValue, 
    masterAdded: masterAdded, 
    historyData: fetchHistoryData(scope) 
  };
}

/**
 * 複数取引の一括登録処理 (Proposal 38)
 */
function registerBulkTransactions(transactions, scope = 'all') {
  if (!transactions || !Array.isArray(transactions) || transactions.length === 0) {
    throw new Error("登録対象のデータがありません");
  }

  const results = [];
  let totalMasterAdded = false;

  // 各データを順次登録
  for (const tx of transactions) {
    // tx: { type: 'purchase'|'expense', data: {...} }
    const sheetName = tx.type === 'purchase' ? 'T_仕入' : 'T_経費';
    try {
      const res = registerTransaction(sheetName, tx.data, 'none'); // ここでは履歴取得をスキップ
      results.push({ id: res.id, status: 'success' });
      if (res.masterAdded) totalMasterAdded = true;
    } catch (e) {
      results.push({ status: 'error', message: e.toString() });
      console.error('Bulk item error:', e.toString());
    }
  }

  return {
    status: 'success',
    results: results,
    masterAdded: totalMasterAdded,
    historyData: fetchHistoryData(scope) // 最後に一括で最新履歴を返す
  };
}

/**
 * ステータス更新処理
 */
/**
 * 取引履歴の更新（ステータス変更など）
 * @param {string} scope 最新履歴を取得するスコープ（新規追加）
 */
function updateTransaction(id, updates, scope = 'all') {
  const prefix = id.substring(0, 1);
  const sheetMap = { 'P': 'T_仕入', 'E': 'T_経費', 'M': 'T_製造', 'S': 'T_販売' };
  const sheetName = sheetMap[prefix];
  const sheet = SS.getSheetByName(sheetName);
  
  const data = DataCache.getValues(sheetName); // キャッシュ利用
  const rowIndex = data.findIndex(row => row[0].toString() === id.toString());
  if (rowIndex === -1) throw new Error("ID not found: " + id);

  const headers = data[0];
  const statusColIndex = headers.indexOf('ステータス');
  const oldStatus = String(data[rowIndex][statusColIndex] || "").trim();
  const newStatus = updates.status ? String(updates.status).trim() : null;

  if (newStatus) {
    sheet.getRange(rowIndex + 1, statusColIndex + 1).setValue(newStatus);
    
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    let autofillLabel = "";
    if (sheetName === 'T_仕入' && newStatus === '入庫済み') autofillLabel = '入庫日';
    else if (sheetName === 'T_経費' && ['完了', 'キャンセル', '返品完了'].includes(newStatus)) autofillLabel = '完了日';
    else if (sheetName === 'T_製造') {
        const mapping = { '製造開始': '製造開始日', '製造中': '製造着手日', 'テスト中': 'テスト開始日', '梱包中': '梱包開始日', '完了': '製造完了日' };
        autofillLabel = mapping[newStatus];
    } else if (sheetName === 'T_販売') {
        const mapping = { '発送済み': '発送日', '受取済み': '受取日', '完了': '取引完了日' };
        autofillLabel = mapping[newStatus];
    }
    
    if (autofillLabel) {
        const dateColIndex = headers.indexOf(autofillLabel);
        if (dateColIndex !== -1) {
            const currentVal = data[rowIndex][dateColIndex];
            if (!updates.dates || !updates.dates[autofillLabel]) {
                if (!currentVal) {
                    sheet.getRange(rowIndex + 1, dateColIndex + 1).setValue(todayStr);
                    if (!updates.dates) updates.dates = {};
                    updates.dates[autofillLabel] = todayStr;
                }
            }
        }
    }
  }

  if (updates.dates) {
    for (const label in updates.dates) {
      const colIndex = headers.indexOf(label);
      if (colIndex !== -1) {
        let val = updates.dates[label];
        if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(val)) {
            val = val.replace(/-/g, '/');
        }
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(val);
      }
    }
  }

  // 重要: シート更新後にキャッシュをクリアすることで、直後の handleStatusLogic が最新値を参照できる
  DataCache.clear(sheetName);

  // 仕入 (T_仕入) または 経費 (T_経費) の場合、価格と数量から「単価」を自動計算する
  if (sheetName === 'T_仕入' || sheetName === 'T_経費') {
    const updatedRow = getRowAsObject(sheet, rowIndex + 1);
    const qtyField = '数量';
    const priceField = sheetName === 'T_仕入' ? '価格' : '合計金額';
    const unitPriceColIndex = headers.indexOf('単価');

    if (unitPriceColIndex !== -1) {
      const price = parseFloat(updatedRow[priceField]) || 0;
      const qty = parseFloat(updatedRow[qtyField]) || 0;
      const calculatedUnitPrice = qty > 0 ? roundTo2dp(price / qty) : 0;
      
      sheet.getRange(rowIndex + 1, unitPriceColIndex + 1).setValue(calculatedUnitPrice);
      // 単価更新後にフォーマットを適用
      formatColumn(sheet, rowIndex + 1, '単価', '0.00');
    }
  }

  if (oldStatus !== newStatus) {
    // 単価更新後なので再度キャッシュをクリアして最新行を取得
    SpreadsheetApp.flush();
    DataCache.clear(sheetName);
    const updatedRow = getRowAsObject(sheet, rowIndex + 1);
    handleStatusLogic(sheetName, id, newStatus, updatedRow, oldStatus);
  }
  
  if (sheetName === 'T_経費') formatColumn(sheet, rowIndex + 1, '単価', '0.00');

  // 最終更新日の更新 (処理ルールに従い、最右部の列を更新)
  const updatedColIndex = headers.indexOf('最終更新日');
  if (updatedColIndex !== -1) {
    sheet.getRange(rowIndex + 1, updatedColIndex + 1).setValue(new Date());
  }
  
  DataCache.clear(sheetName); // 書き込み後にキャッシュをクリア

  // 製造や販売の更新なら在庫も再計算
  if (id.startsWith('M') || id.startsWith('S')) {
    updateInventorySummary();
  }

  SpreadsheetApp.flush();

  return { 
    status: 'success', 
    masterAdded: false,
    historyData: fetchHistoryData(scope)
  };
}

/**
 * 製造レコードへの追加部材引き当て処理
 */
function addMaterialToManufacturing(manufacturingId, partName, quantity, reason) {
  const sheetName = 'T_製造';
  const sheet = SS.getSheetByName(sheetName);
  const data = DataCache.getValues(sheetName);
  const rowIndex = data.findIndex(row => row[0].toString() === manufacturingId.toString());
  
  if (rowIndex === -1) throw new Error("Manufacturing ID not found: " + manufacturingId);

  const headers = data[0];
  const itemCol = headers.indexOf('品名');
  const qtyCol = headers.indexOf('数量');
  const unitPriceCol = headers.indexOf('単価');
  const noteCol = headers.indexOf('備考');
  const updatedCol = headers.indexOf('最終更新日');

  const makeQty = parseFloat(data[rowIndex][qtyCol]) || 0;
  const currentUnitPrice = parseFloat(data[rowIndex][unitPriceCol]) || 0;
  const currentTotalCost = currentUnitPrice * makeQty;

  // 1. 在庫引き当て (FIFO)
  StockManager.init();
  const extraCost = processFIFO(manufacturingId, '製造引当(追加)', partName, quantity);
  
  // 在庫管理の備考に理由を記録
  const stockData = StockManager.data;
  const lastStockIdx = StockManager.newRows.length > 0 ? StockManager.newRows.length - 1 : -1;
  if (lastStockIdx !== -1) {
    const invHeaders = StockManager.headers;
    const invNoteCol = invHeaders.indexOf('備考');
    if (invNoteCol !== -1) {
      StockManager.newRows[lastStockIdx][invNoteCol] = reason || "追加消費";
    }
  }

  // 2. 新単価の計算
  const newTotalCost = currentTotalCost + extraCost;
  const newUnitPrice = makeQty > 0 ? roundTo2dp(newTotalCost / makeQty) : currentUnitPrice;

  // 3. シートへの書き込み
  if (unitPriceCol !== -1) {
    sheet.getRange(rowIndex + 1, unitPriceCol + 1).setValue(newUnitPrice);
  }

  // 備考への追記
  if (noteCol !== -1) {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
    const currentNote = (data[rowIndex][noteCol] || "").toString();
    const addNote = `\n[${todayStr}] 追加部材: ${partName} × ${quantity} (理由: ${reason || 'なし'})`;
    sheet.getRange(rowIndex + 1, noteCol + 1).setValue(currentNote + addNote);
  }

  // 最終更新日
  if (updatedCol !== -1) {
    sheet.getRange(rowIndex + 1, updatedCol + 1).setValue(new Date());
  }

  StockManager.flush();
  updateInventorySummary();
  DataCache.clear(sheetName);

  return { 
    status: 'success',
    message: '部材の追加と単価の再計算が完了しました。',
    historyData: fetchHistoryData('manufacturing')
  };
}

/**
 * 在庫処理用のインメモリマネージャー
 */
const StockManager = {
  sheet: null,
  data: null,
  headers: null,
  newRows: [],
  invIdPrefix: '',
  invIdMaxNum: 0,
  modifiedRows: new Set(), // 追加: 変更された行のインデックスを記録
  isDataModified: false, // 追加: 既存行を書き換えたかどうかのフラグ
  changedItems: new Set(), // 追加: 在庫が変動した品目の記録
  
  init: function() {
    const sheetName = 'T_在庫管理';
    this.sheet = SS.getSheetByName(sheetName);
    this.data = DataCache.getValues(sheetName); // キャッシュ利用
    this.headers = this.data[0].map(h => h ? h.toString().trim() : "");
    this.newRows = [];
    this.modifiedRows = new Set(); // 追加: 変更された行のインデックスを記録
    this.isDataModified = false;
    this.changedItems = new Set();
    
    const today = new Date();
    const yyyymmdd = today.getFullYear() + String(today.getMonth() + 1).padStart(2, '0') + String(today.getDate()).padStart(2, '0');
    this.invIdPrefix = `INV-${yyyymmdd}-`;
    
    const ids = this.data.slice(1).map(r => r[0].toString()).filter(id => id.startsWith(this.invIdPrefix));
    this.invIdMaxNum = ids.reduce((max, id) => {
      const parts = id.split('-');
      const num = parseInt(parts[2]);
      return isNaN(num) ? max : Math.max(max, num);
    }, 0);
  },
  
  markModified: function(rowIndex) {
    this.modifiedRows.add(rowIndex);
    this.isDataModified = true;
    
    // 変更された行の品名を記録（再集計のため）
    if (this.data && this.data[rowIndex]) {
      const nameCol = this.headers.indexOf('品名');
      if (nameCol !== -1) {
        this.changedItems.add(this.data[rowIndex][nameCol]);
      }
    }
  },

  getNextInvId: function() {
    this.invIdMaxNum++;
    return this.invIdPrefix + this.invIdMaxNum.toString().padStart(3, '0');
  },
  
  appendRow: function(rowDataObj) {
    const row = this.headers.map(h => {
      if (rowDataObj[h] !== undefined) return rowDataObj[h];
      if (h === '入出庫日' || h === '最終更新日') return new Date();
      return "";
    });
    this.newRows.push(row);
    if (rowDataObj['品名']) {
      this.changedItems.add(rowDataObj['品名']);
    }
  },
  
  flush: function() {
    // 既存行の更新 (差分更新)
    if (this.isDataModified && this.modifiedRows.size > 0) {
      this.modifiedRows.forEach(rowIdx => {
        const rowData = this.data[rowIdx];
        this.sheet.getRange(rowIdx + 1, 1, 1, this.headers.length).setValues([rowData]);
      });
      this.modifiedRows.clear();
      this.isDataModified = false;
      DataCache.clear('T_在庫管理');
    }

    if (this.newRows.length > 0) {
      const startRow = this.sheet.getLastRow() + 1;
      this.sheet.getRange(startRow, 1, this.newRows.length, this.headers.length).setValues(this.newRows);
      
      const dateCol = this.headers.indexOf('入出庫日') + 1;
      const priceCol = this.headers.indexOf('単価') + 1;
      
      if (dateCol > 0) this.sheet.getRange(startRow, dateCol, this.newRows.length, 1).setNumberFormat('yyyy/MM/dd');
      if (priceCol > 0) this.sheet.getRange(startRow, priceCol, this.newRows.length, 1).setNumberFormat('0.00');
      
      // キャッシュ側にマージしておく（後続の集計用）
      this.data = this.data.concat(this.newRows);
      this.newRows = [];
      DataCache.clear('T_在庫管理'); // 書き込みが発生したためクリア
    }
  }
};

/**
 * ビジネスロジックの発火点
 */
function handleStatusLogic(sheetName, id, status, currentData, oldStatus = null) {
  StockManager.init(); // 在庫管理シートを一括メモリ読み込み

  if (sheetName === 'T_仕入') {
    const triggerStatuses = ['入庫済み', '入庫済', '完了'];
    const isNowTrigger = triggerStatuses.includes(String(status || "").trim());
    const wasTrigger = oldStatus ? triggerStatuses.includes(String(oldStatus).trim()) : false;
    if (isNowTrigger && !wasTrigger) {
      recordInventory(id, '仕入入庫', currentData['品名'], currentData['数量'], currentData['単価'] || (currentData['価格']/currentData['数量']) || 0, currentData['仕入先']);
    }
  }
  
  if (sheetName === 'T_製造') {
    const s = String(status || "").trim();
    const os = String(oldStatus || "").trim();
    if (s === '製造開始' && os !== '製造開始') {
      const makeQty = parseFloat(currentData['数量']) || 0;
      const itemName = currentData['品名'];
      const totalMaterialCost = processManufacturingBOM(id, itemName, makeQty);
      
      if (makeQty > 0) {
        const unitCost = roundTo2dp(totalMaterialCost / makeQty);
        const sheet = SS.getSheetByName('T_製造');
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
        const colIdx = headers.indexOf('単価');
        const data = sheet.getDataRange().getValues();
        const rowIdx = data.findIndex(r => r[0] === id);
        if (rowIdx !== -1 && colIdx !== -1) {
          sheet.getRange(rowIdx + 1, colIdx + 1).setValue(unitCost);
          formatColumn(sheet, rowIdx + 1, '単価', '0.00');
        }
      }
    }
    
    if (s === '完了' && os !== '完了') {
      const unitCost = parseFloat(currentData['単価']) || 0;
      recordInventory(id, '製造入庫', currentData['品名'], currentData['数量'], unitCost, '自社');
    }
  }

  if (sheetName === 'T_販売') {
    const s = String(status || "").trim();
    const os = String(oldStatus || "").trim();
    const triggerStatuses = ['取引開始', '入金待ち', '受注', '開始'];
    const isNowTrigger = triggerStatuses.includes(s);
    const wasTrigger = triggerStatuses.includes(os);
    
    const isInitialComplete = (s === '完了' && !oldStatus);
    if ((isNowTrigger && !wasTrigger) || isInitialComplete) {
      const isPersonal = currentData['管理対象外'] == 1 || currentData['管理区分'] == 1;
      const productName = currentData['品名'];
      const qty = parseFloat(currentData['数量']) || 0;
      
      const productMaster = SS.getSheetByName('M_商品').getDataRange().getValues();
      const pHeaders = productMaster[0].map(h => h.toString().trim());
      const nameIdx = pHeaders.indexOf('品名');
      const productExists = nameIdx !== -1 ? productMaster.slice(1).some(r => r[nameIdx] === productName) : false;
      
      let unitCost = 0;
      if (isPersonal && !productExists) {
        unitCost = 0; 
      } else {
        try {
          const totalCost = processFIFO(id, '販売出庫', productName, qty);
          unitCost = roundTo2dp(totalCost);
        } catch (e) {
          if (isPersonal) {
            unitCost = 0; 
          } else {
            throw e; 
          }
        }
      }
      
      const shippingMethod = currentData['発送方法'];
      if (shippingMethod) {
        const shippingExists = nameIdx !== -1 ? productMaster.slice(1).some(r => r[nameIdx] === shippingMethod) : false;
        if (shippingExists) {
          try {
            processFIFO(id, '販売出庫(梱包)', shippingMethod, 1);
          } catch (e) {
            throw new Error('梱包材の引当エラー: ' + e.message);
          }
        }
      }
      
      const saleSheet = SS.getSheetByName('T_販売');
      const sHeaders = saleSheet.getRange(1, 1, 1, saleSheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
      const sData = saleSheet.getDataRange().getValues();
      const rowIdx = sData.findIndex(r => r[0] === id);
      if (rowIdx !== -1) {
        const colIdx = sHeaders.indexOf('合計単価');
        if (colIdx !== -1) {
            saleSheet.getRange(rowIdx + 1, colIdx + 1).setValue(unitCost);
            formatColumn(saleSheet, rowIdx + 1, '合計単価', '0.00');
            currentData['合計単価'] = unitCost;
        }
      }
    }
  }

  if (sheetName === 'T_経費') {
    const s = String(status || "").trim();
    const os = String(oldStatus || "").trim();
    if (s === '完了' && os !== '完了') {
      if (currentData['管理対象'] == 1) { 
         const qty = parseFloat(currentData['数量']) || 1;
         const price = parseFloat(currentData['合計金額']) || 0;
         // 処理ルールに基づき、商品区分に「経費」を強制指定
         recordInventory(id, '経費入庫', currentData['品名'], qty, price / qty, currentData['購入先'], '経費');
      }
    }
  }
  
  if (status === 'キャンセル') {
    revertFIFO(id);
  }

  if (status === '完了') {
    syncToLedger(sheetName, id, currentData);
  }

  StockManager.flush(); // ここで1回だけT_在庫管理に上書き・追記
  updateInventorySummary();
  SpreadsheetApp.flush(); // 全ての書き込みを最後に確定
  
  // ビジネスロジックによるシートへの単価書き込み等をキャッシュに同期するため確実にクリア
  DataCache.clear(sheetName);
}

function appendInventoryRow(stockSheet, rowDataObj) {
  // 古いAPI（StockManagerを使用しないレガシー用、今回は呼ばれない想定）
  const headers = stockSheet.getDataRange().getValues()[0].map(h => h ? h.toString().trim() : "");
  const row = headers.map(h => {
    if (rowDataObj[h] !== undefined) return rowDataObj[h];
    if (h === '入出庫日' || h === '最終更新日') return new Date();
    return "";
  });
  stockSheet.appendRow(row);
  const rowIdx = stockSheet.getLastRow();
  formatColumn(stockSheet, rowIdx, '入出庫日', 'yyyy/MM/dd');
  formatColumn(stockSheet, rowIdx, '単価', '0.00');
}


/**
 * FIFO在庫引当エンジン（メモリ上で処理する最適化版）
 */
function processFIFO(triggerId, type, itemName, demandQty) {
  const data = StockManager.data;
  const headers = StockManager.headers;

  // 重要: すでに同じ取引IDかつ同じ区分での引当が存在する場合はスキップ（二重引当の物理防止）
  if (triggerId) {
    const prefix = triggerId.substring(0, 1);
    const idColName = (prefix === 'P' || prefix === 'E') ? '仕入ID' : (prefix === 'M' ? '製造ID' : '販売ID');
    const idColIdx = headers.indexOf(idColName);
    const typeColIdx = headers.indexOf('区分');
    if (idColIdx !== -1 && typeColIdx !== -1) {
      const alreadyExists = data.some(r => r[idColIdx] === triggerId && r[typeColIdx] === type);
      if (alreadyExists) {
        console.log('Skipping duplicate processFIFO for: ' + triggerId + ' / ' + type);
        return 0; 
      }
    }
  }

  const fCol = headers.indexOf('実在庫数量');
  const dCol = headers.indexOf('入出庫日');
  
  let lots = [];
  for (let i = 1; i < data.length; i++) {
    const itemCol = headers.indexOf('品名');
    if (itemCol !== -1 && data[i][itemCol] === itemName && data[i][fCol] > 0) {
      lots.push({ rowIndex: i, data: data[i], date: data[i][dCol] }); // rowIndex is 0-based memory index
    }
  }
  lots.sort((a, b) => new Date(a.date) - new Date(b.date));

  let remaining = demandQty;
  let totalCost = 0;

  for (const lot of lots) {
    const avail = lot.data[fCol];
    const consume = Math.min(remaining, avail);
    
    const newAvail = avail - consume;
    lot.data[fCol] = newAvail;
    lot.data[headers.indexOf('引当完了')] = (newAvail <= 0 ? 1 : 0);
    
    // 修正: メモリ上のデータを更新し、変更行としてマーク
    StockManager.data[lot.rowIndex] = lot.data;
    StockManager.markModified(lot.rowIndex);
    
    const pCol = headers.indexOf('最終更新日');
    if (pCol !== -1) {
      StockManager.data[lot.rowIndex][pCol] = new Date();
    }
    
    const unitPrice = lot.data[headers.indexOf('単価')] || 0;
    totalCost += consume * unitPrice;
    
    const newId = StockManager.getNextInvId();
    const rowObj = {
      '在庫管理ID': newId,
      '品名': itemName,
      '商品区分': lot.data[headers.indexOf('商品区分')] || "",
      '区分': type,
      '数量': -consume,
      '単価': unitPrice,
      '仕入先': lot.data[headers.indexOf('仕入先')] || "",
      '引当元管理ID': lot.data[0],
      '引当完了': 1,
      '実在庫数量': 0,
      '備考': ""
    };
    if (triggerId) {
      if (triggerId.startsWith('P')) rowObj['仕入ID'] = triggerId;
      if (triggerId.startsWith('M')) rowObj['製造ID'] = triggerId;
      if (triggerId.startsWith('S')) rowObj['販売ID'] = triggerId;
    }
    StockManager.appendRow(rowObj); // メモリに追加行を積む

    remaining -= consume;
    if (remaining <= 0) break;
  }

  if (remaining > 0) {
    throw new Error(`在庫不足: ${itemName} が ${remaining} 不足しています。`);
  }
  
  return totalCost;
}

/**
 * 在庫処理の取り消し (FIFOの逆操作)
 */
function revertFIFO(triggerId) {
  const data = StockManager.data;
  const headers = StockManager.headers;
  const invIdCol = headers.indexOf('在庫管理ID');
  const triggerIdCol = headers.indexOf(triggerId.startsWith('P') ? '仕入ID' : (triggerId.startsWith('E') ? '仕入ID' : (triggerId.startsWith('M') ? '製造ID' : '販売ID')));
  const nameCol = headers.indexOf('品名');
  const qtyCol = headers.indexOf('数量');
  const actualQtyCol = headers.indexOf('実在庫数量');
  const priceCol = headers.indexOf('単価');
  const sourceIdCol = headers.indexOf('引当元管理ID');
  const statusCol = headers.indexOf('引当完了');

  // 1. 指定された取引IDに関連する在庫レコードをすべて探す
  const relatedRows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][triggerIdCol] === triggerId) {
      relatedRows.push({ index: i, row: data[i] });
    }
  }

  relatedRows.forEach(item => {
    const row = item.row;
    const qty = parseFloat(row[qtyCol]) || 0;
    const itemName = row[nameCol];
    const unitPrice = parseFloat(row[priceCol]) || 0;
    const sourceId = row[sourceIdCol];

    if (qty > 0) {
      // ケースA: 入庫の取り消し (仕入や製造完了のキャンセル)
      // -> 「取消出庫」レコードを作成し、実在庫を0にする
      const newId = StockManager.getNextInvId();
      const cancelRow = {
        '在庫管理ID': newId,
        '品名': itemName,
        '商品区分': row[headers.indexOf('商品区分')] || "",
        '区分': '取消出庫',
        '数量': -qty,
        '単価': unitPrice,
        '引当元管理ID': row[invIdCol],
        '引当完了': 1,
        '実在庫数量': 0,
        '備考': `取消: ${triggerId}`
      };
      if (triggerId.startsWith('P')) cancelRow['仕入ID'] = triggerId;
      if (triggerId.startsWith('E')) cancelRow['仕入ID'] = triggerId;
      if (triggerId.startsWith('M')) cancelRow['製造ID'] = triggerId;
      if (triggerId.startsWith('S')) cancelRow['販売ID'] = triggerId;
      
      StockManager.appendRow(cancelRow);

      // 元のレコードの実在庫を0にする
      data[item.index][actualQtyCol] = 0;
      data[item.index][statusCol] = 1;
      StockManager.markModified(item.index);

    } else if (qty < 0) {
      // ケースB: 出庫の取り消し (販売や製造消費のキャンセル)
      // -> 「取消入庫」レコードを作成し、在庫を戻す
      const newId = StockManager.getNextInvId();
      const cancelRow = {
        '在庫管理ID': newId,
        '品名': itemName,
        '商品区分': row[headers.indexOf('商品区分')] || "",
        '区分': '取消入庫',
        '数量': Math.abs(qty),
        '単価': unitPrice,
        '引当元管理ID': sourceId,
        '引当完了': 1,
        '実在庫数量': 0,
        '備考': `取消: ${triggerId}`
      };
      if (triggerId.startsWith('P')) cancelRow['仕入ID'] = triggerId;
      if (triggerId.startsWith('E')) cancelRow['仕入ID'] = triggerId;
      if (triggerId.startsWith('M')) cancelRow['製造ID'] = triggerId;
      if (triggerId.startsWith('S')) cancelRow['販売ID'] = triggerId;
      
      StockManager.appendRow(cancelRow);

      // (オプション) 引当元レコードの状態を戻す
      if (sourceId) {
        const sourceRowIdx = data.findIndex(r => r[invIdCol] === sourceId);
        if (sourceRowIdx !== -1) {
          const currentActual = parseFloat(data[sourceRowIdx][actualQtyCol]) || 0;
          const restoredQty = currentActual + Math.abs(qty);
          data[sourceRowIdx][actualQtyCol] = restoredQty;
          data[sourceRowIdx][statusCol] = (restoredQty > 0 ? 0 : 1);
          StockManager.markModified(sourceRowIdx);
        }
      }
    }
  });
}

function processManufacturingBOM(makeId, productName, makeQty) {
  const bomSheet = SS.getSheetByName('M_BOM');
  const bomData = getSheetDataAsObjects(bomSheet);
  const components = bomData.filter(r => r['品名'] === productName);

  let totalCost = 0;
  components.forEach(comp => {
    const needed = (parseFloat(comp['数量']) || 0) * makeQty;
    if (needed > 0) {
      totalCost += processFIFO(makeId, '製造引当', comp['部品'], needed);
    }
  });
  return totalCost;
}

function recordInventory(refId, type, itemName, qty, unitPrice, vendorName = "", forceCategory = null) {
  const newId = StockManager.getNextInvId();
  
  const pSheet = SS.getSheetByName('M_商品');
  const pData = pSheet.getDataRange().getValues();
  const pHeaders = pData[0].map(h => h.toString().trim());
  const nameIdx = pHeaders.indexOf('品名');
  const catIdx = pHeaders.indexOf('カテゴリ');
  const pRow = nameIdx !== -1 ? pData.slice(1).find(r => r[nameIdx] === itemName) : null;
  
  // forceCategoryがあれば優先、なければマスタから取得
  const category = forceCategory || (pRow && catIdx !== -1 ? pRow[catIdx] : ""); 

  const rowObj = {
    '在庫管理ID': newId,
    '品名': itemName,
    '商品区分': category,
    '区分': type,
    '数量': qty,
    '単価': roundTo2dp(unitPrice),
    '仕入先': vendorName,
    '引当完了': 0,
    '実在庫数量': qty,
    '備考': ""
  };
  if (refId) {
    if (refId.startsWith('P') || refId.startsWith('E')) rowObj['仕入ID'] = refId;
    if (refId.startsWith('M')) rowObj['製造ID'] = refId;
    if (refId.startsWith('S')) rowObj['販売ID'] = refId;
  }
  StockManager.appendRow(rowObj);
}

/**
 * 帳簿連携 (T_帳簿)
 */
function syncToLedger(sheetName, id, data) {
  // 帳簿記録の対象外シートのガード
  if (sheetName === 'T_製造' || sheetName === 'T_仕入') return; 

  const ledgerSheet = SS.getSheetByName('T_帳簿');
  let headers = ledgerSheet.getDataRange().getValues()[0].map(h => h.toString().trim());
  
  // K列（事業/個人）が存在しない場合は自動追加
  let typeCol = headers.indexOf('事業/個人');
  if (typeCol === -1) {
    ledgerSheet.getRange(1, headers.length + 1).setValue('事業/個人');
    headers.push('事業/個人');
    typeCol = headers.length - 1;
  }

  let row = new Array(headers.length).fill("");
  
  const dateCol = headers.indexOf('日付');
  const descCol = headers.indexOf('品名');
  const incomeCol = headers.indexOf('売上');
  const costCol = headers.indexOf('仕入');

  // 帳簿の日付決定ロジックの適正化
  if (sheetName === 'T_販売') {
    if (dateCol !== -1) row[dateCol] = data['取引完了日'] || new Date();
  } else if (sheetName === 'T_経費') {
    if (dateCol !== -1) row[dateCol] = data['完了日'] || new Date();
  } else {
    if (dateCol !== -1) row[dateCol] = new Date(); 
  }
  
  const price = parseFloat(data['合計金額'] || data['価格'] || data['販売価格'] || 0);

  if (sheetName === 'T_販売') {
    const isPersonal = (data['管理区分'] == 1 || data['管理対象外'] == 1) ? 1 : 0;
    if (typeCol !== -1) row[typeCol] = isPersonal;

    const buyer = data['売先'] || "";
    const item = data['品名'] || "";
    const qty = parseFloat(data['数量']) || 1;
    let desc = buyer ? `${buyer} ${item}` : item;
    if (qty > 1) {
      desc += "×" + qty;
    }
    if (descCol !== -1) row[descCol] = desc; // 売先 + 品名 (複数時は数量付与)
    
    if (incomeCol !== -1) row[incomeCol] = price; // 売上
    if (costCol !== -1) row[costCol] = parseFloat(data['合計単価']) || 0; // 仕入（原価）
    
    const commFeeCol = headers.indexOf('通信費');
    if (commFeeCol !== -1) {
      row[commFeeCol] = parseFloat(data['送料']) || 0;
    }
    
    const feeCol = headers.indexOf('支払手数料');
    if (feeCol !== -1) {
      row[feeCol] = parseFloat(data['手数料']) || 0;
    }
  } else if (sheetName === 'T_経費') {
    if (typeCol !== -1) row[typeCol] = 0; // 経費は事業用

    const account = data['仕訳']; 
    const qty = parseFloat(data['数量']) || 0;
    const itemName = data['品名'] || "";
    if (descCol !== -1) row[descCol] = (account ? account + " " : "") + itemName + (qty > 1 ? "×" + qty : "");

    const colIndex = headers.indexOf(account);
    if (colIndex !== -1 && colIndex >= 4) { 
      row[colIndex] = price;
    } else {
      // 指定された仕訳が見つからない場合のフォールバック（雑費列などがあればそこへ、なければ最後の方へ）
      const miscCol = headers.indexOf('雑費');
      const targetCol = miscCol !== -1 ? miscCol : (headers.length > 9 ? 9 : headers.length - 1);
      row[targetCol] = price; 
    }
  }

  ledgerSheet.appendRow(row);
  const lastRow = ledgerSheet.getLastRow();
  formatColumn(ledgerSheet, lastRow, '日付', 'yyyy/MM/dd');

  // 日付順（昇順）にソートして整理
  if (lastRow > 1) {
    ledgerSheet.getRange(2, 1, lastRow - 1, headers.length).sort({column: 1, ascending: true});
  }
}

/**
 * 在庫集計の再計算（メモリデータ活用版）
 */
function updateInventorySummary() {
  if (!StockManager.changedItems || StockManager.changedItems.size === 0) {
    return;
  }

  const sumSheet = SS.getSheetByName('T_在庫集計');
  const sumRange = sumSheet.getDataRange();
  const sumData = sumRange.getValues(); 
  const sumHeaders = sumData[0].map(h => h.toString().trim());
  
  const data = StockManager.data || []; 
  const headers = StockManager.headers || [];
  
  if (data.length < 2) return;

  const summary = {};
  const nameCol = headers.indexOf('品名');
  const qtyCol = headers.indexOf('実在庫数量');
  
  const targetItems = StockManager.changedItems;
  targetItems.forEach(name => summary[name] = 0);

  // メモリ上の最新データから集計
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    if (targetItems.has(name)) {
      summary[name] += (parseFloat(data[i][qtyCol]) || 0);
    }
  }
  
  const existingNames = new Set();
  const pData = getSheetDataAsObjects('M_商品');

  const sNameCol = sumHeaders.indexOf('品名');
  const sQtyCol = sumHeaders.indexOf('現在庫数');
  const sUpdateCol = sumHeaders.indexOf('最終更新日');

  let modified = false;

  // 2. 既存行の更新 (メモリ上で一括書き換え)
  for (let i = 1; i < sumData.length; i++) {
    const name = sNameCol !== -1 ? sumData[i][sNameCol] : null; 
    if (!name) continue;
    existingNames.add(name);
    
    if (summary[name] !== undefined) {
      const currentQty = sQtyCol !== -1 ? (parseFloat(sumData[i][sQtyCol]) || 0) : 0;
      const newQty = summary[name];
      if (currentQty !== newQty) {
        if (sQtyCol !== -1) sumData[i][sQtyCol] = newQty;
        if (sUpdateCol !== -1) sumData[i][sUpdateCol] = new Date();
        modified = true;
      }
    }
  }
  
  // 変更があれば一括書き込み
  if (modified) {
    sumRange.setValues(sumData);
  }
  
  // 3. 新規品目の追加
  const newRows = [];
  for (const name of targetItems) {
    if (!existingNames.has(name)) {
      const pRow = pData.find(r => r['品名'] === name);
      const cat = pRow ? pRow['カテゴリ'] : "";
      const useFlag = pRow ? (parseInt(pRow['使用FLG']) === 0 ? 0 : 1) : 1; 

      const newRow = new Array(sumHeaders.length).fill("");
      const setCol = (colName, val) => {
        const idx = sumHeaders.indexOf(colName);
        if (idx !== -1) newRow[idx] = val;
      };

      setCol('優先度', 9);
      setCol('表示順', ((sumData.length + newRows.length) * 10));
      setCol('品名', name);
      setCol('カテゴリ', cat);
      setCol('現在庫数', summary[name]);
      setCol('閾値', 10);
      setCol('最終更新日', new Date());
      setCol('使用FLG', useFlag);
      setCol('最終棚卸日', "");

      newRows.push(newRow);
    }
  }

  if (newRows.length > 0) {
    sumSheet.getRange(sumSheet.getLastRow() + 1, 1, newRows.length, sumHeaders.length).setValues(newRows);
  }
  SpreadsheetApp.flush();
  DataCache.clear('T_在庫集計');
}

/**
 * 在庫の全再集計（全品目を対象に再計算）
 */
function rebuildInventorySummary() {
  StockManager.init();
  const data = StockManager.data;
  const headers = StockManager.headers;
  const nameCol = headers.indexOf('品名');
  
  if (nameCol === -1) throw new Error("品名列が見つかりません");
  
  // すべての品名を抽出して再計算対象にする
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    if (name) StockManager.changedItems.add(name);
  }
  
  updateInventorySummary();
  return { status: 'success', message: '在庫の全再集計が完了しました。' };
}

/**
 * ID生成ヘルパー
 */
function generateNextId(sheet, prefix) {
  if (prefix === 'INV') {
    const today = new Date();
    const yyyymmdd = today.getFullYear() + String(today.getMonth() + 1).padStart(2, '0') + String(today.getDate()).padStart(2, '0');
    const fullPrefix = `INV-${yyyymmdd}-`;
    const data = DataCache.getValues(sheet.getName()); // キャッシュ利用
    if (!data || data.length < 2) return fullPrefix + "001";
    const ids = data.slice(1).map(r => r[0].toString()).filter(id => id.startsWith(fullPrefix));
    const maxNum = ids.reduce((max, id) => {
      const parts = id.split('-');
      const num = parseInt(parts[2]);
      return isNaN(num) ? max : Math.max(max, num);
    }, 0);
    return fullPrefix + (maxNum + 1).toString().padStart(3, '0');
  }

  if (prefix === 'A') {
    const data = DataCache.getValues(sheet.getName()); // キャッシュ利用
    if (!data || data.length < 2) return prefix + "0001";
    const ids = data.slice(1).map(r => r[0].toString()).filter(id => id.startsWith(prefix));
    const maxNum = ids.reduce((max, id) => {
      const num = parseInt(id.replace(prefix, ""));
      return isNaN(num) ? max : Math.max(max, num);
    }, 0);
    return prefix + (maxNum + 1).toString().padStart(4, '0');
  }

  const digits = (prefix === 'S') ? 5 : 4;
  const data = DataCache.getValues(sheet.getName()); // キャッシュ利用
  if (!data || data.length < 2) return prefix + "1".padStart(digits, '0');
  
  const ids = data.slice(1).map(r => r[0].toString());
  const maxNum = ids.reduce((max, id) => {
    const num = parseInt(id.replace(prefix, ""));
    return isNaN(num) ? max : Math.max(max, num);
  }, 0);
  
  return prefix + (maxNum + 1).toString().padStart(digits, '0');
}

/**
 * 履歴データの取得（フィルタリング対応版）
 */
function fetchHistoryData(scope = 'all') {
  console.time('fetchHistoryData');
  
  // 取得対象テーブルの決定
  let tables = [];
  if (scope === 'all' || scope === 'essential') {
    tables = ['T_仕入', 'T_経費', 'T_製造', 'T_販売', 'T_帳簿'];
  } else if (scope === 'purchase') {
    tables = ['T_仕入'];
  } else if (scope === 'expense') {
    tables = ['T_経費'];
  } else if (scope === 'manufacturing') {
    tables = ['T_製造', 'T_在庫集計'];
  } else if (scope === 'sales') {
    tables = ['T_販売', 'T_在庫集計', 'T_帳簿'];
  } else if (scope === 'inventory') {
    tables = ['T_在庫集計'];
  } else if (scope === 'ledger') {
    tables = ['T_帳簿'];
  }

  // 完了ステータスの定義
  const completedStatuses = {
    'T_仕入': ['返品完了', '入庫済み', '入庫済', 'キャンセル'],
    'T_経費': ['完了', '返品完了', 'キャンセル'],
    'T_製造': ['完了', '製造完了', '梱包完了', 'キャンセル'],
    'T_販売': ['完了', 'キャンセル']
  };

  const now = new Date();
  const threshold40 = new Date(now.getTime() - (40 * 24 * 60 * 60 * 1000));
  const threshold180 = new Date(now.getTime() - (180 * 24 * 60 * 60 * 1000));

  const rawData = {};
  tables.forEach(name => {
    const values = DataCache.getValues(name);
    if (!values || values.length < 2) {
      rawData[name] = values;
      return;
    }

    // フィルタリング（起動時以外、かつ履歴テーブルの場合のみ適用）
    if (name.startsWith('T_') && name !== 'T_在庫集計') {
      const headers = values[0];
      const statusIdx = headers.indexOf('ステータス');
      // 日付候補の列を探す
      const dateCols = ['入庫日', '完了日', '取引完了日', '仕入日', '製造完了日', '販売開始日', '取引開始日', '製造開始日'];
      let dateIdx = -1;
      for (let col of dateCols) {
        dateIdx = headers.indexOf(col);
        if (dateIdx !== -1) break;
      }

      const threshold = (name === 'T_製造') ? threshold180 : threshold40;
      const completed = completedStatuses[name] || ['完了', 'キャンセル'];

      const filtered = [headers];
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const status = (statusIdx !== -1) ? (row[statusIdx] || "").toString().trim() : "";
        const isCompleted = completed.includes(status);
        
        // 条件1: 未完了なら必ず残す
        if (!isCompleted && status !== "") {
          filtered.push(row);
          continue;
        }

        // 条件2: 完了済みなら日付をチェック
        if (dateIdx !== -1) {
          const d = new Date(row[dateIdx]);
          if (!isNaN(d.getTime()) && d >= threshold) {
            filtered.push(row);
          }
        } else {
          // 日付列がない場合は念のため残す
          filtered.push(row);
        }
      }
      rawData[name] = filtered;
    } else {
      // 在庫集計などはそのまま
      rawData[name] = values;
    }
  });

  console.timeEnd('fetchHistoryData');
  return { isRaw: true, rawData: rawData, scope: scope };
}

function getRowAsObject(sheet, rowIdx) {
  const sheetName = typeof sheet === 'string' ? sheet : sheet.getName();
  const data = DataCache.getValues(sheetName);
  if (!data || data.length < rowIdx) {
    const headers = SS.getSheetByName(sheetName).getDataRange().getValues()[0];
    const vals = SS.getSheetByName(sheetName).getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    const obj = {};
    headers.forEach((h, i) => obj[h] = vals[i]);
    return obj;
  }
  const headers = data[0];
  const vals = data[rowIdx - 1];
  const obj = {};
  headers.forEach((h, i) => obj[h] = vals[i]);
  return obj;
}

function formatColumn(sheet, rowIdx, headerName, formatString, cachedHeaders = null) {
  const headers = cachedHeaders || sheet.getDataRange().getValues()[0];
  const colIdx = headers.indexOf(headerName);
  if (colIdx !== -1) {
    sheet.getRange(rowIdx, colIdx + 1).setNumberFormat(formatString);
  }
}

function roundTo2dp(val) {
  if (val === null || val === undefined || isNaN(val)) return 0;
  return Math.round(val * 100) / 100;
}

/**
 * 過去履歴の検索 (全データから検索)
 */
function searchHistory(target = 'all', keyword = '', dateFrom = '', dateTo = '') {
  console.time('searchHistory');
  
  let tables = [];
  if (target === 'all') {
    tables = ['T_仕入', 'T_経費', 'T_製造', 'T_販売'];
  } else if (target === 'purchase') {
    tables = ['T_仕入'];
  } else if (target === 'expense') {
    tables = ['T_経費'];
  } else if (target === 'manufacturing') {
    tables = ['T_製造'];
  } else if (target === 'sales') {
    tables = ['T_販売'];
  }

  const resultData = {};
  const kw = keyword.toLowerCase();
  
  let fromTime = 0;
  let toTime = Number.MAX_SAFE_INTEGER;
  if (dateFrom) fromTime = new Date(dateFrom).getTime();
  if (dateTo) {
    const td = new Date(dateTo);
    td.setHours(23, 59, 59, 999);
    toTime = td.getTime();
  }

  tables.forEach(name => {
    const values = DataCache.getValues(name);
    if (!values || values.length < 2) {
      resultData[name] = values || [];
      return;
    }

    const headers = values[0];
    const dateCols = ['入庫日', '完了日', '取引完了日', '仕入日', '製造完了日', '販売開始日', '取引開始日', '製造開始日'];
    let dateIdx = -1;
    for (let col of dateCols) {
      dateIdx = headers.indexOf(col);
      if (dateIdx !== -1) break;
    }

    const filtered = [headers];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      let matchKeyword = false;
      let matchDate = true;

      // キーワードチェック (行全体を文字列化して検索)
      if (kw) {
        const rowStr = row.join(' ').toLowerCase();
        if (rowStr.indexOf(kw) !== -1) {
          matchKeyword = true;
        }
      } else {
        matchKeyword = true;
      }

      // 日付チェック
      if (dateIdx !== -1 && (fromTime > 0 || toTime < Number.MAX_SAFE_INTEGER)) {
        const d = new Date(row[dateIdx]);
        if (!isNaN(d.getTime())) {
          const t = d.getTime();
          if (t < fromTime || t > toTime) matchDate = false;
        }
      }

      if (matchKeyword && matchDate) {
        filtered.push(row);
      }
    }
    resultData[name] = filtered;
  });

  console.timeEnd('searchHistory');
  return { isRaw: true, rawData: resultData, scope: 'search' };
}

function fetchMasterData(includeInactive = false) {
  console.time('fetchMasterData');
  const targetSheets = [
    'M_User', 'M_商品', 'M_仕入先', 'M_支払', 'M_売先', 'M_発送', 
    'M_ステータス', 'M_カテゴリ', 'M_BOM', 'M_画面制御', 'M_経費品名', 'M_仕訳', 'T_在庫集計'
  ];
  const masters = {};
  targetSheets.forEach(name => {
    masters[name] = DataCache.getValues(name);
  });
  console.timeEnd('fetchMasterData');
  return masters;
}

function getSheetDataAsObjects(sheet) {
  const sheetName = typeof sheet === 'string' ? sheet : sheet.getName();
  const data = DataCache.getValues(sheetName); 
  if (!data || data.length < 2) return [];
  const headers = data[0].map(h => (h ? h.toString().trim() : "")); 
  return data.slice(1).filter(r => r.some(v => v !== "")).map(r => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = r[i]; });
    return obj;
  });
}

function applyMasterRules(records, includeInactive = false) {
  let filtered = records;
  if (!includeInactive && records.length > 0 && records[0].hasOwnProperty('使用FLG')) {
    filtered = records.filter(r => r['使用FLG'] == 1 || r['使用FLG'] === true);
  }
  if (filtered.length > 0 && filtered[0].hasOwnProperty('表示順')) {
    filtered.sort((a, b) => (parseInt(a['表示順']) || 0) - (parseInt(b['表示順']) || 0));
  }
  return filtered;
}

function verifyAndAddMaster(sheetName, valueToAdd, extraUsage = null) {
  const data = DataCache.getValues(sheetName); // キャッシュ利用
  if (!data || data.length === 0) return { added: false };
  const headers = data[0].map(h => h.toString().trim());
  
  // 検索対象の列（品名、仕入先など）を特定
  const keyConfig = {
    'M_商品': '品名',
    'M_仕入先': '仕入先',
    'M_売先': '売先',
    'M_発送': '発送方法',
    'M_経費品名': '品名',
    'M_支払': '支払方法',
    'M_仕訳': '仕訳名'
  };
  const keyHeader = keyConfig[sheetName] || headers[1];
  const keyIdx = headers.indexOf(keyHeader);
  
  const existsIdx = data.findIndex((r, idx) => idx > 0 && r[keyIdx] === valueToAdd);
  const sheet = SS.getSheetByName(sheetName);
  
  if (existsIdx !== -1) {
    if (sheetName === 'M_仕入先' && extraUsage !== null) {
      const uCol = headers.indexOf('用途区分');
      if (uCol !== -1) {
        const currentUsage = parseInt(data[existsIdx][uCol]);
        if (!isNaN(currentUsage) && currentUsage !== 3 && currentUsage !== extraUsage) {
          sheet.getRange(existsIdx + 1, uCol + 1).setValue(3);
        } else if (isNaN(currentUsage)) {
          sheet.getRange(existsIdx + 1, uCol + 1).setValue(extraUsage);
        }
      }
    } else if (sheetName === 'M_商品' && extraUsage !== null) {
      const catCol = headers.indexOf('カテゴリ');
      if (catCol !== -1 && !data[existsIdx][catCol]) {
        sheet.getRange(existsIdx + 1, catCol + 1).setValue(extraUsage);
      }
    }
    return { added: false, updated: true };
  }
  
  const newRow = Array(headers.length).fill("");
  
  // 商品IDの採番 (M_商品かつ1列目がIDの場合)
  if (sheetName === 'M_商品' && (headers[0] === '商品ID' || headers[0] === 'ID')) {
    newRow[0] = generateNextId(sheet, 'I');
  }
  
  // 表示順の設定
  const orderIdx = headers.indexOf('表示順');
  if (orderIdx !== -1) {
    const nextOrder = data.length > 1 ? Math.max(...data.slice(1).map(r => parseInt(r[orderIdx]) || 0)) + 10 : 10;
    newRow[orderIdx] = nextOrder;
  }
  
  // キー項目（品名など）の設定
  newRow[keyIdx] = valueToAdd;
  
  // カテゴリや用途区分の設定
  if (sheetName === 'M_仕入先' && extraUsage !== null) {
    const uCol = headers.indexOf('用途区分');
    if (uCol !== -1) newRow[uCol] = extraUsage;
  } else if (sheetName === 'M_商品' && extraUsage !== null) {
    const catCol = headers.indexOf('カテゴリ');
    if (catCol !== -1) newRow[catCol] = extraUsage;
  }
  
  // 使用FLGの設定
  const useFlagCol = headers.indexOf('使用FLG');
  if (useFlagCol !== -1) newRow[useFlagCol] = 1;
  
  sheet.appendRow(newRow);
  DataCache.clear(sheetName);
  return { added: true, updated: false };
}



function updateStockThreshold(itemName, newThreshold) {
  const sheet = SS.getSheetByName('T_在庫集計');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const nameCol = headers.indexOf('品名');
  const thresholdCol = headers.indexOf('閾値');
  const updateCol = headers.indexOf('最終更新日');

  if (nameCol === -1 || thresholdCol === -1) {
    return { status: 'error', message: 'Required columns not found in summary' };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === itemName) {
      sheet.getRange(i + 1, thresholdCol + 1).setValue(newThreshold);
      if (updateCol !== -1) sheet.getRange(i + 1, updateCol + 1).setValue(new Date());
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Item not found in summary' };
}

function updateStockItemStatus(itemName, newStatus) {
  const sheet = SS.getSheetByName('T_在庫集計');
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const nameCol = headers.indexOf('品名');
  const flagCol = headers.indexOf('使用FLG');
  const updateCol = headers.indexOf('最終更新日');

  if (nameCol === -1 || flagCol === -1) {
    return { status: 'error', message: 'Required columns not found in summary' };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === itemName) {
      sheet.getRange(i + 1, flagCol + 1).setValue(newStatus);
      if (updateCol !== -1) sheet.getRange(i + 1, updateCol + 1).setValue(new Date());
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Item not found in summary' };
}

function updateStockBulk(thresholdUpdates, statusUpdates) {
  try {
    const sheet = SS.getSheetByName('T_在庫集計');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const now = new Date();
    
    const itemMap = {};
    const nameCol = headers.indexOf('品名');
    if (nameCol === -1) throw new Error('品名列が見つかりません');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol]) itemMap[data[i][nameCol]] = i + 1;
    }
    
    // 閾値の更新
    if (thresholdUpdates) {
      const thresholdCol = headers.indexOf('閾値');
      const updateCol = headers.indexOf('最終更新日');
      for (const itemName in thresholdUpdates) {
        const rowIdx = itemMap[itemName];
        if (rowIdx && thresholdCol !== -1) {
          sheet.getRange(rowIdx, thresholdCol + 1).setValue(thresholdUpdates[itemName]);
          if (updateCol !== -1) sheet.getRange(rowIdx, updateCol + 1).setValue(now);
        }
      }
    }
    
    // ステータスの更新
    if (statusUpdates) {
      const pSheet = SS.getSheetByName('M_商品');
      const pData = pSheet ? pSheet.getDataRange().getValues() : [];
      const pHeaders = pData[0] || [];
      const pNameCol = pHeaders.indexOf('品名');
      const pFlagCol = pHeaders.indexOf('使用FLG');

      for (const itemName in statusUpdates) {
        const rowIdx = itemMap[itemName];
        const newStatus = statusUpdates[itemName];
        if (rowIdx) {
          sheet.getRange(rowIdx, 8).setValue(newStatus); // H列(8)
          sheet.getRange(rowIdx, 7).setValue(now); // G列(7)
        }

        // M_商品側も同期
        if (pSheet && pNameCol !== -1 && pFlagCol !== -1) {
          const pRowIdx = pData.findIndex((r, idx) => idx > 0 && r[pNameCol] === itemName);
          if (pRowIdx !== -1) {
            pSheet.getRange(pRowIdx + 1, pFlagCol + 1).setValue(newStatus);
          }
        }
      }
    }
    
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * 帳簿データのエクスポート (新しいスプレッドシートとして保存)
 */
function exportLedgerReport(period, includePersonal = false) {
  try {
    const ledgerSheet = SS.getSheetByName('T_帳簿');
    if (!ledgerSheet) throw new Error("帳簿シートが見つかりません。");

    const data = getSheetDataAsObjects(ledgerSheet);
    const [selY, selM] = period.split('-');

    const filtered = data.filter(row => {
      const date = new Date(row['日付']);
      if (isNaN(date.getTime())) return false;
      const y = date.getFullYear();
      const m = date.getMonth() + 1;
      const ym = `${y}-${String(m).padStart(2, '0')}`;

      if (selY !== String(y)) return false;
      if (selM === 'FY') {
        // 全月対象
      } else if (selM === 'Q1') {
        if (m < 1 || m > 3) return false;
      } else if (selM === 'Q2') {
        if (m < 4 || m > 6) return false;
      } else if (selM === 'Q3') {
        if (m < 7 || m > 9) return false;
      } else if (selM === 'Q4') {
        if (m < 10 || m > 12) return false;
      } else if (ym !== period) {
        return false;
      }

      // 個人チェックが無い場合は事業のみ（K列が 1 以外のもの）を対象
      if (!includePersonal) {
        const isPersonal = String(row['事業/個人']).trim() === '1';
        if (isPersonal) return false;
      }

      return true;
    });

    if (filtered.length === 0) return { status: 'error', message: '対象期間にデータがありません。' };

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    let periodName = period;
    if (selM === 'FY') periodName = selY + "年度";
    else if (selM.startsWith('Q')) periodName = selY + "年第" + selM.substring(1) + "四半期";
    else periodName = selY + "年" + selM + "月";

    const fileName = `帳簿レポート_${periodName}_${timestamp}`;
    const newSS = SpreadsheetApp.create(fileName);
    const newSheet = newSS.getSheets()[0];
    newSheet.setName('レポート');

    const reportTitle = `${periodName} 帳簿` + (includePersonal ? ' (事業＋個人)' : ' (事業)');
    newSheet.getRange(1, 1).setValue(reportTitle).setFontSize(14).setFontWeight('bold');

    const rawHeaders = ledgerSheet.getRange(1, 1, 1, 10).getValues()[0];
    const headers = [...rawHeaders, '粗利'];
    newSheet.getRange(2, 1, 1, 11).setValues([headers]).setFontWeight('bold').setBackground('#f3f3f3');

    const rows = filtered.map(r => {
      const s = parseNumber(r['売上']);
      const p = parseNumber(r['仕入']);
      const c = parseNumber(r['通信費']);
      const rs = parseNumber(r['修繕費']);
      const sp = parseNumber(r['消耗品費']);
      const d = parseNumber(r['諸会費']);
      const f = parseNumber(r['支払手数料']);
      const ms = parseNumber(r['雑費']);
      const profit = s - (p + c + rs + sp + d + f + ms);

      return [
        r['日付'], r['品名'], s, p, c, rs, sp, d, f, ms, profit
      ];
    });
    newSheet.getRange(3, 1, rows.length, 11).setValues(rows);

    const lastRow = rows.length + 3;
    newSheet.getRange(lastRow, 1).setValue('合計').setFontWeight('bold');
    
    for (let col = 3; col <= 11; col++) {
      const colLetter = String.fromCharCode(64 + col);
      newSheet.getRange(lastRow, col).setFormula(`=SUM(${colLetter}3:${colLetter}${lastRow - 1})`).setFontWeight('bold');
    }

    newSheet.setColumnWidth(2, 370);
    for (let i = 3; i <= 11; i++) {
        newSheet.setColumnWidth(i, 55);
    }
    
    newSheet.getRange(3, 2, rows.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    newSheet.getRange(3, 1, rows.length + 1, 1).setNumberFormat('yyyy/MM/dd');
    newSheet.getRange(3, 3, rows.length + 1, 9).setNumberFormat('#,##0');
    newSheet.setFrozenRows(2);

    return { status: 'success', url: newSS.getUrl(), fileName: fileName };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * 過去の個人販売データをT_帳簿へ一括移行する管理用関数 (1回限り実行用)
 */
function migratePersonalSalesToLedger() {
  const salesSheet = SS.getSheetByName('T_販売');
  const ledgerSheet = SS.getSheetByName('T_帳簿');
  if (!salesSheet || !ledgerSheet) throw new Error("対象シートが見つかりません。");

  // T_帳簿のヘッダー取得とK列「事業/個人」の確保
  let ledgerHeaders = ledgerSheet.getDataRange().getValues()[0].map(h => h.toString().trim());
  let typeCol = ledgerHeaders.indexOf('事業/個人');
  if (typeCol === -1) {
    ledgerSheet.getRange(1, ledgerHeaders.length + 1).setValue('事業/個人');
    ledgerHeaders.push('事業/個人');
    typeCol = ledgerHeaders.length - 1;
  }

  const dateCol = ledgerHeaders.indexOf('日付');
  const descCol = ledgerHeaders.indexOf('品名');
  const incomeCol = ledgerHeaders.indexOf('売上');
  const costCol = ledgerHeaders.indexOf('仕入');
  const commFeeCol = ledgerHeaders.indexOf('通信費');
  const feeCol = ledgerHeaders.indexOf('支払手数料');

  // 既存T_帳簿データの読み込み
  const ledgerData = ledgerSheet.getDataRange().getValues();
  const existingPersonalRecords = new Set();
  let existingPersonalFlaggedCount = 0;

  for (let r = 1; r < ledgerData.length; r++) {
    const row = ledgerData[r];
    const rowDate = row[dateCol];
    const rowDesc = String(row[descCol] || "").trim();
    const rowIncome = parseNumber(row[incomeCol]);
    const rowType = String(row[typeCol] || "").trim();
    
    const isManualPersonal = rowDesc.includes('個人取引:');
    const isMarkedPersonal = (rowType === '1');

    // 既存データで「個人取引:」が含まれている場合、K列を「1」に更新
    if (isManualPersonal && rowType !== '1') {
      ledgerSheet.getRange(r + 1, typeCol + 1).setValue(1);
      existingPersonalFlaggedCount++;
    }

    // 重複照合の対象は「個人取引」として記録されているレコードのみとする（事業用データとは照合しない）
    if ((isManualPersonal || isMarkedPersonal) && rowDate) {
      const d = new Date(rowDate);
      if (!isNaN(d.getTime())) {
        const dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
        // 品名の正規化（「個人取引:」や空白記号を除去して比較）
        const cleanDesc = rowDesc.replace(/個人取引:/g, "").replace(/\s+/g, "");
        const key = `${dStr}_${rowIncome}_${cleanDesc}`;
        existingPersonalRecords.add(key);
      }
    }
  }

  // T_販売データの読み込み
  const salesObjects = getSheetDataAsObjects(salesSheet);
  const rowsToAdd = [];
  let skippedAdjustCount = 0;
  let skippedDuplicateCount = 0;
  let skippedNotTargetCount = 0;
  let addedCount = 0;

  salesObjects.forEach(item => {
    const status = String(item['ステータス'] || "").trim();
    if (status !== '完了') {
      skippedNotTargetCount++;
      return;
    }

    // 個人取引判定 (管理区分 == 1 または 管理対象外 == 1)
    const isPersonal = (item['管理区分'] == 1 || item['管理対象外'] == 1);
    if (!isPersonal) {
      skippedNotTargetCount++;
      return;
    }

    // 発送（発送方法）が「調整用」のものは移行対象外
    const shipping = String(item['発送方法'] || item['発送'] || "").trim();
    if (shipping === '調整用') {
      skippedAdjustCount++;
      Logger.log(`[スキップ: 調整用] ID: ${item['販売ID']}, 発送: 調整用`);
      return;
    }

    const compDateStr = item['取引完了日'] || "";
    const compDate = new Date(compDateStr);
    if (isNaN(compDate.getTime())) {
      skippedNotTargetCount++;
      return;
    }

    const compDateFormatted = Utilities.formatDate(compDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const cutoffDate = new Date("2026/04/01");

    // 品名の決定ルール
    let itemName = "";
    if (compDate < cutoffDate || String(item['品名'] || "").trim() === '個人取引（在庫管理なし）') {
      // 旧システムデータ: 備考から品名を取得
      itemName = String(item['備考'] || "").trim();
    } else {
      // 新システムデータ: 品名から取得
      itemName = String(item['品名'] || "").trim();
    }
    if (!itemName) itemName = "個人取引";

    const buyer = String(item['売先'] || "").trim();
    const qty = parseFloat(item['数量']) || 1;
    let desc = buyer ? `${buyer} ${itemName}` : itemName;
    if (qty > 1) {
      desc += "×" + qty;
    }

    const price = parseFloat(item['合計金額'] || item['価格'] || item['販売価格'] || 0);

    // 重複チェック（既存の個人取引レコードのみと照合）
    const cleanDesc = desc.replace(/個人取引:/g, "").replace(/\s+/g, "");
    const matchKey = `${compDateFormatted}_${price}_${cleanDesc}`;

    if (existingPersonalRecords.has(matchKey)) {
      skippedDuplicateCount++;
      Logger.log(`[スキップ: 登録済重複] ID: ${item['販売ID']}, 日付: ${compDateFormatted}, 売上: ${price}, 品名: ${desc}`);
      return;
    }

    // T_帳簿用レコードの作成
    let newRow = new Array(ledgerHeaders.length).fill("");
    if (dateCol !== -1) newRow[dateCol] = compDateFormatted;
    if (descCol !== -1) newRow[descCol] = desc;
    if (incomeCol !== -1) newRow[incomeCol] = price;
    if (costCol !== -1) newRow[costCol] = parseFloat(item['合計単価']) || 0;
    if (commFeeCol !== -1) newRow[commFeeCol] = parseFloat(item['送料']) || 0;
    if (feeCol !== -1) newRow[feeCol] = parseFloat(item['手数料']) || 0;
    if (typeCol !== -1) newRow[typeCol] = 1; // 個人

    rowsToAdd.push(newRow);
    existingPersonalRecords.add(matchKey);
    addedCount++;
  });

  if (rowsToAdd.length > 0) {
    const startRow = ledgerSheet.getLastRow() + 1;
    ledgerSheet.getRange(startRow, 1, rowsToAdd.length, ledgerHeaders.length).setValues(rowsToAdd);
    
    // 日付順にソート
    const totalRows = ledgerSheet.getLastRow();
    if (totalRows > 1) {
      ledgerSheet.getRange(2, 1, totalRows - 1, ledgerHeaders.length).sort({column: 1, ascending: true});
    }
  }

  const resultMsg = `移行結果: 新規追加 ${addedCount} 件, スキップ計 ${skippedAdjustCount + skippedDuplicateCount} 件 (内訳: 調整用 ${skippedAdjustCount} 件, 登録済重複 ${skippedDuplicateCount} 件), 既存手動データのK列フラグ設定 ${existingPersonalFlaggedCount} 件`;
  Logger.log(resultMsg);
  return { 
    status: 'success', 
    message: resultMsg, 
    addedCount, 
    skippedAdjustCount, 
    skippedDuplicateCount, 
    existingPersonalFlaggedCount 
  };
}

/**
 * 汎用数値変換 (GAS backend用)
 */
function parseNumber(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[^0-9.-]+/g, "")
                  .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * 商品画像のアップロードとM_商品への反映
 */
function uploadProductImage(itemName, base64Data) {
  try {
    const folder = getOrCreateImageFolder();
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, itemName + ".jpg");

    // 既存の同名ファイルを検索（あれば上書きのために削除）
    const existingFiles = folder.getFilesByName(itemName + ".jpg");
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // thumbnail URL (300px) もしくは 直接の表示URL
    const imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w500";

    // M_商品 シートの更新 (B列が品名、F列が画像URL)
    const pSheet = SS.getSheetByName('M_商品');
    const data = pSheet.getDataRange().getValues();
    const headers = data[0];
    const itemColIdx = headers.indexOf('品名');
    let urlColIdx = headers.indexOf('画像URL');

    // 画像URL列がなければF列(インデックス5)に作成
    if (urlColIdx === -1) {
      urlColIdx = 5;
      pSheet.getRange(1, urlColIdx + 1).setValue('画像URL');
    }

    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][itemColIdx] === itemName) {
        pSheet.getRange(i + 1, urlColIdx + 1).setValue(imageUrl);
        found = true;
        break;
      }
    }

    if (!found) {
      return { status: 'error', message: 'M_商品に該当品名が見つかりません: ' + itemName };
    }

    return { status: 'success', url: imageUrl };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * 画像保存用のフォルダを取得または作成
 */
function getOrCreateImageFolder() {
  const folderName = "統合システム_商品画像";
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}


/**
 * 権限承認専用の関数
 */
function authorizeDrive() {
  const dummy = DriveApp.createFolder("承認用一時フォルダ");
  dummy.setTrashed(true);
  console.log("ドライブの作成・書き込み権限が承認されました。");
}

/**
 * 単品アイテムの在庫充足チェック
 */
function checkStockAvailability(itemName, demandQty) {
  const sheet = SS.getSheetByName('T_在庫集計');
  if (!sheet) return { ok: true }; 

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('品名');
  const qtyCol = headers.indexOf('現在庫数');
  
  if (nameCol === -1 || qtyCol === -1) return { ok: true };

  const row = data.find(r => r[nameCol] === itemName);
  const currentQty = row ? (parseFloat(row[qtyCol]) || 0) : 0;

  if (currentQty < demandQty) {
    return { ok: false, shortage: (demandQty - currentQty).toLocaleString() };
  }
  return { ok: true };
}

/**
 * 製造（BOM）に基づく部品全体の在庫充足チェック
 */
function checkBOMAvailability(productName, makeQty) {
  const bomSheet = SS.getSheetByName('M_BOM');
  if (!bomSheet) return { ok: true };

  const bomData = getSheetDataAsObjects(bomSheet);
  const components = bomData.filter(r => r['品名'] === productName);
  
  if (components.length === 0) return { ok: true };

  for (const comp of components) {
    const compName = comp['部品'];
    const needed = (parseFloat(comp['数量']) || 0) * makeQty;
    const res = checkStockAvailability(compName, needed);
    if (!res.ok) {
      return { ok: false, message: `${compName} があと ${res.shortage} 不足しています` };
    }
  }
  return { ok: true };
}

/**
 * 棚卸登録処理
 * @param {Array} stocktakeData {itemName, actualQty, logicalQty, diff} の配列
 * @param {string} note 備考
 */
function registerStocktake(stocktakeData, note) {
  StockManager.init(); // 在庫管理データをメモリにロード
  const stockSheet = SS.getSheetByName('T_在庫管理');
  const histSheet = SS.getSheetByName('T_棚卸履歴');
  if (!stockSheet || !histSheet) return { status: 'error', message: 'シートが見つかりません' };

  const timestamp = new Date();
  const dateOnly = Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd");
  const adjId = generateNextId(histSheet, 'A');
  
  // 1. 棚卸履歴の保存
  const totalItems = stocktakeData.length;
  const diffItems = stocktakeData.filter(d => d.diffQty !== 0).length;
  
  histSheet.appendRow([
    adjId,
    timestamp,
    totalItems,
    diffItems,
    note
  ]);

  // 2. 在庫調整の実行
  const summarySheet = SS.getSheetByName('T_在庫集計');
  const summaryData = summarySheet.getDataRange().getValues();
  const summaryItemMap = {};
  for (let i = 1; i < summaryData.length; i++) {
    if (summaryData[i][2]) summaryItemMap[summaryData[i][2]] = i + 1; // C列
  }

  stocktakeData.forEach(adj => {
    // 差異の有無に関わらず、最終棚卸日を更新 (I列: インデックス9)
    const rowIdx = summaryItemMap[adj.itemName];
    if (rowIdx) {
      summarySheet.getRange(rowIdx, 9).setValue(timestamp);
    }

    if (adj.diffQty === 0) return; // 差異がないものは在庫管理レコードの追加をスキップ

    if (adj.diffQty > 0) {
      // プラス調整：新規ロットとして追加
      processSurplusAdjustment(adj, adjId, timestamp, dateOnly);
    } else {
      // マイナス調整：既存在庫からFIFOで消し込み
      processShortageAdjustment(adj, adjId, timestamp, dateOnly);
    }
  });

  // 3. 在庫集計の更新（在庫数のみ）
  StockManager.flush(); // メモリ上の変更をシートに反映
  updateInventorySummary();
  SpreadsheetApp.flush();

  return { status: 'success', adjId: adjId };
}

/**
 * プラス調整：最新単価を引き継いで新規入庫
 */
function processSurplusAdjustment(adj, adjId, timestamp, dateOnly) {
  const headers = StockManager.headers;
  const data = StockManager.data;
  
  // マスタからカテゴリ取得（recordInventory関数と同じ方式）
  const pSheet = SS.getSheetByName('M_商品');
  const pData = pSheet ? pSheet.getDataRange().getValues() : [];
  const pHeaders = pData.length > 0 ? pData[0].map(h => h.toString().trim()) : [];
  const nameIdx = pHeaders.indexOf('品名');
  const catIdx = pHeaders.indexOf('カテゴリ');
  const pRow = (nameIdx !== -1) ? pData.slice(1).find(r => r[nameIdx] === adj.itemName) : null;
  const category = (pRow && catIdx !== -1) ? (pRow[catIdx] || "") : "";

  // 最新単価とその元在庫管理IDの取得
  let latestPrice = 0;
  let latestSourceId = "";
  const nameCol = headers.indexOf('品名');
  const priceCol = headers.indexOf('単価');
  const idCol = headers.indexOf('在庫管理ID');
  
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][nameCol] === adj.itemName && (parseFloat(data[i][priceCol]) || 0) > 0) {
      latestPrice = parseFloat(data[i][priceCol]);
      latestSourceId = data[i][idCol] || "";
      break;
    }
  }

  const id = StockManager.getNextInvId();
  const rowObj = {
    '在庫管理ID': id,
    '品名': adj.itemName,
    '商品区分': category,
    '入出庫日': dateOnly,
    '区分': '棚卸入庫',
    '数量': adj.diffQty,
    '単価': latestPrice,
    '棚卸ID': adjId,
    '仕入先': '棚卸調整',
    '引当元管理ID': latestSourceId,
    '引当完了': 0,
    '実在庫数量': adj.diffQty,
    '備考': `実${adj.actualQty}(理${adj.logicalQty})`,
    '最終更新日': timestamp
  };

  StockManager.appendRow(rowObj);
}


/**
 * マイナス調整：既存在庫からFIFOで消し込み
 */
function processShortageAdjustment(adj, adjId, timestamp, dateOnly) {
  const headers = StockManager.headers;
  const data = StockManager.data;
  
  const nameCol = headers.indexOf('品名');
  const actualQtyCol = headers.indexOf('実在庫数量');
  const unitPriceCol = headers.indexOf('単価');
  const statusCol = headers.indexOf('引当完了');
  const idCol = headers.indexOf('在庫管理ID');
  
  let shortageLeft = Math.abs(adj.diffQty);
  const adjustments = [];

  // 有効なロットを古い順に探す
  for (let i = 1; i < data.length; i++) {
    if (shortageLeft <= 0) break;
    
    if (data[i][nameCol] === adj.itemName && (parseFloat(data[i][actualQtyCol]) || 0) > 0) {
      const available = parseFloat(data[i][actualQtyCol]);
      const consume = Math.min(available, shortageLeft);
      const unitPrice = parseFloat(data[i][unitPriceCol]) || 0;
      const sourceId = data[i][idCol];

      // 元行の実在庫を減らす
      const newActual = available - consume;
      data[i][actualQtyCol] = newActual;
      if (newActual === 0 && statusCol !== -1) {
        data[i][statusCol] = 1;
      }
      StockManager.markModified(i);

      // 調整行のデータを準備
      adjustments.push({
        sourceId: sourceId,
        qty: -consume,
        price: unitPrice
      });

      shortageLeft -= consume;
    }
  }

  // マスタからカテゴリ取得
  const pSheet = SS.getSheetByName('M_商品');
  const pData = pSheet ? pSheet.getDataRange().getValues() : [];
  const pHeaders = pData.length > 0 ? pData[0].map(h => h.toString().trim()) : [];
  const nameIdx = pHeaders.indexOf('品名');
  const catIdx = pHeaders.indexOf('カテゴリ');
  const pRow = (nameIdx !== -1) ? pData.slice(1).find(r => r[nameIdx] === adj.itemName) : null;
  const category = (pRow && catIdx !== -1) ? (pRow[catIdx] || "") : "";

  // 調整行（証拠）の追加
  adjustments.forEach(item => {
    const id = StockManager.getNextInvId();
    const rowObj = {
      '在庫管理ID': id,
      '品名': adj.itemName,
      '商品区分': category,
      '入出庫日': dateOnly,
      '区分': '棚卸出庫',
      '数量': item.qty,
      '単価': item.price,
      '棚卸ID': adjId,
      '仕入先': '棚卸調整',
      '引当元管理ID': item.sourceId,
      '引当完了': 1,
      '実在庫数量': 0,
      '備考': `実${adj.actualQty}(理${adj.logicalQty})`,
      '最終更新日': timestamp
    };

    StockManager.appendRow(rowObj);
  });
}

/**
 * マスタデータの汎用更新
 * 第一列をIDとして検索し、該当行のデータを更新する
 */
function updateMasterRecord(masterName, id, updates) {
  try {
    const sheet = SS.getSheetByName(masterName);
    if (!sheet) throw new Error("Master sheet not found: " + masterName);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    
    // マスタごとのキー項目を定義
    const keyConfig = {
      'M_商品': '品名',
      'M_仕入先': '仕入先',
      'M_売先': '売先',
      'M_発送': '発送方法',
      'M_経費品名': '品名',
      'M_支払': '支払方法',
      'M_仕訳': '仕訳名',
      'M_BOM': '品名', // BOMは品名で検索してから部品をチェック（後述）
      'T_在庫集計': '品名'
    };

    const keyHeader = keyConfig[masterName] || headers[0];
    const keyColIdx = headers.indexOf(keyHeader);
    
    let rowIndex = -1;
    if (masterName === 'M_BOM' && updates.hasOwnProperty('部品')) {
      // BOMは複合キー対応
      rowIndex = data.findIndex((row, idx) => 
        idx > 0 && row[headers.indexOf('品名')] === id && row[headers.indexOf('部品')] === updates['部品']
      );
    } else {
      rowIndex = data.findIndex((row, idx) => idx > 0 && row[keyColIdx].toString() == id.toString());
    }
    
    if (rowIndex === -1) throw new Error("ID not found: " + id + " in " + masterName);
    
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([[...data[rowIndex]]]); // 元の行データを一旦確保（同期用）
    
    for (const key in updates) {
      const colIdx = headers.indexOf(key);
      if (colIdx !== -1) {
        sheet.getRange(rowIndex + 1, colIdx + 1).setValue(updates[key]);
      }
    }

    // --- 使用FLGの同期ロジック (提案7) ---
    if (updates.hasOwnProperty('使用FLG')) {
      const newStatus = updates['使用FLG'];
      
      if (masterName === 'M_商品') {
        const itemName = data[rowIndex][headers.indexOf('品名')];
        if (itemName) {
          updateStockItemStatus(itemName, newStatus);
        }
      } else if (masterName === 'T_在庫集計') {
        const nameCol = headers.indexOf('品名');
        const itemName = (nameCol !== -1) ? data[rowIndex][nameCol] : null; 
        const pSheet = SS.getSheetByName('M_商品');
        if (pSheet) {
          const pData = pSheet.getDataRange().getValues();
          const pHeaders = pData[0];
          const nameCol = pHeaders.indexOf('品名');
          const flagCol = pHeaders.indexOf('使用FLG');
          if (nameCol !== -1 && flagCol !== -1) {
            const pRowIdx = pData.findIndex((r, idx) => idx > 0 && r[nameCol] === itemName);
            if (pRowIdx !== -1) {
              pSheet.getRange(pRowIdx + 1, flagCol + 1).setValue(newStatus);
            }
          }
        }
      }
    }
    
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * マスタデータの新規追加
 */
function addMasterRecord(masterName, updates) {
  try {
    const sheet = SS.getSheetByName(masterName);
    if (!sheet) throw new Error("Master sheet not found: " + masterName);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    
    const newRow = headers.map(h => {
      if (updates.hasOwnProperty(h)) return updates[h];
      if (h === '最終更新日') return new Date();
      if (h === '使用FLG') return 1; // デフォルト有効
      return "";
    });
    
    sheet.appendRow(newRow);

    // M_商品追加時はT_在庫集計にも初期行を作成
    if (masterName === 'M_商品') {
       updateInventorySummary(updates['品名']);
    }

    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

 
