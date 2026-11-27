document.addEventListener('DOMContentLoaded', async () => {
    const APP_VERSION = '2026.11.27.4'; // version.jsonからbuild.ps1で同期
    console.log('System initialization started... (' + APP_VERSION + ')');

    // ---- API Configuration ----
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbzY_IuQn_5tG_bGBP459xIPoekQoeEdxEo5nfIO3yeC8chmyHicQNzjgseGXtlol2q0SA/exec';
    let currentAuthKey = '';
    let sessionToken = '';
    const rememberedLoginKey = 'inventory_login_session';
    const rememberLoginOption = document.getElementById('remember-login-option');
    const rememberLogin = document.getElementById('remember-login');
    const logout = document.getElementById('logout-btn');
    const supportsRememberedLogin = !window.google?.script?.run &&
        Boolean(rememberLoginOption && rememberLogin && logout);
    localStorage.removeItem('inventory_auth_key');

    // Global state
    let currentMasters = {};
    let lastHistoryData = {}; // 帳簿表示用などにデータを保持
    let lastRawData = {};    // サーバーから届いた生データを保持（マージ用）
    let ledgerYear = new Date().getFullYear();
    let ledgerPeriod = String(new Date().getMonth() + 1).padStart(2, '0'); // デフォルトは「今月」
    let ledgerIncludePersonal = false; // 月次明細の個人合算フラグ（デフォルト: 事業のみ）
    let dashboardSelectedMonth = null; // グラフ内の選択月。共通期間・帳票とは独立

    let isStocktakeMode = false;
    let stocktakeData = {}; // { itemName: { actual: number, diff: number } }
    const STOCKTAKE_CYCLE_DAYS = 30; // サイクルカウントの基準日数

    let pendingStockChanges = {
        thresholds: {}, // itemName -> newValue
        statuses: {}    // itemName -> newValue (0 or 1)
    };

    let html5QrCode = null; // Scanner instance

    // 買い物カゴ (Proposal 38)
    window.registrationBasket = [];

    // 買い物カゴ (Proposal 38)
    window.registrationBasket = [];

    // マスタ編集用スキーマ定義 (提案7・マスター管理強化)
    const MASTER_SCHEMAS = {
        'M_商品': {
            key: '品名',
            fields: [
                { name: '商品ID', type: 'text', visible: true, editable: false, required: true },
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '品名', type: 'text', visible: true, editable: false, required: true },
                { name: 'カテゴリ', type: 'select', visible: true, editable: true, options: ['パーツ', 'パーツ2', '単体商品', '商品', '経費'], required: true },
                { name: '説明', type: 'textarea', visible: true, editable: true, required: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true },
                { name: '画像URL', visible: false, required: false },
                { name: '保管場所', type: 'select', options: ['1_棚上：メイン', '2_棚中：サブ1', '3_棚下：梱包1', '4_押入上1：部品1', '5_押入上2：部品2', '6_押入上3：梱包2', '7_押入下1：サブ2', '8_押入下2：部品箱'], visible: true, editable: true, required: false },
                { name: 'QR/バーコード', type: 'text', visible: true, editable: true, required: false },
                { name: '販売単価', type: 'number', visible: true, editable: true, required: false }
            ]
        },
        'M_仕入先': {
            key: '仕入先',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '仕入先', type: 'text', visible: true, editable: false, required: true },
                { name: '用途区分', type: 'select', visible: true, editable: true, options: [{ v: 1, l: '1:仕入のみ' }, { v: 2, l: '2:経費のみ' }, { v: 3, l: '3:両方' }], required: true },
                { name: '説明', type: 'textarea', visible: true, editable: true, required: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'M_売先': {
            key: '売先',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '売先', type: 'text', visible: true, editable: false, required: true },
                { name: '手数料率', type: 'number', visible: true, editable: true, required: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'M_発送': {
            key: '発送方法',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '発送方法', type: 'text', visible: true, editable: false, required: true },
                { name: '送料', type: 'number', visible: true, editable: true, required: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'M_BOM': {
            key: '品名',
            fields: [
                { name: '品名', type: 'select', visible: true, editable: false, refMaster: 'M_商品', filter: (r) => r['カテゴリ'] === '商品', required: true },
                { name: '部品', type: 'select', visible: true, editable: false, refMaster: 'M_商品', filter: (r) => ['パーツ', '単体商品'].includes(r['カテゴリ']), required: true },
                { name: '数量', type: 'number', visible: true, editable: true, required: true },
                { name: '説明', type: 'textarea', visible: true, editable: true, required: false }
            ]
        },
        'M_経費品名': {
            key: '品名',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '品名', type: 'text', visible: true, editable: false, required: true },
                { name: 'デフォルト仕訳', type: 'select', visible: true, editable: true, refMaster: 'M_仕訳', required: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'T_在庫集計': {
            key: '品名',
            fields: [
                { name: '優先度', type: 'number', visible: true, editable: true, required: true },
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '品名', type: 'text', visible: true, editable: false, required: true },
                { name: 'カテゴリ', type: 'text', visible: true, editable: false, required: true },
                { name: '現在庫数', type: 'number', visible: true, editable: false, required: true },
                { name: '閾値', type: 'number', visible: true, editable: true, required: true },
                { name: '最終更新日', visible: false, required: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true },
                { name: '最終棚卸日', type: 'text', visible: true, editable: false, required: false },
                { name: '商品ID', type: 'text', visible: true, editable: false, required: true },
                { name: '保管場所', type: 'select', options: ['1_棚上：メイン', '2_棚中：サブ1', '3_棚下：梱包1', '4_押入上1：部品1', '5_押入上2：部品2', '6_押入上3：梱包2', '7_押入下1：サブ2', '8_押入下2：部品箱'], visible: true, editable: false, required: false },
                { name: 'QR/バーコード', type: 'text', visible: true, editable: false, required: false }
            ]
        },
        'M_支払': {
            key: '支払方法',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '支払方法', type: 'text', visible: true, editable: false, required: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'M_ステータス': {
            key: 'ステータス名称',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true, required: true },
                { name: '対象機能', type: 'text', visible: true, editable: true, required: true },
                { name: '画面名称', type: 'text', visible: true, editable: true, required: true },
                { name: 'ステータス名称', type: 'text', visible: true, editable: true, required: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        },
        'M_画面制御': {
            key: '要素ID',
            fields: [
                { name: '対象機能', type: 'text', visible: true, editable: true, required: true },
                { name: '画面名称', type: 'text', visible: true, editable: true, required: true },
                { name: '要素ID', type: 'text', visible: true, editable: true, required: true },
                { name: '画面上の項目名', type: 'text', visible: true, editable: true, required: true },
                { name: 'タイプ', type: 'text', visible: true, editable: true, required: true },
                { name: '参照マスタ', type: 'text', visible: true, editable: true, required: true },
                { name: '固定値の内容', type: 'text', visible: true, editable: true, required: false },
                { name: '説明', type: 'textarea', visible: true, editable: true, required: false },
                { name: '抽出条件', type: 'textarea', visible: true, editable: true, required: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true, required: true }
            ]
        }
    };

    // 棚卸管理用の状態
    let stocktakeSession = {
        verifiedItems: new Set(), // 現在のセッションで「済」にした品名
        uncheckedOnly: false      // フィルタ状態
    };

    const historyBulkState = new Map();
    const historyFilters = new Map();
    const historyRenderIds = new Map();
    const historySheets = { purchase: "T_仕入", expense: "T_経費", manufacturing: "T_製造", sales: "T_販売" };
    let historySearchSequence = 0;
    let transactionDetailSequence = 0;
    let transactionDetailOpener = null;
    let shippingChangeState = { id: '', record: null, options: [], preview: null, pending: null, busy: false, sequence: 0 };
    let purchaseCorrectionState = { id: '', record: null, preview: null, pending: null, busy: false, sequence: 0 };

    try {
        // ---- 1. Event Listener Registrations (Essential UI) ----
        setupNavigation();
        setupToggleLogics();
        setupSettingsListeners();
        setupAnalysisAndSearchListeners();
        setupHistoryFilters();
        setupHistoryBulkStatus();
        setupShippingChange_();
        setupPurchaseCorrection_();
        setupScrollTopStockBar();
        setupScannerListeners(); // スキャナーはデータロードを待たずに即座に有効化
        setupInventoryCheckListeners();

        // ---- 2. System Initialization (Data Fetching) ----
        // 認証チェック (GAS環境以外の場合)
        // クイック検索モーダルのイベント設定
        const qsCloseBtn = document.getElementById('quick-search-close-btn');
        if (qsCloseBtn) qsCloseBtn.onclick = closeQuickSearch;

        const qsInput = document.getElementById('quick-search-input');
        if (qsInput) qsInput.oninput = () => { if (typeof window.applyQuickSearchFilter === 'function') window.applyQuickSearchFilter(); };

        // クイック検索モーダル外クリックで閉じる
        window.onclick = function (event) {
            const modal = document.getElementById('quick-search-modal');
            if (event.target == modal) {
                closeQuickSearch();
            }
        };

        // 初期化完了後に検索ボタンを配置
        initQuickSearchButtons();
        if (supportsRememberedLogin) {
            rememberLoginOption.hidden = false;
            logout.style.display = '';
            logout.onclick = async () => {
                logout.disabled = true;
                try {
                    if (sessionToken) {
                        const result = await fetchAPI('endSession');
                        if (result.status !== 'success' && result.message !== 'Unauthorized') {
                            throw new Error('Logout failed');
                        }
                    }
                    clearLoginSession();
                    window.location.reload();
                } catch {
                    showToast('ログアウトできませんでした。通信状況を確認してください。', 'error');
                } finally {
                    logout.disabled = false;
                }
            };
        }
        if (await restoreLoginSession()) {
            setupLoginHandlers(null);
        } else {
            await new Promise(resolve => {
                setupLoginHandlers(resolve);
                showLoginModal();
            });
        }
        // 認証済みの場合、即座にコンテナを表示 (読み込み中アニメーションを表示させるため)
        const appContainer = document.querySelector('.app-container');
        if (appContainer) {
            appContainer.style.display = 'flex';
        }

        // 1. まずキャッシュからマスタと履歴を読み込んで即座にUIを構築
        loadMastersFromCache();
        loadHistoryFromCache();

        // 2. 最新データをバックグラウンドで非同期に取得してサイレント更新
        initSystem('all').then(() => {
            console.log("Background data load completed.");
        }).catch(err => {
            console.warn("Background load failed:", err);
        });

        // ---- 3. Image Feature Initializations ----
        setupImagePreviewListeners();
        setupStockUpdateListeners();
        setupStockFilters();

        console.log("System initialization completed successfully.");
    } catch (error) {
        console.error("Critical System Error:", error);

        // 認証エラー(401)の場合はアラートを出さずhandleUnauthorizedに任せる
        if (error.message.includes("Unauthorized") || error.message.includes("401")) {
            handleUnauthorized();
        } else {
            alert("システムの起動中に致命的なエラーが発生しました。\n\nエラー内容: " + error.message);
        }

        // エラー時はコンテナを隠したままにする
        const appContainer = document.querySelector('.app-container');
        if (appContainer) appContainer.style.display = 'none';

        setupNavigation();
    }

    //* --- 利益・進捗バッジ --- */

    async function initSystem(scope = 'all', options = {}) {
        console.log(`Fetching system init data (scope: ${scope}, skipMasters: ${options.skipMasters})...`);
        try {
            const response = await fetchAPI('getInitData', { scope: scope, skipMasters: options.skipMasters });
            if (response.status === 'success') {
                updateVersionDisplay_(response.data.version);
                // マスタの処理（二次元配列をオブジェクトに変換）
                const masters = {};
                for (const key in response.data.masters) {
                    let records = convertRawToObjects(response.data.masters[key]);
                    // 使用FLGと表示順の適用
                    if (key !== 'T_在庫集計' && records.length > 0 && records[0].hasOwnProperty('使用FLG')) {
                        records = records.filter(r => r['使用FLG'] == 1 || r['使用FLG'] === true);
                    }
                    if (records.length > 0 && records[0].hasOwnProperty('表示順')) {
                        records.sort((a, b) => (parseInt(a['表示順']) || 0) - (parseInt(b['表示順']) || 0));
                    }
                    masters[key] = records;
                }
                currentMasters = Object.assign({}, currentMasters, masters);

                // 在庫集計のデータ拡張 (商品ID, 保管場所の紐付け)
                if (currentMasters['T_在庫集計'] && currentMasters['M_商品']) {
                    const itemMap = {};
                    currentMasters['M_商品'].forEach(m => itemMap[m['品名']] = m);
                    currentMasters['T_在庫集計'].forEach(r => {
                        const m = itemMap[r['品名']];
                        if (m) {
                            r['商品ID'] = m['商品ID'] || '';
                            r['保管場所'] = m['保管場所'] || '';
                            r['QR/バーコード'] = m['QR/バーコード'] || '';
                        }
                    });
                }

                saveMastersToCache(currentMasters);
                buildDynamicUI(currentMasters);

                // 履歴データの処理（二次元配列をオブジェクトに変換し、フロント側で集計）
                if (response.data.historyData && response.data.historyData.isRaw) {
                    // 生データをキャッシュに保持（マージ用）
                    lastRawData = Object.assign({}, lastRawData, response.data.historyData.rawData);
                    saveHistoryToCache(lastRawData);

                    const processed = processClientData(lastRawData, scope);
                    lastHistoryData = processed;
                } else {
                    // フォールバック（旧形式の場合）
                    if (scope === 'essential') {
                        lastHistoryData = response.data.historyData;
                    } else {
                        if (response.data.historyData.history) {
                            lastHistoryData.history = Object.assign({}, lastHistoryData.history, response.data.historyData.history);
                        }
                    }
                }

                renderAllHistory(lastHistoryData);

                const activeTab = document.querySelector('.nav-item.active');
                if (activeTab && activeTab.getAttribute('data-target') === 'ledger') {
                    renderLedger();
                }

                attachHistoryListeners();
                return response.data;
            } else {
                throw new Error(response.message || "API側からエラーが返されました");
            }
        } catch (e) {
            console.error("Failed to init system:", e);
            alert("データの取得中にエラーが発生しました。\nスコープ: " + scope + "\n詳細: " + e.message);
            throw e; // 上位のcatchに渡す
        }
    }


    /**
     * 二次元配列をオブジェクトに変換する
     */
    function convertRawToObjects(raw) {
        if (!Array.isArray(raw) || raw.length < 1 || !Array.isArray(raw[0])) {
            console.warn("Invalid raw data format for conversion:", raw);
            return [];
        }
        const headers = raw[0].map(h => (h ? h.toString().trim() : ""));
        // 有効なヘッダーがない場合は空を返す
        if (headers.every(h => h === "")) return [];

        return raw.slice(1).filter(r => Array.isArray(r) && r.some(v => v !== "" && v !== null && v !== undefined)).map(r => {
            const obj = {};
            headers.forEach((h, i) => { if (h) obj[h] = r[i]; });
            return obj;
        });
    }

    /**
     * クライアント側でのデータ集計とフィルタリング
     */
    function parseRecentActivityDate_(value) {
      if (Object.prototype.toString.call(value) === '[object Date]') {
        return Number.isFinite(value.getTime()) ? new Date(value.getTime()) : null;
      }
      if (typeof value !== 'string') return null;
      const match = value.trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(Z|[+-]\d{2}:\d{2})?)?$/);
      if (!match) return null;
      const [, year, month, day, hour = '0', minute = '0', second = '0', fraction = '', zone] = match;
      const values = [Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second), Number(fraction.padEnd(3, '0'))];
      const check = new Date(Date.UTC(...values));
      const actual = [check.getUTCFullYear(), check.getUTCMonth(), check.getUTCDate(), check.getUTCHours(), check.getUTCMinutes(), check.getUTCSeconds(), check.getUTCMilliseconds()];
      if (values.some((value, index) => value !== actual[index])) return null;
      if (!zone) return new Date(check.getTime() - 9 * 60 * 60 * 1000);
      const iso = year + '-' + month.padStart(2, '0') + '-' + day.padStart(2, '0') + 'T' + hour.padStart(2, '0') + ':' + minute + ':' + second + '.' + fraction.padEnd(3, '0') + zone;
      const result = new Date(iso);
      return Number.isFinite(result.getTime()) ? result : null;
    }

    function createRecentAction(sheetName, row, now = new Date()) {
        const info = { 'T_仕入': ['purchase', '仕入ID'], 'T_経費': ['expense', '経費ID'],
            'T_製造': ['manufacturing', '製造ID'], 'T_販売': ['sales', '販売ID'] }[sheetName];
        if (!info || !row[info[1]]) return null;
        let date = parseRecentActivityDate_(row['最終更新日']);
        const isUpdated = Boolean(date);
        if (!date) {
            for (const key of Object.keys(row)) {
                if (key === '最終更新日' || !key.includes('日')) continue;
                const candidate = parseRecentActivityDate_(row[key]);
                if (candidate && candidate <= now && (!date || candidate > date)) date = candidate;
            }
        }
        if (!date) return null;
        const japan = new Date(date.getTime() + 9 * 60 * 60 * 1000);
        const pad = value => String(value).padStart(2, '0');
        const day = pad(japan.getUTCMonth() + 1) + '/' + pad(japan.getUTCDate());
        const dateStr = isUpdated ? '更新：' + day + ' ' + pad(japan.getUTCHours()) + ':' + pad(japan.getUTCMinutes()) : '取引日：' + day;
        const price = parseFloat(row['価格'] || row['合計金額'] || row['販売価格'] || 0);
        return { type: info[0], id: row[info[1]], date, dateStr, isUpdated,
            itemName: row['品名'], amount: sheetName === 'T_製造' ? 0 : sheetName === 'T_販売' ? price : -price,
            status: String(row['ステータス'] || '').trim(), quantity: row['数量'] || row['製造数量'],
            buyer: sheetName === 'T_販売' ? row['売先'] : null };
    }

    function processClientData(rawData, scope) {
        console.time('Client:processClientData');
        const history = {};
        const summary = { purchase: 0, expense: 0, sales: 0, businessSales: 0, personalSales: 0, expectedSales: 0, completedSalesCount: 0 };
        const recentAll = [];
        const personalSalesByMonth = {};

        const now = new Date();
        const thisYear = now.getFullYear();
        const thisMonth = now.getMonth();
        const thresholdDate = new Date(); thresholdDate.setDate(now.getDate() - 40);
        const mThresholdDate = new Date(); mThresholdDate.setDate(now.getDate() - 180);

        // ステータスの優先度マップ作成 (シート名 -> ステータス名 -> 表示順)
        const statusPriorityMap = {};
        if (currentMasters['M_ステータス']) {
            currentMasters['M_ステータス'].forEach(s => {
                const func = s['対象機能'];
                const name = s['ステータス名称'];
                const order = parseInt(s['表示順']) || 999;
                if (!statusPriorityMap[func]) statusPriorityMap[func] = {};
                statusPriorityMap[func][name] = order;
            });
        }

        const tables = {
            'T_仕入': ['返品完了', '入庫済み', '入庫済', 'キャンセル'],
            'T_経費': ['完了', '返品完了', 'キャンセル'],
            'T_製造': ['完了', 'キャンセル'],
            'T_販売': ['完了', 'キャンセル']
        };

        for (const sheetName in rawData) {
            if (sheetName === 'T_帳簿') continue;

            const records = convertRawToObjects(rawData[sheetName]);
            const excludeList = tables[sheetName] || [];

            // 最新履歴は通常一覧の期間や統計用の日付の有無に依存させない。
            if (tables[sheetName]) records.forEach(row => {
                const action = createRecentAction(sheetName, row, now);
                if (action) recentAll.push(action);
            });

            // 履歴の期間フィルタに依存しない全体集計（個人用売上の全期間集計用）
            if (sheetName === 'T_販売') {
                records.forEach(r => {
                    const status = (r['ステータス'] || "").toString().trim();
                    if (status === '完了' && (r['管理区分'] == 1 || r['管理対象外'] == 1)) {
                        const price = parseFloat(r['価格'] || r['合計金額'] || r['販売価格'] || 0);
                        const cDateStr = r['取引完了日'] || r['販売開始日'] || "";
                        const cDate = new Date(cDateStr);
                        if (!isNaN(cDate.getTime())) {
                            const ym = `${cDate.getFullYear()}-${String(cDate.getMonth() + 1).padStart(2, '0')}`;
                            personalSalesByMonth[ym] = (personalSalesByMonth[ym] || 0) + price;
                        }
                    }
                });
            }

            // フィルタリング（期間とステータス）
            const filtered = records.filter(r => {
                const status = (r['ステータス'] || "").toString().trim();

                // 終了済みステータスの判定
                let isFinished = excludeList.includes(status);

                // 未完了（進行中）の案件は、期間に関わらずすべて表示する
                if (!isFinished) return true;

                // 完了済みのものは、完了日等の日付ベースで期間チェック
                const dateVal = r['入庫日'] || r['完了日'] || r['取引完了日'] || r['仕入日'] || r['製造完了日'] || r['販売開始日'] || r['製造開始日'] || "";
                const d = new Date(dateVal);
                if (isNaN(d.getTime())) return true;

                if (sheetName === 'T_製造') {
                    return d >= mThresholdDate; // 製造の完了分は180日
                }
                return d >= thresholdDate; // その他は40日
            });

            // 統計計算（今月の集計用）
            filtered.forEach(r => {
                const status = (r['ステータス'] || "").toString().trim();
                // 集計用には、完了日よりも「発生日（注文日や登録日）」を優先する（今月の活動として計上するため）
                const dateVal = r['注文日'] || r['登録日'] || r['日付'] || r['仕入日'] || r['開始日'] || r['製造開始日'] || r['販売開始日'] || r['完了日'] || "";
                const d = new Date(dateVal);
                if (isNaN(d.getTime())) return;

                const isThisMonth = d.getFullYear() === thisYear && d.getMonth() === thisMonth;
                const price = parseFloat(r['価格'] || r['合計金額'] || r['販売価格'] || 0);

                if (sheetName === 'T_販売' && status === '完了') {
                    // 販売は「取引完了日」ベースで今月かどうかを判定（開始日が先月でも完了が今月なら今月の売上）
                    const compDate = new Date(r['取引完了日'] || dateVal);
                    if (!isNaN(compDate.getTime()) && compDate.getFullYear() === thisYear && compDate.getMonth() === thisMonth) {
                        const isPersonal = (r['管理区分'] == 1 || r['管理対象外'] == 1);
                        if (isPersonal) summary.personalSales += price;
                        else summary.businessSales += price;
                        summary.sales += price;
                        summary.completedSalesCount++;
                    }
                } else if (isThisMonth) {
                    // 仕入と経費は発生日（dateVal）ベース
                    if (sheetName === 'T_仕入' && (status === '入庫済み' || status === '入庫済')) {
                        summary.purchase += price;
                    } else if (sheetName === 'T_経費' && status === '完了') {
                        summary.expense += price;
                    }
                }

                // 売上見込みの計算（進行中の販売、または今月完了した販売）
                if (sheetName === 'T_販売' && status !== 'キャンセル') {
                    if (status !== '完了') {
                        summary.expectedSales += price; // 進行中
                    } else {
                        // 完了の場合は今月の売上のみを見込みに含める
                        const compDate = new Date(r['取引完了日'] || "");
                        if (!isNaN(compDate.getTime()) && compDate.getFullYear() === thisYear && compDate.getMonth() === thisMonth) {
                            summary.expectedSales += price;
                        }
                    }
                }

            });

            const funcMap = { 'T_仕入': '仕入', 'T_経費': '経費', 'T_製造': '製造', 'T_販売': '販売' };
            const priorities = statusPriorityMap[funcMap[sheetName]] || {};

            history[sheetName] = filtered
                .filter(r => !excludeList.includes((r['ステータス'] || "").toString().trim()))
                .sort((a, b) => {
                    // 1. ステータス優先度 (昇順: マスタの表示順)
                    const pA = priorities[(a['ステータス'] || "").toString().trim()] || 999;
                    const pB = priorities[(b['ステータス'] || "").toString().trim()] || 999;
                    if (pA !== pB) return pA - pB;

                    // 2. 日付 (昇順: 古い順)
                    const dA = new Date(a['入庫日'] || a['完了日'] || a['取引完了日'] || a['仕入日'] || a['製造完了日'] || a['販売開始日'] || a['製造開始日'] || 0);
                    const dB = new Date(b['入庫日'] || b['完了日'] || b['取引完了日'] || b['仕入日'] || b['製造完了日'] || b['販売開始日'] || b['製造開始日'] || 0);
                    return dA - dB;
                });
        }

        const recentActions = recentAll
            .sort((a, b) => b.date - a.date)
            .slice(0, 5);

        // T_帳簿の個人レコード（事業/個人 == 1）をpersonalSalesByMonthにマージ集計
        if (rawData['T_帳簿']) {
            const ledgerList = convertRawToObjects(rawData['T_帳簿']);
            const ledgerPersonalByMonth = {};
            ledgerList.forEach(r => {
                if (String(r['事業/個人']).trim() === '1') {
                    const price = parseFloat(r['売上'] || 0);
                    const d = new Date(r['日付']);
                    if (!isNaN(d.getTime())) {
                        const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                        ledgerPersonalByMonth[ym] = (ledgerPersonalByMonth[ym] || 0) + price;
                    }
                }
            });
            // T_帳簿に個人データがある月はT_帳簿の値を優先採用
            for (const ym in ledgerPersonalByMonth) {
                personalSalesByMonth[ym] = ledgerPersonalByMonth[ym];
            }
        }

        console.timeEnd('Client:processClientData');
        return {
            history: history,
            summary: summary,
            ledger: convertRawToObjects(rawData['T_帳簿']),
            recentActions: recentActions,
            personalSalesByMonth: personalSalesByMonth
        };
    }

    function formatDate(date) {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
    }

    /**
     * マスタデータのキャッシュ処理
     */
    function loadMastersFromCache() {
        try {
            const cached = localStorage.getItem('inventory_masters_cache');
            if (cached) {
                currentMasters = JSON.parse(cached);
                console.log("Masters loaded from cache.");
                buildDynamicUI(currentMasters);
            }
        } catch (e) {
            console.warn("Failed to load cache:", e);
        }
    }

    function saveMastersToCache(masters) {
        try {
            localStorage.setItem('inventory_masters_cache', JSON.stringify(masters));
        } catch (e) {
            console.warn("Failed to save cache:", e);
        }
    }

    /**
     * 履歴データのキャッシュ処理 (Proposal 43)
     */
    function loadHistoryFromCache() {
        try {
            const cached = localStorage.getItem('inventory_history_cache');
            if (cached) {
                lastRawData = JSON.parse(cached);
                console.log("History loaded from cache.");
                lastHistoryData = processClientData(lastRawData, 'all');
                renderAllHistory(lastHistoryData);
                
                const activeTab = document.querySelector('.nav-item.active');
                if (activeTab && activeTab.getAttribute('data-target') === 'ledger') {
                    if (typeof renderLedger === 'function') renderLedger();
                }
                return true;
            }
        } catch (e) {
            console.warn("Failed to load history cache:", e);
        }
        return false;
    }

    function saveHistoryToCache(rawData) {
        try {
            localStorage.setItem('inventory_history_cache', JSON.stringify(rawData));
        } catch (e) {
            console.warn("Failed to save history cache:", e);
        }
    }

    function renderAllHistory(data) {
        console.time('Client:renderAllHistory');
        const history = data.history || {};
        const summary = data.summary || {};

        // サマリー金額の更新
        const pTotal = document.getElementById('purchase-monthly-total');
        if (pTotal) pTotal.textContent = `¥${Math.round(summary.purchase || 0).toLocaleString()}`;
        const eTotal = document.getElementById('expense-monthly-total');
        if (eTotal) eTotal.textContent = `¥${Math.round(summary.expense || 0).toLocaleString()}`;
        const sTotal = document.getElementById('sales-monthly-total');
        if (sTotal) {
            const businessSales = summary.businessSales || 0;
            const totalSales = summary.sales || 0;
            const personalSales = summary.personalSales || 0;
            if (personalSales > 0) {
                sTotal.innerHTML = `¥${Math.round(businessSales).toLocaleString()}<span style="font-size:0.75em;color:var(--text-muted);margin-left:4px;">(¥${Math.round(totalSales).toLocaleString()})</span>`;
            } else {
                sTotal.textContent = `¥${Math.round(businessSales).toLocaleString()}`;
            }
        }

        const sExpected = document.getElementById('sales-expected-total');
        if (sExpected) sExpected.textContent = `¥${Math.round(summary.expectedSales || 0).toLocaleString()}`;

        const sCompleted = document.getElementById('sales-completed-count');
        if (sCompleted) sCompleted.textContent = (summary.completedSalesCount || 0);

        renderHistoryCards('purchase', history['T_仕入']);
        renderHistoryCards('expense', history['T_経費']);
        renderHistoryCards('manufacturing', history['T_製造']);
        renderHistoryCards('sales', history['T_販売']);

        // ナビゲーションバッジの更新
        updateNavBadges();

        // 帳簿(ダッシュボード)がアクティブなら再描画し、在庫アラートの状態(製造中/仕入中)を最新にする
        const activeTab = document.querySelector('.nav-item.active');
        if (activeTab && activeTab.getAttribute('data-target') === 'ledger') {
            renderLedger();
        }
        // 履歴の更新後、通常の現在庫カードにも最新の進行中数量を反映する。
        // 棚卸入力中のカードは再生成せず、入力内容と操作位置を維持する。
        if (activeTab?.getAttribute('data-target') === 'inventory' && !isStocktakeMode) applyStockFilters();
    }

    function setupHistoryFilters() {
        for (const [tab, sheet] of Object.entries(historySheets)) {
            const panel = document.getElementById(tab + "-history-filter");
            const clear = document.querySelector(".history-filter-clear[data-history-tab=" + tab + "]");
            if (!panel || !clear) continue;
            const fields = [...panel.querySelectorAll("input, select")];
            const redraw = () => renderHistoryCards(tab, lastHistoryData?.history?.[sheet]);
            panel.querySelector(".history-filter-apply").onclick = () => {
                const filter = Object.fromEntries(fields.map(field => [field.name, field.value.trim()]));
                if (filter.from && filter.to && filter.from > filter.to) {
                    alert("終了日は開始日以降を指定してください。");
                    return;
                }
                historyFilters.set(tab, filter);
                panel.open = false;
                redraw();
            };
            clear.onclick = () => {
                historyFilters.delete(tab);
                fields.forEach(field => { field.value = ""; });
                panel.open = false;
                redraw();
            };
        }
    }

    function bulkStatusSpec_(scope) {
      const specs = {
        purchase: { sheet: 'T_仕入', id: '仕入ID', prefix: 'P', feature: '仕入',
          ranks: { '購入予定': 0, '注文済み': 1, '仕入中': 1, '入庫済み': 2, '入庫済': 2 },
          targets: ['注文済み', '入庫済み', '入庫済'], dates: { '注文済み': '仕入日', '入庫済み': '入庫日', '入庫済': '入庫日' } },
        manufacturing: { sheet: 'T_製造', id: '製造ID', prefix: 'M', feature: '製造',
          ranks: { '製造開始': 0, '製造中': 1, 'テスト中': 2, '梱包中': 3, '完了': 4 },
          targets: ['製造中', 'テスト中', '梱包中', '完了'], dates: { '製造中': '製造着手日', 'テスト中': 'テスト開始日', '梱包中': '梱包開始日', '完了': '製造完了日' } },
        sales: { sheet: 'T_販売', id: '販売ID', prefix: 'S', feature: '販売',
          ranks: { '入金待ち': 0, '取引開始': 1, '受注': 1, '開始': 1, '発送済み': 2, '受取済み': 3, '完了': 4 },
          targets: ['取引開始', '受注', '発送済み', '受取済み', '完了'], dates: { '発送済み': '発送日', '受取済み': '受取日', '完了': '取引完了日' } },
        expense: { sheet: 'T_経費', id: '経費ID', prefix: 'E', feature: '経費',
          ranks: { '購入予定': 0, '注文済み': 1, '完了': 2 },
          targets: ['注文済み', '完了'], dates: { '完了': '完了日' } }
      };
      return specs[scope] || null;
    }

    function prepareBulkStatusChange_(scope, row, target, date) {
      const spec = bulkStatusSpec_(scope);
      if (!spec || !spec.targets.includes(target)) throw new Error('対象外の変更先です。');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '') || !Number.isFinite(Date.parse(date)) || new Date(date).toISOString().slice(0, 10) !== date) throw new Error('日付を正しく指定してください。');
      const old = String(row['ステータス'] || '').trim();
      if (!Number.isFinite(spec.ranks[old]) || ['完了', '入庫済み', '入庫済'].includes(old)) throw new Error('完了済み・取消済み等の履歴は対象外です。');
      if (spec.ranks[target] <= spec.ranks[old]) throw new Error('同じ状態・前の状態への変更は対象外です。');
      if (scope === 'purchase' && old === '購入予定' && target !== '注文済み') throw new Error('購入予定は先に実購入日を指定して注文済みへ変更してください。');
      if (scope === 'sales' && target === '完了' && old !== '受取済み') throw new Error('販売の完了は受取済みの履歴が対象です。');
      if (scope === 'sales' && old === '入金待ち' && !['取引開始', '受注'].includes(target)) throw new Error('入金待ちは先に取引開始へ変更してください。');
      const required = fields => fields.forEach(field => {
        if (row[field] === undefined || row[field] === null || String(row[field]).trim() === '') throw new Error(field + 'が未入力です。個別に保存してください。');
      });
      required(['品名', '数量']);
      if (!Number.isFinite(Number(row['数量'])) || Number(row['数量']) < 1) throw new Error('数量を確認してください。');
      const money = (value, label, positive = false) => {
        if (value === '' || value === null || value === undefined || !Number.isFinite(Number(value)) || (positive ? Number(value) <= 0 : Number(value) < 0)) throw new Error(label + 'を確認してください。');
      };
      if (scope === 'purchase') {
        required(['仕入先', '支払方法', '区分']);
        money(row['価格'], '価格', true);
        if (!(old === '購入予定' && target === '注文済み')) required(['仕入日']);
      } else if (scope === 'expense') {
        required(['注文日', '購入先', '支払方法', '仕訳']);
        money(row['合計金額'], '合計金額', true);
      } else if (scope === 'manufacturing') {
        required(['製造開始日']);
        money(row['単価'], '単価');
      } else {
        required(['売先', '発送方法']);
        money(row['価格'] ?? row['販売価格'], '販売価格');
        money(row['送料'], '送料');
      }
      const dates = {};
      const field = spec.dates[target];
      const replacePurchaseDate = scope === 'purchase' && old === '購入予定' && target === '注文済み';
      if (field && (replacePurchaseDate || !row[field])) dates[field] = date;
      return { status: target, dates };
    }


    function getHistoryBulkState(tab) {
        if (!historyBulkState.has(tab)) historyBulkState.set(tab, {
            mode: false, selected: new Set(), busy: false, target: '',
            date: new Date().toLocaleDateString('sv-SE'), pending: null, results: [], preview: ''
        });
        return historyBulkState.get(tab);
    }

    function historyBulkStorageKey(tab) {
        return 'inventory_status_batch_v1_' + GAS_URL.split('/').slice(-2, -1)[0] + '_' + tab;
    }

    function getHistoryBulkReason(tab, row) {
        const spec = bulkStatusSpec_(tab);
        const state = getHistoryBulkState(tab);
        const id = String(row[spec.id] || '');
        if (!new RegExp('^' + spec.prefix + '[0-9]+$').test(id)) return 'IDを確認してください。';
        if (state.results.some(result => result.id === id && ['review', 'blocked'].includes(result.state))) return '以前の更新結果を照合してください。';
        const status = String(row['ステータス'] || '').trim();
        if (!Number.isFinite(spec.ranks[status]) || ['完了', '入庫済み', '入庫済'].includes(status)) return '完了済み・取消済み等の履歴は対象外です。';
        if (state.target) {
            try { prepareBulkStatusChange_(tab, row, state.target, state.date); }
            catch (error) { return error.message; }
        }
        return '';
    }

    function updateHistoryBulkControls(tab) {
        const root = document.getElementById(tab + '-bulk-controls');
        if (!root) return;
        const state = getHistoryBulkState(tab);
        const spec = bulkStatusSpec_(tab);
        const query = name => root.querySelector('[data-bulk="' + name + '"]');
        const savedTarget = state.target;
        const active = (currentMasters['M_ステータス'] || []).filter(row => row['使用FLG'] == 1 && row['対象機能'] === spec.feature).map(row => row['ステータス名称']);
        const targets = spec.targets.filter(value => active.includes(value));
        query('target').replaceChildren(new Option('変更先を選択', ''), ...targets.map(value => new Option(value, value)));
        if (savedTarget && !targets.includes(savedTarget)) query('target').appendChild(new Option(savedTarget, savedTarget));
        query('target').value = savedTarget;
        query('date').value = state.date;
        const field = spec.dates[state.target];
        query('date-label').textContent = tab === 'purchase' && state.target === '注文済み' ? '実購入日（仕入日を置換）' : field ? field + '（空欄のみ設定）' : state.target ? '日付の変更なし' : '変更日';
        query('enter').hidden = state.mode;
        query('body').hidden = !state.mode;
        query('count').textContent = '選択：' + state.selected.size + '件 ／ 最大20件';
        query('check').textContent = state.selected.size + '件の変更内容を確認する';
        for (const button of root.querySelectorAll('button')) button.disabled = state.busy || Boolean(state.pending);
        query('target').disabled = query('date').disabled = state.busy || Boolean(state.pending);
        query('date').disabled ||= Boolean(state.target && !field);
        query('check').disabled ||= !state.selected.size || !state.target;
        query('retry').hidden = !state.pending;
        query('retry').disabled = state.busy;
    }

    function showHistoryBulkResults(tab, message = '') {
        const root = document.getElementById(tab + '-bulk-controls');
        const state = getHistoryBulkState(tab);
        const resultBox = root.querySelector('[data-bulk="result"]');
        resultBox.textContent = message;
        const labels = { updated: '更新済み', skipped: '対象外', blocked: '要確認', review: '要確認', waiting: '未実行' };
        for (const result of state.results) {
            const item = document.createElement('div');
            item.textContent = result.id + '：' + (labels[result.state] || '要確認') + (result.reason ? ' — ' + result.reason : '');
            resultBox.appendChild(item);
        }
    }

    function setupHistoryBulkStatus() {
        for (const tab of ['purchase', 'manufacturing', 'sales', 'expense']) {
            const root = document.getElementById(tab + '-bulk-controls');
            if (!root) continue;
            const state = getHistoryBulkState(tab);
            try {
                const saved = JSON.parse(localStorage.getItem(historyBulkStorageKey(tab)) || 'null');
                if (saved?.scope === tab && Array.isArray(saved.items) && saved.items.length <= 20) {
                    state.pending = saved; state.mode = true; state.target = saved.target; state.date = saved.date;
                }
            } catch (_) {}
            root.innerHTML = `
                <button type="button" data-bulk="enter" class="history-bulk-button">一括選択</button>
                <div data-bulk="body" hidden>
                    <div class="history-bulk-actions"><button type="button" data-bulk="all" class="history-bulk-button">表示中から選択</button><button type="button" data-bulk="exit" class="history-bulk-button">選択を終了</button></div>
                    <p data-bulk="count" class="history-bulk-count"></p>
                    <div class="history-inputs-grid">
                        <div class="input-group mini"><label for="${tab}-bulk-target">変更先</label><select id="${tab}-bulk-target" data-bulk="target"></select></div>
                        <div class="input-group mini"><label for="${tab}-bulk-date" data-bulk="date-label">変更日</label><input id="${tab}-bulk-date" data-bulk="date" type="date"></div>
                    </div>
                    <p class="history-bulk-hint">同じ日付で変更する履歴を選択してください。既存の日付は維持します。購入予定から注文済みへ変更する仕入日は置き換えます。</p>
                    <button type="button" data-bulk="check" class="history-bulk-button history-bulk-primary">変更内容を確認する</button>
                    <div data-bulk="preview" class="history-bulk-preview" hidden></div>
                    <button type="button" data-bulk="retry" class="history-bulk-button" hidden>送信済みの更新結果を確認する</button>
                </div>
                <div data-bulk="result" class="history-bulk-result" role="status"></div>`;
            const query = name => root.querySelector('[data-bulk="' + name + '"]');
            const redraw = () => renderHistoryCards(tab, lastHistoryData?.history?.[historySheets[tab]]);
            const invalidate = () => { state.preview = ''; query('preview').hidden = true; };
            query('enter').onclick = () => { if (state.busy || state.pending) return; state.mode = true; state.selected.clear(); invalidate(); redraw(); };
            query('exit').onclick = () => { if (state.busy || state.pending) return; state.mode = false; state.selected.clear(); invalidate(); redraw(); };
            query('all').onclick = () => {
                if (state.busy || state.pending) return;
                for (const row of filterHistoryRows(tab, lastHistoryData?.history?.[historySheets[tab]] || [], historyFilters.get(tab) || {})) {
                    if (state.selected.size >= 20) break;
                    if (!getHistoryBulkReason(tab, row)) state.selected.add(String(row[bulkStatusSpec_(tab).id]));
                }
                invalidate(); redraw();
            };
            query('target').onchange = () => { state.target = query('target').value; invalidate(); redraw(); };
            query('date').onchange = () => { state.date = query('date').value; invalidate(); redraw(); };
            const makeBatch = () => {
                const spec = bulkStatusSpec_(tab);
                const rows = filterHistoryRows(tab, lastHistoryData?.history?.[spec.sheet] || [], historyFilters.get(tab) || {});
                return { scope: tab, target: state.target, date: state.date, items: rows.filter(row => state.selected.has(String(row[spec.id])))
                    .map(row => ({ id: String(row[spec.id]), status: String(row['ステータス'] || '').trim() })) };
            };
            query('check').onclick = () => {
                if (state.busy || state.pending || !state.target || !state.selected.size) return;
                const batch = makeBatch();
                const rows = lastHistoryData?.history?.[historySheets[tab]] || [];
                const spec = bulkStatusSpec_(tab);
                try {
                    const changes = batch.items.map(item => {
                        const row = rows.find(row => String(row[spec.id]) === item.id);
                        const updates = prepareBulkStatusChange_(tab, row, batch.target, batch.date);
                        return { row, item, updates };
                    });
                    state.preview = JSON.stringify(batch);
                    query('preview').replaceChildren();
                    for (const change of changes) {
                        const line = document.createElement('p');
                        const dates = Object.entries(change.updates.dates).map(([field, value]) => field + '：' + value).join(' ／ ');
                        line.textContent = change.item.id + ' ' + change.row['品名'] + '：' + change.item.status + ' → ' + batch.target + (dates ? ' ／ ' + dates : ' ／ 既存の日付を維持');
                        query('preview').appendChild(line);
                    }
                    const execute = document.createElement('button');
                    execute.type = 'button'; execute.className = 'history-bulk-button history-bulk-primary';
                    execute.textContent = 'この内容で' + batch.items.length + '件を更新する';
                    execute.onclick = () => {
                        if (state.busy || state.pending || state.preview !== JSON.stringify(makeBatch())) return;
                        batch.requestId = Array.from(crypto.getRandomValues(new Uint8Array(16)), value => value.toString(16).padStart(2, '0')).join('');
                        try { localStorage.setItem(historyBulkStorageKey(tab), JSON.stringify(batch)); }
                        catch (_) { showHistoryBulkResults(tab, '端末に更新内容を保存できないため、送信を中止しました。'); return; }
                        state.pending = batch;
                        send();
                    };
                    query('preview').appendChild(execute);
                    query('preview').hidden = false;
                } catch (error) { showHistoryBulkResults(tab, error.message); }
            };
            async function send() {
                if (state.busy || !state.pending) return;
                state.busy = true; invalidate(); updateHistoryBulkControls(tab); redraw();
                try {
                    const response = await fetchAPI('updateStatusesBulk', { batch: state.pending });
                    if (response.status === 'rejected' && response.safeToEdit) {
                        localStorage.removeItem(historyBulkStorageKey(tab));
                        state.pending = null; state.target = ''; state.selected.clear();
                        showHistoryBulkResults(tab, response.message);
                        return;
                    }
                    if (response.status !== 'success') throw new Error(response.message || '更新結果が不明です。');
                    state.results = response.results;
                    // 保存済み送信内容を削除できるまでは、新しいバッチを作らない。
                    localStorage.removeItem(historyBulkStorageKey(tab));
                    state.pending = null; state.selected.clear();
                    showHistoryBulkResults(tab, '更新結果');
                } catch (error) {
                    showHistoryBulkResults(tab, error.message + ' 同じ送信内容で結果を確認してください。在庫・帳簿の確認が必要な履歴は再実行しません。');
                } finally {
                    try { await initSystem('all'); } catch (_) { showToast('履歴の更新に失敗しました。同期してください。', 'error'); }
                    state.busy = false; redraw(); updateHistoryBulkControls(tab);
                }
            }
            query('retry').onclick = send;
            updateHistoryBulkControls(tab);
            if (state.pending) showHistoryBulkResults(tab, '送信済みの更新があります。同じ内容で結果を確認してください。');
        }
    }

    function filterHistoryRows(tab, data, filter) {
        const text = value => String(value ?? "").trim();
        const normalize = value => text(value).normalize("NFKC").toLowerCase();
        const keyword = normalize(filter.keyword);
        return data.filter(row => {
            if (filter.status && text(row["ステータス"]) !== filter.status) return false;
            const name = row["品名"] || row["完成品名"];
            if (filter.item && text(name) !== filter.item) return false;
            if (filter.buyer && text(row["売先"]) !== filter.buyer) return false;
            const vendor = tab === "purchase" ? row["仕入先"] : row["購入先"];
            if (keyword && ![name, vendor].some(value => normalize(value).includes(keyword))) return false;
            if (!filter.from && !filter.to) return true;
            const dates = { purchase: row["仕入日"] || row["発注日"], expense: row["登録日"] || row["注文日"],
                manufacturing: row["製造開始日"], sales: row["販売開始日"] || row["取引開始日"] };
            const date = formatISODate(dates[tab]);
            return Boolean(date) && (!filter.from || date >= filter.from) && (!filter.to || date <= filter.to);
        });
    }

    function updateHistoryFilterOptions(tab, data, filter) {
        const panel = document.getElementById(tab + "-history-filter");
        if (!panel) return;
        const columns = { status: row => row["ステータス"], item: row => row["品名"] || row["完成品名"], buyer: row => row["売先"] };
        for (const select of panel.querySelectorAll("select")) {
            const selected = panel.open ? select.value : filter[select.name] || "";
            const values = new Set(data.map(row => String(columns[select.name](row) ?? "").trim()).filter(Boolean));
            if (selected) values.add(selected);
            select.replaceChildren(new Option("すべて", ""), ...[...values].sort().map(value => new Option(value, value)));
            select.value = selected;
        }
    }

    function updateNavBadges() {
        if (!lastHistoryData) return;

        const history = lastHistoryData.history || {};

        // 各タブの未完了カウント定義
        const counts = {
            purchase: (history['T_仕入'] || []).filter(r =>
                !['入庫済み', '返金済み', '返品完了', 'キャンセル'].includes(r['ステータス'])
            ).length,
            expense: (history['T_経費'] || []).filter(r =>
                !['完了', '返金済み', '返品完了', 'キャンセル'].includes(r['ステータス'])
            ).length,
            manufacturing: (history['T_製造'] || []).filter(r =>
                !['完了', 'キャンセル'].includes(r['ステータス'])
            ).length,
            sales: (history['T_販売'] || []).filter(r =>
                !['完了', 'キャンセル'].includes(r['ステータス'])
            ).length
        };

        // 帳簿（在庫アラート）のカウント
        if (currentMasters && currentMasters['T_在庫集計']) {
            counts.ledger = currentMasters['T_在庫集計'].filter(p => {
                const stock = parseFloat(p['現在庫数']) || 0;
                const threshold = parseFloat(p['閾値']) || 0;
                const useFlag = parseInt(p['使用FLG']) !== 0; // 0以外（1や空）は有効
                // 閾値が0のものはアラート対象外とする（表示側のロジックに合わせる）
                return threshold > 0 && useFlag && stock <= threshold;
            }).length;
        }

        // DOMに反映
        const navItems = document.querySelectorAll('.nav-item');
        navItems.forEach(item => {
            const target = item.getAttribute('data-target');
            const badge = item.querySelector('.nav-badge');
            if (!badge) return;

            const count = counts[target] || 0;
            if (count > 0) {
                badge.textContent = count > 99 ? '99+' : count;
                badge.classList.add('visible');
            } else {
                badge.classList.remove('visible');
                // アニメーション完了後にテキストをクリア
                setTimeout(() => { if (!badge.classList.contains('visible')) badge.textContent = ''; }, 200);
            }
        });
        console.timeEnd('Client:renderAllHistory');
    }

    /**
     * API通信の共通処理 (GAS本番環境とローカル環境を自動判別)
     */
    async function fetchAPI(action, payload = {}) {
        const bodyData = Object.assign({ action, key: currentAuthKey, sessionToken }, payload);

        if (!bodyData.key && !bodyData.sessionToken && action !== 'getInitData') {
            console.warn(`[fetchAPI] Warning: No auth key for action: ${action}`);
        }

        // GAS本番環境 (google.script.run が存在する場合)
        if (typeof google !== 'undefined' && google.script && google.script.run) {
            return new Promise((resolve, reject) => {
                google.script.run
                    .withSuccessHandler(res => {
                        if (!res) reject(new Error("Server returned null response"));
                        else if (action !== 'checkAuth' && res.status === 'error' && res.message.includes('Unauthorized')) {
                            handleUnauthorized();
                            reject(new Error(res.message));
                        }
                        else resolve(res);
                    })
                    .withFailureHandler(err => reject(new Error(err.message || "Server connection failed")))
                    .apiEntryPoint(bodyData);
            });
        }

        // ローカル環境 or フォールバック (fetch を使用)
        try {
            const response = await fetch(GAS_URL, {
                method: 'POST',
                body: JSON.stringify(bodyData)
            });

            if (!response.ok) {
                if (response.status === 401) {
                    handleUnauthorized();
                }
                throw new Error("Network response was not ok: " + response.status);
            }

            const result = await response.json();
            if (!['checkAuth', 'startSession', 'endSession'].includes(action) && result.status === 'error' && result.message.includes('Unauthorized')) {
                console.error(`[fetchAPI] Unauthorized error for action: ${action}`, result);
                handleUnauthorized();
            }
            return result;
        } catch (error) {
            console.error("Fetch Error:", error);
            throw error;
        }
    }

    /**
     * 認証エラー時の処理
     */
    function clearLoginSession() {
        currentAuthKey = '';
        sessionToken = '';
        try {
            localStorage.removeItem(rememberedLoginKey);
        } catch {
            // 保存領域を使えない場合もメモリ上の認証情報は破棄する。
        }
    }

    async function restoreLoginSession() {
        if (!supportsRememberedLogin) return false;
        try {
            const saved = JSON.parse(localStorage.getItem(rememberedLoginKey));
            if (!saved || typeof saved.token !== 'string' || !(saved.expiresAt > Date.now())) {
                clearLoginSession();
                return false;
            }
            const result = await fetchAPI('checkAuth', { sessionToken: saved.token });
            if (result.status !== 'success') {
                if (result.message === 'Unauthorized') clearLoginSession();
                return false;
            }
            sessionToken = saved.token;
            return true;
        } catch {
            return false;
        }
    }

    function handleUnauthorized() {
        clearLoginSession();
        const appContainer = document.querySelector('.app-container');
        if (appContainer) appContainer.style.display = 'none';
        showLoginModal();
    }

    /**
     * ログインモーダルの表示
     */
    function showLoginModal() {
        const modal = document.getElementById('login-modal');
        if (modal) {
            modal.style.display = 'block';
            document.getElementById('login-password').focus();
        }
    }

    /**
     * ログイン関連のイベントハンドラ設定
     */
    function setupLoginHandlers(onFirstLogin) {
        const submitBtn = document.getElementById('login-submit-btn');
        const passwordInput = document.getElementById('login-password');
        const errorMsg = document.getElementById('login-error');

        if (!submitBtn || !passwordInput) return;

        const handleLogin = async () => {
            const password = passwordInput.value;
            if (!password || submitBtn.disabled) return;

            submitBtn.disabled = true;
            submitBtn.textContent = '認証中...';
            errorMsg.style.display = 'none';

            try {
                const remember = supportsRememberedLogin && rememberLogin.checked;
                const token = remember
                    ? Array.from(crypto.getRandomValues(new Uint8Array(32)), byte => byte.toString(16).padStart(2, '0')).join('')
                    : '';
                const result = await fetchAPI(remember ? 'startSession' : 'checkAuth', { key: password, sessionToken: token });

                if (result.status === 'success') {
                    // 認証成功
                    clearLoginSession();
                    if (remember) {
                        sessionToken = token;
                        try {
                            localStorage.setItem(rememberedLoginKey, JSON.stringify({ token, expiresAt: result.expiresAt }));
                        } catch {
                            showToast('この端末ではログイン状態を保存できません。', 'error');
                        }
                    } else {
                        currentAuthKey = password;
                    }
                    passwordInput.value = '';
                    document.getElementById('login-modal').style.display = 'none';

                    if (onFirstLogin) {
                        onFirstLogin();
                        onFirstLogin = null;
                    } else {
                        const appContainer = document.querySelector('.app-container');
                        if (appContainer) appContainer.style.display = 'flex';
                        await initSystem('all');
                    }
                } else {
                    errorMsg.textContent = result.message || 'パスワードが正しくありません';
                    errorMsg.style.display = 'block';
                }
            } catch (e) {
                console.error("Login error:", e);
                errorMsg.textContent = '通信エラーが発生しました';
                errorMsg.style.display = 'block';
            } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'ログイン';
            }
        };

        if (submitBtn) submitBtn.onclick = handleLogin;
        if (passwordInput) {
            passwordInput.onkeypress = (e) => {
                if (e.key === 'Enter') handleLogin();
            };
        }
    }

    function setupNavigation() {
        const navItems = document.querySelectorAll('.nav-item');
        const tabContents = document.querySelectorAll('.tab-content');

        if (navItems.length === 0) console.warn("No nav-items found");

        navItems.forEach(item => {
            item.addEventListener('click', () => {
                // Remove active classes
                navItems.forEach(nav => nav.classList.remove('active'));
                tabContents.forEach(tab => tab.classList.remove('active'));

                // Add active classes
                item.classList.add('active');
                const targetId = item.getAttribute('data-target');
                const targetTab = document.getElementById(`tab-${targetId}`);
                if (targetTab) {
                    targetTab.classList.add('active');
                }

                // 帳簿タブがアクティブになった際のレンダリング
                if (targetId === 'ledger') {
                    renderLedger();
                }

                // 棚卸モード中に他タブへ移動したらモードを終了する
                if (targetId !== 'inventory' && isStocktakeMode) {
                    isStocktakeMode = false;
                    stocktakeData = {};
                    const stocktakeBar = document.getElementById('stocktake-bar');
                    if (stocktakeBar) stocktakeBar.style.display = 'none';
                    const stocktakeToggle = document.getElementById('stocktake-mode-toggle');
                    if (stocktakeToggle) {
                        stocktakeToggle.classList.remove('active');
                        stocktakeToggle.innerHTML = '<ion-icon name="checkbox-outline"></ion-icon> 棚卸開始';
                    }
                }

                // Scroll to top
                const contentArea = document.getElementById('content-area');
                if (contentArea) contentArea.scrollTo({ top: 0, behavior: 'smooth' });
            });
        });
    }

    function setupToggleLogics() {
        // Sales Type Toggle
        const salesTypeRadios = document.querySelectorAll('input[name="sales-type"]');
        const productNameLabel = document.getElementById('product-name-label');
        const personalHint = document.querySelector('.personal-only');

        salesTypeRadios.forEach(radio => {
            radio.addEventListener('change', (e) => {
                const isPersonal = e.target.value === 'personal';
                const modeText = document.getElementById('sale-item-mode-text');
                if (modeText) {
                    modeText.textContent = isPersonal ?
                        '(個人用: 自由入力可能)' :
                        '(事業用: リストから選択のみ)';
                }
                if (personalHint) personalHint.style.display = isPersonal ? 'block' : 'none';
                toggleSaleItemInput(isPersonal);
            });
        });

        // Side Menu Logic (提案32)
        const sideMenu = document.getElementById('side-menu');
        const menuOverlay = document.getElementById('menu-overlay');
        const menuOpenBtn = document.getElementById('menu-open-btn');
        const menuCloseBtn = document.getElementById('menu-close-btn');

        if (menuOpenBtn) {
            menuOpenBtn.addEventListener('click', () => {
                sideMenu.classList.add('open');
                menuOverlay.classList.add('visible');
            });
        }

        if (menuCloseBtn) {
            menuCloseBtn.addEventListener('click', () => {
                sideMenu.classList.remove('open');
                menuOverlay.classList.remove('visible');
            });
        }

        if (menuOverlay) {
            menuOverlay.addEventListener('click', () => {
                sideMenu.classList.remove('open');
                menuOverlay.classList.remove('visible');
            });
        }

        // メニュー内アイテムのクリック処理
        const menuSettingsBtn = document.getElementById('menu-settings-btn');
        const menuMasterBtn = document.getElementById('menu-master-btn');
        const settingsModal = document.getElementById('settings-modal');

        if (menuSettingsBtn) {
            menuSettingsBtn.addEventListener('click', () => {
                sideMenu.classList.remove('open');
                menuOverlay.classList.remove('visible');
                if (settingsModal) {
                    settingsModal.classList.add('active');
                    renderMasterList('system');
                }
            });
        }

        if (menuMasterBtn) {
            menuMasterBtn.addEventListener('click', () => {
                sideMenu.classList.remove('open');
                menuOverlay.classList.remove('visible');
                if (settingsModal) {
                    settingsModal.classList.add('active');
                    renderMasterList('master');
                }
            });
        }

        const menuRebuildStockBtn = document.getElementById('menu-rebuild-stock-btn');
        if (menuRebuildStockBtn) {
            menuRebuildStockBtn.addEventListener('click', async () => {
                sideMenu.classList.remove('open');
                menuOverlay.classList.remove('visible');

                if (!confirm('在庫の全再集計を実行しますか？\n(スプレッドシートの手修正内容が在庫合計に反映されます)')) return;

                setLoading(true, '在庫データを再計算中...');
                try {
                    const res = await fetchAPI('rebuildInventorySummary');
                    if (res.status === 'success') {
                        showToast(res.message, 'success');
                        // 再集計後に最新データを取得
                        await initSystem();
                    } else {
                        throw new Error(res.message);
                    }
                } catch (e) {
                    console.error("Rebuild error:", e);
                    showToast('再集計に失敗しました: ' + e.message, 'error');
                } finally {
                    setLoading(false);
                }
            });
        }

        // Sync Button
        const syncBtn = document.getElementById('sync-btn');
        if (syncBtn) {
            syncBtn.addEventListener('click', () => {
                const icon = syncBtn.querySelector('ion-icon');
                if (icon) icon.classList.add('spinning');
                setLoading(true, 'システムデータを同期中...');
                initSystem().then(() => {
                    if (icon) icon.classList.remove('spinning');
                    setLoading(false);
                }).catch(e => {
                    if (icon) icon.classList.remove('spinning');
                    setLoading(false);
                    showToast('同期に失敗しました', 'error');
                });
            });
        }

        // Ledger Switcher
        const switcherItems = document.querySelectorAll('.switcher-item');
        const dashboardView = document.getElementById('ledger-dashboard-view');
        const detailsView = document.getElementById('ledger-details-view');
        const inventoryAlerts = document.getElementById('inventory-alerts-container');

        switcherItems.forEach(item => {
            item.addEventListener('click', () => {
                switcherItems.forEach(si => si.classList.remove('active'));
                item.classList.add('active');

                const view = item.getAttribute('data-view');
                if (dashboardView && detailsView) {
                    dashboardView.style.display = (view === 'dashboard') ? 'block' : 'none';
                    detailsView.style.display = (view === 'dashboard') ? 'none' : 'block';
                }
                // 在庫アラートの表示切り替え（ダッシュボード時のみ）
                if (inventoryAlerts) {
                    inventoryAlerts.style.display = (view === 'dashboard') ? 'block' : 'none';
                }
                renderLedger(); // 切り替え時に再描画
            });
        });

        // Period Filters
        const dashboardsPeriod = document.getElementById('dashboard-period-select');
        const detailsPeriod = document.getElementById('details-period-select');
        [dashboardsPeriod, detailsPeriod].forEach(p => {
            if (p) p.addEventListener('change', () => renderLedger());
        });

        // Date & Quantity Defaults
        const todayStr = new Date().toISOString().split('T')[0];
        const dateInputs = ['purchase-date', 'expense-date', 'manufacturing-date', 'sales-date'];
        dateInputs.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = todayStr;
        });

        const qtyInputs = ['buy-quantity', 'exp-quantity', 'make-quantity', 'sale-quantity'];
        qtyInputs.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = id === 'sale-quantity' ? '1' : '';
        });

        // Inventory Search & Filter Logic
        const stockSearchInput = document.getElementById('stock-search-input');
        const stockSearchClear = document.getElementById('stock-search-clear');


        if (stockSearchInput && stockSearchClear) {
            stockSearchInput.addEventListener('input', (e) => {
                const term = e.target.value.trim();
                stockSearchClear.classList.toggle('visible', term.length > 0);
                applyStockFilters();
            });

            stockSearchClear.addEventListener('click', () => {
                stockSearchInput.value = '';
                stockSearchClear.classList.remove('visible');
                applyStockFilters();
            });
        }

        // All filter chip logic is in setupStockFilters

        // Transaction Buttons
        setupTransactionSubmitters();
        setupAutocompleteListeners();

        // Stocktake Mode Toggle
        const stocktakeToggle = document.getElementById('stocktake-mode-toggle');
        const stocktakeBar = document.getElementById('stocktake-bar');

        if (stocktakeToggle) {
            stocktakeToggle.addEventListener('click', () => {
                isStocktakeMode = !isStocktakeMode;
                stocktakeToggle.classList.toggle('active', isStocktakeMode);
                stocktakeToggle.innerHTML = isStocktakeMode ?
                    '<ion-icon name="close-outline"></ion-icon> 棚卸を終了' :
                    '<ion-icon name="checkbox-outline"></ion-icon> 棚卸開始';

                if (stocktakeBar) stocktakeBar.style.display = isStocktakeMode ? 'flex' : 'none';

                // 棚卸モード開始時にセッション状態をリセット
                stocktakeData = {};
                stocktakeSession.verifiedItems.clear();

                // 再描画
                applyStockFilters();
                updateStocktakeSummary();
            });
        }

        // Add Material Modal Listeners (Proposal 15)
        const addMatClose = document.getElementById('add-material-close-btn');
        const addMatCancel = document.getElementById('add-material-cancel');
        const addMatSubmit = document.getElementById('add-material-submit');
        const addMatModal = document.getElementById('add-material-modal');

        if (addMatClose) addMatClose.addEventListener('click', () => addMatModal.classList.remove('active'));
        if (addMatCancel) addMatCancel.addEventListener('click', () => addMatModal.classList.remove('active'));
        if (addMatSubmit) {
            addMatSubmit.addEventListener('click', () => handleMaterialSubmission());
        }

        // Stocktake Submit
        const stocktakeSubmit = document.getElementById('stocktake-submit');
        if (stocktakeSubmit) {
            stocktakeSubmit.addEventListener('click', async () => {
                const verifiedNames = Array.from(stocktakeSession.verifiedItems);
                if (verifiedNames.length === 0) return alert("確認済みの品目がありません。品目横のチェックアイコンを押して確認済みにしてください。");
                if (!confirm(`${verifiedNames.length}件の棚卸を確定します。よろしいですか？`)) return;

                stocktakeSubmit.disabled = true;
                const originalHTML = stocktakeSubmit.innerHTML;
                stocktakeSubmit.innerHTML = '<ion-icon name="sync-outline" class="spinning"></ion-icon> 登録中...';
                try {
                    const targets = verifiedNames.map(n => {
                        const saved = stocktakeData[n];
                        const currentQty = parseFloat((currentMasters['T_在庫集計'].find(m => m['品名'] === n) || {})['現在庫数']) || 0;
                        return {
                            itemName: n,
                            actualQty: saved ? saved.actual : currentQty,
                            logicalQty: currentQty,
                            diffQty: saved ? saved.diff : 0
                        };
                    });
                    const res = await fetchAPI('registerStocktake', { stocktakeData: targets, note: document.getElementById('stocktake-note').value });
                    if (res.status === 'success') {
                        showToast("棚卸が完了しました。");
                        isStocktakeMode = false;
                        const noteInput = document.getElementById('stocktake-note');
                        if (noteInput) noteInput.value = '';
                        if (stocktakeToggle) {
                            stocktakeToggle.classList.remove('active');
                            stocktakeToggle.innerHTML = '<ion-icon name="checkbox-outline"></ion-icon> 棚卸開始';
                        }
                        if (stocktakeBar) stocktakeBar.style.display = 'none';

                        // セッションリセット
                        stocktakeData = {};
                        stocktakeSession.verifiedItems.clear();

                        await initSystem();
                    } else throw new Error(res.message);
                } catch (e) {
                    alert("エラー: " + e.message);
                } finally {
                    stocktakeSubmit.disabled = false;
                    stocktakeSubmit.innerHTML = originalHTML;
                }
            });
        }
    }

    function buildDynamicUI(masters) {
        const controls = masters['M_画面制御'] || [];
        document.querySelectorAll('[data-quantity-select="true"]').forEach(input => {
            if (!controls.some(control => control['要素ID'] === input.id && control['タイプ'] === 'quantity-select')) {
                removeQuantitySelect_(input);
                delete input.dataset.quantitySelect;
            }
        });
        controls.forEach(ctrl => {
            const elmId = ctrl['要素ID'];
            if (ctrl['タイプ'] === 'quantity-select') {
                setupQuantitySelect_(ctrl);
                return;
            }
            let elm = document.getElementById(elmId);
            if (!elm) return;

            const type = ctrl['タイプ'];
            const masterName = ctrl['参照マスタ'];

            // サジェスト形式（suggest/suggest-strict）が指定されているのに HTML が SELECT の場合、INPUT に置換する
            if ((type === 'suggest' || type === 'suggest-strict') && elm.tagName === 'SELECT') {
                const input = document.createElement('input');
                input.type = 'text';
                input.id = elmId;
                const listId = elmId + '-list';
                input.setAttribute('list', listId);
                input.placeholder = (ctrl['画面上の項目名'] || '品名') + 'を検索・入力';

                // datalistがない場合は作成
                if (!document.getElementById(listId)) {
                    const dl = document.createElement('datalist');
                    dl.id = listId;
                    elm.parentNode.appendChild(dl);
                }

                // プレビュー用のIDなどを保持している可能性を考慮し、元の要素のクラスなども一部引き継ぐ
                input.className = elm.className;

                elm.parentNode.replaceChild(input, elm);
                // 変数 elm を新しい要素に更新して以降の処理へ
                elm = input;
            }

            // タイプが select/fixed なのに HTML が INPUT の場合、SELECT に置換する
            if ((type === 'select' || type === 'fixed') && elm.tagName === 'INPUT') {
                const select = document.createElement('select');
                select.id = elmId;
                select.className = elm.className;
                if (elm.value) {
                    const opt = document.createElement('option');
                    opt.value = elm.value;
                    opt.textContent = elm.value;
                    opt.selected = true;
                    select.appendChild(opt);
                }
                elm.parentNode.replaceChild(select, elm);
                elm = select;
            }

            if (type === 'fixed') {
                const options = (ctrl['固定値の内容'] || "").split(',').map(s => s.trim());
                populateElement(elm, type, options);
            }
            else if (type === 'select' || type === 'suggest' || type === 'suggest-strict') {
                let options = [];
                if (masterName) {
                    if (masterName === 'M_ステータス') {
                        const statuses = (masters['M_ステータス'] || [])
                            .filter(s => (parseInt(s['使用FLG']) === 1 || s['使用FLG'] == 1) && s['対象機能'] === ctrl['対象機能'] && s['画面名称'] === ctrl['画面名称'])
                            .sort((a, b) => (a['表示順'] || 0) - (b['表示順'] || 0));
                        options = statuses.map(s => s['ステータス名称']);
                    } else {
                        let masterData = masters[masterName] || [];
                        if (ctrl['抽出条件']) {
                            masterData = parseAndApplyFilter(masterData, ctrl['抽出条件'], masters);
                        }
                        if (masterData && masterData.length > 0) {
                            const excludeFields = ['表示順', '使用FLG', 'カテゴリ', '手数料率', '送料', '用途区分', '説明', 'デフォルト仕訳', '役割（タイプ）', '対象機能', '画面名称', '商品ID', '仕入先ID', '売先ID', '発送ID', '仕訳ID', '画像URL'];
                            const priorityFields = ['品名', '完成品名', '名称', 'ステータス名称', '仕入先', '売先', '発送方法', '仕訳名'];

                            let keyField = priorityFields.find(k => masterData[0].hasOwnProperty(k));
                            if (!keyField) {
                                keyField = Object.keys(masterData[0]).find(k => !excludeFields.includes(k));
                            }

                            if (keyField) options = masterData.map(r => r[keyField]);
                        }
                    }
                } else if (ctrl['固定値の内容']) {
                    options = ctrl['固定値の内容'].split(',').map(s => s.trim());
                }

                if (options.length > 0) {
                    if (elm.tagName === 'INPUT' && (type === 'suggest' || type === 'suggest-strict')) {
                        elm.type = 'text'; // サジェストの表示安定性を優先してテキスト型に変更
                        if (!elm.hasAttribute('list')) {
                            const listId = elmId + '-list';
                            elm.setAttribute('list', listId);
                            if (!document.getElementById(listId)) {
                                const dl = document.createElement('datalist');
                                dl.id = listId;
                                elm.parentNode.appendChild(dl);
                            }
                        }
                    }
                    populateElement(elm, type, options);
                }

                if (type === 'suggest' || type === 'suggest-strict') {
                    addClearButton(elm);
                    elm.dataset.type = type;
                    elm.dataset.master = masterName || "";
                }
            }
        });

        // 在庫一覧の初期描画: フィルタ設定を考慮して描画
        applyStockFilters();

        // 動的生成・差し替え後の要素にリスナーを再設定 (提案対応)
        setupAutocompleteListeners();
        setupInventoryCheckListeners();
    }

    function quantityOptions_(control) {
        if (!control || control['タイプ'] !== 'quantity-select') return [];
        return [...new Set(String(control['固定値の内容'] || '').split(',').map(value => value.trim())
            .filter(value => /^\d+$/.test(value) && Number.isSafeInteger(Number(value)) && Number(value) >= 1)
            .map(value => String(Number(value))))];
    }

    function removeQuantitySelect_(input) {
        const wrapper = document.getElementById(input.id + '-quantity-control');
        if (wrapper) wrapper.replaceWith(input);
        input.hidden = false; input.style.marginTop = '';
        if (input.dataset.quantityPadding !== undefined) input.style.setProperty('padding-right', input.dataset.quantityPadding, input.dataset.quantityPaddingPriority || '');
        delete input.dataset.quantityMode; delete input.dataset.quantityCustomValue;
        delete input.dataset.quantityPadding;
        delete input.dataset.quantityPaddingPriority;
    }

    function setQuantityMode_(input, custom) {
        const choice = document.getElementById(input.id + '-choice');
        const back = document.getElementById(input.id + '-quantity-back');
        if (!choice || !back) return;
        choice.hidden = custom; input.hidden = !custom; back.hidden = !custom;
        input.dataset.quantityMode = custom ? 'input' : 'select';
        input.style.marginTop = '';
        input.style.setProperty('padding-right', custom ? '40px' : (input.dataset.quantityPadding || ''), custom ? 'important' : (input.dataset.quantityPaddingPriority || ''));
    }

    function syncQuantityChoice_(input, resetMode = false) {
        const choice = input && document.getElementById(input.id + '-choice');
        if (!input || !choice) return;
        const listed = Array.from(choice.options).some(option => option.value === input.value);
        choice.value = listed ? input.value : 'other';
        setQuantityMode_(input, (!resetMode && input.dataset.quantityMode === 'input') || !listed);
    }

    function setupQuantitySelect_(control) {
        let input = document.getElementById(control['要素ID']);
        if (!input || !['INPUT', 'SELECT'].includes(input.tagName)) return;
        if (input.tagName === 'SELECT') {
            const previous = input;
            input = document.createElement('input'); input.id = previous.id; input.className = previous.className;
            input.type = 'number'; input.value = previous.value; input.disabled = previous.disabled;
            previous.replaceWith(input);
        }
        input.type = 'number'; input.min = '1'; input.step = '1'; input.inputMode = 'numeric';
        input.dataset.quantitySelect = 'true';
        const label = control['画面上の項目名'] || '数量';
        input.setAttribute('aria-label', label + 'の直接入力');
        const options = quantityOptions_(control);
        let choice = document.getElementById(input.id + '-choice');
        if (!options.length) {
            removeQuantitySelect_(input);
            return;
        }
        if (!choice) {
            const wrapper = document.createElement('div');
            wrapper.id = input.id + '-quantity-control';
            wrapper.style.cssText = 'position:relative;width:100%;min-width:0';
            input.before(wrapper); wrapper.appendChild(input);
            input.dataset.quantityPadding = input.style.paddingRight;
            input.dataset.quantityPaddingPriority = input.style.getPropertyPriority('padding-right');
            choice = document.createElement('select');
            choice.id = input.id + '-choice';
            choice.style.cssText = 'width:100%;min-width:0;max-width:100%;box-sizing:border-box';
            input.before(choice);
            const back = document.createElement('button');
            back.id = input.id + '-quantity-back'; back.type = 'button';
            back.textContent = '▼'; back.title = '数量候補を選ぶ'; back.setAttribute('aria-label', '数量候補を選ぶ');
            back.setAttribute('aria-controls', choice.id);
            back.style.cssText = 'position:absolute;right:0;top:0;bottom:0;width:36px;min-height:0;margin:0;padding:0;border:0;background:transparent;color:inherit;font-size:13px;cursor:pointer';
            wrapper.appendChild(back);
            back.addEventListener('click', () => {
                input.dataset.quantityCustomValue = input.value;
                input.value = ''; choice.value = '';
                setQuantityMode_(input, false);
                input.dispatchEvent(new Event('input', { bubbles: true }));
                choice.focus();
            });
            choice.addEventListener('change', () => {
                const custom = choice.value === 'other';
                input.value = custom ? (input.dataset.quantityCustomValue || '') : choice.value;
                setQuantityMode_(input, custom);
                input.dispatchEvent(new Event('input', { bubbles: true }));
                if (custom) input.focus();
            });
        }
        choice.setAttribute('aria-label', label);
        input.placeholder = '数量を入力';
        choice.replaceChildren(new Option('数量を選択', ''), ...options.map(value => new Option(value, value)), new Option('その他（直接入力）', 'other'));
        syncQuantityChoice_(input);
    }

    /**
     * 在庫一覧のフィルタを適用して再描画する
     */
    function applyStockFilters() {
        const stockSearchInput = document.getElementById('stock-search-input');
        const showHiddenChip = document.querySelector('.filter-chip[data-filter="hidden"]');
        const thresholdZeroChip = document.querySelector('.filter-chip[data-filter="threshold-zero"]');
        const uncheckedOnlyChip = document.querySelector('.filter-chip[data-filter="unchecked"]');
        const inStockOnlyChip = document.querySelector('.filter-chip[data-filter="in-stock"]');
        const noStockOnlyChip = document.querySelector('.filter-chip[data-filter="no-stock"]');
        const stockAlertChip = document.querySelector('.filter-chip[data-filter="stock-alert"]');

        const term = (stockSearchInput ? stockSearchInput.value : "").toLowerCase().trim();
        const showHidden = showHiddenChip ? showHiddenChip.classList.contains('active') : false;
        const thresholdZeroOnly = thresholdZeroChip ? thresholdZeroChip.classList.contains('active') : false;
        const uncheckedOnly = uncheckedOnlyChip ? uncheckedOnlyChip.classList.contains('active') : false;
        const inStockOnly = inStockOnlyChip ? inStockOnlyChip.classList.contains('active') : false;
        const noStockOnly = noStockOnlyChip ? noStockOnlyChip.classList.contains('active') : false;
        const stockAlertOnly = stockAlertChip ? stockAlertChip.classList.contains('active') : false;

        // カテゴリチップの選択状態を取得
        const activeCategoryChips = document.querySelectorAll('.dynamic-chips-wrapper .filter-chip.active');
        const activeCategories = Array.from(activeCategoryChips).map(c => c.getAttribute('data-category'));

        const allStockProducts = currentMasters['T_在庫集計'] || [];

        // カテゴリチップの動的生成 (フィルタ適用前に現在のマスタから抽出)
        generateCategoryChips(allStockProducts);

        const filtered = allStockProducts.filter(p => {
            const name = String(p['品名'] || "").toLowerCase();
            const category = String(p['カテゴリ'] || p['商品区分'] || "").toLowerCase();

            // マスタからJANコードとIDを補完して検索対象にする
            const itemInMaster = (currentMasters['M_商品'] || []).find(m => m['品名'] === p['品名']);
            const id = String((itemInMaster && itemInMaster['商品ID']) || p['商品ID'] || "").toLowerCase();
            const barcode = String((itemInMaster && itemInMaster['QR/バーコード']) || p['QR/バーコード'] || "").toLowerCase();
            const location = String(p['保管場所'] || (itemInMaster && itemInMaster['保管場所']) || "").toLowerCase();

            const useFlag = parseInt(p['使用FLG']) !== 0;
            const threshold = parseFloat(p['閾値']) || 0;
            const stock = parseFloat(p['現在庫数']) || 0;
            const isUnchecked = !p['最終棚卸日'];

            const matchesSearch = name.includes(term) || category.includes(term) || id.includes(term) || barcode.includes(term) || location.includes(term);
            const matchesVisibility = showHidden || useFlag;
            const matchesThreshold = !thresholdZeroOnly || threshold === 0;
            const matchesUnchecked = !uncheckedOnly || isUnchecked;
            const matchesInStock = !inStockOnly || stock > 0;
            const matchesNoStock = !noStockOnly || stock === 0;
            const matchesStockAlert = !stockAlertOnly || (stock <= threshold && threshold > 0);

            // カテゴリフィルタ (複数選択時はOR条件)
            const matchesCategory = activeCategories.length === 0 || activeCategories.some(ac => category.includes(ac.toLowerCase()));

            return matchesSearch && matchesVisibility && matchesThreshold && matchesUnchecked && matchesInStock && matchesNoStock && matchesStockAlert && matchesCategory;
        });

        // 優先度と表示順によるソート (提案10)
        filtered.sort((a, b) => {
            const prioA = parseInt(a['優先度']) || 999;
            const prioB = parseInt(b['優先度']) || 999;

            if (prioA !== prioB) {
                return prioA - prioB; // 優先度が小さい(1に近い)ものを上に
            }

            // 優先度が同じなら表示順
            const orderA = parseInt(a['表示順']) || 999;
            const orderB = parseInt(b['表示順']) || 999;
            if (orderA !== orderB) {
                return orderA - orderB; // 表示順が小さい(1に近い)ものを上に
            }

            // 表示順も同じなら品名順
            return (a['品名'] || "").localeCompare(b['品名'] || "", 'ja');
        });

        renderStockList(filtered);
    }

    /**
     * 在庫データからカテゴリを抽出し、フィルタチップを生成する
     */
    function generateCategoryChips(allStock) {
        const container = document.getElementById('dynamic-category-chips');
        if (!container) return;

        // すでに選択されているカテゴリを記憶
        const activeCats = Array.from(container.querySelectorAll('.filter-chip.active')).map(c => c.getAttribute('data-category'));

        // ユニークなカテゴリを抽出 (空文字は除外)
        const categories = [...new Set(allStock.map(p => p['カテゴリ'] || p['商品区分'] || '').filter(c => c))].sort();

        // チップのHTML生成
        container.innerHTML = categories.map(cat => `
            <button class="filter-chip ${activeCats.includes(cat) ? 'active' : ''}" data-category="${cat}">
                ${cat}
            </button>
        `).join('');

        // イベントリスナーの再設定
        container.querySelectorAll('.filter-chip').forEach(chip => {
            chip.addEventListener('click', () => {
                chip.classList.toggle('active');
                applyStockFilters();
            });
        });
    }

    /**
     * 最新の生データから在庫マスタを同期し、UIをリフレッシュする
     */
    function refreshInventoryUI() {
        if (lastRawData['T_在庫集計']) {
            console.log("Syncing inventory masters from raw data...");
            currentMasters['T_在庫集計'] = convertRawToObjects(lastRawData['T_在庫集計']);
            applyStockFilters();
        }
    }

    function populateElement(elm, type, options) {
        if (elm.tagName === 'SELECT') {
            const currentVal = elm.value;
            elm.innerHTML = '<option value="">--- 選択してください ---</option>';
            options.forEach(opt => {
                const o = document.createElement('option');
                o.value = opt;
                o.textContent = opt;
                elm.appendChild(o);
            });
            if (currentVal) elm.value = currentVal;
        } else if (elm.tagName === 'DATALIST') {
            elm.innerHTML = '';
            options.forEach(opt => {
                const o = document.createElement('option');
                o.value = opt;
                elm.appendChild(o);
            });
        } else if (elm.hasAttribute('list')) {
            const listId = elm.getAttribute('list');
            if (listId) {
                const datalist = document.getElementById(listId);
                if (datalist) {
                    datalist.innerHTML = '';
                    options.forEach(opt => {
                        const o = document.createElement('option');
                        o.value = opt;
                        datalist.appendChild(o);
                    });
                }
            }
        }
    }

    /**
     * クイック検索モーダルを開く (Proposal 37)
     */
    window.openQuickSearch = function (inputId) {
        const input = document.getElementById(inputId);
        if (!input) return;

        let options = [];
        if (input.tagName === 'SELECT') {
            options = Array.from(input.options)
                .map(o => o.value)
                .filter(val => val !== "");
        } else {
            const listId = input.getAttribute('list');
            const datalist = document.getElementById(listId);
            if (datalist) {
                options = Array.from(datalist.options).map(o => o.value);
            }
        }

        if (options.length === 0) return;

        const title = input.previousElementSibling ? input.previousElementSibling.textContent : "項目を選択";

        document.getElementById('quick-search-title').textContent = title;
        document.getElementById('quick-search-input').value = "";
        window.currentQuickSearchInputId = inputId;
        window.currentQuickSearchOptions = options;

        renderQuickSearchList(options);
        document.getElementById('quick-search-modal').classList.add('active');
        setTimeout(() => document.getElementById('quick-search-input').focus(), 100);
    };

    /**
     * クイック検索リストをレンダリング
     */
    function renderQuickSearchList(options) {
        const container = document.getElementById('quick-search-list');
        container.innerHTML = options.map(opt => `
            <div class="quick-search-item" onclick="selectQuickSearchItem('${opt.replace(/'/g, "\\'")}')">
                <span>${opt}</span>
                <ion-icon name="chevron-forward-outline"></ion-icon>
            </div>
        `).join('');
    }

    /**
     * クイック検索のフィルタリング
     */
    window.applyQuickSearchFilter = function () {
        const keyword = document.getElementById('quick-search-input').value.toLowerCase();
        const filtered = window.currentQuickSearchOptions.filter(opt =>
            opt.toLowerCase().includes(keyword)
        );
        renderQuickSearchList(filtered);
    };

    /**
     * 項目選択
     */
    window.selectQuickSearchItem = function (value) {
        const input = document.getElementById(window.currentQuickSearchInputId);
        if (input) {
            input.value = value;
            // イベントを発火させて変更を通知
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
        }
        closeQuickSearch();
    };

    function closeQuickSearch() {
        document.getElementById('quick-search-modal').classList.remove('active');
    }

    /**
     * 全てのサジェスト入力欄に検索ボタンを付与
     */
    function initQuickSearchButtons() {
        document.querySelectorAll('input[list]').forEach(input => {
            // 既にボタンがあるか、特定の除外対象（マスタ編集用など）でなければ追加
            if (input.parentElement.classList.contains('input-with-icon')) return;
            if (input.id.startsWith('master-')) return;

            const wrapper = document.createElement('div');
            wrapper.className = 'input-with-icon';
            input.parentNode.insertBefore(wrapper, input);
            wrapper.appendChild(input);

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'icon-input-btn mini-search-btn';
            btn.innerHTML = '<ion-icon name="search-outline"></ion-icon>';
            btn.onclick = () => window.openQuickSearch(input.id);
            wrapper.appendChild(btn);
        });
    }

    function toggleSaleItemInput(isPersonal) {
        let saleItemField = document.getElementById('sale-item');
        if (!saleItemField) return;

        const container = saleItemField.closest('.input-with-icon');
        if (!container) return;

        const currentVal = saleItemField.value;
        const modeText = document.getElementById('sale-item-mode-text');

        if (isPersonal && saleItemField.tagName === 'SELECT') {
            const input = document.createElement('input');
            input.type = 'text';
            input.id = 'sale-item';
            input.placeholder = '品名を自由入力';
            input.value = currentVal;

            container.replaceChild(input, saleItemField);
            if (modeText) modeText.textContent = '(個人用: 自由入力可)';
        } else if (!isPersonal && saleItemField.tagName === 'INPUT') {
            const masters = currentMasters['M_画面制御'] || [];
            const ctrl = masters.find(m => m['要素ID'] === 'sale-item');
            const type = ctrl ? ctrl['タイプ'] : 'select';

            if (type === 'select') {
                const select = document.createElement('select');
                select.id = 'sale-item';
                const productMaster = (currentMasters['M_商品'] || [])
                    .filter(p => p['カテゴリ'] === '商品' || p['カテゴリ'] === '単体商品');
                select.innerHTML = '<option value="">--- 選択してください ---</option>';
                productMaster.forEach(p => {
                    const o = document.createElement('option');
                    o.value = p['品名'];
                    o.textContent = p['品名'];
                    select.appendChild(o);
                });
                select.value = currentVal;

                container.replaceChild(select, saleItemField);
                if (modeText) modeText.textContent = '(事業用: リストから選択のみ)';
            } else {
                saleItemField.placeholder = '品名を入力または検索';
                if (modeText) modeText.textContent = '(事業用: リストから選択のみ)';
            }
        }

        setupImagePreviewListeners();
    }

    /**
     * 入力欄にクリアボタン（×）を追加するヘルパー
     */
    function addClearButton(elm) {
        if (!elm || elm.dataset.hasClear === "true") return;

        const wrapper = document.createElement('div');
        wrapper.className = 'input-clear-wrapper';

        const parent = elm.parentElement;
        if (!parent) return;

        parent.insertBefore(wrapper, elm);
        wrapper.appendChild(elm);

        const btn = document.createElement('button');
        btn.className = 'input-clear-btn';
        btn.innerHTML = '<ion-icon name="close-outline"></ion-icon>';
        btn.setAttribute('type', 'button'); // 送信防止
        btn.setAttribute('tabindex', '-1'); // タブ移動でスキップ
        wrapper.appendChild(btn);

        const toggleBtn = () => {
            btn.classList.toggle('visible', elm.value.length > 0);
        };

        elm.addEventListener('input', toggleBtn);
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            elm.value = '';
            toggleBtn();
            elm.focus();
            // 入力イベントを発火させて連動するロジックを動かす
            elm.dispatchEvent(new Event('input'));
        });

        elm.dataset.hasClear = "true";
        toggleBtn(); // 初回表示判定
    }

    function setupAutocompleteListeners() {
        const expItem = document.getElementById('exp-item');
        if (expItem && !expItem.dataset.acBound) {
            expItem.dataset.acBound = "true";
            expItem.addEventListener('input', (e) => {
                const val = e.target.value;
                if (!currentMasters || !currentMasters['M_経費品名']) return;
                const expenses = currentMasters['M_経費品名'];
                const match = expenses.find(r => r['品名'] === val);
                if (match && match['デフォルト仕訳']) {
                    const descInput = document.getElementById('exp-account');
                    if (descInput) descInput.value = match['デフォルト仕訳'];
                }
            });
        }

        // 販売：発送方法の選択時に送料を自動セット
        const saleShipping = document.getElementById('sale-shipping');
        if (saleShipping && !saleShipping.dataset.acBound) {
            saleShipping.dataset.acBound = "true";
            saleShipping.addEventListener('change', (e) => {
                const val = e.target.value;
                if (!currentMasters || !currentMasters['M_発送'] || !val) return;
                const shippingMaster = currentMasters['M_発送'];
                const match = shippingMaster.find(r => r['発送方法'] === val);
                if (match && match['送料'] !== undefined) {
                    const costInput = document.getElementById('sale-shipping-cost');
                    if (costInput) costInput.value = match['送料'];
                }
            });
        }
    }

    /**
     * 製造・販売：入力変更時に在庫チェックを走らせるリスナー群
     */
    function setupInventoryCheckListeners() {
        // 製造：BOM在庫チェック
        const makeItem = document.getElementById('make-item');
        const makeQty = document.getElementById('make-quantity');

        if (makeItem && !makeItem.dataset.bomBound) {
            makeItem.dataset.bomBound = "true";
            const bomCheckUpdate = () => {
                if (typeof updateBOMCheck === 'function') {
                    updateBOMCheck(makeItem.value, document.getElementById('make-quantity')?.value || '');
                }
            };
            makeItem.addEventListener(makeItem.tagName === 'INPUT' ? 'input' : 'change', bomCheckUpdate);
        }
        if (makeQty && !makeQty.dataset.quantityBomBound) {
            makeQty.dataset.quantityBomBound = 'true';
            const update = () => updateBOMCheck(document.getElementById('make-item')?.value || '', makeQty.value);
            makeQty.addEventListener('input', update);
            makeQty.addEventListener('change', update);
        }

        // 販売：単品在庫チェック
        const saleItem = document.getElementById('sale-item');
        const saleQty = document.getElementById('sale-quantity');

        if (saleItem && !saleItem.dataset.stockBound) {
            saleItem.dataset.stockBound = "true";
            const saleCheckUpdate = () => {
                const itemName = saleItem.value;
                if (typeof updateSaleStockCheck === 'function') {
                    updateSaleStockCheck(itemName, document.getElementById('sale-quantity')?.value || '');
                }

                // 価格の自動補完 (提案35対応 & バグ修正)
                if (itemName) {
                    const priceInput = document.getElementById('sale-price');
                    if (priceInput) {
                        // 品名が変更された、または価格が空の場合に補完を実行
                        const isItemChanged = itemName !== saleItem.dataset.lastAutoItem;
                        const isPriceEmpty = !priceInput.value;
                        const isAutoPrice = priceInput.value === priceInput.dataset.lastAutoPrice;

                        if (isItemChanged || isPriceEmpty || isAutoPrice) {
                            let suggestedPrice = null;

                            // 1. 直近の販売履歴から検索
                            const salesHistory = lastRawData ? lastRawData['T_販売'] : null;
                            if (salesHistory && salesHistory.length > 1) {
                                const headers = salesHistory[0];
                                const nameIdx = headers.indexOf('品名');
                                const priceIdx = headers.indexOf('販売価格') !== -1 ? headers.indexOf('販売価格') : headers.indexOf('価格');
                                const statusIdx = headers.indexOf('ステータス');

                                for (let i = salesHistory.length - 1; i >= 1; i--) {
                                    const row = salesHistory[i];
                                    if (row[nameIdx] === itemName && row[statusIdx] === '完了' && row[priceIdx]) {
                                        suggestedPrice = row[priceIdx];
                                        break;
                                    }
                                }
                            }

                            // 2. 履歴になければ商品マスタから取得
                            if (!suggestedPrice && currentMasters && currentMasters['M_商品']) {
                                const product = currentMasters['M_商品'].find(p => p['品名'] === itemName);
                                if (product && product['販売単価']) {
                                    suggestedPrice = product['販売単価'];
                                }
                            }

                            if (suggestedPrice) {
                                priceInput.value = suggestedPrice;
                                priceInput.dataset.lastAutoPrice = suggestedPrice; // 自動入力値を記憶
                                saleItem.dataset.lastAutoItem = itemName; // 補完時の品名を記憶
                                priceInput.dispatchEvent(new Event('change'));
                            }
                        }
                    }
                }
            };
            saleItem.addEventListener(saleItem.tagName === 'INPUT' ? 'input' : 'change', saleCheckUpdate);
        }
        // 品名の再生成に影響されず、数量リスナーは同じINPUTへ一度だけ登録する。
        if (saleQty && !saleQty.dataset.quantityStockBound) {
            saleQty.dataset.quantityStockBound = 'true';
            const update = () => updateSaleStockCheck(document.getElementById('sale-item')?.value || '', saleQty.value);
            saleQty.addEventListener('input', update);
            saleQty.addEventListener('change', update);
        }
    }

    /**
     * 販売時の在庫チェック (提案対応)
     */
    function updateSaleStockCheck(itemName, quantity) {
        const container = document.getElementById('sale-stock-check-container');
        const resultDiv = document.getElementById('sale-stock-check-result');
        if (!container || !resultDiv) return;

        if (!itemName || !quantity) {
            container.style.display = 'none';
            return;
        }

        const qty = parseFloat(quantity) || 0;
        const stockItem = (currentMasters['T_在庫集計'] || []).find(s => s['品名'] === itemName);

        if (!stockItem) {
            container.style.display = 'none';
            return;
        }

        container.style.display = 'block';
        const currentQty = parseFloat(stockItem['現在庫数']) || 0;
        const remaining = currentQty - qty;
        const isOk = remaining >= 0;

        // 案1: 品名を除外したコンパクトな表示
        resultDiv.innerHTML = `
            <div class="bom-item-status" style="font-weight: 600;">
                <span class="bom-item-qty ${isOk ? 'ok' : 'ng'}" style="flex: none; min-width: auto; padding: 2px 6px;">
                    <ion-icon name="${isOk ? 'checkmark-circle' : 'warning'}" style="vertical-align: middle; margin-right: 4px;"></ion-icon>
                    ${isOk ? '在庫OK' : '在庫不足'}
                </span>
                <span style="font-size: 11px; color: var(--text-muted); margin-left: 8px;">
                    (現在庫: ${currentQty} → ${isOk ? '残り' : '不足'}: ${Math.abs(remaining)})
                </span>
            </div>
        `;
    }

    function setupTransactionSubmitters() {
        const configs = [
            { btnId: 'buy-submit', sheet: 'T_仕入', fields: { date: 'purchase-date', status: 'buy-status-entry', vendor: 'buy-vendor', item: 'buy-item', price: 'buy-price', quantity: 'buy-quantity', payment: 'buy-payment', category: 'buy-category', note: 'buy-note' } },
            { btnId: 'exp-submit', sheet: 'T_経費', fields: { date: 'expense-date', status: 'exp-status-entry', account: 'exp-account', vendor: 'exp-vendor', item: 'exp-item', price: 'exp-price', quantity: 'exp-quantity', payment: 'exp-payment', note: 'exp-note' } },
            { btnId: 'make-submit', sheet: 'T_製造', fields: { date: 'manufacturing-date', item: 'make-item', quantity: 'make-quantity', note: 'make-note' }, extra: { status: '製造開始' } },
            { btnId: 'sale-submit', sheet: 'T_販売', fields: { date: 'sales-date', buyer: 'sale-buyer', item: 'sale-item', quantity: 'sale-quantity', price: 'sale-price', shipping: 'sale-shipping', shippingCost: 'sale-shipping-cost', trackingNumber: 'sale-tracking-number', note: 'sale-note' }, extra: { status: '取引開始' } }
        ];

        configs.forEach(conf => {
            const btns = [document.getElementById(conf.btnId)];

            // 販売登録の場合は「入金待ち」ボタンも対象に含める
            if (conf.btnId === 'sale-submit') {
                const pendingBtn = document.getElementById('sale-submit-pending');
                if (pendingBtn) btns.push(pendingBtn);
            }

            btns.forEach(btn => {
                if (!btn || btn.dataset.bound) return;
                btn.dataset.bound = "true";
                btn.addEventListener('click', async () => {
                    // ボタンIDによってステータスを上書き
                    const statusOverride = (btn.id === 'sale-submit-pending') ? { status: '入金待ち' } : {};
                    const payload = conf.extra ? { ...conf.extra, ...statusOverride } : { ...statusOverride };

                    for (const key in conf.fields) {
                        const el = document.getElementById(conf.fields[key]);
                        if (el) {
                            let val = el.value;
                            if (key === 'quantity') {
                                val = val.trim() === '' ? NaN : Number(val);
                            } else if (el.type === 'number') {
                                // 数値項目が空の場合は NaN を維持（バリデーションで検知するため）
                                val = val === '' ? NaN : parseFloat(val);
                            } else if (typeof val === 'string') {
                                val = val.trim();
                            }
                            payload[key] = val;
                        }
                    }

                    // 経費タブの特別なフラグ取得
                    if (conf.btnId === 'exp-submit') {
                        const stockCb = document.getElementById('exp-is-stock');
                        payload.isStock = (stockCb && stockCb.checked) ? 1 : 0;
                        const rctCb = document.getElementById('exp-receipt');
                        payload.receipt = (rctCb && rctCb.checked) ? 1 : 0;
                    }

                    // Add Sale Type and Shipping Payer for sales
                    if (conf.btnId === 'sale-submit') {
                        const typeRadio = document.querySelector('input[name="sales-type"]:checked');
                        payload.type = typeRadio ? typeRadio.value : 'business';

                        const payerRadio = document.querySelector('input[name="sale-shipping-payer"]:checked');
                        payload.shippingPayer = payerRadio ? payerRadio.value : '出品者';
                    }

                    // --- 登録前バリデーション ---
                    const error = validateData(conf.sheet, payload);
                    if (error) {
                        alert(`入力エラー:\n${error}`);
                        return;
                    }

                    await handleSubmission(btn, conf.sheet, payload);
                });
            });
        });
    }

    /**
     * 共通バリデーションロジック
     */
    function validateData(sheet, data, isUpdate = false) {
        const configs = [
            { btnId: 'buy-submit', sheet: 'T_仕入', fields: { date: 'purchase-date', status: 'buy-status-entry', vendor: 'buy-vendor', item: 'buy-item', price: 'buy-price', quantity: 'buy-quantity', payment: 'buy-payment', category: 'buy-category', note: 'buy-note' } },
            { btnId: 'exp-submit', sheet: 'T_経費', fields: { date: 'expense-date', status: 'exp-status-entry', account: 'exp-account', vendor: 'exp-vendor', item: 'exp-item', price: 'exp-price', quantity: 'exp-quantity', payment: 'exp-payment', note: 'exp-note' } },
            { btnId: 'make-submit', sheet: 'T_製造', fields: { date: 'manufacturing-date', item: 'make-item', quantity: 'make-quantity', note: 'make-note' } },
            { btnId: 'sale-submit', sheet: 'T_販売', fields: { date: 'sales-date', buyer: 'sale-buyer', item: 'sale-item', quantity: 'sale-quantity', price: 'sale-price', shipping: 'sale-shipping', shippingCost: 'sale-shipping-cost', trackingNumber: 'sale-tracking-number', note: 'sale-note' } }
        ];
        const confFields = (configs.find(c => c.sheet === sheet) || {}).fields || {};

        const rules = {
            'T_仕入': [
                { key: 'date', label: '仕入日', required: true },
                { key: 'status', label: 'ステータス', required: true },
                { key: 'vendor', label: '仕入先', required: true },
                { key: 'item', label: '品名', required: true },
                { key: 'price', label: '価格', required: true, min: 0 },
                { key: 'quantity', label: '数量', required: true, min: 1 },
                { key: 'payment', label: '支払方法', required: true },
                { key: 'category', label: '区分', required: true }
            ],
            'T_経費': [
                { key: 'date', label: '注文日', required: true },
                { key: 'status', label: 'ステータス', required: true },
                { key: 'account', label: '仕訳', required: true },
                { key: 'vendor', label: '購入先', required: true },
                { key: 'item', label: '品名・内容', required: true },
                { key: 'price', label: '価格', required: true, min: 0 },
                { key: 'quantity', label: '数量', required: true, min: 1 },
                { key: 'payment', label: '支払方法', required: true }
            ],
            'T_製造': [
                { key: 'date', label: '製造開始日', required: true },
                { key: 'item', label: '品名', required: true },
                { key: 'quantity', label: '製造数量', required: true, min: 1 }
            ],
            'T_販売': [
                { key: 'date', label: '販売開始日', required: true },
                { key: 'buyer', label: '売先', required: true },
                { key: 'item', label: '販売品名', required: true },
                { key: 'quantity', label: '数量', required: true, min: 1 },
                { key: 'price', label: '販売価格', required: true, min: 0 },
                { key: 'shipping', label: '発送方法', required: true },
                { key: 'shippingCost', label: '送料合計', required: true, min: 0 }
            ]
        };

        const sheetRules = rules[sheet];
        if (!sheetRules) return null;
        const quantityControl = document.getElementById(confFields.quantity);
        const plannedEmpty = data.status === '購入予定' && Number.isNaN(data.quantity);
        if (!isUpdate && (sheet === 'T_販売' || quantityControl?.dataset.quantitySelect === 'true') && !plannedEmpty && (!Number.isSafeInteger(data.quantity) || data.quantity < 1)) {
            return '数量は1以上の整数で入力してください。';
        }

        for (const rule of sheetRules) {
            const val = data[rule.key];

            // 必須チェック
            if (rule.required) {
                // 購入予定の場合はすべての必須チェックをスキップ (提案30: 柔軟なワークフロー)
                const isPlanned = (data.status === '購入予定');

                if (!isPlanned) {
                    if (val === undefined || val === null || val === '' || (typeof val === 'number' && isNaN(val))) {
                        return `「${rule.label}」を入力してください。`;
                    }
                }
            }

            // 数値チェック（最小値）
            if (rule.min !== undefined && typeof val === 'number') {
                // 購入予定の場合は数値の最小値チェックをスキップ
                const isPlanned = (data.status === '購入予定');

                if (!isPlanned) {
                    if (val < rule.min) {
                        return `「${rule.label}」は ${rule.min} 以上の数値を入力してください。`;
                    }
                }
            }

            // suggest-strict のバリデーション (新規登録時のみ実行)
            const el = document.getElementById(confFields[rule.key]);
            if (el && el.dataset.type === 'suggest-strict' && !isUpdate) {
                // 販売タブで「個人用」が選択されている場合は、マスタチェックをスキップする
                const isPersonalSales = (sheet === 'T_販売' && data.type === 'personal');
                if (!isPersonalSales) {
                    const masterName = el.dataset.master;
                    const masterData = currentMasters[masterName] || [];

                    // 除外フィールドを考慮して品名などのキーフィールドを特定
                    const excludeFields = ['表示順', '使用FLG', 'カテゴリ', '手数料率', '送料', '用途区分', '説明', 'デフォルト仕訳', '役割（タイプ）', '対象機能', '画面名称', '商品ID', '仕入先ID', '売先ID', '発送ID', '仕訳ID', '画像URL'];
                    const priorityFields = ['品名', '完成品名', '名称', 'ステータス名称', '仕入先', '売先', '発送方法', '仕訳名'];

                    let keyField = priorityFields.find(k => masterData.length > 0 && masterData[0].hasOwnProperty(k));
                    if (!keyField && masterData.length > 0) {
                        keyField = Object.keys(masterData[0]).find(k => !excludeFields.includes(k));
                    }

                    if (keyField) {
                        const exists = masterData.some(r => String(r[keyField]) === String(val));
                        if (!exists) {
                            return `「${rule.label}」にリスト外の値（${val}）が入力されています。マスタに登録済みの名称を選択してください。`;
                        }
                    }
                }
            }
        }

        // 特別なビジネスルール: 販売時の在庫チェック（簡易版 - フロントにある最新データで確認）
        // 履歴更新時は既に引当済みのためスキップ
        if (sheet === 'T_販売' && !isUpdate) {
            const isPersonal = data.type === 'personal';
            const productName = data.item;
            const qty = data.quantity;

            // マスタに存在する商品のみチェック（個人用の新規品名はスルー）
            const stockItem = (currentMasters['T_在庫集計'] || []).find(s => s['品名'] === productName);
            if (stockItem) {
                const currentQty = parseFloat(stockItem['現在庫数']) || 0;
                if (qty > currentQty) {
                    return `在庫不足です。「${productName}」の現在庫は ${currentQty} です。`;
                }
            }
        }

        return null;
    }

    async function handleSubmission(btn, sheet, data) {
        const originalHTML = btn.innerHTML;
        btn.innerHTML = '<ion-icon name="sync-outline" class="spinning"></ion-icon> 登録中...';
        btn.disabled = true;
        setLoading(true, 'データを登録中...');

        const refreshScope = (sheet === 'T_仕入') ? 'purchase' :
            (sheet === 'T_経費') ? 'expense' :
                (sheet === 'T_製造') ? 'manufacturing' :
                    (sheet === 'T_販売') ? 'sales' : 'all';

        try {
            const response = await fetchAPI('registerTransaction', { sheet: sheet, data: data, scope: refreshScope });
            if (response.status === 'success') {
                btn.innerHTML = '<ion-icon name="checkmark-outline"></ion-icon> 完了';
                btn.style.background = 'var(--accent-green)';

                // 入力フォームの初期化
                const tabContent = btn.closest('.tab-content');
                if (tabContent) {
                    const inputs = tabContent.querySelectorAll('input:not([type="radio"]):not([type="checkbox"]), select, textarea');
                    inputs.forEach(input => {
                        input.value = '';
                    });

                    // 販売タブの場合は数量を1にリセット (提案35)
                    if (sheet === 'T_販売') {
                        const qtyInput = document.getElementById('sale-quantity');
                        if (qtyInput) qtyInput.value = '1';

                        // 品名・数量をリセット後に input イベントを発火させて
                        // 在庫OKの表示コンテナを初期化する
                        const itemInput = document.getElementById('sale-item');
                        if (itemInput) {
                            itemInput.dispatchEvent(new Event('input', { bubbles: true }));
                            itemInput.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                    }

                    // 商品写真プレビューもクリア
                    const previews = tabContent.querySelectorAll('.input-image-preview');
                    previews.forEach(p => {
                        p.innerHTML = '';
                        p.classList.remove('visible');
                    });

                    // 日付等の初期値再セット
                    const defaultDates = tabContent.querySelectorAll('input[type="date"]');
                    const d = new Date();
                    const ymd = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                    defaultDates.forEach(el => el.value = ymd);

                    const defaultQtys = tabContent.querySelectorAll('input[type="number"]');
                    defaultQtys.forEach(el => { if (el.id.includes('quantity') && el.id !== 'sale-quantity') el.value = ''; });
                    tabContent.querySelectorAll('[data-quantity-select="true"]').forEach(input => {
                        delete input.dataset.quantityCustomValue;
                        syncQuantityChoice_(input, true);
                    });
                }
                // 履歴と在庫状況を即時更新してUIに反映 (高速化対応)
                if (response.newRecord && response.sheetName) {
                    console.time('Client:IncrementalUpdate(New)');

                    // 1. 履歴データの生データを更新
                    if (!lastRawData[response.sheetName]) lastRawData[response.sheetName] = [];
                    lastRawData[response.sheetName].push(response.newRecord);

                    // 2. 在庫集計が同梱されている場合は更新
                    if (response.inventorySummary) {
                        currentMasters['T_在庫集計'] = response.inventorySummary;
                    }

                    // 3. 全データを再集計・再描画
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);

                    console.timeEnd('Client:IncrementalUpdate(New)');
                } else if ((response.data && response.data.historyData) || response.historyData) {
                    // 履歴データが含まれている場合の処理
                    const hData = response.historyData || response.data.historyData;
                    lastRawData = Object.assign({}, lastRawData, hData.rawData);
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);
                    refreshInventoryUI(); // 在庫一覧も更新
                } else {
                    // 最終手段
                    await fetchHistory(refreshScope);
                }

                // マスタ情報の更新（新規マスタ追加があった時のみ）
                if (response.masterAdded) {
                    console.log("Master added, refreshing system data...");
                    initSystem();
                }

                showToast(`登録が完了しました`);

                setTimeout(() => {
                    btn.innerHTML = originalHTML;
                    btn.style.background = '';
                    btn.disabled = false;
                }, 2000);
            } else {
                throw new Error(response.message);
            }
        } catch (e) {
            console.error(e);
            showToast("登録失敗: " + e.message, 'error');
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        } finally {
            setLoading(false);
        }
    }

    function collectInventoryProgress(history = {}) {
        const totals = new Map();
        const add = (row, field, quantity) => {
            const name = String(row['品名'] || row['完成品名'] || row['商品名'] || '').trim();
            const value = Number(quantity);
            if (!name || !Number.isFinite(value) || value <= 0) return;
            if (!totals.has(name)) totals.set(name, { purchase: 0, manufacturing: 0, expense: 0, planned: 0 });
            totals.get(name)[field] += value;
        };
        for (const row of history['T_仕入'] || []) {
            const status = String(row['ステータス'] || '').trim();
            if (status === '注文済み') add(row, 'purchase', row['数量']);
            if (status === '購入予定') add(row, 'planned', row['数量']);
        }
        for (const row of history['T_製造'] || []) {
            if (['製造開始', '製造中', 'テスト中', '梱包中'].includes(String(row['ステータス'] || '').trim())) {
                add(row, 'manufacturing', row['数量'] ?? row['製造数量']);
            }
        }
        for (const row of history['T_経費'] || []) {
            const status = String(row['ステータス'] || '').trim();
            if (status === '注文済み' && Number(row['管理対象']) === 1) add(row, 'expense', row['数量']);
            // 購入予定は既存アラートと同じく、仕入・経費を合算する。
            if (status === '購入予定') add(row, 'planned', row['数量']);
        }
        return totals;
    }

    function inventoryProgressBadges(progress) {
        if (!progress) return '';
        const badges = [
            ['purchase', 'pur', 'cart-outline', '仕入中'],
            ['manufacturing', 'mfg', 'hammer-outline', '製造中'],
            ['expense', 'exp', 'receipt-outline', '経費注文中'],
            ['planned', 'planned', 'calendar-outline', '購入予定']
        ];
        return badges.filter(([field]) => progress[field] > 0).map(([field, cls, icon, title]) => {
            const label = title + ' ' + progress[field].toLocaleString() + '個';
            return `<span class="alert-processing-badge ${cls}" title="${escapeHtml(label)}"><ion-icon name="${icon}"></ion-icon><span class="inventory-progress-text">${escapeHtml(label)}</span></span>`;
        }).join('');
    }

    function renderStockList(stockData) {
        const container = document.getElementById('stock-list-body');
        const badge = document.querySelector('.count-badge');
        if (!container) return;

        container.innerHTML = '';
        let safeData = stockData || [];

        // フィルタ適用：未確認のみ（サイクルカウント基準）
        if (stocktakeSession.uncheckedOnly) {
            const now = new Date();
            const cycleDays = 30; // 30日を基準とする

            safeData = safeData.filter(row => {
                const lastDateVal = row['最終棚卸日'];
                let isRecent = false;
                if (lastDateVal) {
                    const d = new Date(lastDateVal);
                    const diffTime = Math.abs(now - d);
                    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                    isRecent = diffDays <= cycleDays;
                }
                const isVerifiedInSession = stocktakeSession.verifiedItems.has(row['品名']);
                return !isRecent && !isVerifiedInSession;
            });
        }

        if (badge) badge.textContent = safeData.length + '点';

        if (safeData.length === 0) {
            const msg = isStocktakeMode ? (stocktakeSession.uncheckedOnly ? 'サイクル期間内の未確認品目はありません' : '対象品目なし') : (currentMasters['T_在庫集計'] ? '表示できる在庫はありません' : '在庫データを取得できませんでした。');
            container.innerHTML = `<div style="text-align:center; padding: 40px; color: var(--text-secondary);">${msg}</div>`;
            return;
        }

        const now = new Date();

        const progressByItem = isStocktakeMode ? new Map() : collectInventoryProgress(lastHistoryData?.history);
        safeData.forEach(row => {
            const itemName = row['品名'];
            const category = row['カテゴリ'] || row['商品区分'] || '';
            const rawQty = parseFloat(row['現在庫数']);
            const stockQty = isNaN(rawQty) ? 0 : rawQty;

            const rawThreshold = parseFloat(row['閾値']);
            const threshold = isNaN(rawThreshold) ? 10 : rawThreshold;
            const useFlag = parseInt(row['使用FLG']) !== 0;

            const lastStocktakeDate = row['最終棚卸日'] ? new Date(row['最終棚卸日']) : null;
            const isRecent = lastStocktakeDate && (Math.abs(now - lastStocktakeDate) / (1000 * 60 * 60 * 24) <= 30);
            const isVerified = isRecent || stocktakeSession.verifiedItems.has(itemName);

            const card = document.createElement('div');
            card.className = 'stock-item-card';
            if (!useFlag) card.classList.add('inactive-row');
            if (isVerified && isStocktakeMode) card.classList.add('verified-row');

            const itemInMaster = (currentMasters['M_商品'] || []).find(m => m['品名'] === itemName);
            const imageUrl = safeHttpsUrl(itemInMaster?.['画像URL']);
            const itemID = (itemInMaster && itemInMaster['商品ID']) || row['商品ID'] || '';
            const location = row['保管場所'] || (itemInMaster && itemInMaster['保管場所']) || '';
            const barcode = (itemInMaster && itemInMaster['QR/バーコード']) || row['QR/バーコード'] || '';

            card.setAttribute('data-item-name', itemName);
            card.setAttribute('data-item-id', itemID);
            card.setAttribute('data-barcode', barcode);
            card.setAttribute('data-location', location);

            if (isStocktakeMode) {
                // ---- 棚卸モードのカード ----
                const saved = stocktakeData[itemName];
                const currentActual = saved ? saved.actual : stockQty;
                const currentDiff = saved ? saved.diff : 0;

                let diffText = currentDiff > 0 ? '+' + currentDiff : String(currentDiff);
                let diffClass = currentDiff > 0 ? 'diff-plus' : (currentDiff < 0 ? 'diff-minus' : 'diff-zero');

                card.innerHTML = `
                    <div class="card-header">
                        <div class="product-thumb-container">
                            ${imageUrl ? `<img src="${escapeHtml(imageUrl)}" loading="lazy">` : `<ion-icon name="image-outline"></ion-icon>`}
                        </div>
                        <div class="card-main-info">
                            <div class="card-product-name">${escapeHtml(itemName)}</div>
                            <div class="card-category-badge">${escapeHtml(category)}</div>
                            <div class="card-sub-info">
                                ${itemID ? `<span class="id-badge">${escapeHtml(itemID)}</span>` : ''}
                                ${barcode ? `<span class="barcode-badge"><ion-icon name="barcode-outline"></ion-icon>${escapeHtml(barcode)}</span>` : ''}
                                ${location ? location.split(/[／/，,、\s]+/).filter(s => s).map(loc => `<span class="location-badge"><ion-icon name="location-outline"></ion-icon>${escapeHtml(loc)}</span>`).join('') : ''}
                                ${lastStocktakeDate ? `前回: ${lastStocktakeDate.toLocaleDateString()} ` : ''}
                                ${isRecent ? `<span class="stocktake-done-badge"><ion-icon name="checkmark"></ion-icon>完了</span>` : ''}
                            </div>
                        </div>
                        <button class="verify-btn ${isVerified ? 'verified' : ''}" title="${isVerified ? '未確認に戻す' : '確認済みにする'}">
                            <ion-icon name="${isVerified ? 'checkmark-done' : 'checkmark'}"></ion-icon>
                        </button>
                    </div>
                    <div class="card-controls">
                        <div class="control-group">
                            <span class="control-label">現在庫</span>
                            <span class="control-value">${stockQty}</span>
                        </div>
                        <div class="control-group" style="align-items: center;">
                            <span class="control-label">実在庫入力</span>
                            <div class="stepper-placeholder"></div>
                        </div>
                        <div class="control-group" style="align-items: flex-end;">
                            <span class="control-label">差異</span>
                            <span class="diff-badge ${diffClass}">${diffText}</span>
                        </div>
                    </div>
                `;

                const stepperPlaceholder = card.querySelector('.stepper-placeholder');
                const verifyBtn = card.querySelector('.verify-btn');

                const stepper = createStepper(currentActual, (newVal) => {
                    const diff = newVal - stockQty;
                    stocktakeData[itemName] = { actual: newVal, diff: diff };

                    if (!stocktakeSession.verifiedItems.has(itemName)) {
                        stocktakeSession.verifiedItems.add(itemName);
                        verifyBtn.classList.add('verified');
                        card.classList.add('verified-row');
                        verifyBtn.querySelector('ion-icon').setAttribute('name', 'checkmark-done');
                    }

                    const diffBadge = card.querySelector('.diff-badge');
                    diffBadge.textContent = diff > 0 ? '+' + diff : String(diff);
                    diffBadge.className = 'diff-badge ' + (diff > 0 ? 'diff-plus' : (diff < 0 ? 'diff-minus' : 'diff-zero'));

                    updateStocktakeSummary();
                });
                stepperPlaceholder.appendChild(stepper);

                verifyBtn.addEventListener('click', () => {
                    const currentlyVerified = stocktakeSession.verifiedItems.has(itemName);
                    if (currentlyVerified) {
                        stocktakeSession.verifiedItems.delete(itemName);
                        verifyBtn.classList.remove('verified');
                        card.classList.remove('verified-row');
                        verifyBtn.querySelector('ion-icon').setAttribute('name', 'checkmark');
                        const input = stepper.querySelector('input');
                        input.value = stockQty;
                        delete stocktakeData[itemName];
                        const diffBadge = card.querySelector('.diff-badge');
                        diffBadge.textContent = '0';
                        diffBadge.className = 'diff-badge diff-zero';
                    } else {
                        stocktakeSession.verifiedItems.add(itemName);
                        verifyBtn.classList.add('verified');
                        card.classList.add('verified-row');
                        verifyBtn.querySelector('ion-icon').setAttribute('name', 'checkmark-done');
                        if (!stocktakeData[itemName]) {
                            stocktakeData[itemName] = { actual: stockQty, diff: 0 };
                        }
                    }
                    updateStocktakeSummary();
                });

            } else {
                // ---- 通常モードのカード ----
                const isDanger = stockQty < threshold;
                const pendingThreshold = pendingStockChanges.thresholds[itemName];
                const displayThreshold = (pendingThreshold !== undefined) ? pendingThreshold : threshold;
                const pendingStatus = pendingStockChanges.statuses[itemName];
                const currentUseFlag = (pendingStatus !== undefined) ? (pendingStatus !== 0) : useFlag;

                const isThresholdDirty = pendingThreshold !== undefined;
                const isStatusDirty = pendingStatus !== undefined;
                const isRowDirty = isThresholdDirty || isStatusDirty;

                card.innerHTML = `
                    <div class="card-header">
                        <div class="product-thumb-container">
                            ${imageUrl ? `<img src="${escapeHtml(imageUrl)}" loading="lazy">` : `<ion-icon name="image-outline"></ion-icon>`}
                        </div>
                        <div class="card-main-info">
                            <div class="card-product-name clickable" data-name="${escapeHtml(itemName)}" data-category="${escapeHtml(category)}" onclick="navigateToTransactionForm(this.dataset.name, this.dataset.category)">${escapeHtml(itemName)}</div>
                            <div class="card-category-badge">${escapeHtml(category)}</div>
                            <div class="card-sub-info">
                                ${itemID ? `<span class="id-badge">${escapeHtml(itemID)}</span>` : ''}
                                ${barcode ? `<span class="barcode-badge"><ion-icon name="barcode-outline"></ion-icon>${escapeHtml(barcode)}</span>` : ''}
                                ${location ? location.split(/[／/，,、\s]+/).filter(s => s).map(loc => `<span class="location-badge"><ion-icon name="location-outline"></ion-icon>${escapeHtml(loc)}</span>`).join('') : ''}
                            </div>
                        </div>
                        <div class="status-icon-wrap">
                            <ion-icon name="${isDanger ? 'alert-circle' : 'checkmark-circle'}" 
                                        class="status-icon ${isDanger ? 'danger' : 'success'}"
                                        title="${isDanger ? '在庫不足' : '在庫あり'}"></ion-icon>
                        </div>
                    </div>
                    <div class="card-controls">
                        <div class="control-group">
                            <span class="control-label">現在庫</span>
                            <span class="control-value" style="color: ${isDanger ? '#ef4444' : 'var(--accent-green)'}">${stockQty}</span>
                        </div>
                        <div class="control-group" style="align-items: center;">
                            <span class="control-label">アラート閾値</span>
                            <div class="stepper-placeholder"></div>
                        </div>
                        <div class="card-actions action-col" style="margin-top: 4px;">
                            <button class="camera-btn" data-name="${escapeHtml(itemName)}" onclick="triggerPhotoUpload(this.dataset.name)" title="写真を登録/変更" style="margin-left:0;">
                                <ion-icon name="camera-outline"></ion-icon>
                            </button>
                            <button class="update-mini-btn no-text btn-save-threshold ${isThresholdDirty ? 'is-dirty' : ''}" title="閾値の変更を取り消す" aria-label="閾値の変更を取り消す">
                                <ion-icon name="arrow-undo-outline"></ion-icon>
                            </button>
                            <button class="update-mini-btn no-text secondary-action btn-archive-item ${isStatusDirty ? 'is-dirty' : ''}" title="${currentUseFlag ? '廃盤にする' : '復活させる'}">
                                <ion-icon name="${currentUseFlag ? 'archive-outline' : 'refresh-outline'}"></ion-icon>
                            </button>
                        </div>
                    </div>
                `;

                const progressBadges = inventoryProgressBadges(progressByItem.get(String(itemName || '').trim()));
                if (progressBadges) {
                    const progressRow = document.createElement('div');
                    progressRow.className = 'inventory-progress-badges';
                    progressRow.setAttribute('aria-label', '進行中の数量（現在庫には含みません）');
                    progressRow.innerHTML = progressBadges;
                    card.appendChild(progressRow);
                }

                if (isRowDirty) card.classList.add('is-dirty-row');

                const stepperPlaceholder = card.querySelector('.stepper-placeholder');
                const statusIconWrap = card.querySelector('.status-icon-wrap');
                const stockValCell = card.querySelector('.control-value');
                const saveBtn = card.querySelector('.btn-save-threshold');
                const archiveBtn = card.querySelector('.btn-archive-item');

                const stepper = createStepper(displayThreshold, (newThreshold) => {
                    const original = threshold;
                    if (!isNaN(newThreshold) && newThreshold !== original) {
                        pendingStockChanges.thresholds[itemName] = newThreshold;
                    } else {
                        delete pendingStockChanges.thresholds[itemName];
                    }
                    const danger = stockQty < (isNaN(newThreshold) ? threshold : newThreshold);
                    statusIconWrap.innerHTML = `
                        <ion-icon name="${danger ? 'alert-circle' : 'checkmark-circle'}" 
                                   class="status-icon ${danger ? 'danger' : 'success'}"
                                   title="${danger ? '在庫不足' : '在庫あり'}"></ion-icon>
                    `;
                    stockValCell.style.color = danger ? '#ef4444' : 'var(--accent-green)';
                    updateStockUpdateBar();
                    const isDirty = pendingStockChanges.thresholds[itemName] !== undefined || pendingStockChanges.statuses[itemName] !== undefined;
                    card.classList.toggle('is-dirty-row', isDirty);
                    saveBtn.classList.toggle('is-dirty', pendingStockChanges.thresholds[itemName] !== undefined);
                });
                stepperPlaceholder.appendChild(stepper);

                saveBtn.addEventListener('click', () => {
                    if (pendingStockChanges.thresholds[itemName] !== undefined) {
                        const original = threshold;
                        stepper.querySelector('input').value = original;
                        delete pendingStockChanges.thresholds[itemName];
                        const danger = stockQty < original;
                        statusIconWrap.innerHTML = `
                            <ion-icon name="${danger ? 'alert-circle' : 'checkmark-circle'}" 
                                       class="status-icon ${danger ? 'danger' : 'success'}"
                                       title="${danger ? '在庫不足' : '在庫あり'}"></ion-icon>
                        `;
                        stockValCell.style.color = danger ? '#ef4444' : 'var(--accent-green)';
                        const isDirty = pendingStockChanges.statuses[itemName] !== undefined;
                        card.classList.toggle('is-dirty-row', isDirty);
                        saveBtn.classList.toggle('is-dirty', false);
                        updateStockUpdateBar();
                    }
                });

                archiveBtn.addEventListener('click', () => {
                    const currentStatus = (pendingStockChanges.statuses[itemName] !== undefined) ? pendingStockChanges.statuses[itemName] : (useFlag ? 1 : 0);
                    const newStatus = currentStatus === 1 ? 0 : 1;
                    if (newStatus === (useFlag ? 1 : 0)) {
                        delete pendingStockChanges.statuses[itemName];
                    } else {
                        pendingStockChanges.statuses[itemName] = newStatus;
                    }
                    const isNowActive = newStatus === 1;
                    archiveBtn.querySelector('ion-icon').setAttribute('name', isNowActive ? 'archive-outline' : 'refresh-outline');
                    archiveBtn.title = isNowActive ? '廃盤にする' : '復活させる';
                    const isDirty = pendingStockChanges.thresholds[itemName] !== undefined || pendingStockChanges.statuses[itemName] !== undefined;
                    card.classList.toggle('is-dirty-row', isDirty);
                    archiveBtn.classList.toggle('is-dirty', pendingStockChanges.statuses[itemName] !== undefined);
                    updateStockUpdateBar();
                });
            }
            container.appendChild(card);
        });
    }


    /**
     * ▲▼ボタン付きのステッパーUIを作成する
     */
    function createStepper(initialValue, onChange) {
        // ... (existing createStepper implementation)
        const container = document.createElement('div');
        container.className = 'stepper-container';

        const btnMinus = document.createElement('button');
        btnMinus.className = 'stepper-btn';
        btnMinus.innerHTML = '<ion-icon name="remove-outline"></ion-icon>';

        const input = document.createElement('input');
        input.type = 'number';
        input.className = 'stepper-input';
        input.value = initialValue;

        const btnPlus = document.createElement('button');
        btnPlus.className = 'stepper-btn';
        btnPlus.innerHTML = '<ion-icon name="add-outline"></ion-icon>';

        const update = () => {
            let val = parseFloat(input.value) || 0;
            if (val < 0) val = 0;
            input.value = val;
            onChange(val);
        };

        btnMinus.addEventListener('click', (e) => {
            e.stopPropagation();
            input.value = (parseFloat(input.value) || 0) - 1;
            update();
        });

        btnPlus.addEventListener('click', (e) => {
            e.stopPropagation();
            input.value = (parseFloat(input.value) || 0) + 1;
            update();
        });

        input.addEventListener('change', update);
        input.addEventListener('click', (e) => e.stopPropagation());

        container.appendChild(btnMinus);
        container.appendChild(input);
        container.appendChild(btnPlus);

        return container;
    }

    function updateStocktakeSummary() {
        const verifiedCountElm = document.getElementById('stocktake-verified-count');
        const diffCountElm = document.getElementById('stocktake-diff-count');

        // セッションで「済」にした品目数
        const verifiedCount = stocktakeSession.verifiedItems.size;

        // 差異がある品目数（stocktakeDataに登録されているもののうち、差異が0でないもの）
        const diffCount = Object.keys(stocktakeData).filter(n => stocktakeData[n].diff !== 0).length;

        if (verifiedCountElm) verifiedCountElm.textContent = verifiedCount;
        if (diffCountElm) diffCountElm.textContent = diffCount;

        const submitBtn = document.getElementById('stocktake-submit');
        if (submitBtn) submitBtn.disabled = (verifiedCount === 0);
    }

    /**
     * ---- スキャン機能のロジック ----
     */
    function setupScannerListeners() {
        const topScanBtn = document.getElementById('stock-scan-btn-top');
        const bottomScanBtn = document.getElementById('stocktake-scan-btn');
        const closeBtn = document.getElementById('scanner-close-btn');
        const fileBtn = document.getElementById('scanner-file-btn');
        const fileInput = document.getElementById('qr-file-input');

        if (topScanBtn) topScanBtn.addEventListener('click', () => startScanner());
        if (bottomScanBtn) bottomScanBtn.addEventListener('click', () => startScanner());
        if (closeBtn) closeBtn.addEventListener('click', () => stopScanner());

        if (fileBtn) fileBtn.addEventListener('click', () => fileInput.click());
        if (fileInput) {
            fileInput.addEventListener('change', e => {
                if (e.target.files.length === 0) return;
                const file = e.target.files[0];
                if (!html5QrCode) html5QrCode = new Html5Qrcode("reader");

                showToast('画像を解析中...', 'info');
                html5QrCode.scanFile(file, true)
                    .then(decodedText => {
                        onScanSuccess(decodedText);
                        stopScanner();
                    })
                    .catch(err => {
                        alert("QRコードを認識できませんでした。別の角度から撮影してください。");
                        console.error("Scan file error:", err);
                    });
            });
        }
    }

    function startScanner() {
        const overlay = document.getElementById('scanner-overlay');
        overlay.style.display = 'flex';

        if (typeof showToast === 'function') showToast('カメラを起動しています...', 'info');

        if (!html5QrCode) {
            html5QrCode = new Html5Qrcode("reader");
        }

        const config = { fps: 10, qrbox: { width: 250, height: 250 } };

        html5QrCode.start({ facingMode: "environment" }, config, onScanSuccess)
            .then(() => {
                if (typeof showToast === 'function') showToast('スキャナーが起動しました', 'success');
            })
            .catch(err => {
                console.error("Scanner start error:", err);
                const errorName = err.name || "";
                let msg = "カメラの起動に失敗しました。";
                if (errorName === "NotAllowedError") msg += "\nカメラの使用許可を確認してください。";
                else if (errorName === "NotFoundError") msg += "\nカメラが見つかりません。";
                else msg += "\n詳細: " + err;

                alert(msg);
                overlay.style.display = 'none';
            });
    }

    function stopScanner() {
        const overlay = document.getElementById('scanner-overlay');
        overlay.style.display = 'none';

        if (html5QrCode && html5QrCode.isScanning) {
            html5QrCode.stop().catch(err => console.error("Scanner stop error:", err));
        }
    }

    function onScanSuccess(decodedText, decodedResult) {
        console.log(`Scan Result: ${decodedText}`);
        stopScanner();

        // 1. 場所QRコードの判定 (LOC-XXXX, LOC:XXXX, または名称そのまま)
        const locationOptions = [
            '1_棚上：メイン', '2_棚中：サブ1', '3_棚下：梱包1',
            '4_押入上1：部品1', '5_押入上2：部品2', '6_押入上3：梱包2',
            '7_押入下1：サブ2', '8_押入下2：部品箱'
        ];

        let targetLoc = null;
        if (decodedText.startsWith('LOC-')) {
            targetLoc = decodedText.replace('LOC-', '');
        } else if (decodedText.startsWith('LOC:')) {
            targetLoc = decodedText.replace('LOC:', '');
        } else if (locationOptions.includes(decodedText)) {
            targetLoc = decodedText;
        }

        if (targetLoc) {
            const searchInput = document.getElementById('stock-search-input');
            if (searchInput) {
                // 在庫タブを表示
                const invTabBtn = document.querySelector('.nav-item[data-target="inventory"]');
                if (invTabBtn) invTabBtn.click();

                // 検索窓に値をセット
                searchInput.value = targetLoc;
                searchInput.dispatchEvent(new Event('input')); // 検索実行

                if (typeof showToast === 'function') showToast(`場所「${targetLoc}」で絞り込みました`, 'success');
            }
            return;
        }

        // 2. JANコードまたは商品IDを検索窓に入力して「在庫タブ」を表示
        const searchInput = document.getElementById('stock-search-input');
        if (searchInput) {
            // 在庫タブに切り替え
            const invTabBtn = document.querySelector('.nav-item[data-target="inventory"]');
            if (invTabBtn) invTabBtn.click();

            // 検索窓に値をセット
            searchInput.value = decodedText;
            searchInput.dispatchEvent(new Event('input')); // 検索実行

            if (typeof showToast === 'function') showToast(`スキャン結果: ${decodedText} で検索しました`, 'success');
        }
    }

    function setupStockUpdateListeners() {
        const thresholdBtn = document.getElementById('bulk-update-threshold');
        const statusBtn = document.getElementById('bulk-update-status');

        if (thresholdBtn) {
            thresholdBtn.addEventListener('click', () => handleBulkUpdate('thresholds'));
        }
        if (statusBtn) {
            statusBtn.addEventListener('click', () => handleBulkUpdate('statuses'));
        }
    }

    function updateStockUpdateBar() {
        const bar = document.getElementById('stock-update-bar');
        if (!bar) return;

        const countBadge = document.getElementById('pending-update-count');
        const thresholdBtn = document.getElementById('bulk-update-threshold');
        const statusBtn = document.getElementById('bulk-update-status');

        const thresholdCount = Object.keys(pendingStockChanges.thresholds).length;
        const statusCount = Object.keys(pendingStockChanges.statuses).length;
        const totalCount = new Set([...Object.keys(pendingStockChanges.thresholds), ...Object.keys(pendingStockChanges.statuses)]).size;

        if (totalCount > 0) {
            bar.style.display = 'flex';
            if (countBadge) countBadge.textContent = totalCount;
            if (thresholdBtn) thresholdBtn.disabled = (thresholdCount === 0);
            if (statusBtn) statusBtn.disabled = (statusCount === 0);
        } else {
            bar.style.display = 'none';
        }
        syncScrollTopStockBar();
    }

    function syncScrollTopStockBar() {
        const bar = document.getElementById('stock-update-bar');
        const button = document.getElementById('scroll-to-top-btn');
        if (!button) return;
        if (!bar || getComputedStyle(bar).display === 'none' || bar.offsetHeight === 0) {
            button.style.removeProperty('--stock-update-button-bottom');
            return;
        }
        // offsetHeightを使い、操作欄の登場アニメーション中のtransformは測定に含めない。
        const bottom = parseFloat(getComputedStyle(bar).bottom) || 0;
        button.style.setProperty('--stock-update-button-bottom', (bottom + bar.offsetHeight + 12) + 'px');
    }

    function setupScrollTopStockBar() {
        const bar = document.getElementById('stock-update-bar');
        if (!bar || !document.getElementById('scroll-to-top-btn')) return;
        if (typeof ResizeObserver !== 'undefined') {
            const observer = new ResizeObserver(syncScrollTopStockBar);
            observer.observe(bar, { box: 'border-box' });
        }
        window.addEventListener('resize', syncScrollTopStockBar);
        syncScrollTopStockBar();
    }

    async function handleBulkUpdate(type) {
        const updates = pendingStockChanges[type];
        const count = Object.keys(updates).length;
        if (count === 0) return;

        const confirmMsg = type === 'thresholds'
            ? `${count}件の閾値変更を確定しますか？`
            : `${count}件の廃盤/復活設定を確定しますか？`;

        if (!confirm(confirmMsg)) return;

        const btn = document.getElementById(type === 'thresholds' ? 'bulk-update-threshold' : 'bulk-update-status');
        const originalHTML = btn.innerHTML;
        btn.innerHTML = '<ion-icon name="sync-outline" class="spinning"></ion-icon> 処理中...';
        btn.disabled = true;

        try {
            const payload = {
                thresholdUpdates: type === 'thresholds' ? updates : null,
                statusUpdates: type === 'statuses' ? updates : null
            };

            const res = await fetchAPI('updateStockBulk', payload);
            if (res.status === 'success') {
                showToast(`${count}件の更新を完了しました`);

                // ローカルマスタの更新
                for (const itemName in updates) {
                    const masterIdx = currentMasters['T_在庫集計'].findIndex(m => m['品名'] === itemName);
                    if (masterIdx !== -1) {
                        const key = type === 'thresholds' ? '閾値' : '使用FLG';
                        currentMasters['T_在庫集計'][masterIdx][key] = updates[itemName];
                    }
                }

                // 変更内容をクリア
                pendingStockChanges[type] = {};
                updateStockUpdateBar();
                updateNavBadges();

                // 再描画（検索条件を維持するためinputイベントを発火）
                const searchInput = document.getElementById('stock-search-input');
                if (searchInput) searchInput.dispatchEvent(new Event('input'));
            } else {
                throw new Error(res.message);
            }
        } catch (e) {
            alert('一括更新に失敗しました: ' + e.message);
        } finally {
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        }
    }

    // ---- History Management ----
    /**
     * 履歴データの更新
     * @param {string} scope 'all', 'purchase', 'expense', 'manufacturing', 'sales', 'inventory', 'ledger'
     */
    async function fetchHistory(scope = 'all') {
        try {
            const response = await fetchAPI('getHistory', { scope: scope });
            if (response.status === 'success') {
                if (response.data.isRaw) {
                    console.time('Client:MergeAndRender');
                    // 取得した部分的なデータを既存の生データにマージ
                    lastRawData = Object.assign({}, lastRawData, response.data.rawData);

                    // マージ後の全データを使用して、フィルタリングと集計を再実行
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);
                    refreshInventoryUI(); // 在庫一覧も更新
                    console.timeEnd('Client:MergeAndRender');
                } else {
                    // フォールバック（旧形式）
                    const processed = processClientData(response.data.rawData || response.data);
                    lastHistoryData = processed;
                    renderAllHistory(processed);
                    refreshInventoryUI(); // 在庫一覧も更新
                }
            }
        } catch (e) {
            console.error("Failed to fetch history:", e);
        }
    }

    /**
     * 発送方法と追跡番号から運送会社の追跡URLを生成する
     */
    function getTrackingUrl(method, number) {
        if (!number) return null;
        const numberStr = String(number);
        const cleanNumber = numberStr.replace(/[^0-9A-Za-z]/g, '');
        if (!cleanNumber) return null;

        const m = (method || "").toLowerCase();

        // ヤマト運輸判定
        if (m.includes('らくらく') || m.includes('宅急便') || m.includes('ヤマト') || m.includes('クロネコ')) {
            return 'https://toi.kuronekoyamato.co.jp/cgi-bin/tneko?init&number1=' + cleanNumber;
        }

        // 日本郵便判定
        if (m.includes('ゆうゆう') || m.includes('ゆうパケット') || m.includes('郵便') || m.includes('レターパック') || m.includes('定形') || m.includes('クリックポスト') || m.includes('特定記録')) {
            return 'https://trackings.post.japanpost.jp/services/srv/search/direct?searchKind=S004&locale=ja&reqCodeNo1=' + cleanNumber;
        }

        // 判別できない場合は、一般的な桁数から推測（12桁はヤマト・郵便どちらもあり得るが、一旦ヤマトへ）
        if (cleanNumber.length === 12) {
            return 'https://toi.kuronekoyamato.co.jp/cgi-bin/tneko?init&number1=' + cleanNumber;
        }

        return null;
    }

    function renderHistoryCards(tab, data) {
        const container = document.getElementById(`${tab}-history-list`);
        if (!container) return;
        const renderId = (historyRenderIds.get(tab) || 0) + 1;
        historyRenderIds.set(tab, renderId);

        // データ読み込み中
        if (data === undefined) {
            container.innerHTML = `
                <div class="empty-history" style="padding: 40px; text-align: center; color: var(--text-muted);">
                    <ion-icon name="sync-outline" class="spinning" style="font-size: 24px; margin-bottom: 8px;"></ion-icon>
                    <div>履歴データを読み込み中...</div>
                </div>
            `;
            return;
        }

        container.innerHTML = '';

        // サマリーカードの件数バッジを更新
        const countEl = document.getElementById(`${tab}-history-count`);
        const count = (data && data.length) ? data.length : 0;
        if (countEl) countEl.textContent = `${count}件`;

        // 進行中集計の更新
        if (tab === 'manufacturing' || tab === 'sales') {
            const statusEl = document.getElementById(`${tab}-history-status`);
            if (statusEl && data && data.length > 0) {
                const statusCounts = {};
                data.forEach(item => {
                    const s = item['ステータス'] || '不明';
                    statusCounts[s] = (statusCounts[s] || 0) + 1;
                });
                const statusStr = Object.entries(statusCounts).map(([s, n]) => `${s}:${n}件`).join(' / ');
                statusEl.textContent = tab === 'sales' ? `${data.length}件(${statusStr})` : statusStr;
            } else if (statusEl) {
                statusEl.textContent = tab === 'sales' ? '発送・取引待ちはありません' : '進行中の製造はありません';
            }
        }

        const filter = historyFilters.get(tab) || {};
        const renderData = filterHistoryRows(tab, data || [], filter);
        const activeCount = Object.values(filter).filter(Boolean).length;
        const filterCount = document.getElementById(tab + "-history-filter-count");
        const clear = document.querySelector(".history-filter-clear[data-history-tab=" + tab + "]");
        if (filterCount) filterCount.textContent = renderData.length + "件 / 全" + count + "件" + (activeCount ? " ・ 条件" + activeCount + "件" : "");
        if (clear) clear.hidden = !activeCount;
        updateHistoryFilterOptions(tab, data || [], filter);
        const bulk = getHistoryBulkState(tab);
        {
            const visibleIds = new Set(renderData.filter(row => !getHistoryBulkReason(tab, row)).map(row => String(row[bulkStatusSpec_(tab).id])));
            for (const id of bulk.selected) {
                if (!visibleIds.has(id)) bulk.selected.delete(id);
            }
            // 同期・絞り込みによる再描画後は古い確認画面を再利用しない。
            bulk.preview = '';
            const preview = document.querySelector('#' + tab + '-bulk-controls [data-bulk="preview"]');
            if (preview) preview.hidden = true;
            updateHistoryBulkControls(tab);
        }

        if (!data || data.length === 0) {
            container.innerHTML = '<p class="empty-msg">現在進行中の履歴はありません</p>';
            return;
        }

        if (!renderData || renderData.length === 0) {
            container.innerHTML = '<p class="empty-msg">条件に一致する履歴はありません</p>';
            return;
        }

        // --- 描画最適化ロジック ---
        const CHUNK_SIZE = 50; // 1回あたりに描画する件数
        let index = 0;

        function renderChunk() {
            if (historyRenderIds.get(tab) !== renderId) return;
            const fragment = document.createDocumentFragment();
            const limit = Math.min(index + CHUNK_SIZE, renderData.length);

            for (; index < limit; index++) {
                const item = renderData[index];
                const card = createHistoryCardElement(tab, item);
                if (bulk.mode) {
                    const id = String(item[bulkStatusSpec_(tab).id]);
                    const reason = getHistoryBulkReason(tab, item);
                    const label = document.createElement('label');
                    label.className = 'history-bulk-selection';
                    const checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.checked = bulk.selected.has(id);
                    checkbox.disabled = bulk.busy || Boolean(bulk.pending) || Boolean(reason);
                    checkbox.onchange = () => {
                        if (!checkbox.checked) bulk.selected.delete(id);
                        else if (bulk.selected.size < 20) bulk.selected.add(id);
                        else checkbox.checked = false;
                        bulk.preview = '';
                        const preview = document.querySelector('#' + tab + '-bulk-controls [data-bulk="preview"]');
                        if (preview) preview.hidden = true;
                        updateHistoryBulkControls(tab);
                    };
                    label.append(checkbox, reason ? '対象外：' + reason : ' 一括変更の対象にする');
                    card.prepend(label);
                }
                fragment.appendChild(card);
            }

            container.appendChild(fragment);

            if (index < renderData.length) {
                requestAnimationFrame(renderChunk);
            }
        }

        // 最初のチャンクを即座に描画し、残りをアニメーションフレームで行う
        renderChunk();
    }

    /**
     * 履歴カード単体要素の生成
     */
    function createHistoryCardElement(tab, item) {
        const card = document.createElement('div');
        card.className = 'history-card glass-card';

        const idKey = Object.keys(item).find(k => k.endsWith('ID')) || 'ID';
        const id = item[idKey];
        const itemName = item['品名'] || item['完成品名'] || '';
        const price = item['合計金額'] || item['価格'] || item['販売価格'] || item['小計'] || item['売上'] || 0;
        const formattedPrice = typeof price === 'number' ? price.toLocaleString() : price;

        // M_画面制御の設定を取得するヘルパー
        const getFieldSetting = (logicalId) => {
            const masters = currentMasters['M_画面制御'] || [];
            return masters.find(m => m['要素ID'] === logicalId) || {};
        };

        // 入力フィールド生成ヘルパー
        const createInputHtml = (label, header, value, logicalId, typeOverride = null, isReadOnly = false) => {
            const setting = getFieldSetting(logicalId);
            const type = typeOverride || setting['タイプ'] || 'text';

            if (isReadOnly) {
                const displayVal = (type === 'number' && typeof value === 'number') ? value.toLocaleString() : (value || '-');
                return `<div class="input-group mini"><label>${label}</label><div class="static-value">${escapeHtml(displayVal)}</div></div>`;
            }

            // datalist ID の解決 (index.html の定義に合わせる)
            let listId = '';
            if (logicalId === 'buy-vendor') listId = 'supplier-list';
            else if (logicalId === 'buy-payment') listId = 'payment-list';
            else if (logicalId === 'exp-vendor') listId = 'vendor-list';
            else if (logicalId === 'exp-account') listId = 'expense-category-list';
            else if (logicalId === 'exp-payment') listId = 'expense-payment-list';
            else if (setting['要素ID']) listId = setting['要素ID'] + '-list';

            if (type === 'number') {
                // 0 の場合は空にして placeholder で 0 を出すことで 0999 問題を回避
                const valAttr = (value === 0 || value === '0' || !value) ? '' : value;
                return `<div class="input-group mini"><label>${label}</label><input type="number" class="note-input" data-header="${header}" value="${escapeHtml(valAttr)}" placeholder="0" step="any"></div>`;
            } else if (type === 'date') {
                return `<div class="input-group mini"><label>${label}</label><input type="date" class="date-input" data-header="${header}" value="${formatISODate(value)}"></div>`;
            } else {
                const listAttr = listId ? `list="${escapeHtml(listId)}"` : '';
                return `<div class="input-group mini"><label>${label}</label><input type="text" class="note-input" data-header="${header}" value="${escapeHtml(value)}" ${listAttr}></div>`;
            }
        };

        let innerHTML = `
            <div class="history-card-header">
                <div class="header-left">
                    <span class="history-id">${id}</span>
                    ${(tab === 'sales' && item['売先']) ? `<span class="badge badge-buyer">${escapeHtml(item['売先'])}</span>` : ''}
                    ${(tab === 'sales' && item['送料負担区分'] == 1) ? '<span class="badge badge-warning">落札者負担</span>' : ''}
                    ${(tab === 'expense' && item['管理対象'] == 1) ? '<span class="badge badge-info"><ion-icon name="cube-outline"></ion-icon> 在庫対象</span>' : ''}
                    ${(tab === 'expense' && item['レシート'] == 1) ? '<span class="badge badge-secondary"><ion-icon name="receipt-outline"></ion-icon> レシート有</span>' : ''}
                    ${item['ステータス'] === '購入予定' ? '<span class="badge badge-planned">購入予定</span>' : ''}
                    ${(tab === 'manufacturing' && !['完了', 'キャンセル'].includes(item['ステータス'])) ? '<span class="badge badge-info">パーツ引当済</span>' : ''}
                    ${(tab === 'manufacturing' && (item['備考'] || '').includes('追加部材:')) ? '<span class="badge badge-added">部材追加あり</span>' : ''}
                    ${(tab === 'sales' && (item['管理対象外'] == 1 || item['管理区分'] == 1)) ? '<span class="badge badge-personal">個人利用</span>' : ''}
                    ${(tab === 'sales' && item['管理対象外'] != 1 && item['管理区分'] != 1 && !['キャンセル'].includes(item['ステータス'])) ? '<span class="badge badge-info">在庫引当済</span>' : ''}
                </div>
                <div class="header-right" style="display: flex; gap: 8px;">
                    ${(tab === 'manufacturing' && !['完了', 'キャンセル'].includes(item['ステータス'])) ? `
                        <button class="add-material-btn" data-id="${escapeHtml(id)}" title="部材を追加消費する">
                            <ion-icon name="add-circle-outline"></ion-icon>
                            <span>部材追加</span>
                        </button>
                    ` : ''}
                    <button class="update-mini-btn" data-id="${escapeHtml(id)}" data-tab="${tab}" data-item="${escapeHtml(itemName)}"
                        data-category="${escapeHtml(item['区分'] || '')}" data-account="${escapeHtml(item['仕訳'] || '')}"
                        data-current-status="${escapeHtml(item['ステータス'] || '')}"
                        data-quantity="${escapeHtml(item['数量'] || item['製造数量'] || 0)}"
                        data-price="${escapeHtml(item['価格'] || item['合計金額'] || item['販売価格'] || 0)}"
                        data-vendor="${escapeHtml(item['仕入先'] || item['購入先'] || '')}"
                        data-payment="${escapeHtml(item['支払方法'] || '')}"
                        data-buyer="${escapeHtml(item['売先'] || '')}"
                        data-shipping="${escapeHtml(item['発送方法'] || '')}"
                        data-is-personal="${(item['管理対象外'] == 1 || item['管理区分'] == 1) ? 'true' : 'false'}"
                        title="変更を保存">
                        <span>保存</span>
                        <ion-icon name="save-outline"></ion-icon>
                    </button>
                </div>
            </div>
        `;


        const itemInMaster = (currentMasters['M_商品'] || []).find(m => m['品名'] === itemName);
        const imageUrl = safeHttpsUrl(itemInMaster?.['画像URL']);
        const thumbHtml = imageUrl ? `<div class="product-thumb-container" style="width:30px; height:30px; margin-right:8px; border-radius:4px;"><img src="${escapeHtml(imageUrl)}" loading="lazy"></div>` : '';

        if (tab === 'purchase') {
            const isPlanned = (item['ステータス'] === '購入予定');
            const isPending = (isPlanned && (!item['価格'] || !item['数量']));
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">
                    ${thumbHtml}${escapeHtml(itemName)}
                    ${isPending ? '<span class="badge badge-pending" style="margin-left:8px; font-size:10px;">情報未確定</span>' : ''}
                </div>
                <div class="history-inputs-grid">
                    ${createInputHtml('仕入日', '仕入日', item['仕入日'] || item['発注日'], 'buy-date', 'date', false)}
                    ${createInputHtml('入庫日', '入庫日', item['入庫日'], 'buy-date', 'date', false)}
                    <div class="purchase-amounts">
                    ${createInputHtml('価格', '価格', item['価格'], 'buy-price', 'number', !isPlanned)}
                    ${createInputHtml('数量', '数量', item['数量'], 'buy-quantity', 'number', !isPlanned)}
                    ${createInputHtml('単価', '単価', item['単価'], '', 'number', true)}
                    </div>
                    ${!isPlanned && !['入庫済み','入庫済','完了','キャンセル','返品完了','返金済み'].includes(item['ステータス']) ? `
                    <details class="purchase-refund" data-quantity="${escapeHtml(item['数量'])}" data-price="${escapeHtml(item['価格'])}">
                        <summary>不足分の返金を反映する</summary>
                        <div class="history-inputs-grid">
                            <label>実入荷数 <input type="number" min="1" step="1"></label>
                            <label>返金額（円） <input type="number" min="0.01" step="0.01"></label>
                        </div>
                        <output>実入荷数と返金額を入力してください。</output>
                    </details>` : ''}
                    ${createInputHtml('仕入先', '仕入先', item['仕入先'], 'buy-vendor', null, !isPlanned)}
                    ${createInputHtml('支払方法', '支払方法', item['支払方法'], 'buy-payment', null, !isPlanned)}
                    ${generateCategorySelect('区分', '仕入', item['区分'], !isPlanned)}
                    <div class="input-group mini">${generateStatusSelect(id, '仕入', item['ステータス'])}</div>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.1); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${escapeHtml(item['備考'] || item['note'] || '')}</textarea>
                </div>
            `;
        } else if (tab === 'expense') {
            const isPlanned = (item['ステータス'] === '購入予定');
            const isPending = (isPlanned && (!item['合計金額'] || !item['数量']));
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">
                    ${thumbHtml}${escapeHtml(itemName)}
                    ${isPending ? '<span class="badge badge-pending" style="margin-left:8px; font-size:10px;">情報未確定</span>' : ''}
                </div>
                <div class="history-inputs-grid">
                    ${createInputHtml('注文日', '注文日', item['注文日'] || item['登録日'], 'expense-date', 'date', false)}
                    ${createInputHtml('完了日', '完了日', item['完了日'], 'expense-date', 'date', false)}
                    ${createInputHtml('金額', '合計金額', item['合計金額'], 'exp-price', 'number', !isPlanned)}
                    ${createInputHtml('数量', '数量', item['数量'], 'exp-quantity', 'number', !isPlanned)}
                    ${createInputHtml('購入先', '購入先', item['購入先'], 'exp-vendor', null, !isPlanned)}
                    ${createInputHtml('仕訳', '仕訳', item['仕訳'], 'exp-account', null, !isPlanned)}
                    ${createInputHtml('支払方法', '支払方法', item['支払方法'], 'exp-payment', null, !isPlanned)}
                    <div class="input-group mini">${generateStatusSelect(id, '経費', item['ステータス'])}</div>
                </div>
                <div class="history-inputs-grid" style="grid-template-columns: 1fr; margin-top: 5px;">
                    <label style="display: flex; align-items: center; gap: 8px; font-size: 0.9em; color: var(--text-secondary); cursor: pointer;">
                        <input type="checkbox" class="chk-input" data-header="レシート" ${item['レシート'] == 1 ? 'checked' : ''} style="width:auto; transform:scale(1.1);">
                        レシート提出済みにする
                    </label>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.1); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${escapeHtml(item['備考'] || item['note'] || '')}</textarea>
                </div>
            `;
        } else if (tab === 'manufacturing') {
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}<button type="button" class="history-detail-link">${escapeHtml(itemName)}</button> 数量:${escapeHtml(item['数量'] || item['製造数量'])} 単価:${escapeHtml(item['単価'] || 0)}</div>
                <div class="process-steps-container">
                    <label class="group-label">工程進捗管理</label>
                    <div class="process-dates-grid">
                        <div class="date-input-mini"><label>開始</label><input type="date" class="date-input" data-header="製造開始日" value="${formatISODate(item['製造開始日'])}"></div>
                        <div class="date-input-mini"><label>着手</label><input type="date" class="date-input" data-header="製造着手日" value="${formatISODate(item['製造着手日'])}"></div>
                        <div class="date-input-mini"><label>テスト</label><input type="date" class="date-input" data-header="テスト開始日" value="${formatISODate(item['テスト開始日'])}"></div>
                        <div class="date-input-mini"><label>梱包</label><input type="date" class="date-input" data-header="梱包開始日" value="${formatISODate(item['梱包開始日'])}"></div>
                        <div class="date-input-mini"><label>完了</label><input type="date" class="date-input" data-header="製造完了日" value="${formatISODate(item['製造完了日'])}"></div>
                        <div class="date-input-mini">${generateStatusSelect(id, '製造', item['ステータス'])}</div>
                    </div>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.12); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${escapeHtml(item['備考'] || item['note'] || item['メモ'] || '')}</textarea>
                </div>
            `;
        } else if (tab === 'sales') {
            const unitPrice = item['合計単価'] || item['単価'] || 0;
            const totalPrice = item['販売価格'] || item['価格'] || item['合計金額'] || 0;
            const formattedUnit = typeof unitPrice === 'number' ? unitPrice.toLocaleString() : unitPrice;
            const formattedTotal = typeof totalPrice === 'number' ? totalPrice.toLocaleString() : totalPrice;
            const trackingUrl = getTrackingUrl(item['発送方法'], item['追跡番号']);
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}${escapeHtml(itemName)} 数量:${escapeHtml(item['数量'])} 価格:¥${escapeHtml(formattedTotal)}(単価:¥${escapeHtml(formattedUnit)}) 送料:¥${escapeHtml((item['送料'] || 0).toLocaleString())}</div>
                <div class="history-inputs-grid" style="margin-bottom: 8px;">
                    ${createInputHtml('販売開始日', '販売開始日', item['販売開始日'], 'sales-date', 'date', false)}
                    <div class="input-group mini"><label>発送方法</label><div class="static-value">${escapeHtml(item['発送方法'] || '-')}</div></div>
                </div>
                ${['入金待ち', '取引開始', '受注', '開始'].includes(String(item['ステータス'] || '').trim()) && !item['発送日'] && !item['受取日'] && !item['取引完了日'] ? '<button type="button" class="shipping-change-open history-detail-link">発送方法を変更</button>' : ''}
                <div class="process-steps-container">
                    <label class="group-label">取引工程日付</label>
                    <div class="process-dates-grid" style="grid-template-columns: repeat(2, 1fr);">
                        <div class="date-input-mini"><label>発送日</label><input type="date" class="date-input" data-header="発送日" value="${formatISODate(item['発送日'])}"></div>
                        <div class="date-input-mini"><label>受取日</label><input type="date" class="date-input" data-header="受取日" value="${formatISODate(item['受取日'])}"></div>
                        <div class="date-input-mini"><label>完了</label><input type="date" class="date-input" data-header="取引完了日" value="${formatISODate(item['取引完了日'])}"></div>
                        <div class="date-input-mini">${generateStatusSelect(id, '販売', item['ステータス'])}</div>
                    </div>
                </div>
                <div class="history-inputs-grid" style="grid-template-columns: 1fr; margin-top: 8px;">
                    <div class="input-group mini">
                        <label>追跡番号</label>
                        <div style="display: flex; gap: 4px; align-items: center;">
                            <input type="text" class="note-input" data-header="追跡番号" value="${escapeHtml(item['追跡番号'] || '')}" placeholder="追跡番号" style="flex: 1;">
                            ${trackingUrl ? `<a href="${trackingUrl}" target="_blank" class="mini-icon-link" title="配送状況を確認"><ion-icon name="navigate-circle-outline"></ion-icon></a>` : ''}
                        </div>
                    </div>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.12); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${escapeHtml(item['備考'] || item['note'] || '')}</textarea>
                </div>
            `;
        }

        card.innerHTML = innerHTML;
        const refund = card.querySelector('.purchase-refund');
        refund?.addEventListener('input', () => {
            const values = readPurchaseRefund(refund);
            refund.querySelector('output').textContent = values
                ? `返金後の価格 ${(values.cents / 100).toFixed(2)}円 ／ 入庫単価 ${(Math.round(values.cents / values.quantity) / 100).toFixed(2)}円`
                : '実入荷数と返金額を確認してください。';
        });

        const detailLink = card.querySelector('.history-detail-link');
        if (tab === 'manufacturing' && detailLink) detailLink.addEventListener('click', () => openTransactionDetail_('manufacturing', id, detailLink));
        card.querySelector('.shipping-change-open')?.addEventListener('click', () => openShippingChange_(String(id)));

        // クリアボタンの適用
        const textInputs = card.querySelectorAll('input[type="text"], input[type="number"]');
        textInputs.forEach(input => addClearButton(input));

        // 保存ボタンにイベントをバインド
        const saveBtn = card.querySelector('.update-mini-btn');
        if (saveBtn) {
            saveBtn.addEventListener('click', () => handleHistorySave(saveBtn));
        }

        // 部材追加ボタンにイベントをバインド (Proposal 15)
        const addMatBtn = card.querySelector('.add-material-btn');
        if (addMatBtn) {
            addMatBtn.addEventListener('click', () => showAddMaterialModal(id));
        }

        return card;
    }

    /**
     * 部材追加モーダルの表示 (Proposal 15)
     */
    let currentManufacturingIdForAdd = null;
    function showAddMaterialModal(id) {
        currentManufacturingIdForAdd = id;
        const modal = document.getElementById('add-material-modal');
        if (!modal) return;

        // 初期化
        document.getElementById('add-material-item').value = '';
        document.getElementById('add-material-quantity').value = '1';
        document.getElementById('add-material-reason').value = '';

        modal.classList.add('active');
    }

    async function handleMaterialSubmission() {
        const id = currentManufacturingIdForAdd;
        const item = document.getElementById('add-material-item').value.trim();
        const quantity = parseFloat(document.getElementById('add-material-quantity').value);
        const reason = document.getElementById('add-material-reason').value.trim();

        if (!item) return alert("部材名を入力してください。");
        if (isNaN(quantity) || quantity <= 0) return alert("有効な数量を入力してください。");

        if (!confirm(`${item} を ${quantity} 個、追加で引き当てます。よろしいですか？\n(この操作により製造単価が再計算されます)`)) return;

        const btn = document.getElementById('add-material-submit');
        const originalHTML = btn.innerHTML;
        btn.innerHTML = '<ion-icon name="sync-outline" class="spinning"></ion-icon> 実行中...';
        btn.disabled = true;

        try {
            const res = await fetchAPI('addMaterialToManufacturing', {
                manufacturingId: id,
                item: item,
                quantity: quantity,
                reason: reason
            });

            if (res.status === 'success') {
                showToast("部材を追加し、単価を再計算しました。");
                document.getElementById('add-material-modal').classList.remove('active');

                // 履歴データを更新して再描画
                if (res.historyData && res.historyData.rawData) {
                    lastRawData = Object.assign({}, lastRawData, res.historyData.rawData);
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);
                } else {
                    await fetchHistory('manufacturing');
                }
            } else {
                throw new Error(res.message);
            }
        } catch (e) {
            alert("エラー: " + e.message);
        } finally {
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        }
    }

    function generateStatusSelect(id, actionName, currentStatus) {
        // M_画面制御から該当機能の履歴用ステータス設定を探す
        const screenCtrl = (currentMasters['M_画面制御'] || []).find(c =>
            c['対象機能'] === actionName &&
            c['要素ID'].endsWith('-status-') &&
            c['参照マスタ'] === 'M_ステータス'
        );
        const targetScreen = screenCtrl ? screenCtrl['画面名称'] : '';

        const statuses = (currentMasters['M_ステータス'] || [])
            .filter(s => s['対象機能'] === actionName && s['画面名称'] === targetScreen && parseInt(s['使用FLG']) === 1)
            .sort((a, b) => (a['表示順'] || 0) - (b['表示順'] || 0));

        let options = statuses.map(s => {
            const selected = s['ステータス名称'] === currentStatus ? 'selected' : '';
            return `<option value="${escapeHtml(s['ステータス名称'])}" ${selected}>${escapeHtml(s['ステータス名称'])}</option>`;
        }).join('');

        return `<label>ステータス</label><select class="status-select">${options}</select>`;
    }

    /**
     * 区分（M_区分）のプルダウンを生成する
     */
    function generateCategorySelect(label, actionName, currentValue, isReadOnly) {
        if (isReadOnly) {
            return `<div class="input-group mini"><label>${label}</label><div class="static-value">${escapeHtml(currentValue || '-')}</div></div>`;
        }

        // M_画面制御から設定を取得を試みる
        const masters = currentMasters['M_画面制御'] || [];
        const ctrl = masters.find(m =>
            (m['画面名称'] === actionName || m['対象機能'] === actionName) &&
            (m['画面上の項目名'] === label || m['項目名'] === label)
        );

        let options = '<option value="">選択してください</option>';

        if (ctrl && (ctrl['タイプ'] === 'fixed' || ctrl['型'] === 'fixed')) {
            const fixedStr = ctrl['固定値'] || ctrl['固定値の内容'] || '';
            const fixedValues = fixedStr.split(',').map(v => v.trim()).filter(v => v);
            options += fixedValues.map(val => {
                const selected = val === currentValue ? 'selected' : '';
                return `<option value="${escapeHtml(val)}" ${selected}>${escapeHtml(val)}</option>`;
            }).join('');
        } else {
            const refMasterName = ctrl ? ctrl['参照マスタ'] : 'M_区分';
            const masterData = currentMasters[refMasterName] || currentMasters['M_カテゴリ'] || currentMasters['M_区分'] || [];

            const categories = masterData
                .filter(c => {
                    // 使用FLGが1のもののみ
                    if (parseInt(c['使用FLG']) === 0) return false;
                    // 対象機能列があれば絞り込み、なければすべて通す
                    if (c['対象機能'] && c['対象機能'] !== actionName) return false;
                    return true;
                })
                .sort((a, b) => (a['表示順'] || 0) - (b['表示順'] || 0));

            options += categories.map(c => {
                // カラム名の揺れに対応 (区分名 / 区分名称 / 項目名 / カテゴリ)
                const val = c['区分名'] || c['区分名称'] || c['項目名'] || c['仕訳名'] || c['カテゴリ'] || Object.values(c)[1];
                if (!val) return '';
                const selected = val === currentValue ? 'selected' : '';
                return `<option value="${escapeHtml(val)}" ${selected}>${escapeHtml(val)}</option>`;
            }).join('');
        }

        return `<div class="input-group mini"><label>${label}</label><select class="note-input" data-header="${label}">${options}</select></div>`;
    }

    function readPurchaseRefund(refund) {
        const [quantityInput, amountInput] = refund.querySelectorAll('input');
        const quantity = Number(quantityInput.value);
        const amount = Number(amountInput.value);
        const originalQuantity = Number(refund.dataset.quantity);
        const originalPrice = Number(refund.dataset.price);
        const cents = Math.round(originalPrice * 100) - Math.round(amount * 100);
        if (!quantityInput.checkValidity() || !amountInput.checkValidity() ||
            !Number.isInteger(quantity) || quantity < 1 || !(quantity < originalQuantity) ||
            !(amount > 0) || !(cents > 0)) return null;
        return { quantity, amount, originalQuantity, originalPrice, cents };
    }

    async function handleHistorySave(btn) {
        const originalHTML = btn.innerHTML;
        const id = btn.getAttribute('data-id');
        const card = btn.closest('.history-card');
        const statusSelect = card.querySelector('.status-select');
        const newStatus = statusSelect ? statusSelect.value : '';
        const oldStatus = btn.getAttribute('data-current-status') || '';

        // ステータス変更時の確認 (提案33)
        if (newStatus !== oldStatus) {
            const criticalStatuses = ['完了', 'キャンセル', '入庫済み', '発送済み'];
            if (criticalStatuses.includes(newStatus)) {
                if (!confirm(`ステータスを「${newStatus}」に変更してもよろしいですか？`)) {
                    return;
                }
            }
        }

        // 追加: カード内の日付入力も取得
        const dateUpdates = {};
        card.querySelectorAll('.date-input').forEach(input => {
            const header = input.getAttribute('data-header');
            if (header) {
                dateUpdates[header] = input.value;
            }
        });

        // 備考・追跡番号等のテキスト入力の取得
        card.querySelectorAll('.note-input').forEach(input => {
            const header = input.getAttribute('data-header');
            if (header) {
                let val = input.value;
                // 数値項目は数値化
                if (['価格', '数量', '合計金額'].includes(header)) {
                    val = parseFloat(val) || 0;
                }
                dateUpdates[header] = val;
            }
        });

        // 静的表示テキストも取得（非編集モード時の値を保持するため）
        card.querySelectorAll('.static-value').forEach(div => {
            const label = div.previousElementSibling;
            if (label && label.tagName === 'LABEL') {
                const header = label.textContent;
                if (header && dateUpdates[header] === undefined) {
                    let val = div.textContent;
                    if (val === '-') val = '';
                    // 通貨記号やカンマを除去して数値化
                    if (['価格', '数量', '合計金額'].includes(header)) {
                        val = parseFloat(val.replace(/[¥,]/g, '')) || 0;
                    }
                    dateUpdates[header] = val;
                }
            }
        });

        const refund = card.querySelector('.purchase-refund[open]');
        const purchaseRefund = refund ? readPurchaseRefund(refund) : null;
        if (refund && (!purchaseRefund || newStatus !== '入庫済み')) {
            alert('実入荷数と返金額を確認し、入庫済みを選んでください。');
            return;
        }

        // 保存前のバリデーション (提案30)
        const currentTab = btn.getAttribute('data-tab');
        const sheetNameMap = {
            'purchase': 'T_仕入',
            'expense': 'T_経費',
            'manufacturing': 'T_製造',
            'sales': 'T_販売'
        };
        const sheetName = sheetNameMap[currentTab];

        if (sheetName) {
            const isPlanned = (newStatus === '購入予定');
            // 基礎データの構築
            const finalData = {
                status: newStatus,
                item: btn.getAttribute('data-item'),
                ...dateUpdates,
                note: dateUpdates['備考'],
                receipt: dateUpdates['レシート']
            };

            // シートごとの内部キーマッピング (validateData 用)
            if (currentTab === 'purchase') {
                finalData.date = dateUpdates['仕入日'] || dateUpdates['発注日'];
                finalData.vendor = dateUpdates['仕入先'];
                finalData.price = (isPlanned && (dateUpdates['価格'] === '' || dateUpdates['価格'] === undefined)) ? 0 : (parseFloat(dateUpdates['価格']) || 0);
                finalData.quantity = (isPlanned && (dateUpdates['数量'] === '' || dateUpdates['数量'] === undefined)) ? 1 : (parseFloat(dateUpdates['数量']) || 0);
                if (purchaseRefund) {
                    finalData.price = purchaseRefund.cents / 100;
                    finalData.quantity = purchaseRefund.quantity;
                }
                finalData.payment = dateUpdates['支払方法'];
                finalData.category = dateUpdates['区分'] || btn.getAttribute('data-category');
            } else if (currentTab === 'expense') {
                finalData.date = dateUpdates['注文日'] || dateUpdates['登録日'];
                finalData.vendor = dateUpdates['購入先'];
                finalData.price = (isPlanned && (dateUpdates['合計金額'] === '' || dateUpdates['合計金額'] === undefined)) ? 0 : (parseFloat(dateUpdates['合計金額']) || 0);
                finalData.quantity = (isPlanned && (dateUpdates['数量'] === '' || dateUpdates['数量'] === undefined)) ? 1 : (parseFloat(dateUpdates['数量']) || 0);
                finalData.payment = dateUpdates['支払方法'];
                finalData.account = dateUpdates['仕訳'] || btn.getAttribute('data-account');
            } else if (currentTab === 'manufacturing') {
                finalData.date = dateUpdates['製造開始日'];
                finalData.quantity = parseFloat(btn.getAttribute('data-quantity')) || 0;
            } else if (currentTab === 'sales') {
                finalData.date = dateUpdates['販売開始日'];
                finalData.buyer = dateUpdates['売先'] || btn.getAttribute('data-buyer');
                finalData.shipping = dateUpdates['発送方法'] || btn.getAttribute('data-shipping');
                finalData.shippingCost = parseFloat(dateUpdates['送料']) || 0;
                finalData.quantity = parseFloat(btn.getAttribute('data-quantity')) || 0;
                finalData.price = parseFloat(btn.getAttribute('data-price')) || 0;
                finalData.type = btn.getAttribute('data-is-personal') === 'true' ? 'personal' : 'business';
            }

            const error = validateData(sheetName, finalData, true);
            if (error) {
                alert(error);
                btn.innerHTML = originalHTML;
                btn.disabled = false;
                setLoading(false);
                return;
            }
        }

        // チェックボックスの取得
        card.querySelectorAll('.chk-input').forEach(chk => {
            const header = chk.getAttribute('data-header');
            if (header) {
                dateUpdates[header] = chk.checked ? 1 : 0;
            }
        });

        btn.innerHTML = '<ion-icon name="sync-outline" class="spinning"></ion-icon>';
        btn.disabled = true;
        setLoading(true, 'データを更新中...');

        const refreshScope = id.startsWith('A') ? 'purchase' :
            id.startsWith('E') ? 'expense' :
                id.startsWith('M') ? 'manufacturing' :
                    id.startsWith('S') ? 'sales' : 'all';

        try {
            const response = await fetchAPI('updateTransaction', {
                id: id,
                updates: {
                    status: newStatus,
                    dates: dateUpdates,
                    ...(purchaseRefund ? { purchaseRefund } : {})
                },
                scope: refreshScope
            });

            if (response.status === 'success') {
                btn.innerHTML = '<ion-icon name="checkmark-outline"></ion-icon>';
                showToast(`${id} の更新が完了しました`);

                // 応答に同梱された差分データで即時更新
                if (response.newRecord && response.sheetName) {
                    console.time('Client:IncrementalUpdate(Update)');

                    // 1. 該当レコードの生データを特定して置換
                    const sheetData = lastRawData[response.sheetName] || [];
                    const idColIdx = 0; // 常に1列目がID
                    const targetId = id;
                    const rowIndex = sheetData.findIndex(r => r[idColIdx] == targetId);

                    if (rowIndex !== -1) {
                        sheetData[rowIndex] = response.newRecord;
                    } else {
                        sheetData.push(response.newRecord);
                    }

                    // 2. 在庫集計が同梱されている場合は更新
                    if (response.inventorySummary) {
                        currentMasters['T_在庫集計'] = response.inventorySummary;
                    }

                    // 3. 全データを再集計・再描画
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);

                    console.timeEnd('Client:IncrementalUpdate(Update)');
                } else if ((response.data && response.data.historyData) || response.historyData) {
                    // 履歴データが含まれている場合の処理
                    const hData = response.historyData || response.data.historyData;
                    lastRawData = Object.assign({}, lastRawData, hData.rawData);
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    renderAllHistory(processed);
                    refreshInventoryUI(); // 在庫一覧も更新
                } else {
                    // フォールバック（少し待ってから再読み込み）
                    setTimeout(() => fetchHistory(refreshScope), 1000);
                }
            } else {
                throw new Error(response.message);
            }
        } catch (e) {
            console.error(e);
            showToast("更新失敗: " + e.message, 'error');
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        } finally {
            setLoading(false);
        }
    }

    async function updateStatus(btn, id, newStatus, currentItem) {
        // ... (existing updateStatus body, keeping it same but just for context if needed)
    }

    /**
     * Ledger (帳簿) タブのレンダリング
     */
    function renderLedger() {
        if (!lastHistoryData || !lastHistoryData.ledger) return;

        // 共通フィルタの描画/更新
        updateLedgerPeriodFilters(lastHistoryData.ledger);

        // 在庫アラートの表示
        renderInventoryAlerts();

        const ledger = lastHistoryData.ledger;
        const personalSalesByMonth = ledgerIncludePersonal ? (lastHistoryData.personalSalesByMonth || {}) : {};
        const recentActions = lastHistoryData.recentActions || [];

        // 選択されている期間文字列 (例: "2026-04", "2026-FY", "2026-Q1")
        const selectedMonth = `${ledgerYear}-${ledgerPeriod}`;

        // 1. 集計用ヘルパー関数
        function getTotalsFor(targetY, targetP) {
            const res = { sales: 0, cost: 0, profit: 0, personalSales: 0 };
            ledger.forEach(row => {
                const d = new Date(row['日付']);
                if (isNaN(d.getTime())) return;
                const y = d.getFullYear();
                const m = d.getMonth() + 1;

                let match = false;
                if (String(y) === String(targetY)) {
                    if (targetP === 'FY') match = true;
                    else if (targetP === 'Q1') match = (m >= 1 && m <= 3);
                    else if (targetP === 'Q2') match = (m >= 4 && m <= 6);
                    else if (targetP === 'Q3') match = (m >= 7 && m <= 9);
                    else if (targetP === 'Q4') match = (m >= 10 && m <= 12);
                    else if (String(m).padStart(2, '0') === targetP) match = true;
                }
                if (match) {
                    const isPersonal = String(row['事業/個人']).trim() === '1';
                    if (isPersonal) return; // 事業用集計のため個人はスキップ
                    const s = parseNumber(row['売上']);
                    const cost = parseNumber(row['仕入']) + parseNumber(row['通信費']) + parseNumber(row['修繕費']) +
                        parseNumber(row['消耗品費']) + parseNumber(row['諸会費']) + parseNumber(row['支払手数料']) + parseNumber(row['雑費']);
                    res.sales += s;
                    res.cost += cost;
                    res.profit += (s - cost);
                }
            });

            // 個人用売上の集計
            for (const ym in personalSalesByMonth) {
                const [y, m] = ym.split('-');
                if (String(y) === String(targetY)) {
                    let match = false;
                    if (targetP === 'FY') match = true;
                    else if (targetP === 'Q1') match = (parseInt(m) >= 1 && parseInt(m) <= 3);
                    else if (targetP === 'Q2') match = (parseInt(m) >= 4 && parseInt(m) <= 6);
                    else if (targetP === 'Q3') match = (parseInt(m) >= 7 && parseInt(m) <= 9);
                    else if (targetP === 'Q4') match = (parseInt(m) >= 10 && parseInt(m) <= 12);
                    else if (m === targetP) match = true;

                    if (match) res.personalSales += personalSalesByMonth[ym];
                }
            }
            return res;
        }

        // ホームの月選択だけを集計カードに適用。グラフと月次明細の期間は維持する。
        const [summaryYear, summaryPeriod] = dashboardSelectedMonth
            ? dashboardSelectedMonth.split('-') : [ledgerYear, ledgerPeriod];
        // 現在・前月・前年の特定
        const current = getTotalsFor(summaryYear, summaryPeriod);

        let prevY = summaryYear, prevP = summaryPeriod;
        if (summaryPeriod === 'FY') { prevY = parseInt(summaryYear) - 1; prevP = 'FY'; }
        else if (summaryPeriod === 'Q1') { prevY = parseInt(summaryYear) - 1; prevP = 'Q4'; }
        else if (summaryPeriod.startsWith('Q')) { prevP = 'Q' + (parseInt(summaryPeriod[1]) - 1); }
        else {
            let m = parseInt(summaryPeriod);
            if (m === 1) { prevY = parseInt(summaryYear) - 1; prevP = '12'; }
            else { prevP = String(m - 1).padStart(2, '0'); }
        }
        const prev = getTotalsFor(prevY, prevP);
        const lastYear = getTotalsFor(parseInt(summaryYear) - 1, summaryPeriod);

        // 成長率計算
        function calcRate(curr, base) {
            if (!base || base === 0) return null;
            return ((curr - base) / Math.abs(base)) * 100;
        }

        function getGrowthHTML(rate, label) {
            if (rate === null) return `<span class="growth-badge"><span class="growth-label">${label}:</span><span class="growth-val">---</span></span>`;
            const isUp = rate > 0;
            const icon = isUp ? '↑' : (rate < 0 ? '↓' : '');
            const cls = isUp ? 'up' : (rate < 0 ? 'down' : '');
            return `<span class="growth-badge ${cls}"><span class="growth-label">${label}:</span><span class="growth-val">${icon}${Math.abs(rate).toFixed(1)}%</span></span>`;
        }

        // 2. DOM反映：サマリーカード
        const summaryCards = document.querySelector('.ledger-top-summary');
        if (summaryCards) {
            const isFY = (summaryPeriod === 'FY');
            const prevLabel = isFY ? '前年比' : (summaryPeriod.startsWith('Q') ? '前期比' : '前月比');
            const yearLabel = '前年比';

            const salesPrevHTML = getGrowthHTML(calcRate(current.sales, prev.sales), prevLabel);
            const salesYearHTML = isFY ? '' : getGrowthHTML(calcRate(current.sales, lastYear.sales), yearLabel);

            const profitPrevHTML = getGrowthHTML(calcRate(current.profit, prev.profit), prevLabel);
            const profitYearHTML = isFY ? '' : getGrowthHTML(calcRate(current.profit, lastYear.profit), yearLabel);

            const pSales = current.personalSales || 0;
            const salesTotalHTML = `¥${Math.round(current.sales).toLocaleString()}${pSales > 0 ? `<br><span style="font-size:12px;color:var(--text-muted);">(¥${Math.round(current.sales + pSales).toLocaleString()})</span>` : ''}`;
            const profitTotalHTML = `¥${Math.round(current.profit).toLocaleString()}${pSales > 0 ? `<br><span style="font-size:12px;color:var(--text-muted);">(¥${Math.round(current.profit + pSales).toLocaleString()})</span>` : ''}`;

            summaryCards.innerHTML = `
                <div class="dashboard-summary-period" aria-live="polite">集計対象：${summaryYear}年${isFY ? '通期' : (summaryPeriod.startsWith('Q') ? '第' + summaryPeriod[1] + '四半期' : Number(summaryPeriod) + '月')}</div>
                <div class="mini-summary-card income">
                    <label>売上合計</label>
                    <div class="val">${salesTotalHTML}</div>
                    <div class="growth-container">${salesPrevHTML}${salesYearHTML}</div>
                </div>
                <div class="mini-summary-card expense">
                    <label>経費合計</label>
                    <div class="val">¥${Math.round(current.cost).toLocaleString()}</div>
                </div>
                <div class="mini-summary-card profit">
                    <label>純利益</label>
                    <div class="val">${profitTotalHTML}</div>
                    <div class="growth-container">${profitPrevHTML}${profitYearHTML}</div>
                </div>
            `;
        }

        // 3. グラフ用の月次データ集計
        const monthlySum = {};
        ledger.forEach(row => {
            const d = new Date(row['日付']);
            if (isNaN(d.getTime())) return;
            const ym = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            if (!monthlySum[ym]) monthlySum[ym] = { sales: 0, cost: 0 };
            const isPersonal = String(row['事業/個人']).trim() === '1';
            const s = parseNumber(row['売上']);
            const cost = parseNumber(row['仕入']) + parseNumber(row['通信費']) + parseNumber(row['修繕費']) +
                parseNumber(row['消耗品費']) + parseNumber(row['諸会費']) + parseNumber(row['支払手数料']) + parseNumber(row['雑費']);
            if (!isPersonal) {
                // 事業のみグラフの売上・費用に計上（個人売上はpersonalSalesByMonthから積み上げ描画）
                monthlySum[ym].sales += s;
                monthlySum[ym].cost += cost;
            }
        });

        // フィルタリング対象の明細（月次明細用）
        const filteredLedger = [];
        ledger.forEach(row => {
            const isPersonal = String(row['事業/個人']).trim() === '1';
            // 個人チェックOFF（デフォルト）の場合は個人レコードを除外
            if (!ledgerIncludePersonal && isPersonal) return;

            const d = new Date(row['日付']);
            if (isNaN(d.getTime())) return;
            const y = d.getFullYear();
            const m = d.getMonth() + 1;
            if (String(y) === String(ledgerYear)) {
                let match = false;
                if (ledgerPeriod === 'FY') match = true;
                else if (ledgerPeriod === 'Q1') match = (m >= 1 && m <= 3);
                else if (ledgerPeriod === 'Q2') match = (m >= 4 && m <= 6);
                else if (ledgerPeriod === 'Q3') match = (m >= 7 && m <= 9);
                else if (ledgerPeriod === 'Q4') match = (m >= 10 && m <= 12);
                else if (String(m).padStart(2, '0') === ledgerPeriod) match = true;
                if (match) filteredLedger.push(row);
            }
        });

        // グラフ描画：トレンドの決定
        const chartBars = document.querySelector('.chart-bars');
        const legend = document.querySelector('.chart-legend');
        legend?.querySelector('.dot.personal')?.closest('.legend-item')?.remove();
        if (chartBars && Object.keys(monthlySum).length > 0) {
            const allMonths = Object.keys(monthlySum).sort();
            let monthsToShow = [];

            if (ledgerPeriod === 'FY') {
                monthsToShow = allMonths.filter(k => k.startsWith(ledgerYear));
            } else if (ledgerPeriod.startsWith('Q')) {
                const q = parseInt(ledgerPeriod[1]);
                const startM = (q - 1) * 3 + 1;
                const endM = q * 3;
                monthsToShow = allMonths.filter(k => {
                    const [y, m] = k.split('-').map(Number);
                    return y === parseInt(ledgerYear) && m >= startM && m <= endM;
                });
            } else {
                // 月次選択時：直近6ヶ月トレンド表示
                const targetYM = `${ledgerYear}-${ledgerPeriod}`;
                const idx = allMonths.indexOf(targetYM);
                if (idx !== -1) {
                    monthsToShow = allMonths.slice(Math.max(0, idx - 5), idx + 1);
                } else {
                    // データがない場合は直近6ヶ月
                    monthsToShow = allMonths.slice(-6);
                }
            }

            // 個人用売上を含めたmaxValの計算
            const maxVal = Math.max(...monthsToShow.map(m => {
                const totalSales = monthlySum[m].sales + (personalSalesByMonth[m] || 0);
                return Math.max(totalSales, monthlySum[m].cost);
            }), 10000);

            const yAxis = document.querySelector('.chart-y-axis');
            if (yAxis) {
                const spans = yAxis.querySelectorAll('span');
                if (spans.length >= 3) {
                    const topVal = maxVal >= 1000 ? (maxVal / 1000).toFixed(0) + 'k' : maxVal.toString();
                    const midVal = (maxVal / 2) >= 1000 ? (maxVal / 2000).toFixed(0) + 'k' : (maxVal / 2).toString();
                    spans[0].textContent = topVal;
                    spans[1].textContent = midVal;
                    spans[2].textContent = '0';
                }
            }

            chartBars.innerHTML = monthsToShow.map(m => {
                const businessSales = monthlySum[m].sales;
                const personalSales = personalSalesByMonth[m] || 0;
                const totalSales = businessSales + personalSales;
                const totalSalesHeight = (totalSales / maxVal) * 100;
                const businessPct = totalSales > 0 ? (businessSales / totalSales) * 100 : 0;
                const personalPct = totalSales > 0 ? (personalSales / totalSales) * 100 : 0;
                const businessHeight = (businessSales / maxVal) * 100;
                const cHeight = (monthlySum[m].cost / maxVal) * 100;
                const label = m.split('-')[1] + '月';
                const selected = m === (dashboardSelectedMonth || selectedMonth);
                const groupAttrs = `class="bar-group dashboard-month-option ${selected ? 'is-selected' : ''}" data-month="${m}" role="button" tabindex="0" aria-label="${m.split('-')[0]}年${Number(m.split('-')[1])}月の集計を表示" aria-pressed="${selected}"`;
                // バグ報告11修正: バー2本構成（1本目: 売上+個人積み上げ、2本目: 費用）
                if (personalSales > 0) {
                    return `
                    <div ${groupAttrs}>
                        <div class="bar-stack" style="height: ${totalSalesHeight}%;">
                            <div class="bar-segment income" style="height: ${businessPct}%;"></div>
                            <div class="bar-segment personal" style="height: ${personalPct}%;"></div>
                        </div>
                        <div class="bar expense" style="height: ${cHeight}%;"></div>
                        <span class="bar-label" style="font-size:0.75em;">${label}</span>
                    </div>
                `;
                } else {
                    return `
                    <div ${groupAttrs}>
                        <div class="bar income" style="height: ${businessHeight}%;"></div>
                        <div class="bar expense" style="height: ${cHeight}%;"></div>
                        <span class="bar-label" style="font-size:0.75em;">${label}</span>
                    </div>
                `;
                }
            }).join('');

            chartBars.querySelectorAll('[data-month]').forEach(group => {
                const selectMonth = () => {
                    if (!monthsToShow.includes(group.dataset.month)) return;
                    dashboardSelectedMonth = group.dataset.month;
                    renderLedger();
                };
                group.onclick = selectMonth;
                group.onkeydown = event => {
                    if (event.key === 'Enter' || event.key === ' ') {
                        event.preventDefault();
                        const month = group.dataset.month;
                        selectMonth();
                        chartBars.querySelector('[data-month="' + month + '"]')?.focus();
                    }
                };
            });

            // 凡例に「個人」を追加（個人用データがある場合）
            const hasPersonalData = monthsToShow.some(m => personalSalesByMonth[m] > 0);
            if (legend && hasPersonalData) {
                const personalLegend = document.createElement('div');
                personalLegend.className = 'legend-item';
                personalLegend.innerHTML = '<div class="dot personal"></div>個人';
                legend.appendChild(personalLegend);
            }
        } else if (chartBars) {
            chartBars.innerHTML = ''; // データが空になった際に古い月選択を残さない
        }

        // 最新取引履歴 (Dashboard専用)
        const recentDashboard = document.getElementById('ledger-dashboard-view');
        if (recentDashboard && recentDashboard.style.display !== 'none') {
            const listCard = recentDashboard.querySelector('.ledger-list');
            if (listCard) {
                listCard.querySelectorAll('.transaction-item, .empty-history').forEach(el => el.remove());
                const reportBtn = listCard.querySelector('button');
                recentActions.forEach(action => {
                    const itemDiv = document.createElement('div');
                    itemDiv.className = 'transaction-item';
                    const isIncome = action.type === 'sales';
                    const isExpense = ['purchase', 'expense'].includes(action.type);

                    // CSSのクラス名に合わせる (income, purchase, expense, manufacturing)
                    const colorClass = action.type === 'sales' ? 'income' : action.type;

                    const symbol = isIncome ? '+' : (isExpense ? '-' : '');
                    const iconName = action.type === 'sales' ? 'pricetags-outline' :
                        (action.type === 'purchase' ? 'cart-outline' :
                            (action.type === 'expense' ? 'cash-outline' : 'hammer-outline'));

                    const isCompleted = ['完了', '入庫済み', '入庫済'].includes(action.status);

                    itemDiv.innerHTML = `
                        <div class="type-icon ${colorClass}">
                            <ion-icon name="${iconName}"></ion-icon>
                        </div>
                        <div class="item-info">
                            <div class="item-header-meta">
                                <span class="item-type-badge">${getLabelPrefix(action.type)}</span>
                                <span class="item-status-badge ${isCompleted ? 'completed' : ''}">${action.status || '取引中'}</span>
                                ${isIncome && action.buyer ? `<span class="badge badge-buyer">${action.buyer}</span>` : ''}
                            </div>
                            <h4 class="item-name-text clickable" onclick="scrollToTransaction('${action.type}', '${action.id}')">${action.itemName}</h4>
                            <div class="item-sub-meta">
                                <span>${action.dateStr}</span>
                                <span>数量: ${action.quantity || '-'}</span>
                                <span>ID: ${action.id || '-'}</span>
                            </div>
                        </div>
                        <div class="item-amount ${isIncome ? 'positive' : (action.type === 'manufacturing' ? 'neutral' : 'negative')}">
                            ${action.type === 'manufacturing' ? '<span class="no-amount">-</span>' : symbol + '¥' + Math.abs(action.amount).toLocaleString()}
                        </div>
                    `;
                    if (reportBtn) listCard.insertBefore(itemDiv, reportBtn);
                    else listCard.appendChild(itemDiv);
                });
            }
        }

        // 明細リスト (Details用)
        const ledgerBody = document.getElementById('ledger-list-body');
        if (ledgerBody) {
            let tableSalesTotal = 0;
            let tablePurchaseTotal = 0;
            let tableCommTotal = 0;
            let tableRepairTotal = 0;
            let tableSuppliesTotal = 0;
            let tableDuesTotal = 0;
            let tableFeeTotal = 0;
            let tableMiscTotal = 0;
            let tableCostTotal = 0;
            let tableProfitTotal = 0;

            ledgerBody.innerHTML = filteredLedger.length > 0 ? filteredLedger.map(row => {
                const s = parseNumber(row['売上']);
                const p = parseNumber(row['仕入']);
                const c = parseNumber(row['通信費']);
                const r = parseNumber(row['修繕費']);
                const sp = parseNumber(row['消耗品費']);
                const d = parseNumber(row['諸会費']);
                const f = parseNumber(row['支払手数料']);
                const ms = parseNumber(row['雑費']);

                const rowCost = p + c + r + sp + d + f + ms;
                const rowProfit = s - rowCost;

                tableSalesTotal += s;
                tablePurchaseTotal += p;
                tableCommTotal += c;
                tableRepairTotal += r;
                tableSuppliesTotal += sp;
                tableDuesTotal += d;
                tableFeeTotal += f;
                tableMiscTotal += ms;
                tableCostTotal += rowCost;
                tableProfitTotal += rowProfit;

                const itemName = row['品名'] || "-";

                return `
                    <tr>
                        <td class="sticky-col first">${formatDate(row['日付'], 'short')}</td>
                        <td class="sticky-col second" title="${escapeHtml(itemName)}" onclick="this.classList.toggle('expanded')">${escapeHtml(itemName)}</td>
                        <td>${s > 0 ? '¥' + s.toLocaleString() : "-"}</td>
                        <td>${p > 0 ? '¥' + p.toLocaleString() : "-"}</td>
                        <td>${c > 0 ? '¥' + c.toLocaleString() : "-"}</td>
                        <td>${r > 0 ? '¥' + r.toLocaleString() : "-"}</td>
                        <td>${sp > 0 ? '¥' + sp.toLocaleString() : "-"}</td>
                        <td>${d > 0 ? '¥' + d.toLocaleString() : "-"}</td>
                        <td>${f > 0 ? '¥' + f.toLocaleString() : "-"}</td>
                        <td>${ms > 0 ? '¥' + ms.toLocaleString() : "-"}</td>
                        <td class="profit-col">¥${rowProfit.toLocaleString()}</td>
                    </tr>
                `;
            }).join('') : `<tr><td colspan="11" style="text-align:center; padding:20px;">選択された期間のデータはありません。</td></tr>`;

            const updateEl = (id, val) => {
                const el = document.getElementById(id);
                const rounded = Math.round(val || 0);
                if (el) el.textContent = rounded > 0 ? '¥' + rounded.toLocaleString() : (rounded < 0 ? '-¥' + Math.abs(rounded).toLocaleString() : '0');
            };

            // 明細表のフッターは、各項目の合計値と正確に一致させる (バグ報告15対応)
            updateEl('ledger-total-sales', tableSalesTotal);
            updateEl('ledger-total-purchase', tablePurchaseTotal);
            updateEl('ledger-total-comm', tableCommTotal);
            updateEl('ledger-total-repair', tableRepairTotal);
            updateEl('ledger-total-supplies', tableSuppliesTotal);
            updateEl('ledger-total-dues', tableDuesTotal);
            updateEl('ledger-total-fee', tableFeeTotal);
            updateEl('ledger-total-misc', tableMiscTotal);
            updateEl('ledger-total-profit', tableProfitTotal);
        }

        if (!window.ledgerInitialized) {
            const exportBtn = document.getElementById('export-ledger-btn');
            if (exportBtn) {
                exportBtn.addEventListener('click', handleLedgerExport);
            }
            window.ledgerInitialized = true;
        }
    }

    /**
     * 統一期間フィルタの描画
     */
    function updateLedgerPeriodFilters(ledger) {
        const container = document.getElementById('unified-ledger-filter');
        if (!container) return;

        // 年の抽出
        const years = new Set();
        ledger.forEach(row => {
            const d = new Date(row['日付']);
            if (!isNaN(d.getTime())) years.add(d.getFullYear());
        });
        if (years.size === 0) years.add(new Date().getFullYear());
        const sortedYears = Array.from(years).sort((a, b) => b - a);

        // コンテンツ生成
        let html = `
            <div class="filter-group">
                <div class="filter-row">
                    <span class="filter-label">対象年:</span>
                    <div class="segment-group">
                        ${sortedYears.map(y => `
                            <button class="segment-btn ${ledgerYear == y ? 'active' : ''}" data-year="${y}">${y}</button>
                        `).join('')}
                    </div>
                </div>
                <div class="filter-row">
                    <span class="filter-label">期間:</span>
                    <div class="segment-group">
                        <button class="segment-btn ${ledgerPeriod == 'FY' ? 'active' : ''}" data-period="FY">通期</button>
                        <button class="segment-btn ${ledgerPeriod == 'Q1' ? 'active' : ''}" data-period="Q1">Q1</button>
                        <button class="segment-btn ${ledgerPeriod == 'Q2' ? 'active' : ''}" data-period="Q2">Q2</button>
                        <button class="segment-btn ${ledgerPeriod == 'Q3' ? 'active' : ''}" data-period="Q3">Q3</button>
                        <button class="segment-btn ${ledgerPeriod == 'Q4' ? 'active' : ''}" data-period="Q4">Q4</button>
                    </div>
                    <select class="mini-month-select" id="ledger-month-picker">
                        <option value="">月を選択...</option>
                        ${[...Array(12)].map((_, i) => {
            const val = String(i + 1).padStart(2, '0');
            return `<option value="${val}" ${ledgerPeriod == val ? 'selected' : ''}>${i + 1}月</option>`;
        }).join('')}
                    </select>
                    <label class="ledger-personal-check-label" title="チェックを入れると月次明細に個人取引も合算します">
                        <input type="checkbox" id="ledger-include-personal-check" ${ledgerIncludePersonal ? 'checked' : ''}>
                        <span>個人</span>
                    </label>
                </div>
            </div>
        `;

        container.innerHTML = html;

        // イベントバインド
        container.querySelectorAll('[data-year]').forEach(btn => {
            btn.onclick = () => {
                ledgerYear = parseInt(btn.dataset.year);
                dashboardSelectedMonth = null;
                renderLedger();
            };
        });
        container.querySelectorAll('[data-period]').forEach(btn => {
            btn.onclick = () => {
                ledgerPeriod = btn.dataset.period;
                dashboardSelectedMonth = null;
                renderLedger();
            };
        });
        const monthPicker = document.getElementById('ledger-month-picker');
        if (monthPicker) {
            monthPicker.onchange = () => {
                if (monthPicker.value) {
                    ledgerPeriod = monthPicker.value;
                    dashboardSelectedMonth = null;
                    renderLedger();
                }
            };
        }
        const personalCheck = document.getElementById('ledger-include-personal-check');
        if (personalCheck) {
            personalCheck.onchange = (e) => {
                ledgerIncludePersonal = e.target.checked;
                renderLedger();
            };
        }
    }

    /**
     * 帳簿データをエクスポート
     */
    async function handleLedgerExport() {
        const typeText = ledgerIncludePersonal ? '（事業＋個人）' : '（事業のみ）';
        if (!confirm(`${ledgerYear}年${ledgerPeriod}の帳簿レポート${typeText}をスプレッドシートに出力しますか？`)) return;

        const statusArea = document.getElementById('ledger-export-status');
        if (!statusArea) return;

        const period = `${ledgerYear}-${ledgerPeriod}`;
        const btn = document.getElementById('export-ledger-btn');
        const originalContent = btn.innerHTML;

        try {
            // ローディング表示
            btn.disabled = true;
            btn.innerHTML = '<ion-icon name="refresh-outline" class="spin"></ion-icon>';
            setLoading(true, '帳簿レポートを生成中...');
            statusArea.style.display = 'block';
            statusArea.className = 'export-status-area info';
            statusArea.innerHTML = `<p>出力中... (${period} ${typeText})</p>`;

            const response = await fetchAPI('exportLedgerReport', { period: period, includePersonal: !!ledgerIncludePersonal });

            if (response.status === 'success') {
                const reportUrl = safeHttpsUrl(response.url);
                if (!reportUrl) throw new Error('帳票の出力先URLが無効です。');
                showToast('レポート出力が完了しました');
                statusArea.className = 'export-status-area success';
                statusArea.innerHTML = `
                    <p>出力完了: ${escapeHtml(response.fileName)}</p>
                    <a href="${escapeHtml(reportUrl)}" target="_blank" rel="noopener noreferrer" class="download-link">
                        <ion-icon name="open-outline"></ion-icon> スプレッドシートを開く
                    </a>
                `;
            } else {
                throw new Error(response.message);
            }
        } catch (e) {
            showToast('エクスポート失敗: ' + e.message, 'error');
            statusArea.className = 'export-status-area error';
            statusArea.innerHTML = `<p>出力エラー: ${escapeHtml(e.message)}</p>`;
        } finally {
            btn.disabled = false;
            btn.innerHTML = originalContent;
            setLoading(false);
        }
    }

    function getLabelPrefix(type) {
        const map = { 'purchase': '仕入', 'expense': '経費', 'sales': '販売', 'manufacturing': '製造' };
        return map[type] || '不明';
    }

    function formatDate(date, formatType = "full") {
        if (!date) return "-";
        const d = new Date(date);
        if (isNaN(d.getTime())) return date;

        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');

        if (formatType === "short") {
            return `${month}/${day}`;
        }
        return `${year}/${month}/${day}`;
    }

    function formatISODate(date) {
        if (!date) return "";
        const d = new Date(date);
        if (isNaN(d.getTime())) return "";

        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function parseNumber(val) {
        if (val === undefined || val === null || val === "") return 0;
        if (typeof val === 'number') return val;
        // ¥記号、カンマ、全角数字などを除去・変換して数値化
        const cleaned = String(val).replace(/[^0-9.-]+/g, "")
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        const num = parseFloat(cleaned);
        return isNaN(num) ? 0 : num;
    }

    function attachHistoryListeners() {
        // 在庫一覧の保存ボタンなどは既存のままでOK
    }

    // --- Image Features Helpers ---

    let currentUploadItemName = null;

    /**
     * 各タブの品名選択リスナーを設定
     */
    function setupImagePreviewListeners() {
        const itemMappings = [
            { id: 'buy-item', preview: 'buy-item-preview' },
            { id: 'exp-item', preview: 'exp-item-preview' },
            { id: 'make-item', preview: 'make-item-preview' },
            { id: 'sale-item', preview: 'sale-item-preview' }
        ];

        itemMappings.forEach(map => {
            const elm = document.getElementById(map.id);
            if (!elm || elm.dataset.hasPreviewListener === "true") return;

            const handler = () => updateProductImagePreview(map.preview, elm.value);
            elm.addEventListener('input', handler);
            elm.addEventListener('change', handler);
            elm.dataset.hasPreviewListener = "true";
        });
    }

    /**
     * 入力された品名に基づいて画像をプレビュー表示
     */
    function updateProductImagePreview(previewId, itemName) {
        const previewElm = document.getElementById(previewId);
        if (!previewElm || !currentMasters['M_商品']) return;

        const product = currentMasters['M_商品'].find(r => r['品名'] === itemName);
        const imageUrl = safeHttpsUrl(product?.['画像URL']);

        if (imageUrl) {
            const image = document.createElement('img');
            image.src = imageUrl;
            image.addEventListener('click', () => window.showImageModal(imageUrl));
            previewElm.replaceChildren(image);
            previewElm.classList.add('visible');
        } else {
            previewElm.innerHTML = '';
            previewElm.classList.remove('visible');
        }
    }

    /**
     * 写真アップロードのトリガー
     */
    window.triggerPhotoUpload = function (itemName) {
        currentUploadItemName = itemName;
        document.getElementById('hidden-photo-input').click();
    };

    /**
     * ファイルが選択された時の処理
     */
    window.handlePhotoSelected = async function (event) {
        const file = event.target.files[0];
        if (!file || !currentUploadItemName) return;

        try {
            setLoading(true, '画像をアップロード中...');
            const resizedBase64 = await resizeImage(file, 800, 800);

            // UI上のフィードバック（ボタンを無効化するなど）
            const syncBtn = document.getElementById('sync-btn');
            if (syncBtn) syncBtn.classList.add('spinning');

            const response = await fetchAPI('uploadProductImage', {
                itemName: currentUploadItemName,
                base64: resizedBase64
            });

            if (response.status === 'success') {
                console.log("Photo uploaded successfully:", response.url);
                // マスタデータをローカルで更新
                const product = currentMasters['M_商品'].find(r => r['品名'] === currentUploadItemName);
                if (product) product['画像URL'] = response.url;

                // フィルタを適用した状態で表示を更新
                applyStockFilters();
                // プレビュー表示中のものがあれば更新
                setupImagePreviewListeners(); // 再描画は不要だがキャッシュ更新の意味で

                showToast(`「${currentUploadItemName}」の写真を登録しました`);
            } else {
                throw new Error(response.message);
            }
        } catch (e) {
            showToast("写真のアップロードに失敗しました: " + e.message, 'error');
        } finally {
            if (document.getElementById('sync-btn')) {
                document.getElementById('sync-btn').classList.remove('spinning');
            }
            setLoading(false);
            event.target.value = ''; // Inputをリセット
        }
    };

    /**
     * ブラウザ側での画像リサイズ（Canvas使用）
     */
    function resizeImage(file, maxWidth, maxHeight) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => {
                    let width = img.width;
                    let height = img.height;

                    if (width > height) {
                        if (width > maxWidth) {
                            height *= maxWidth / width;
                            width = maxWidth;
                        }
                    } else {
                        if (height > maxHeight) {
                            width *= maxHeight / height;
                            height = maxHeight;
                        }
                    }

                    const canvas = document.createElement('canvas');
                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, width, height);
                    resolve(canvas.toDataURL('image/jpeg', 0.8)); // 画質0.8でJPEG圧縮
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        });
    }

    /**
     * 大画面プレビューの表示/非表示
     */
    function escapeHtml(value) {
        const entities = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
        return String(value ?? '').replace(/[&<>"']/g, char => entities[char]);
    }

    function safeHttpsUrl(value) {
        try {
            const url = new URL(value);
            return url.protocol === 'https:' && !url.username && !url.password ? url.href : '';
        } catch {
            return '';
        }
    }

    document.addEventListener('click', event => {
        const thumbnail = event.target.closest?.('.product-thumb-container img');
        if (thumbnail) window.showImageModal(thumbnail.src);
    });

    window.showImageModal = function (url) {
        url = safeHttpsUrl(url);
        if (!url) return;
        const modal = document.getElementById('image-modal');
        const content = document.getElementById('image-modal-content');
        content.src = url;
        modal.classList.add('active');
    };

    window.closeImageModal = function () {
        document.getElementById('image-modal').classList.remove('active');
    };

    /**
     * 発送前と確認できる販売数量を品名ごとに集計する（表示用）。
     */
    function collectPendingSales_(sales = []) {
        const totals = new Map();
        if (!Array.isArray(sales)) return totals;
        sales.forEach(sale => {
            const status = String(sale['ステータス'] || '').trim();
            if (!['入金待ち', '取引開始', '受注', '開始'].includes(status)) return;
            if (['発送日', '受取日', '取引完了日'].some(key => sale[key] !== undefined && sale[key] !== null && String(sale[key]).trim() !== '')) return;
            const name = String(sale['品名'] || '').trim();
            const rawQuantity = sale['数量'];
            const quantity = typeof rawQuantity === 'number' || typeof rawQuantity === 'string' ? Number(rawQuantity) : NaN;
            if (!name || !Number.isFinite(quantity) || quantity <= 0) return;
            totals.set(name, (totals.get(name) || 0) + quantity);
        });
        return totals;
    }

    // 在庫アラートダッシュボードの描画
    function renderInventoryAlerts() {
        const container = document.getElementById('inventory-alerts-container');
        if (!container) return;

        const stockData = currentMasters['T_在庫集計'] || [];
        const alerts = stockData.filter(row => {
            const qty = parseFloat(row['現在庫数']) || 0;
            const threshold = parseFloat(row['閾値']) || 0;
            // 閾値が0のものはアラート対象外とする
            if (threshold === 0) return false;
            // 廃盤(使用FLG=0)以外で、在庫が閾値以下のものを抽出
            return qty <= threshold && (parseInt(row['使用FLG']) !== 0);
        });

        if (alerts.length === 0) {
            container.innerHTML = `
                <div class="no-alerts-msg">
                    <ion-icon name="checkmark-circle-outline" style="font-size:24px; color:var(--accent-green); margin-bottom:8px; display:block; margin-left:auto; margin-right:auto;"></ion-icon>
                    現在、在庫不足の品目はありません。
                </div>
            `;
            return;
        }

        // 重要度順（在庫切れ > 閾値以下）にソート
        alerts.sort((a, b) => {
            const qtyA = parseFloat(a['現在庫数']) || 0;
            const qtyB = parseFloat(b['現在庫数']) || 0;
            return qtyA - qtyB;
        });

        let html = `
            <div class="alert-heading">
                <h3><ion-icon name="alert-circle"></ion-icon> 在庫アラート (${alerts.length})</h3>
            </div>
            <div class="alert-cards-container">
        `;

        const progressByItem = collectInventoryProgress(lastHistoryData?.history);
        const pendingSalesByItem = collectPendingSales_(lastHistoryData?.history?.['T_販売']);

        alerts.forEach(item => {
            const itemName = item['品名'];
            const qty = parseFloat(item['現在庫数']) || 0;
            const threshold = parseFloat(item['閾値']) || 0;
            const isCritical = qty === 0;
            const cardClass = isCritical ? 'critical' : 'warning';

            // 在庫は動かさず、発送前と確認できる販売の数量だけを表示する。
            const salesTotal = pendingSalesByItem.get(String(itemName || '').trim()) || 0;

            const progressBadges = inventoryProgressBadges(progressByItem.get(String(itemName || '').trim()));
            const salesBadge = salesTotal > 0 ? `<span class="alert-processing-badge sales"><ion-icon name="send-outline"></ion-icon>発送待(${salesTotal})</span>` : '';
            const statusBadge = progressBadges + salesBadge;

            // M_商品からカテゴリ情報を取得
            const product = (currentMasters['M_商品'] || []).find(m => {
                const name = (m['品名'] || '').toString().trim();
                return name === itemName.trim();
            });
            const category = product ? product['カテゴリ'] : "";

            const isMade = (category === '商品' || category === 'パーツ2');
            const isExpense = (category === '経費');

            let actionBtn = '';
            if (isMade) {
                actionBtn = `<button class="alert-btn make" data-name="${escapeHtml(itemName)}" onclick="jumpToTab('manufacturing', this.dataset.name)"><ion-icon name="hammer-outline"></ion-icon>製造へ</button>`;
            } else if (isExpense) {
                actionBtn = `<button class="alert-btn expense" data-name="${escapeHtml(itemName)}" onclick="jumpToTab('expense', this.dataset.name)"><ion-icon name="cash-outline"></ion-icon>経費へ</button>`;
            } else {
                actionBtn = `<button class="alert-btn purchase" data-name="${escapeHtml(itemName)}" onclick="jumpToTab('purchase', this.dataset.name)"><ion-icon name="cart-outline"></ion-icon>仕入へ</button>`;
            }

            html += `
                <div class="alert-card ${cardClass}">
                    <div class="alert-info">
                        <h4>${escapeHtml(itemName)}</h4>
                        <div class="alert-status-row">
                            <span class="alert-current-qty">現在: <b>${qty}</b></span>
                            <span class="alert-threshold">閾値: ${threshold}</span>
                            <span class="alert-category">${category}</span>
                        </div>
                        ${statusBadge ? `<div class="inventory-progress-badges" aria-label="進行中の数量（現在庫には含みません）">${statusBadge}</div>` : ''}
                    </div>
                    <div class="alert-actions">
                        ${actionBtn}
                    </div>
                </div>
            `;
        });

        html += `</div>`;
        container.innerHTML = html;
    }

    /**
     * 特定のタブへ遷移し、品目を選択済みにする (Proposal 36 拡張)
     */
    window.jumpToTab = function (tabId, itemName) {
        const navItem = document.querySelector(`.nav-item[data-target="${tabId}"]`);
        if (navItem) {
            navItem.click(); // タブ切り替え

            // タブ切り替えのアニメーション待ち
            setTimeout(() => {
                const map = { purchase: 'buy-item', manufacturing: 'make-item', expense: 'exp-item' };
                const inputId = map[tabId];
                const el = document.getElementById(inputId);
                if (el) {
                    el.value = itemName;

                    // 履歴から詳細データを自動セット (Proposal 36)
                    autoPopulateFromHistory(tabId, itemName);

                    // 入力イベントを発火させて連動するロジック（プレビュー表示など）を動かす
                    el.dispatchEvent(new Event('input'));
                    el.dispatchEvent(new Event('change'));

                    // スムーズにスクロールして目立たせる
                    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    el.focus();
                }
            }, 150);
        }
    };

    /**
     * 履歴から最新の値を自動補完する (Proposal 36共通)
     */
    function autoPopulateFromHistory(tabId, itemName) {
        if (tabId === 'purchase' || tabId === 'expense') {
            const historyKey = tabId === 'purchase' ? 'T_仕入' : 'T_経費';
            const historyData = (window.lastHistoryData && window.lastHistoryData.history) ? window.lastHistoryData.history[historyKey] : [];
            const latest = historyData.find(r => r['品名'] === itemName);

            if (latest) {
                if (tabId === 'purchase') {
                    setElementValue('buy-vendor', latest['仕入先']);
                    setElementValue('buy-price', latest['価格']);
                    setElementValue('buy-quantity', latest['数量']);
                    setElementValue('buy-payment', latest['支払方法']);
                    setElementValue('buy-category', latest['区分']);
                } else if (tabId === 'expense') {
                    setElementValue('exp-account', latest['仕訳']);
                    setElementValue('exp-vendor', latest['購入先']);
                    setElementValue('exp-price', latest['合計金額']);
                    setElementValue('exp-quantity', latest['数量']);
                    setElementValue('exp-payment', latest['支払方法']);
                    const manageFlag = document.getElementById('exp-is-stock');
                    if (manageFlag) manageFlag.checked = (latest['管理対象'] == 1);
                }
            }
        } else if (tabId === 'manufacturing') {
            setElementValue('make-quantity', "1");
        }
    }

    // ヘルパー: 要素が存在する場合のみ値をセットしてイベントを発火
    function setElementValue(id, value) {
        const el = document.getElementById(id);
        if (el && value !== undefined && value !== null) {
            el.value = value;
            syncQuantityChoice_(el, true);
            el.dispatchEvent(new Event('input'));
            el.dispatchEvent(new Event('change'));
        }
    }

    /**
     * マスタの抽出条件を解析してデータをフィルタリングする (Proposal 5)
     * @param {Array} data フィルタリング対象の配列
     * @param {string} filterStr 抽出条件文字列 (例: カテゴリ:商品,パーツ & 在庫:1以上)
     * @param {Object} allMasters 全マスタデータ (在庫参照用)
     */
    function parseAndApplyFilter(data, filterStr, allMasters) {
        if (!filterStr || !data || data.length === 0) return data;

        // & で AND 条件を分割
        const conditions = filterStr.split('&').map(s => s.trim());
        let filteredData = [...data];

        conditions.forEach(cond => {
            const parts = cond.split(':');
            if (parts.length < 2) return;

            const key = parts[0].trim();
            const valStr = parts[1].trim();

            if (key === '在庫' || key === '現在庫数') {
                // 在庫数によるフィルタリング (T_在庫集計を参照)
                const stockData = allMasters['T_在庫集計'] || [];
                const match = valStr.match(/([<>=]+|以上|以下|超|未満)?\s*(\d+)/);
                if (match) {
                    const op = match[1] || '>=';
                    const targetVal = parseFloat(match[2]);

                    const inStockItems = stockData.filter(stock => {
                        const currentVal = parseFloat(stock['現在庫数']) || 0;
                        if (op === '>=' || op === '以上') return currentVal >= targetVal;
                        if (op === '<=' || op === '以下') return currentVal <= targetVal;
                        if (op === '>' || op === '超') return currentVal > targetVal;
                        if (op === '<' || op === '未満') return currentVal < targetVal;
                        return currentVal === targetVal;
                    }).map(stock => stock['品名']);

                    filteredData = filteredData.filter(item => {
                        const itemName = item['品名'] || item['商品名'] || item['完成品名'];
                        return inStockItems.includes(itemName);
                    });
                }
            } else if (key === 'マスタ統合') {
                // 特殊処理: 指定された別マスタをマージする
                const otherMasterName = valStr;
                const otherData = (allMasters[otherMasterName] || []).filter(r => (parseInt(r['使用FLG']) || 0) === 1);
                filteredData = [...filteredData, ...otherData];
            } else {
                // 通常の列名によるフィルタリング
                const allowedValues = valStr.split(',').map(v => v.trim());
                filteredData = filteredData.filter(item => {
                    const itemVal = (item[key] || '').toString().trim();
                    return allowedValues.includes(itemVal);
                });
            }
        });

        return filteredData;
    }

    /**
     * 在庫一覧のフィルタ設定
     */
    function setupStockFilters() {
        const chips = document.querySelectorAll('.filter-chip');
        chips.forEach(chip => {
            chip.addEventListener('click', () => {
                chip.classList.toggle('active');

                if (chip.getAttribute('data-filter') === 'unchecked') {
                    stocktakeSession.uncheckedOnly = chip.classList.contains('active');
                }

                applyStockFilters();
            });
        });
    }

    /**
     * ---- 設定モーダル・マスタ管理ロジック ----
     */
    function setupSettingsListeners() {
        const modal = document.getElementById('settings-modal');
        const settingsBtn = document.getElementById('settings-btn');
        const closeBtn = document.getElementById('settings-close-btn');
        const backBtn = document.getElementById('settings-back-btn');
        const selectView = document.getElementById('master-select-view');
        const editorView = document.getElementById('master-editor-view');
        const title = document.getElementById('settings-title');

        if (settingsBtn) {
            settingsBtn.addEventListener('click', () => {
                modal.classList.add('active');
                renderMasterList();
            });
        }

        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                modal.classList.remove('active');
            });
        }

        if (backBtn) {
            backBtn.addEventListener('click', () => {
                editorView.style.display = 'none';
                selectView.style.display = 'block';
                backBtn.style.display = 'none';
                title.textContent = 'システム設定・マスタ管理';
            });
        }

        // モーダル外クリックで閉じる
        modal.addEventListener('click', (e) => {
            if (e.target === modal) modal.classList.remove('active');
        });

        // 二重モーダルの閉じる制御
        const editModal = document.getElementById('master-edit-modal');
        document.getElementById('master-edit-close-btn').addEventListener('click', () => editModal.classList.remove('active'));
        document.getElementById('master-edit-cancel').addEventListener('click', () => editModal.classList.remove('active'));
    }

    function renderMasterList(mode = 'all') {
        const container = document.querySelector('.master-list-grid');
        const title = document.getElementById('settings-title');
        if (!container) return;

        if (title) {
            title.textContent = (mode === 'master') ? 'マスタ管理' :
                (mode === 'system') ? 'システム設定' : '設定・マスタ管理';
        }

        const masterConfig = {
            'M_商品': { title: '商品・パーツ', icon: 'cube-outline', type: 'master' },
            'M_仕入先': { title: '仕入先・購入先', icon: 'business-outline', type: 'master' },
            'M_売先': { title: '販売先(顧客)', icon: 'people-outline', type: 'master' },
            'M_BOM': { title: '製造レシピ(BOM)', icon: 'construct-outline', type: 'master' },
            'M_発送': { title: '配送・送料', icon: 'bus-outline', type: 'master' },
            'M_経費品名': { title: '経費科目名', icon: 'receipt-outline', type: 'master' },
            'M_支払': { title: '支払方法', icon: 'wallet-outline', type: 'master' },
            'M_ステータス': { title: 'ステータス定義', icon: 'flag-outline', type: 'system' },
            'M_画面制御': { title: '画面入力制御', icon: 'options-outline', type: 'system' },
            'T_在庫集計': { title: 'アラート設定(在庫閾値)', icon: 'notifications-outline', type: 'system' }
        };

        container.innerHTML = '';
        Object.keys(masterConfig).forEach(mkey => {
            const conf = masterConfig[mkey];

            // モードによるフィルタリング
            if (mode !== 'all' && conf.type !== mode) return;

            const card = document.createElement('div');
            card.className = 'master-card';
            card.innerHTML = `
                <ion-icon name="${conf.icon}"></ion-icon>
                <span>${conf.title}</span>
            `;
            card.addEventListener('click', () => openMasterEditor(mkey, conf.title));
            container.appendChild(card);
        });
    }

    async function openMasterEditor(masterKey, masterTitle) {
        const selectView = document.getElementById('master-select-view');
        const editorView = document.getElementById('master-editor-view');
        const backBtn = document.getElementById('settings-back-btn');
        const title = document.getElementById('settings-title');
        const head = document.getElementById('master-table-head');
        const body = document.getElementById('master-table-body');

        selectView.style.display = 'none';
        editorView.style.display = 'block';
        backBtn.style.display = 'block';
        title.textContent = masterTitle;
        title.dataset.currentKey = masterKey;

        // 読み込み開始時にヘッダーとボディをクリア
        head.innerHTML = '';
        body.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 20px;">読み込み中...</td></tr>`;

        try {
            // 全データを取得するためにAPIを叩く（無効データも含める）
            const res = await fetchAPI('getMasters', { includeInactive: true });
            const rawData = res.data[masterKey] || [];
            if (rawData.length < 1) {
                body.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 20px;">データがありません</td></tr>`;
                return;
            }

            const headers = rawData[0];
            const data = rawData.slice(1).map(row => {
                const obj = {};
                headers.forEach((h, i) => obj[h] = row[i]);
                return obj;
            });

            // ヘッダー生成
            const schema = MASTER_SCHEMAS[masterKey];
            const rawKeys = Object.keys(data[0]).filter(k => k !== '最終更新日' && k.trim() !== "");
            let keys = [];

            if (schema) {
                console.log(`Applying schema for ${masterKey}:`, schema);
                // 1. シートにある列のうち、表示すべきものだけを抽出 (トリムして比較)
                keys = rawKeys.filter(rk => {
                    const cleanRK = rk.trim();
                    const f = schema.fields.find(field => field.name.trim() === cleanRK);
                    return f ? f.visible !== false : true;
                });

                // 2. スキーマの定義順に従って並び替え
                keys.sort((a, b) => {
                    const idxA = schema.fields.findIndex(f => f.name.trim() === a.trim());
                    const idxB = schema.fields.findIndex(f => f.name.trim() === b.trim());
                    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                    if (idxA !== -1) return -1;
                    if (idxB !== -1) return 1;
                    return 0;
                });
            } else {
                console.warn(`No schema found for ${masterKey}`);
                keys = rawKeys;
            }

            head.innerHTML = `<tr><th class="sticky-col">操作</th>${keys.map(k => `<th>${k}</th>`).join('')}</tr>`;

            // ボディ生成
            renderMasterTableBody(data, keys, masterKey);

            // 検索・追加ボタンのイベント再設定
            const searchInput = document.getElementById('master-search-input');
            searchInput.value = '';
            searchInput.oninput = (e) => {
                const term = e.target.value.toLowerCase();
                const filtered = data.filter(row =>
                    Object.values(row).some(v => String(v).toLowerCase().includes(term))
                );
                renderMasterTableBody(filtered, keys, masterKey);
            };

            const addBtn = document.getElementById('master-add-btn');
            addBtn.onclick = () => showMasterEntryModal(masterKey, keys, null);

        } catch (e) {
            body.innerHTML = `<tr><td colspan="5" style="color:red; text-align:center; padding: 20px;">取得エラー: ${escapeHtml(e.message)}</td></tr>`;
        }
    }

    function renderMasterTableBody(data, keys, masterKey) {
        const body = document.getElementById('master-table-body');
        const activeMasterKey = masterKey || Object.keys(currentMasters).find(k => k === document.getElementById('settings-title').dataset.currentKey) || "";
        const schema = MASTER_SCHEMAS[activeMasterKey];

        body.innerHTML = '';
        data.forEach((row, idx) => {
            const tr = document.createElement('tr');
            const isActive = row['使用FLG'] == 1 || row['使用FLG'] === true;
            if (!isActive) tr.classList.add('master-row-inactive');

            let html = keys.map(k => {
                let val = row[k] !== undefined ? row[k] : '';
                // スイッチ項目の場合は表示を工夫
                const f = schema ? schema.fields.find(field => field.name.trim() === k.trim()) : null;
                if (f && f.type === 'switch') {
                    val = val == 1 ? '<span class="status-badge success">有効</span>' : '<span class="status-badge danger">無効</span>';
                }
                return `<td>${val}</td>`;
            }).join('');

            // アクションボタン（ステータス切替 ＆ 編集）
            const tdActions = document.createElement('td');
            tdActions.className = 'action-cell';

            // 更新用キーの取得
            const keyCol = schema ? schema.key : keys[0];
            const rowId = row[keyCol];

            const btnStatus = document.createElement('button');
            btnStatus.className = 'btn-edit-master';
            btnStatus.style.color = isActive ? 'var(--accent-green)' : 'var(--text-muted)';
            btnStatus.title = isActive ? '無効にする' : '有効にする';
            btnStatus.innerHTML = `<ion-icon name="${isActive ? 'eye-outline' : 'eye-off-outline'}"></ion-icon>`;
            btnStatus.onclick = () => toggleMasterStatus(activeMasterKey, rowId, !isActive);

            const btnEdit = document.createElement('button');
            btnEdit.className = 'btn-edit-master';
            btnEdit.innerHTML = `<ion-icon name="create-outline"></ion-icon>`;
            btnEdit.onclick = () => showMasterEntryModal(activeMasterKey, keys, row);

            tdActions.classList.add('sticky-col');
            tdActions.appendChild(btnStatus);
            tdActions.appendChild(btnEdit);

            tr.appendChild(tdActions);
            tr.insertAdjacentHTML('beforeend', html);
            body.appendChild(tr);
        });
    }

    async function toggleMasterStatus(masterKey, id, newStatus) {
        if (!confirm(`この項目のステータスを変更しますか？`)) return;
        setLoading(true, 'マスタを更新中...');
        try {
            const res = await fetchAPI('updateMasterRecord', {
                masterName: masterKey,
                id: id,
                updates: { '使用FLG': newStatus ? 1 : 0 }
            });
            if (res.status === 'success') {
                showToast('マスタを更新しました');
                openMasterEditor(masterKey, document.getElementById('settings-title').textContent); // 再読み込み
            } else throw new Error(res.message);
        } catch (e) {
            showToast('更新失敗: ' + e.message, 'error');
        } finally {
            setLoading(false);
        }
    }

    function showMasterEntryModal(masterKey, keys, rowData) {
        const modal = document.getElementById('master-edit-modal');
        const fields = document.getElementById('master-edit-fields');
        const form = document.getElementById('master-edit-form');
        const title = document.getElementById('master-edit-title');
        const schema = MASTER_SCHEMAS[masterKey];

        title.textContent = rowData ? `${masterKey} の編集` : `${masterKey} への新規追加`;
        fields.innerHTML = '';
        modal.classList.add('active');

        // 表示対象の列を取得
        let targetKeys = keys;
        if (schema) {
            targetKeys = keys.filter(k => {
                const f = schema.fields.find(field => field.name === k);
                return f ? f.visible !== false : true;
            });
        }

        targetKeys.forEach((key) => {
            const group = document.createElement('div');
            group.className = 'input-group';

            const fieldConfig = schema ? schema.fields.find(f => f.name === key) : null;
            const value = rowData ? rowData[key] : '';
            const isId = fieldConfig ? (schema.key === key) : (keys.indexOf(key) === 0);
            const isEditable = fieldConfig ? (rowData ? fieldConfig.editable : true) : true;

            let inputHtml = '';
            const type = fieldConfig ? fieldConfig.type : (typeof value === 'number' ? 'number' : 'text');

            if (type === 'switch') {
                const checked = (value == 1 || value === true) ? 'checked' : '';
                inputHtml = `
                    <div style="display:flex; align-items:center; gap:10px; padding:10px 0;">
                        <span style="font-size:14px;">${value == 1 ? '有効' : '無効'}</span>
                        <label class="switch-ui">
                            <input type="checkbox" name="${key}" ${checked} onchange="this.previousElementSibling.textContent = this.checked ? '有効' : '無効'">
                            <span class="slider round"></span>
                        </label>
                    </div>
                `;
            } else if (type === 'select') {
                let options = [];
                if (fieldConfig.options) {
                    options = fieldConfig.options;
                } else if (fieldConfig.refMaster) {
                    const refData = currentMasters[fieldConfig.refMaster] || [];
                    const filteredRef = fieldConfig.filter ? refData.filter(fieldConfig.filter) : refData;
                    options = filteredRef.map(r => r['品名'] || r[Object.keys(r)[0]]);
                }

                inputHtml = `
                    <select name="${key}" ${!isEditable ? 'disabled class="readonly-field"' : ''}>
                        <option value="">選択してください</option>
                        ${options.map(opt => {
                    const val = typeof opt === 'object' ? opt.v : opt;
                    const lbl = typeof opt === 'object' ? opt.l : opt;
                    return `<option value="${val}" ${val == value ? 'selected' : ''}>${lbl}</option>`;
                }).join('')}
                    </select>
                `;
            } else if (type === 'textarea') {
                inputHtml = `<textarea name="${key}" ${!isEditable ? 'readonly class="readonly-field"' : ''}>${escapeHtml(value !== undefined ? value : '')}</textarea>`;
            } else if (key === 'QR/バーコード') {
                inputHtml = `
                    <div class="input-with-icon">
                        <input type="text" 
                               name="${key}" 
                               value="${value !== undefined ? value : ''}"
                               ${!isEditable ? 'readonly class="readonly-field"' : ''}>
                        <button type="button" class="icon-input-btn" onclick="triggerMasterScanner(this)" title="スキャン">
                            <ion-icon name="camera-outline"></ion-icon>
                        </button>
                    </div>
                `;
            } else {
                inputHtml = `
                    <input type="${type}" 
                           name="${key}" 
                           value="${value !== undefined ? value : ''}"
                           ${!isEditable ? 'readonly class="readonly-field"' : ''}
                           ${key === '使用FLG' ? 'placeholder="1:有効, 0:無効"' : ''}>
                `;
            }

            const labelSuffix = fieldConfig && fieldConfig.required === false ? ' <span style="font-size:10px; color:#999;">(任意)</span>' : '';
            group.innerHTML = `<label>${key}${labelSuffix}</label>${inputHtml}`;
            fields.appendChild(group);
        });

        form.onsubmit = async (e) => {
            e.preventDefault();
            const formData = new FormData(form);
            const updates = {};
            let hasError = false;

            // スキーマの定義に従って値を収集
            targetKeys.forEach(key => {
                const fieldConfig = schema ? schema.fields.find(f => f.name === key) : null;
                const type = fieldConfig ? fieldConfig.type : 'text';
                let value = '';

                if (type === 'switch') {
                    const checkbox = form.querySelector(`input[name="${key}"]`);
                    value = checkbox.checked ? 1 : 0;
                } else {
                    const element = form.querySelector(`[name="${key}"]`);
                    value = element ? element.value.trim() : '';
                }

                const isRequired = fieldConfig ? fieldConfig.required !== false : true;
                if (isRequired && value === "" && type !== 'switch') {
                    if (!hasError) showToast(`「${key}」を入力してください。`, 'error');
                    hasError = true;
                }

                // 型の変換
                if (type === 'number' || key === '表示順' || key === '使用FLG' || key === '手数料率') {
                    const num = parseFloat(value);
                    if (isNaN(num)) {
                        if (!hasError) showToast(`「${key}」には数値を入力してください。`, 'error');
                        hasError = true;
                    }
                    updates[key] = num;
                } else {
                    updates[key] = value;
                }
            });

            if (hasError) return;

            // 編集モードならキーを取得
            const rowId = rowData ? rowData[schema ? schema.key : keys[0]] : null;
            const isNew = !rowData;

            const btn = document.getElementById('master-edit-submit');
            const originalText = btn.textContent;
            btn.disabled = true;
            btn.textContent = '保存中...';
            setLoading(true, 'マスタデータを保存中...');

            try {
                let res;
                if (rowData) {
                    res = await fetchAPI('updateMasterRecord', {
                        masterName: masterKey,
                        id: rowId,
                        updates: updates
                    });
                } else {
                    res = await fetchAPI('addMasterRecord', {
                        masterName: masterKey,
                        updates: updates
                    });
                }

                if (res.status === 'success') {
                    showToast('保存しました');
                    modal.classList.remove('active');
                    openMasterEditor(masterKey, document.getElementById('settings-title').textContent);
                } else throw new Error(res.message);
            } catch (err) {
                showToast('保存失敗: ' + err.message, 'error');
            } finally {
                btn.disabled = false;
                btn.textContent = originalText;
                setLoading(false);
            }
        };
    }

    // マスタ編集用スキャナー起動
    window.triggerMasterScanner = function (btn) {
        const input = btn.previousElementSibling;
        const originalOnScanSuccess = window.onScanSuccess;

        // 一時的なスキャン成功時処理
        window.onScanSuccess = function (decodedText) {
            input.value = decodedText;
            stopScanner();
            showToast('バーコードを読み取りました', 'success');
            // 元の処理に戻す
            window.onScanSuccess = originalOnScanSuccess;
        };

        startScanner();
    };

    /**
     * トースト通知を表示する
     */
    function showToast(message, type = 'success') {
        const container = document.getElementById('toast-container');
        if (!container) return;

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        const icon = type === 'success' ? 'checkmark-circle-outline' : 'alert-circle-outline';

        toast.innerHTML = `
            <ion-icon name="${icon}"></ion-icon>
            <span>${message}</span>
        `;

        container.appendChild(toast);

        // アニメーション終了後に削除
        setTimeout(() => {
            toast.remove();
        }, 3000);

        // モバイルならバイブレーション
        if (type === 'success' && navigator.vibrate) {
            navigator.vibrate(50);
        }
    }

    /**
     * グローバルローダーの表示/非表示
     */
    function setLoading(isLoading, message = '処理中...') {
        const loader = document.getElementById('global-loader');
        const text = document.getElementById('loader-text');
        if (!loader || !text) return;

        if (isLoading) {
            text.textContent = message;
            loader.classList.add('active');
        } else {
            loader.classList.remove('active');
        }
    }

    /**
     * 特定の取引レコードまでスクロールし、ハイライト表示する
     */
    window.scrollToTransaction = (type, id) => {
        // 通常一覧の表示期間外にある最近更新された履歴は、既存検索で参照する。
        const idKey = { purchase: '仕入ID', expense: '経費ID', manufacturing: '製造ID', sales: '販売ID' }[type];
        const rows = lastHistoryData.history?.[historySheets[type]] || [];
        if (idKey && !rows.some(row => String(row[idKey]) === String(id))) {
            const modal = document.getElementById('history-search-modal');
            if (modal) {
                document.getElementById('history-search-target').value = type;
                document.getElementById('history-search-keyword').value = id;
                document.getElementById('history-search-from').value = '';
                document.getElementById('history-search-to').value = '';
                modal.classList.add('active');
                executeHistorySearch(String(id));
                return;
            }
        }
        // 1. タブの切り替え
        const navItem = document.querySelector(`.nav-item[data-target="${type}"]`);
        if (navItem) navItem.click();

        // 2. カードの特定とスクロール
        // 履歴カードの描画完了を少し待つ
        setTimeout(() => {
            const container = document.getElementById('content-area');
            const idBadges = document.querySelectorAll('.history-id');
            let targetCard = null;
            for (const badge of idBadges) {
                if (badge.textContent.trim() === id) {
                    targetCard = badge.closest('.history-card');
                    break;
                }
            }

            if (targetCard && container) {
                // スムーズスクロール
                targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
                // ハイライト効果
                targetCard.classList.add('highlight-flash');
                setTimeout(() => targetCard.classList.remove('highlight-flash'), 3000);
            } else {
                console.warn(`Target card with ID ${id} not found.`);
            }
        }, 300); // 描画時間を考慮して少し長めに待機
    };

    /**
     * 在庫一覧からカテゴリに応じた取引画面へ遷移し、対象品目を選択する (提案11改)
     */
    /**
     * 在庫一覧からカテゴリに応じた取引画面へ遷移し、対象品目を選択する (提案11改)
     */
    window.navigateToTransactionForm = (itemName, category) => {
        const cleanCategory = (category || "").toString().trim();

        // カテゴリが「商品」または「単体商品」の場合は選択モーダルを表示 (提案18/25 B案・提案49)
        if (cleanCategory === '商品' || cleanCategory === '単体商品') {
            showActionSelectionModal(itemName, cleanCategory);
            return;
        }

        // それ以外は従来通り直接遷移
        let targetTab = 'purchase';
        let inputId = 'buy-item';

        if (cleanCategory === 'パーツ2') {
            targetTab = 'manufacturing';
            inputId = 'make-item';
        } else if (cleanCategory === '経費') {
            targetTab = 'expense';
            inputId = 'exp-item';
        } else if (cleanCategory === 'パーツ') {
            targetTab = 'purchase';
            inputId = 'buy-item';
        }

        executeNavigation(targetTab, inputId, itemName);
    };

    /**
     * 実際の遷移処理を実行する
     */
    function executeNavigation(targetTab, inputId, itemName) {
        // 1. タブを切り替える
        const navItem = document.querySelector(`.nav-item[data-target="${targetTab}"]`);
        if (navItem) navItem.click();

        // 2. 品目を選択する
        const input = document.getElementById(inputId);
        if (input) {
            input.value = itemName;

            // 履歴から詳細データを自動セット (Proposal 36)
            autoPopulateFromHistory(targetTab, itemName);

            input.dispatchEvent(new Event('change'));
            input.dispatchEvent(new Event('input')); // datalist用
        }

        // 3. 入力フォームまでスクロールして強調
        setTimeout(() => {
            const form = document.querySelector(`#tab-${targetTab} .card`);
            if (form) {
                form.scrollIntoView({ behavior: 'smooth', block: 'start' });
                form.style.transition = 'box-shadow 0.5s';
                form.style.boxShadow = '0 0 20px var(--primary-color)';
                setTimeout(() => form.style.boxShadow = '', 1500);
            }
        }, 100);
    }

    /**
     * 遷移先選択モーダルを表示する (提案18/25 B案)
     */
    function showActionSelectionModal(itemName, category = '商品') {
        let overlay = document.getElementById('action-selection-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.id = 'action-selection-overlay';
            overlay.className = 'action-selection-overlay';
            overlay.innerHTML = `
                <div class="action-selection-card">
                    <div class="action-selection-title">取引の登録</div>
                    <div class="action-selection-item-name" id="action-selection-item-name"></div>
                    <div class="action-btn-group">
                        <button class="action-select-btn manufacturing" id="action-btn-primary">
                            <ion-icon name="hammer-outline"></ion-icon>製造登録へ
                        </button>
                        <button class="action-select-btn sales" id="action-btn-sales">
                            <ion-icon name="cart-outline"></ion-icon>販売登録へ
                        </button>
                        <button class="action-select-btn cancel" id="action-btn-cancel">キャンセル</button>
                    </div>
                </div>
            `;
            document.body.appendChild(overlay);
        }

        const nameEl = document.getElementById('action-selection-item-name');
        if (nameEl) nameEl.textContent = itemName;

        const primaryBtn = document.getElementById('action-btn-primary');
        const salesBtn = document.getElementById('action-btn-sales');
        
        // 提案49: カテゴリに応じたボタン表示の切り替え
        if (category === '単体商品') {
            primaryBtn.innerHTML = '<ion-icon name="cart-outline"></ion-icon>仕入登録へ';
            primaryBtn.className = 'action-select-btn purchase';
        } else {
            primaryBtn.innerHTML = '<ion-icon name="hammer-outline"></ion-icon>製造登録へ';
            primaryBtn.className = 'action-select-btn manufacturing';
        }

        const close = () => overlay.classList.remove('active');

        primaryBtn.onclick = () => {
            close();
            if (category === '単体商品') {
                executeNavigation('purchase', 'buy-item', itemName);
            } else {
                executeNavigation('manufacturing', 'make-item', itemName);
            }
        };
        salesBtn.onclick = () => {
            close();
            executeNavigation('sales', 'sale-item', itemName);
        };
        document.getElementById('action-btn-cancel').onclick = close;
        overlay.onclick = (e) => { if (e.target === overlay) close(); };

        setTimeout(() => overlay.classList.add('active'), 10);
    }

    /**
     * 製造開始画面でのBOM在庫チェック (提案対応)
     */
    function updateBOMCheck(itemName, targetQtyStr) {
        const container = document.getElementById('bom-check-container');
        const resultList = document.getElementById('bom-check-result');
        if (!container || !resultList) return;

        const targetQty = parseFloat(targetQtyStr);
        if (!itemName || isNaN(targetQty) || targetQty <= 0) {
            container.style.display = 'none';
            resultList.innerHTML = '';
            return;
        }

        // BOMデータの取得 (M_BOM)
        const bomData = (currentMasters['M_BOM'] || []).filter(b => b['品名'] === itemName);
        if (bomData.length === 0) {
            container.style.display = 'none';
            resultList.innerHTML = '';
            return;
        }

        container.style.display = 'block';
        resultList.innerHTML = '';
        let allOk = true;

        bomData.forEach(bom => {
            const componentName = bom['部品'];
            const neededPerOne = parseFloat(bom['数量']) || 0;
            const totalNeeded = neededPerOne * targetQty;

            // 現在庫の取得 (T_在庫集計)
            const stockItem = (currentMasters['T_在庫集計'] || []).find(s => s['品名'] === componentName);
            const currentStock = stockItem ? parseFloat(stockItem['現在庫数']) || 0 : 0;

            const isOk = currentStock >= totalNeeded;
            if (!isOk) allOk = false;

            const itemDiv = document.createElement('div');
            itemDiv.className = 'bom-item-status';
            itemDiv.innerHTML = `
                <span class="bom-item-name">${componentName}</span>
                <span class="bom-item-qty ${isOk ? 'ok' : 'ng'}">
                    ${totalNeeded} / 在庫:${currentStock}
                </span>
            `;
            resultList.appendChild(itemDiv);
        });

        const summary = document.createElement('div');
        summary.className = `bom-status-summary ${allOk ? 'ok' : 'ng'}`;
        summary.innerHTML = allOk ?
            '<ion-icon name="checkmark-circle" style="vertical-align: middle; margin-right: 4px;"></ion-icon> 在庫はすべて足りています' :
            '<ion-icon name="warning-outline" style="vertical-align: middle; margin-right: 4px;"></ion-icon> 一部の部品が不足しています';
        resultList.appendChild(summary);
    }

    function fxMoneyUnits_(value, decimals) {
      const text = String(value ?? '').trim();
      const pattern = decimals === 2 ? /^\d+(?:\.\d{1,2})?$/ : /^\d+$/;
      if (!pattern.test(text)) throw new Error(decimals === 2 ? 'USDは小数第2位まで入力してください。' : '円金額は0以上の整数で入力してください。');
      const units = Math.round(Number(text) * (decimals === 2 ? 100 : 1));
      if (!Number.isSafeInteger(units) || units > 1000000000) throw new Error('金額が入力可能な範囲を超えています。');
      return units;
    }

    function calculateFxAmounts_(group) {
      if (!group || !Array.isArray(group.lines) || group.lines.length < 1 || group.lines.length > 20) throw new Error('明細は1〜20件で入力してください。');
      const usdTotal = fxMoneyUnits_(group.usdTotal, 2);
      const yenTotal = fxMoneyUnits_(group.yenTotal, 0);
      if (usdTotal <= 0 || yenTotal <= 0) throw new Error('決済総額は0より大きい金額を入力してください。');
      const cents = group.lines.map(line => fxMoneyUnits_(line.usd, 2));
      if (cents.some(value => value <= 0)) throw new Error('商品合計USDは0より大きい金額を入力してください。');
      const sum = cents.reduce((a, b) => a + b, 0);
      if (sum !== usdTotal) throw new Error('商品合計USDと決済総額USDが一致しません。差額（決済−商品）：' + ((usdTotal - sum) / 100).toFixed(2) + ' USD');
      const products = cents.map(value => value * yenTotal);
      if (products.some(value => !Number.isSafeInteger(value))) throw new Error('計算可能な金額の範囲を超えています。');
      const automatic = products.map(value => (value - value % usdTotal) / usdTotal);
      const order = products.map((value, index) => ({ index, remainder: value % usdTotal }))
        .sort((a, b) => b.remainder - a.remainder || a.index - b.index);
      const remaining = yenTotal - automatic.reduce((a, b) => a + b, 0);
      for (let i = 0; i < remaining; i++) automatic[order[i].index]++;
      const prices = group.lines.map((line, index) => line.yen === '' || line.yen === undefined || line.yen === null ? automatic[index] : fxMoneyUnits_(line.yen, 0));
      if (prices.reduce((a, b) => a + b, 0) !== yenTotal) throw new Error('商品合計円と決済総額円が一致しません。円の修正額を確認してください。');
      return { usdTotal, yenTotal, cents, automatic, prices };
    }


    // PROP-062: 通常の円登録・買い物カゴとは独立した外貨決済画面
    function setupForeignPurchase() {
        const modal = document.getElementById('fx-purchase-modal');
        if (!modal) return;
        const el = id => document.getElementById('fx-' + id);
        const storageKey = 'inventory_fx_draft_v1_' + GAS_URL.split('/').slice(-2, -1)[0];
        const blank = () => ({ date: new Date().toLocaleDateString('sv-SE'), vendor: 'AliExpress', payment: 'Revolut', usdTotal: '', yenTotal: '', lines: [] });
        let state = { group: blank(), locked: false, done: false, results: [] };
        let previewSignature = '';
        let busy = false;
        let opener = null;
        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
            if (saved && Array.isArray(saved.group?.lines) && saved.group.lines.length <= 20) state = saved;
        } catch (_) { /* 保存が使えない場合は、送信前の保存でエラーとして案内する */ }
        const persist = () => localStorage.setItem(storageKey, JSON.stringify(state));
        const optionsFrom = id => Array.from(document.getElementById(id)?.options || [])
            .filter(option => option.value).map(option => option.value);
        const selectField = (name, title, value, options) => `<label>${title}<select data-field="${name}"><option value="">選択してください</option>${options.map(option => `<option value="${escapeHtml(option)}" ${option === value ? 'selected' : ''}>${escapeHtml(option)}</option>`).join('')}</select></label>`;
        const inputField = (name, title, value, attributes = '') => `<label>${title}<input data-field="${name}" type="text" value="${escapeHtml(value ?? '')}" ${attributes}></label>`;
        const invalidatePreview = () => {
            previewSignature = '';
            el('preview').hidden = true;
            if (!state.locked) el('confirm').textContent = '登録内容を確認する';
        };
        function showResults(message = '') {
            el('results').textContent = message;
            (state.results || []).forEach((result, index) => {
                const item = document.createElement('div');
                const label = { completed: '登録済み', pending: '要確認（処理結果未確定）', review: '要確認', waiting: '未処理' }[result.state] || '要確認';
                item.textContent = `${index + 1}. ${state.group.lines[index]?.item || ''}：${label}${result.id ? ' (' + result.id + ')' : ''}`;
                el('results').appendChild(item);
            });
        }
        function updateControls() {
            el('inputs').disabled = state.locked || busy;
            el('confirm').disabled = state.done || busy;
            el('new').disabled = busy || (state.locked && !state.done);
            el('close').disabled = busy;
            if (state.locked) el('confirm').textContent = state.done ? '登録完了' : '同じ決済の結果を確認する';
            el('new').textContent = state.done ? '次の決済を入力' : '入力を破棄して新しい決済を入力';
        }
        function updateCalculation() {
            try {
                const auto = calculateFxAmounts_({ ...state.group, lines: state.group.lines.map(line => ({ ...line, yen: '' })) });
                el('rate').textContent = formatFxRate_(auto);
                el('lines').querySelectorAll('.fx-line').forEach((card, index) => {
                    const line = state.group.lines[index];
                    const yen = line.yen === '' ? auto.automatic[index] : fxMoneyUnits_(line.yen, 0);
                    const unit = Number(line.quantity) > 0 ? (yen / Number(line.quantity)).toFixed(2) : '—';
                    card.querySelector('.fx-line-total').textContent = `円換算合計：${yen.toLocaleString()}円 ／ 単価：${unit}円`;
                });
                const amounts = calculateFxAmounts_(state.group);
                el('check').textContent = `USD合計 ${(amounts.usdTotal / 100).toFixed(2)} ／ 円合計 ${amounts.yenTotal.toLocaleString()}円：決済額と一致（円未満は自動調整）`;
            } catch (error) {
                el('check').textContent = error.message;
                el('rate').textContent = '';
                el('lines').querySelectorAll('.fx-line-total').forEach(item => { item.textContent = '合計が一致すると円換算額を表示します'; });
            }
        }
        function renderLines() {
            el('lines').innerHTML = state.group.lines.map((line, index) => {
                const purchase = line.type === 'purchase';
                return `<section class="fx-line" data-index="${index}">
                    <h3 class="fx-section-title">商品${index + 1}</h3>
                    <div class="fx-grid">
                        ${selectField('type', '登録先', line.type, ['purchase', 'expense']).replace('>purchase<', '>仕入<').replace('>expense<', '>経費<')}
                        ${selectField('status', 'ステータス', line.status, optionsFrom(purchase ? 'buy-status-entry' : 'exp-status-entry'))}
                        ${inputField('item', '品名', line.item, 'list="fx-item-list" maxlength="200"')}
                        ${inputField('quantity', '数量', line.quantity, 'inputmode="numeric"')}
                        ${inputField('usd', '商品合計（USD）', line.usd, 'inputmode="decimal"')}
                        ${inputField('yen', '円合計の修正（空欄は自動計算）', line.yen, 'inputmode="numeric"')}
                        ${purchase ? selectField('category', '区分', line.category, optionsFrom('buy-category')) : inputField('account', '仕訳', line.account, 'list="expense-category-list" maxlength="200"')}
                    </div>
                    ${purchase ? '' : `<div class="fx-check-row"><label><input data-field="isStock" type="checkbox" ${line.isStock ? 'checked' : ''}>在庫管理対象</label><label><input data-field="receipt" type="checkbox" ${line.receipt ? 'checked' : ''}>レシートあり</label></div>`}
                    <label>備考<textarea data-field="note" maxlength="2000">${escapeHtml(line.note || '')}</textarea></label>
                    <div class="fx-line-total"></div>
                    <button type="button" class="primary-btn secondary-btn fx-remove" data-remove="${index}">明細を削除</button>
                </section>`;
            }).join('');
            el('add').disabled = state.group.lines.length >= 20;
            updateCalculation();
            updateControls();
        }
        function addLine() {
            if (state.locked || state.group.lines.length >= 20) return;
            const defaults = optionsFrom('buy-status-entry');
            state.group.lines.push({ type: 'purchase', item: '', quantity: '', usd: '', yen: '', status: defaults[0] || '',
                category: document.getElementById('buy-category')?.value || '', account: '', isStock: 0, receipt: 0, note: '' });
            invalidatePreview();
            renderLines();
        }
        function readCommon() {
            for (const [field, id] of [['date', 'date'], ['vendor', 'vendor'], ['payment', 'payment'], ['usdTotal', 'usd-total'], ['yenTotal', 'yen-total']]) state.group[field] = el(id).value.trim();
        }
        function validateGroup() {
            readCommon();
            const amounts = calculateFxAmounts_(state.group);
            if (!/^\d{4}-\d{2}-\d{2}$/.test(state.group.date) || !state.group.vendor || !state.group.payment) throw new Error('購入日・購入先・支払方法を入力してください。');
            state.group.lines.forEach((line, index) => {
                if (!Number.isSafeInteger(Number(line.quantity)) || Number(line.quantity) < 1) throw new Error(`商品${index + 1}の数量は1以上の整数で入力してください。`);
                const data = { ...line, date: state.group.date, vendor: state.group.vendor, payment: state.group.payment, quantity: Number(line.quantity), price: amounts.prices[index] };
                const error = validateData(line.type === 'purchase' ? 'T_仕入' : 'T_経費', data);
                if (error) throw new Error(`商品${index + 1}：${error}`);
            });
            return amounts;
        }
        async function submit() {
            if (busy || state.done) return;
            const previouslyLocked = state.locked;
            try {
                if (!state.locked) {
                    const amounts = validateGroup();
                    const signature = JSON.stringify(state.group);
                    if (previewSignature !== signature) {
                        el('preview').innerHTML = '<h3 class="fx-section-title">登録内容の確認</h3>' + state.group.lines.map((line, index) => `<p>${index + 1}. ${escapeHtml(line.item)} ／ ${line.type === 'purchase' ? '仕入' : '経費'} ／ ${escapeHtml(line.status)} ／ ${line.quantity}個 ／ ${amounts.prices[index].toLocaleString()}円</p>`).join('') + `<p>合計：${amounts.yenTotal.toLocaleString()}円</p>`;
                        el('preview').hidden = false;
                        el('confirm').textContent = 'この内容でまとめて登録';
                        previewSignature = signature;
                        return;
                    }
                    if (!state.group.id) state.group.id = Array.from(crypto.getRandomValues(new Uint8Array(16)), value => value.toString(16).padStart(2, '0')).join('');
                    state.locked = true;
                }
                // 保存失敗時は送信しない。通信切断後も同じ番号・同じ内容で結果を確認する。
                try { persist(); } catch (_) { state.locked = previouslyLocked; throw new Error('入力内容を端末に保存できないため登録を中止しました。ブラウザの保存設定を確認してください。'); }
                busy = true;
                updateControls();
                showResults('決済番号：' + state.group.id + '　登録結果を確認中…');
                const response = await fetchAPI('registerForeignPurchase', { group: state.group });
                if (response.status === 'rejected') {
                    if (response.safeToEdit) { state.locked = false; delete state.group.id; invalidatePreview(); }
                    showResults(response.message);
                } else if (response.status === 'success' || response.status === 'review') {
                    state.done = response.status === 'success';
                    state.results = response.results;
                    showResults(state.done ? 'すべて登録しました。' : '要確認の明細があります。仕入・経費の履歴と備考の決済番号を照合してください。登録済み明細は再送しません。');
                } else throw new Error(response.message || '登録結果を取得できませんでした。');
                try { persist(); } catch (_) { /* 送信前の決済番号・内容は保存済み */ }
                if (state.done || response.status === 'review') {
                    try { await initSystem(); } catch (_) { showToast('履歴の更新に失敗しました。同期ボタンで再取得してください。', 'error'); }
                }
            } catch (error) {
                showResults(error.message + (state.locked ? '　同じ決済の結果を確認してください。新しい番号で登録し直さないでください。' : ''));
            } finally {
                busy = false;
                updateControls();
            }
        }
        document.querySelectorAll('.fx-open-btn').forEach(button => button.addEventListener('click', () => {
            opener = button;
            for (const [field, id] of [['date', 'date'], ['vendor', 'vendor'], ['payment', 'payment'], ['usdTotal', 'usd-total'], ['yenTotal', 'yen-total']]) el(id).value = state.group[field] || '';
            const names = [...(currentMasters['M_商品'] || []), ...(currentMasters['M_経費品名'] || [])].map(row => row['品名']).filter(Boolean);
            el('item-list').innerHTML = [...new Set(names)].map(name => `<option value="${escapeHtml(name)}"></option>`).join('');
            if (!state.group.lines.length && !state.locked) addLine();
            renderLines();
            showResults(state.done ? '登録済みの決済です。' : (state.locked ? '送信済みの決済です。同じ決済の結果を確認できます。' : ''));
            modal.classList.add('active');
            el('close').focus();
        }));
        el('close').onclick = () => { if (!busy) { modal.classList.remove('active'); opener?.focus(); } };
        el('add').onclick = () => { addLine(); try { persist(); } catch (_) {} };
        el('lines').addEventListener('click', event => {
            const button = event.target.closest('[data-remove]');
            if (!button || state.locked) return;
            state.group.lines.splice(Number(button.dataset.remove), 1);
            invalidatePreview(); renderLines(); try { persist(); } catch (_) {}
        });
        el('inputs').addEventListener('input', event => {
            if (state.locked) return;
            const card = event.target.closest('[data-index]');
            if (card && event.target.dataset.field) {
                const line = state.group.lines[Number(card.dataset.index)];
                line[event.target.dataset.field] = event.target.type === 'checkbox' ? (event.target.checked ? 1 : 0) : event.target.value;
                if (event.target.dataset.field === 'type') {
                    line.status = optionsFrom(line.type === 'purchase' ? 'buy-status-entry' : 'exp-status-entry')[0] || '';
                    renderLines();
                }
            } else readCommon();
            invalidatePreview(); updateCalculation(); showResults(); try { persist(); } catch (_) {}
        });
        el('confirm').onclick = submit;
        el('new').onclick = () => {
            if (busy || (state.locked && !state.done) || !confirm(state.done ? '次の決済を入力しますか？' : '入力内容を破棄しますか？')) return;
            try { localStorage.removeItem(storageKey); } catch (_) { showResults('保存済み入力を削除できません。'); return; }
            state = { group: blank(), locked: false, done: false, results: [] };
            invalidatePreview();
            for (const [field, id] of [['date', 'date'], ['vendor', 'vendor'], ['payment', 'payment'], ['usdTotal', 'usd-total'], ['yenTotal', 'yen-total']]) el(id).value = state.group[field];
            addLine(); showResults();
        };
    }
    setupForeignPurchase();

    // --- Registration Basket Logic (Proposal 38) ---

    /**
     * 現在の入力内容を買い物カゴに追加
     */
    window.addItemToBasket = function (type) {
        const prefix = type === 'purchase' ? 'buy' : 'exp';
        const data = {};

        // 共通・個別項目の取得
        if (type === 'purchase') {
            data.date = document.getElementById('purchase-date').value;
            data.status = document.getElementById('buy-status-entry').value;
            data.vendor = document.getElementById('buy-vendor').value;
            data.item = document.getElementById('buy-item').value;
            data.price = parseFloat(document.getElementById('buy-price').value) || 0;
            data.quantity = parseFloat(document.getElementById('buy-quantity').value) || 0;
            data.payment = document.getElementById('buy-payment').value;
            data.category = document.getElementById('buy-category').value;
            data.note = document.getElementById('buy-note').value;
        } else {
            data.date = document.getElementById('expense-date').value;
            data.status = document.getElementById('exp-status-entry').value;
            data.account = document.getElementById('exp-account').value;
            data.vendor = document.getElementById('exp-vendor').value;
            data.item = document.getElementById('exp-item').value;
            data.price = parseFloat(document.getElementById('exp-price').value) || 0;
            data.quantity = parseFloat(document.getElementById('exp-quantity').value) || 0;
            data.payment = document.getElementById('exp-payment').value;
            data.receipt = document.getElementById('exp-receipt').checked ? 1 : 0;
            data.isStock = document.getElementById('exp-is-stock').checked ? 1 : 0;
            data.note = document.getElementById('exp-note').value;
        }

        // 基本バリデーション
        if (!data.item || data.price <= 0 || data.quantity <= 0) {
            alert("品名、価格、数量を正しく入力してください。");
            return;
        }

        // カゴに追加
        window.registrationBasket.push({ type, data });

        // 入力欄のクリア（共通項目以外）
        document.getElementById(`${prefix}-item`).value = "";
        document.getElementById(`${prefix}-price`).value = "";
        document.getElementById(`${prefix}-quantity`).value = "";
        document.getElementById(`${prefix}-note`).value = "";
        const previewEl = document.getElementById(`${prefix}-item-preview`);
        if (previewEl) previewEl.innerHTML = "";

        renderBasket();

        // フィードバック
        const btn = document.getElementById(`${prefix}-add-basket`);
        if (btn) {
            const originalText = btn.innerHTML;
            btn.innerHTML = '<ion-icon name="checkmark-done-outline"></ion-icon> 追加しました';
            setTimeout(() => { btn.innerHTML = originalText; }, 1500);
        }
    };

    /**
     * 買い物カゴの内容をレンダリング
     */
    function renderBasket() {
        const sections = ['purchase', 'expense'];
        const basket = window.registrationBasket;
        const total = basket.reduce((sum, item) => sum + item.data.price, 0);

        sections.forEach(tab => {
            const container = document.getElementById(`basket-section-${tab}`);
            const list = document.getElementById(`basket-list-${tab}`);
            const totalEl = document.getElementById(`basket-total-${tab}`);
            if (!container || !list || !totalEl) return;

            if (basket.length === 0) {
                container.classList.remove('active');
                return;
            }

            container.classList.add('active');
            totalEl.textContent = `¥${total.toLocaleString()}`;

            list.innerHTML = basket.map((item, index) => `
                <div class="basket-item">
                    <div class="basket-item-info">
                        <div class="basket-item-name">
                            <span class="basket-badge ${item.type}">${item.type === 'purchase' ? '仕入' : '経費'}</span>
                            ${item.data.item}
                        </div>
                        <div class="basket-item-details">
                            <span>単価: ¥${Math.round(item.data.price / item.data.quantity).toLocaleString()}</span>
                            <span>数量: ${item.data.quantity}</span>
                        </div>
                    </div>
                    <div class="basket-item-price">¥${item.data.price.toLocaleString()}</div>
                    <div class="basket-item-remove" onclick="removeFromBasket(${index})">
                        <ion-icon name="close-circle-outline"></ion-icon>
                    </div>
                </div>
            `).join('');
        });
    }

    /**
     * カゴからアイテムを削除
     */
    window.removeFromBasket = function (index) {
        window.registrationBasket.splice(index, 1);
        renderBasket();
    };

    /**
     * カゴを空にする
     */
    window.clearBasket = function () {
        if (!confirm("リストの内容をすべて破棄しますか？")) return;
        window.registrationBasket = [];
        renderBasket();
    };

    /**
     * 一括登録を実行
     */
    window.submitBulkRegistration = async function () {
        if (window.registrationBasket.length === 0) return;

        const total = window.registrationBasket.reduce((sum, item) => sum + item.data.price, 0);
        if (!confirm(`${window.registrationBasket.length}件の明細（合計 ¥${total.toLocaleString()}）を一括登録しますか？`)) return;

        setLoading(true, "一括登録中...");
        try {
            const result = await fetchAPI('registerBulk', {
                transactions: window.registrationBasket,
                scope: 'all'
            });

            if (result.status === 'success') {
                showToast("一括登録が完了しました。");
                window.registrationBasket = [];
                renderBasket();

                // 履歴・在庫データの更新フローを既存処理(handleSubmission)と同期
                if (result.historyData && result.historyData.rawData) {
                    // 生データをマージ
                    lastRawData = Object.assign({}, lastRawData, result.historyData.rawData);
                    // データを加工・再集計
                    const processed = processClientData(lastRawData);
                    lastHistoryData = processed;
                    // 各履歴画面を描画
                    renderAllHistory(processed);
                    // 在庫一覧UIを更新
                    refreshInventoryUI();
                }

                // 新規マスタ追加があった場合
                if (result.masterAdded) {
                    initSystem();
                }

                // フォームをリセット（共通項目も含む）
                const buyVendor = document.getElementById('buy-vendor');
                if (buyVendor) buyVendor.value = "";
                const expVendor = document.getElementById('exp-vendor');
                if (expVendor) expVendor.value = "";

            } else {
                showToast("登録エラー: " + result.message, 'error');
            }
        } catch (e) {
            console.error(e);
            showToast("通信エラー: " + e.toString(), 'error');
        } finally {
            setLoading(false);
        }
    };

    // イベントリスナーの設定
    const buyAddBtn = document.getElementById('buy-add-basket');
    if (buyAddBtn) buyAddBtn.onclick = () => window.addItemToBasket('purchase');

    const expAddBtn = document.getElementById('exp-add-basket');
    if (expAddBtn) expAddBtn.onclick = () => window.addItemToBasket('expense');

    document.querySelectorAll('.bulk-submit-btn').forEach(btn => {
        btn.onclick = window.submitBulkRegistration;
    });
    document.querySelectorAll('.bulk-clear-btn').forEach(btn => {
        btn.onclick = window.clearBasket;
    });

    /**
     * ---- スクロール位置に応じたトップへ戻るボタンの制御 (提案42) ----
     */
    const contentArea = document.getElementById('content-area');
    const scrollBtn = document.getElementById('scroll-to-top-btn');

    if (contentArea && scrollBtn) {
        contentArea.addEventListener('scroll', () => {
            if (contentArea.scrollTop > 300) {
                scrollBtn.classList.add('visible');
            } else {
                scrollBtn.classList.remove('visible');
            }
        });

        scrollBtn.addEventListener('click', () => {
            contentArea.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        });
    }

    function historyDay_(date) {
        const jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
        return jst.getUTCFullYear() + '/' + String(jst.getUTCMonth() + 1).padStart(2, '0') + '/' + String(jst.getUTCDate()).padStart(2, '0');
    }

    function historyDateInfo_(type, row) {
        const specs = {
            purchase: [['仕入', ['仕入日']], ['入庫', ['入庫日']]],
            expense: [['注文', ['注文日']], ['完了', ['完了日']]],
            manufacturing: [['開始', ['製造開始日']], ['完了', ['製造完了日']]],
            sales: [['開始', ['販売開始日', '取引開始日']], ['完了', ['取引完了日', '完了日']]]
        };
        const dates = (specs[type] || []).map(([label, keys]) => {
            for (const key of keys) {
                const date = parseRecentActivityDate_(row[key]);
                if (date) return { label, value: historyDay_(date) };
            }
            return { label, value: '' };
        });
        return { dates, searchDay: dates[1]?.value || dates[0]?.value || '' };
    }

    function purchaseCorrectionNotice_(response, id) {
        const inputIssue = ['INVALID_INPUT', 'NO_CHANGE', 'OUT_OF_RANGE', 'STALE_PREVIEW'].includes(response?.code);
        const expected = response?.status === 'rejected' && response?.code && response.code !== 'INFO_UNAVAILABLE' && response?.unchanged === true;
        const title = expected ? (inputIssue ? '訂正内容を確認してください。' : 'この仕入は訂正できません。') : '仕入情報を取得できませんでした。';
        const message = expected ? response.message : '画面を開き直してください。';
        return title + '\n' + message + '\n仕入ID：' + id + '\nこの操作ではデータを変更していません。';
    }

    function renderPurchaseCorrection_() {
        const state = purchaseCorrectionState;
        const el = key => document.getElementById('purchase-correction-' + key);
        if (!el('modal')) return;
        const locked = state.busy || Boolean(state.pending) || Boolean(state.done) || !state.record;
        for (const key of ['quantity', 'price', 'reason', 'preview']) el(key).disabled = locked;
        el('save').hidden = !state.preview && !state.pending;
        el('save').disabled = state.busy || Boolean(state.done);
        el('save').textContent = state.pending ? '送信済みの訂正結果を確認' : 'この内容で訂正する';
        el('resume').hidden = !state.pending && !state.storageError;
    }

    function closePurchaseCorrection_() {
        purchaseCorrectionState.sequence++;
        document.getElementById('purchase-correction-modal')?.classList.remove('active');
    }

    async function openPurchaseCorrection_(id) {
        const state = purchaseCorrectionState;
        if (state.busy) return;
        const el = key => document.getElementById('purchase-correction-' + key);
        const sequence = ++state.sequence;
        state.id = state.pending?.id || id; state.record = null; state.preview = null; state.done = false;
        el('modal').classList.add('active'); el('close').focus();
        el('current').textContent = '仕入ID：' + state.id;
        el('message').textContent = ''; el('preview-content').hidden = true;
        for (const key of ['quantity', 'price', 'reason']) el(key).value = '';
        if (state.pending) {
            for (const key of ['quantity', 'price', 'reason']) el(key).value = state.pending[key];
            el('message').textContent = '送信済みの訂正を確認してください。訂正番号：' + state.pending.requestId;
            renderPurchaseCorrection_(); return;
        }
        if (state.storageError) {
            el('message').textContent = '保存された訂正情報を読み出せません。新しい訂正を送らず、仕入・入庫と在庫を確認してください。';
            renderPurchaseCorrection_(); return;
        }
        state.busy = true; el('message').textContent = '入庫ロットが未使用か確認中...'; renderPurchaseCorrection_();
        try {
            const response = await fetchAPI('getPurchaseCorrectionContext', { id: state.id });
            if (sequence !== state.sequence) return;
            if (response.status !== 'success') {
                el('message').textContent = purchaseCorrectionNotice_(response, state.id); return;
            }
            state.record = response.data.record;
            el('quantity').value = response.data.quantity; el('price').value = response.data.price;
            el('current').textContent = '仕入ID：' + state.id + ' ／ ' + state.record['品名'] + ' ／ 入庫ID：' + response.data.lotId;
            el('message').textContent = '';
        } catch (_) { if (sequence === state.sequence) el('message').textContent = purchaseCorrectionNotice_(null, state.id); }
        finally { state.busy = false; renderPurchaseCorrection_(); }
    }

    async function previewPurchaseCorrection_() {
        const state = purchaseCorrectionState;
        if (state.busy || state.pending || !state.record || state.done) return;
        const el = key => document.getElementById('purchase-correction-' + key);
        const correction = { id: state.id, quantity: el('quantity').value, price: el('price').value, reason: el('reason').value };
        const sequence = ++state.sequence;
        state.preview = null; state.busy = true; el('preview-content').hidden = true; el('message').textContent = '訂正内容と未使用状態を確認中...'; renderPurchaseCorrection_();
        try {
            const response = await fetchAPI('previewPurchaseCorrection', { correction });
            if (sequence !== state.sequence) return;
            if (response.status !== 'success') {
                el('message').textContent = purchaseCorrectionNotice_(response, state.id);
                if (response.status === 'rejected' && !['INVALID_INPUT', 'NO_CHANGE', 'OUT_OF_RANGE', 'STALE_PREVIEW'].includes(response.code)) state.record = null;
                return;
            }
            state.preview = response.data;
            const preview = state.preview;
            const diff = preview.after.quantity - preview.before.quantity;
            const lines = [ '対象ロット：' + preview.lotId,
                '数量・対象ロットの在庫：' + preview.before.quantity + ' → ' + preview.after.quantity + '（差分 ' + (diff > 0 ? '+' : '') + diff + '）',
                'この品名の現在庫合計：' + Number(preview.stockBefore).toLocaleString('ja-JP', { maximumFractionDigits: 10 }) + ' → ' + Number(preview.stockAfter).toLocaleString('ja-JP', { maximumFractionDigits: 10 }),
                '仕入総額：' + historyMoney_(preview.before.price) + ' → ' + historyMoney_(preview.after.price),
                '単価：' + historyMoney_(preview.before.unit) + ' → ' + historyMoney_(preview.after.unit),
                '理由：' + (preview.reason || '未記入') ];
            el('preview-content').innerHTML = '<ul>' + lines.map(line => '<li>' + escapeHtml(line) + '</li>').join('') + '</ul><p>訂正前後の値を備考へ追記します。元の備考と入庫IDは維持します。</p>';
            el('preview-content').hidden = false; el('message').textContent = '';
        } catch (_) { if (sequence === state.sequence) el('message').textContent = purchaseCorrectionNotice_(null, state.id); }
        finally { state.busy = false; renderPurchaseCorrection_(); }
    }

    async function savePurchaseCorrection_() {
        const state = purchaseCorrectionState;
        if (state.busy || state.done || (!state.pending && !state.preview)) return;
        const el = key => document.getElementById('purchase-correction-' + key);
        if (!state.pending) {
            const preview = state.preview;
            const pending = { requestId: Array.from(crypto.getRandomValues(new Uint8Array(16)), byte => byte.toString(16).padStart(2, '0')).join(''),
                id: state.id, quantity: String(preview.after.quantity), price: String(preview.after.price), reason: preview.reason, token: preview.token };
            try { localStorage.setItem('inventory_purchase_correction_pending', JSON.stringify(pending)); }
            catch (_) { el('message').textContent = '送信内容を端末に保存できないため訂正していません。'; return; }
            state.pending = pending;
        }
        state.busy = true; el('message').textContent = '訂正結果を確認中...'; renderPurchaseCorrection_();
        const sequence = ++state.sequence;
        try {
            const response = await fetchAPI('correctPurchase', { correction: state.pending });
            if (response.status === 'success' || (response.status === 'rejected' && response.safeToEdit === true)) {
                try { localStorage.removeItem('inventory_purchase_correction_pending'); } catch (_) { /* 同じ番号での再確認は再更新しない */ }
                state.pending = null; state.preview = null; el('preview-content').hidden = true;
                if (response.status === 'success') {
                    state.done = true; el('message').textContent = '数量・仕入総額と在庫を訂正しました。';
                    document.getElementById('transaction-detail-modal')?.classList.remove('active');
                    try {
                        if (!Array.isArray(response.inventorySummary)) throw Error('在庫集計がありません。');
                        lastRawData['T_在庫集計'] = response.inventorySummary;
                        saveHistoryToCache(lastRawData);
                        syncShippingInventory_(response.inventorySummary);
                        const data = await initSystem('purchase', { skipMasters: true });
                        if (!data?.historyData) throw Error('仕入履歴がありません。');
                        if (document.getElementById('history-search-modal')?.classList.contains('active')) await executeHistorySearch();
                    } catch (_) { el('message').textContent = '訂正は完了しました。画面の再取得に失敗したため、再読み込みしてください。'; }
                } else {
                    el('message').textContent = purchaseCorrectionNotice_(response, state.id);
                    if (response.status === 'rejected' && !['INVALID_INPUT', 'NO_CHANGE', 'OUT_OF_RANGE', 'STALE_PREVIEW'].includes(response.code)) state.record = null;
                }
            } else if (sequence === state.sequence) {
                const message = response.status === 'review' ? response.message : '訂正結果を確認できませんでした。新しい番号で訂正せず、同じ番号の結果を確認してください。';
                el('message').textContent = message + '\n仕入ID：' + state.id + '\n訂正番号：' + state.pending.requestId;
            }
        } catch (_) {
            el('message').textContent = '通信が途切れました。新しい番号で訂正せず「送信済みの訂正結果を確認」を押してください。訂正番号：' + state.pending.requestId;
        } finally { state.busy = false; renderPurchaseCorrection_(); }
    }

    function setupPurchaseCorrection_() {
        const el = key => document.getElementById('purchase-correction-' + key);
        if (!el('modal')) return;
        try {
            const saved = localStorage.getItem('inventory_purchase_correction_pending');
            if (saved) {
                const pending = JSON.parse(saved);
                if (!/^[a-f0-9]{32}$/.test(pending?.requestId || '') || !/^P[0-9]+$/.test(pending?.id || '') || typeof pending.token !== 'string' || typeof pending.reason !== 'string') throw Error('invalid');
                purchaseCorrectionState.pending = pending;
            }
        } catch (_) { purchaseCorrectionState.storageError = true; }
        for (const key of ['quantity', 'price', 'reason']) el(key).addEventListener('input', () => {
            if (purchaseCorrectionState.busy || purchaseCorrectionState.pending) return;
            purchaseCorrectionState.preview = null; el('preview-content').hidden = true; el('message').textContent = ''; renderPurchaseCorrection_();
        });
        el('preview').addEventListener('click', previewPurchaseCorrection_);
        el('save').addEventListener('click', savePurchaseCorrection_);
        el('close').addEventListener('click', closePurchaseCorrection_);
        el('resume').addEventListener('click', () => openPurchaseCorrection_(purchaseCorrectionState.pending?.id || ''));
        el('modal').addEventListener('click', event => { if (event.target === el('modal')) closePurchaseCorrection_(); });
        renderPurchaseCorrection_();
    }

    function renderShippingChange_() {
        const state = shippingChangeState;
        const el = key => document.getElementById('shipping-change-' + key);
        if (!el('method')) return;
        el('method').disabled = el('fee').disabled = state.busy || Boolean(state.pending) || Boolean(state.done) || !state.record;
        el('preview').disabled = state.busy || Boolean(state.pending) || Boolean(state.done) || !state.record;
        el('save').hidden = !state.preview && !state.pending;
        el('save').disabled = state.busy || Boolean(state.done);
        el('save').textContent = state.pending ? '送信済みの変更結果を確認' : 'この内容で変更する';
        el('resume').hidden = !state.pending && !state.storageError;
    }

    function closeShippingChange_() {
        shippingChangeState.sequence++;
        document.getElementById('shipping-change-modal')?.classList.remove('active');
    }

    async function openShippingChange_(id) {
        const state = shippingChangeState;
        if (state.busy) return;
        const el = key => document.getElementById('shipping-change-' + key);
        state.id = state.pending?.id || id;
        state.record = null; state.preview = null; state.done = false;
        const sequence = ++state.sequence;
        el('modal').classList.add('active');
        el('preview-content').hidden = true;
        el('method').replaceChildren(); el('fee').value = '';
        el('current').textContent = '販売ID：' + state.id;
        el('message').textContent = '';
        el('close').focus();
        if (state.pending) {
            el('method').appendChild(new Option(state.pending.shipping || '未指定', state.pending.shipping));
            el('fee').value = state.pending.fee;
            el('message').textContent = '送信済みの変更を確認してください。変更番号：' + state.pending.requestId;
            renderShippingChange_(); return;
        }
        if (state.storageError) {
            el('message').textContent = '保存された変更情報を読み出せません。新しい変更を送らず、販売と在庫の状態を確認してください。';
            renderShippingChange_(); return;
        }
        state.busy = true; el('message').textContent = '現在の販売情報を確認中...'; renderShippingChange_();
        try {
            const response = await fetchAPI('getShippingChangeOptions', { id: state.id });
            if (sequence !== state.sequence) return;
            if (response.status !== 'success') throw Error(response.message || '販売情報を取得できませんでした。');
            state.record = response.data.record; state.options = response.data.options;
            const old = state.record['発送方法'] || '';
            const names = [...new Set(['', old, ...state.options.map(option => option.shipping)])];
            el('method').replaceChildren(...names.map(name => new Option(name || '未指定（発送資材なし）', name)));
            el('method').value = old; el('fee').value = state.record['送料'] ?? 0;
            el('current').textContent = '販売ID：' + state.id + ' ／ 現在：' + (old || '未指定') + ' ／ 送料：' + historyMoney_(state.record['送料']);
            el('message').textContent = '';
        } catch (error) {
            if (sequence === state.sequence) el('message').textContent = error.message;
        } finally { state.busy = false; renderShippingChange_(); }
    }

    async function previewShippingChange_() {
        const state = shippingChangeState;
        if (state.busy || state.pending || !state.record || state.done) return;
        const el = key => document.getElementById('shipping-change-' + key);
        const sequence = ++state.sequence;
        state.busy = true; state.preview = null;
        el('preview-content').hidden = true; el('message').textContent = '在庫と変更内容を確認中...'; renderShippingChange_();
        try {
            const response = await fetchAPI('previewShippingChange', { id: state.id, shipping: el('method').value, fee: el('fee').value });
            if (sequence !== state.sequence) return;
            if (response.status !== 'success') throw Error(response.message || '変更内容を確認できませんでした。');
            state.preview = response.data;
            const preview = state.preview;
            const lines = [ '発送方法：' + (preview.oldShipping || '未指定') + ' → ' + (preview.shipping || '未指定'),
                '送料：' + historyMoney_(preview.oldFee) + ' → ' + historyMoney_(preview.fee) ];
            for (const [label, entries] of [['返却', preview.returns], ['引当', preview.consumes]]) {
                for (const entry of entries) lines.push(label + '：' + entry.name + ' × ' + entry.quantity + '個（元ロット：' + entry.lot + '）');
            }
            if (!preview.returns.length && !preview.consumes.length) lines.push('発送資材の在庫増減はありません。');
            el('preview-content').innerHTML = '<ul>' + lines.map(line => '<li>' + escapeHtml(line) + '</li>').join('') + '</ul><p>販売商品自体の引当と帳簿は変更しません。</p>';
            el('preview-content').hidden = false; el('message').textContent = '';
        } catch (error) { if (sequence === state.sequence) el('message').textContent = error.message; }
        finally { state.busy = false; renderShippingChange_(); }
    }

    function syncShippingInventory_(raw) {
        if (!Array.isArray(raw) || !Array.isArray(raw[0])) throw Error('最新の在庫集計を取得できませんでした。');
        const previous = new Map((currentMasters['T_在庫集計'] || []).map(row => [row['品名'], row]));
        const products = new Map((currentMasters['M_商品'] || []).map(row => [row['品名'], row]));
        currentMasters['T_在庫集計'] = convertRawToObjects(raw).map(row => {
            const merged = { ...previous.get(row['品名']), ...row };
            const product = products.get(row['品名']);
            if (product) for (const key of ['商品ID', '保管場所', 'QR/バーコード']) merged[key] = product[key] || merged[key] || '';
            return merged;
        });
        saveMastersToCache(currentMasters);
        // 保存中に棚卸へ移動していても、入力中のカードを作り直さない。
        if (!isStocktakeMode) applyStockFilters();
        if (document.querySelector('.nav-item.active')?.getAttribute('data-target') === 'ledger') renderLedger();
        else renderInventoryAlerts();
    }

    async function saveShippingChange_() {
        const state = shippingChangeState;
        if (state.busy || state.done || (!state.preview && !state.pending)) return;
        const el = key => document.getElementById('shipping-change-' + key);
        if (!state.pending) {
            const preview = state.preview;
            const requestId = Array.from(crypto.getRandomValues(new Uint8Array(16)), byte => byte.toString(16).padStart(2, '0')).join('');
            const pending = { requestId, id: state.id, shipping: preview.shipping, fee: String(preview.fee), token: preview.token };
            try { localStorage.setItem('inventory_shipping_change_pending', JSON.stringify(pending)); }
            catch (_) { el('message').textContent = '送信情報を端末に保存できないため、変更していません。'; return; }
            state.pending = pending;
        }
        state.busy = true; el('message').textContent = '変更結果を確認中...'; renderShippingChange_();
        const sequence = ++state.sequence;
        try {
            const response = await fetchAPI('changeShipping', { change: state.pending });
            if (response.status === 'success' || (response.status === 'rejected' && response.safeToEdit === true)) {
                try { localStorage.removeItem('inventory_shipping_change_pending'); } catch (_) { /* 同じ番号なら再確認しても再更新しない */ }
                state.pending = null; state.preview = null;
                if (response.status === 'success') {
                    state.done = true;
                    el('message').textContent = '発送方法と送料を変更しました。';
                    try {
                        const data = await initSystem('sales', { skipMasters: true });
                        syncShippingInventory_(data?.historyData?.rawData?.['T_在庫集計']);
                    }
                    catch (_) { el('message').textContent = '変更は完了しました。履歴の再取得に失敗したため、画面を再読み込みしてください。'; }
                } else el('message').textContent = response.message;
                el('preview-content').hidden = true;
            } else if (sequence === state.sequence) {
                el('message').textContent = (response.message || '結果が要確認です。') + ' 変更番号：' + state.pending.requestId;
            }
        } catch (_) {
            el('message').textContent = '通信が途切れました。新しい変更を送らず「送信済みの変更結果を確認」を押してください。変更番号：' + state.pending.requestId;
        } finally { state.busy = false; renderShippingChange_(); }
    }

    function setupShippingChange_() {
        const el = key => document.getElementById('shipping-change-' + key);
        if (!el('modal')) return;
        try {
            const saved = localStorage.getItem('inventory_shipping_change_pending');
            if (saved) {
                const pending = JSON.parse(saved);
                if (!/^[a-f0-9]{32}$/.test(pending?.requestId || '') || !/^S[0-9]+$/.test(pending?.id || '') || typeof pending.shipping !== 'string' || typeof pending.token !== 'string') throw Error('invalid');
                shippingChangeState.pending = pending;
            }
        } catch (_) { shippingChangeState.storageError = true; }
        const invalidate = () => {
            if (shippingChangeState.busy || shippingChangeState.pending) return;
            shippingChangeState.preview = null; el('preview-content').hidden = true; el('message').textContent = ''; renderShippingChange_();
        };
        el('method').addEventListener('change', () => {
            const state = shippingChangeState;
            if (!state.record || state.pending || state.busy) return;
            const option = state.options.find(option => option.shipping === el('method').value);
            el('fee').value = el('method').value === (state.record['発送方法'] || '') ? (state.record['送料'] ?? 0) : (option?.fee ?? 0);
            invalidate();
        });
        el('fee').addEventListener('input', invalidate);
        el('preview').addEventListener('click', previewShippingChange_);
        el('save').addEventListener('click', saveShippingChange_);
        el('close').addEventListener('click', closeShippingChange_);
        el('resume').addEventListener('click', () => openShippingChange_(shippingChangeState.pending?.id || ''));
        el('modal').addEventListener('click', event => { if (event.target === el('modal')) closeShippingChange_(); });
        renderShippingChange_();
    }

    function updateVersionDisplay_(backendVersion) {
        const el = document.getElementById('app-version');
        if (!el) return;
        el.textContent = '画面 ' + APP_VERSION + ' ／ GAS ' + (backendVersion || '未確認');
        el.title = backendVersion && backendVersion !== APP_VERSION ? '画面とGASのバージョンが異なります。反映状況を確認してください。' : '';
    }

    function formatFxRate_(amounts) {
        // 整数の円・セントから表示を作り、1.005等の浮動小数点の丸め誤差を避ける。
        const numerator = BigInt(amounts.yenTotal) * 10000n;
        const denominator = BigInt(amounts.usdTotal);
        const rounded = (2n * numerator + denominator) / (2n * denominator);
        const rate = String(rounded / 100n) + '.' + String(rounded % 100n).padStart(2, '0');
        return '参考レート：1 USD ≈ ' + rate + '円（表示のみ四捨五入）';
    }

    function historyMoney_(value) {
        if (value === '' || value === undefined || value === null || !Number.isFinite(Number(value))) return '未記録';
        return '¥' + Number(value).toLocaleString('ja-JP', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }

    function historySummary_(type, row) {
        const quantity = Number(row['数量']);
        const lines = [];
        if (type === 'purchase') lines.push('仕入単価：' + historyMoney_(row['単価']));
        if (type === 'sales') {
            const price = row['販売価格'] ?? row['価格'] ?? row['合計金額'];
            lines.push('販売単価：' + historyMoney_(price !== '' && price != null && quantity > 0 ? Number(price) / quantity : null));
            lines.push('発送方法：' + (row['発送方法'] || '未記録'), '送料：' + historyMoney_(row['送料']));
        }
        if (type === 'manufacturing') {
            lines.push('製造単価：' + historyMoney_(row['単価']));
            lines.push('製造原価合計：' + historyMoney_(row['単価'] !== '' && row['単価'] != null && quantity > 0 ? Number(row['単価']) * quantity : null));
        }
        return lines;
    }

    function clearHistorySearch_() {
        historySearchSequence++;
        for (const key of ['keyword', 'from', 'to']) document.getElementById('history-search-' + key).value = '';
        document.getElementById('history-search-target').value = 'all';
        document.getElementById('history-search-results').innerHTML = '<div class="empty-history">条件を指定して検索してください</div>';
        document.getElementById('history-search-results').scrollTop = 0;
        const count = document.getElementById('history-search-count');
        if (count) count.textContent = '';
    }

    function closeTransactionDetail_() {
        transactionDetailSequence++;
        document.getElementById('transaction-detail-modal')?.classList.remove('active');
        transactionDetailOpener?.focus();
    }

    async function openTransactionDetail_(type, id, opener = document.activeElement) {
        const modal = document.getElementById('transaction-detail-modal');
        const content = document.getElementById('transaction-detail-content');
        if (!modal || !content) return;
        const sequence = ++transactionDetailSequence;
        transactionDetailOpener = opener;
        modal.classList.add('active');
        content.textContent = '読み込み中...';
        document.getElementById('transaction-detail-close-btn')?.focus();
        try {
            const response = await fetchAPI('getTransactionDetail', { type, id: String(id) });
            if (sequence !== transactionDetailSequence) return;
            if (response.status !== 'success') throw new Error(response.message || '詳細を取得できませんでした。');
            const data = response.data;
            const row = data.record;
            const info = historyDateInfo_(type, row);
            const field = (label, value) => '<div><dt>' + escapeHtml(label) + '</dt><dd>' + escapeHtml(value === '' || value == null ? '未記録' : String(value)) + '</dd></div>';
            let html = '<h3>' + escapeHtml(row['品名'] || row['摘要'] || '') + '</h3><p>閲覧専用</p><dl class="transaction-detail-fields">';
            html += field('ID', id) + field('ステータス', row['ステータス']) + field('数量', row['数量']);
            for (const date of info.dates) html += field(date.label, date.value || '未記録');
            for (const text of historySummary_(type, row)) {
                const split = text.indexOf('：'); html += field(text.slice(0, split), text.slice(split + 1));
            }
            for (const key of ['価格', '合計金額', '合計単価', '仕入先', '購入先', '売先', '支払方法', '勘定科目', '区分', '管理区分', '追跡番号', '備考']) {
                if (Object.hasOwn(row, key)) html += field(key === '合計単価' ? '原価（総額）' : key, row[key]);
            }
            if (type === 'manufacturing') {
                for (const key of ['製造着手日', 'テスト開始日', '梱包開始日']) {
                    const parsed = parseRecentActivityDate_(row[key]);
                    html += field(key, parsed ? historyDay_(parsed) : '未記録');
                }
            }
            html += '</dl>';
            if (type === 'purchase' && ['入庫済み', '入庫済', '完了'].includes(String(row['ステータス'] || '').trim())) {
                html += '<button type="button" class="history-detail-link" id="purchase-correction-open">数量・仕入総額を訂正</button><p>未使用の入庫か確認してから訂正します。</p>';
            }
            if (type === 'manufacturing') {
                html += '<h3>実際の部材引当・取消記録</h3><p>出庫はマイナス、返却はプラスで表示します。現在のBOMとは別の記録です。</p>';
                html += data.materials.length ? '<div class="detail-table-wrap"><table><thead><tr><th>部材・区分</th><th>数量</th><th>単価</th><th>元ロット</th></tr></thead><tbody>' + data.materials.map(material => '<tr><td>' + escapeHtml(material['品名']) + '<br>' + escapeHtml(material['区分']) + '</td><td>' + escapeHtml(String(material['数量'])) + '</td><td>' + escapeHtml(historyMoney_(material['単価'])) + '</td><td>' + escapeHtml(material['引当元管理ID'] || '未記録') + '</td></tr>').join('') + '</tbody></table></div>' : '<p>引当記録が見つかりません。未使用と断定するものではありません。</p>';
                html += '<h3>現在のBOM（当時の構成を保証しません）</h3>';
                html += data.bom.length ? '<ul>' + data.bom.map(part => '<li>' + escapeHtml(part['部品']) + '：製品1個あたり ' + escapeHtml(String(part['数量'])) + '個</li>').join('') + '</ul>' : '<p>現在のBOMに登録がありません。</p>';
            }
            content.innerHTML = html;
            document.getElementById('purchase-correction-open')?.addEventListener('click', () => openPurchaseCorrection_(String(id)));
        } catch (error) {
            if (sequence === transactionDetailSequence) content.textContent = '詳細を表示できませんでした。' + error.message;
        }
    }

    function setupAnalysisAndSearchListeners() {
        updateVersionDisplay_();
        document.getElementById('history-search-clear-btn')?.addEventListener('click', clearHistorySearch_);
        document.getElementById('transaction-detail-close-btn')?.addEventListener('click', closeTransactionDetail_);
        document.getElementById('transaction-detail-modal')?.addEventListener('click', event => {
            if (event.target.id === 'transaction-detail-modal') closeTransactionDetail_();
        });
        document.addEventListener('keydown', event => {
            if (event.key === 'Escape' && document.getElementById('transaction-detail-modal')?.classList.contains('active')) closeTransactionDetail_();
        });
        const salesBtn = document.getElementById('menu-sales-analysis-btn');
        const historyBtn = document.getElementById('menu-history-search-btn');
        const salesModal = document.getElementById('sales-analysis-modal');
        const historyModal = document.getElementById('history-search-modal');

        if (salesBtn && salesModal) {
            salesBtn.addEventListener('click', () => {
                const sideMenu = document.getElementById('side-menu');
                const menuOverlay = document.getElementById('menu-overlay');
                if (sideMenu) sideMenu.classList.remove('open');
                if (menuOverlay) menuOverlay.classList.remove('visible');
                salesModal.classList.add('active');
                renderSalesAnalysis();
            });
        }
        if (historyBtn && historyModal) {
            historyBtn.addEventListener('click', () => {
                const sideMenu = document.getElementById('side-menu');
                const menuOverlay = document.getElementById('menu-overlay');
                if (sideMenu) sideMenu.classList.remove('open');
                if (menuOverlay) menuOverlay.classList.remove('visible');
                historyModal.classList.add('active');
            });
        }

        const salesCloseBtn = document.getElementById('sales-analysis-close-btn');
        if (salesCloseBtn && salesModal) salesCloseBtn.addEventListener('click', () => salesModal.classList.remove('active'));

        const historyCloseBtn = document.getElementById('history-search-close-btn');
        if (historyCloseBtn && historyModal) historyCloseBtn.addEventListener('click', () => historyModal.classList.remove('active'));

        const salesPeriodSelect = document.getElementById('sales-analysis-period-select');
        if (salesPeriodSelect) {
            salesPeriodSelect.addEventListener('change', () => {
                renderSalesAnalysis();
            });
        }

        const salesTypeSelect = document.getElementById('sales-analysis-type-select');
        if (salesTypeSelect) {
            salesTypeSelect.addEventListener('change', () => {
                renderSalesAnalysis();
            });
        }

        const executeSearchBtn = document.getElementById('history-search-execute-btn');
        if (executeSearchBtn) {
            executeSearchBtn.addEventListener('click', executeHistorySearch);
        }
        
        window.addEventListener('click', (event) => {
            if (salesModal && event.target === salesModal) salesModal.classList.remove('active');
            if (historyModal && event.target === historyModal) historyModal.classList.remove('active');
        });
    }

    function renderSalesAnalysis() {
        if (!lastHistoryData || !lastHistoryData.ledger) {
            if (typeof showToast === 'function') showToast('データを読み込んでいます...', 'info');
            return;
        }

        const period = document.getElementById('sales-analysis-period-select').value;
        const typeSelect = document.getElementById('sales-analysis-type-select');
        const targetType = typeSelect ? typeSelect.value : 'business';
        const ledger = lastHistoryData.ledger;

        let targetData = [];
        const now = new Date();
        const currentY = now.getFullYear();
        const currentM = now.getMonth() + 1;

        ledger.forEach(row => {
            const isPersonal = String(row['事業/個人']).trim() === '1';
            if (targetType === 'business' && isPersonal) return;
            if (targetType === 'personal' && !isPersonal) return;

            const d = new Date(row['日付']);
            if (isNaN(d.getTime())) return;
            const y = d.getFullYear();
            const m = d.getMonth() + 1;

            if (period === 'current_month') {
                if (y === currentY && m === currentM) targetData.push(row);
            } else if (period === 'last_month') {
                const prevM = currentM === 1 ? 12 : currentM - 1;
                const prevY = currentM === 1 ? currentY - 1 : currentY;
                if (y === prevY && m === prevM) targetData.push(row);
            } else if (period === 'current_year') {
                if (y === currentY) targetData.push(row);
            } else {
                targetData.push(row);
            }
        });

        // 1. 商品別ランキング
        const productMap = {};
        targetData.forEach(row => {
            let itemName = row['品名'] || '不明';
            const parts = itemName.split(' ');
            if (parts.length > 1) {
                itemName = parts.slice(1).join(' ');
            }
            if (!productMap[itemName]) productMap[itemName] = { sales: 0, count: 0 };
            const s = parseNumber(row['売上']);
            if (s > 0) {
                productMap[itemName].sales += s;
                productMap[itemName].count++;
            }
        });

        const productRank = Object.entries(productMap)
            .sort((a, b) => b[1].sales - a[1].sales)
            .slice(0, 10);

        const prodContainer = document.getElementById('analysis-product-ranking');
        if (prodContainer) {
            prodContainer.innerHTML = productRank.length > 0 ? productRank.map((item, idx) => `
                <div class="analysis-item">
                    <div class="analysis-item-rank">${idx + 1}</div>
                    <div class="analysis-item-name" title="${escapeHtml(item[0])}">${escapeHtml(item[0])}</div>
                    <div class="analysis-item-value">
                        <div class="analysis-item-amount">¥${Math.round(item[1].sales).toLocaleString()}</div>
                        <div class="analysis-item-sub">${item[1].count}件</div>
                    </div>
                </div>
            `).join('') : '<p class="empty-msg" style="text-align:center; padding: 10px;">データがありません</p>';
        }

        // 2. 売先別ランキング
        const buyerMap = {};
        targetData.forEach(row => {
            let buyerName = 'その他';
            const fullName = row['品名'] || '';
            const parts = fullName.split(' ');
            if (parts.length > 1) {
                buyerName = parts[0];
            }
            if (!buyerMap[buyerName]) buyerMap[buyerName] = { sales: 0, count: 0 };
            const s = parseNumber(row['売上']);
            if (s > 0) {
                buyerMap[buyerName].sales += s;
                buyerMap[buyerName].count++;
            }
        });

        const buyerRank = Object.entries(buyerMap)
            .sort((a, b) => b[1].sales - a[1].sales)
            .slice(0, 5);

        const buyerContainer = document.getElementById('analysis-buyer-ranking');
        if (buyerContainer) {
            buyerContainer.innerHTML = buyerRank.length > 0 ? buyerRank.map((item, idx) => `
                <div class="analysis-item">
                    <div class="analysis-item-rank">${idx + 1}</div>
                    <div class="analysis-item-name" title="${escapeHtml(item[0])}">${escapeHtml(item[0])}</div>
                    <div class="analysis-item-value">
                        <div class="analysis-item-amount">¥${Math.round(item[1].sales).toLocaleString()}</div>
                        <div class="analysis-item-sub">${item[1].count}件</div>
                    </div>
                </div>
            `).join('') : '<p class="empty-msg" style="text-align:center; padding: 10px;">データがありません</p>';
        }

        // 3. 利益率推移 と 4. 年間俯瞰 (月ごとの集計)
        const monthlySum = {};
        targetData.forEach(row => {
            const d = new Date(row['日付']);
            if (isNaN(d.getTime())) return;
            const y = d.getFullYear();
            const m = d.getMonth() + 1;
            const ym = `${y}年${m}月`;

            if (!monthlySum[ym]) monthlySum[ym] = { sales: 0, cost: 0, dateObj: new Date(y, m - 1, 1) };

            const s = parseNumber(row['売上']);
            const c = parseNumber(row['仕入']) + parseNumber(row['通信費']) + parseNumber(row['修繕費']) +
                      parseNumber(row['消耗品費']) + parseNumber(row['諸会費']) + parseNumber(row['支払手数料']) + parseNumber(row['雑費']);

            monthlySum[ym].sales += s;
            monthlySum[ym].cost += c;
        });

        const monthlyArr = Object.entries(monthlySum)
            .sort((a, b) => b[1].dateObj - a[1].dateObj);

        const profitContainer = document.getElementById('analysis-profit-margin');
        const yearlyContainer = document.getElementById('analysis-yearly-summary');

        if (profitContainer) {
            profitContainer.innerHTML = monthlyArr.length > 0 ? monthlyArr.map(item => {
                const s = item[1].sales;
                const c = item[1].cost;
                const p = s - c;
                const margin = s > 0 ? ((p / s) * 100).toFixed(1) : '0.0';
                return `
                    <div class="analysis-item">
                        <div class="analysis-item-name" style="width: 80px; flex: none;">${item[0]}</div>
                        <div class="analysis-item-value" style="flex: 1; display:flex; justify-content:space-between; align-items:center;">
                            <div style="font-size:12px; color:var(--text-muted);">粗利: ¥${Math.round(p).toLocaleString()}</div>
                            <div class="analysis-item-amount" style="color: ${margin >= 0 ? 'var(--accent-green)' : '#ef4444'};">${margin}%</div>
                        </div>
                    </div>
                `;
            }).join('') : '<p class="empty-msg" style="text-align:center; padding: 10px;">データがありません</p>';
        }

        if (yearlyContainer) {
            yearlyContainer.innerHTML = monthlyArr.length > 0 ? monthlyArr.map(item => {
                const s = item[1].sales;
                const c = item[1].cost;
                const p = s - c;
                return `
                    <tr>
                        <td style="font-size:11px;">${item[0]}</td>
                        <td>¥${Math.round(s).toLocaleString()}</td>
                        <td>¥${Math.round(c).toLocaleString()}</td>
                        <td style="color: ${p >= 0 ? 'inherit' : '#ef4444'}">¥${Math.round(p).toLocaleString()}</td>
                    </tr>
                `;
            }).join('') : '<tr><td colspan="4" style="text-align:center; padding: 10px;">データがありません</td></tr>';
        }
    }

    async function executeHistorySearch(exactId = null) {
        const sequence = ++historySearchSequence;
        const keyword = document.getElementById('history-search-keyword').value;
        const fromDate = document.getElementById('history-search-from').value;
        const toDate = document.getElementById('history-search-to').value;
        const target = document.getElementById('history-search-target').value;
        const resultsContainer = document.getElementById('history-search-results');
        const count = document.getElementById('history-search-count');
        if (count) count.textContent = '';
        if (fromDate && toDate && fromDate > toDate) {
            resultsContainer.textContent = '開始日は終了日以前を指定してください。';
            return;
        }

        resultsContainer.innerHTML = `
            <div class="empty-history" style="padding: 40px; text-align: center; color: var(--text-muted);">
                <ion-icon name="sync-outline" class="spinning" style="font-size: 24px; margin-bottom: 8px;"></ion-icon>
                <div>検索中...</div>
            </div>
        `;

        try {
            const response = await fetchAPI('searchHistory', { target: target, keyword: keyword, dateFrom: fromDate, dateTo: toDate });
            if (sequence !== historySearchSequence) return;
            if (response.status === 'success') {
                const rawData = response.data.rawData;
                let merged = [];
                for (let tableName in rawData) {
                    if (rawData[tableName].length > 1) {
                        const parsed = convertRawToObjects(rawData[tableName]);
                        let tType = '';
                        if (tableName === 'T_仕入') tType = 'purchase';
                        else if (tableName === 'T_経費') tType = 'expense';
                        else if (tableName === 'T_製造') tType = 'manufacturing';
                        else if (tableName === 'T_販売') tType = 'sales';
                        
                        parsed.forEach(row => {
                            row.type = tType;
                            merged.push(row);
                        });
                    }
                }

                // ホームからのID指定だけ完全一致にし、通常のキーワード検索は維持する。
                if (typeof exactId === 'string' && exactId) {
                    merged = merged.filter(row => String(row['販売ID'] || row['仕入ID'] || row['経費ID'] || row['製造ID']) === exactId);
                }
                if (merged.length === 0) {
                    if (count) count.textContent = '該当 0件';
                    resultsContainer.innerHTML = `
                        <div class="empty-history" style="padding: 40px; text-align: center; color: var(--text-muted);">
                            <p>該当する履歴が見つかりませんでした。</p>
                        </div>
                    `;
                    return;
                }

                merged.sort((a, b) => historyDateInfo_(b.type, b).searchDay.localeCompare(historyDateInfo_(a.type, a).searchDay));

                if (count) count.textContent = merged.length > 50 ? '該当 ' + merged.length + '件 ／ 日付順に先頭50件を表示。条件を絞って検索してください。' : '該当 ' + merged.length + '件';
                merged = merged.slice(0, 50);
                resultsContainer.innerHTML = '';
                
                merged.forEach(action => {
                    const itemDiv = document.createElement('div');
                    itemDiv.className = 'transaction-item';
                    
                    const isIncome = action.type === 'sales';
                    const isExpense = ['purchase', 'expense'].includes(action.type);
                    const colorClass = action.type === 'sales' ? 'income' : action.type;
                    const symbol = isIncome ? '+' : (isExpense ? '-' : '');
                    const iconName = action.type === 'sales' ? 'pricetags-outline' :
                        (action.type === 'purchase' ? 'cart-outline' :
                            (action.type === 'expense' ? 'cash-outline' : 'hammer-outline'));

                    const status = action['ステータス'] || '不明';
                    const isCompleted = ['完了', '入庫済み', '入庫済'].includes(status);
                    
                    const itemName = action['品名'] || action['摘要'] || action['完成品名'] || '不明';
                    const id = action['販売ID'] || action['仕入ID'] || action['経費ID'] || action['製造ID'] || '-';
                    const qty = action['数量'] || action['製造数量'] || '-';
                    const amount = parseNumber(action['価格']) || parseNumber(action['合計金額']) || parseNumber(action['売上']) || parseNumber(action['仕入']) || 0;
                    
                    const dateStr = historyDateInfo_(action.type, action).dates.map(date => date.label + '：' + (date.value || '未記録')).join(' ／ ');
                    
                    const buyer = action['売先'] || '';
                    
                    const labelPrefix = action.type === 'purchase' ? '仕入' : (action.type === 'expense' ? '経費' : (action.type === 'manufacturing' ? '製造' : '販売'));

                    itemDiv.innerHTML = `
                        <div class="type-icon ${colorClass}">
                            <ion-icon name="${iconName}"></ion-icon>
                        </div>
                        <div class="item-info">
                            <div class="item-header-meta">
                                <span class="item-type-badge">${labelPrefix}</span>
                                <span class="item-status-badge ${isCompleted ? 'completed' : ''}">${escapeHtml(status)}</span>
                                ${isIncome && buyer ? `<span class="badge badge-buyer">${escapeHtml(buyer)}</span>` : ''}
                            </div>
                            <h4><button type="button" class="history-detail-link item-name-text">${escapeHtml(itemName)}</button></h4>
                            <div class="item-sub-meta">
                                <span>${escapeHtml(dateStr)}</span>
                                <span>数量: ${escapeHtml(qty)}</span>
                                <span>ID: ${escapeHtml(id)}</span>
                            </div>
                            <div class="history-search-summary">${historySummary_(action.type, action).map(line => '<span>' + escapeHtml(line) + '</span>').join('')}</div>
                        </div>
                        <div class="item-amount ${isIncome ? 'positive' : (action.type === 'manufacturing' ? 'neutral' : 'negative')}">
                            ${action.type === 'manufacturing' ? '<span class="no-amount">-</span>' : symbol + '¥' + Math.abs(amount).toLocaleString()}
                        </div>
                    `;
                    const detailLink = itemDiv.querySelector('.history-detail-link');
                    detailLink.addEventListener('click', () => openTransactionDetail_(action.type, id, detailLink));
                    resultsContainer.appendChild(itemDiv);
                });
                resultsContainer.scrollTop = 0;
                
            } else {
                resultsContainer.textContent = '検索に失敗しました。条件を確認して再度検索してください。';
                if (typeof showToast === 'function') showToast('検索に失敗しました: ' + response.message, 'error');
            }
        } catch (e) {
            if (sequence !== historySearchSequence) return;
            resultsContainer.textContent = '検索エラーが発生しました。再度検索してください。';
            console.error(e);
            if (typeof showToast === 'function') showToast('検索エラーが発生しました', 'error');
        }
    }

});
