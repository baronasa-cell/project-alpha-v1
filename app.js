document.addEventListener('DOMContentLoaded', async () => {
    console.log("System initialization started... (v2.2-QR-LayoutFixed)");
    
    // ---- API Configuration ----
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbzexidaVzlRQ1_StDZo6Oo_oOt9TtX33Nk2sPwbo-oDzuRW6_Tbt2_zQxlxv-Ctr4jZuA/exec';
    let currentAuthKey = localStorage.getItem('inventory_auth_key') || '';
    let useMock = false;

    // Global state
    let currentMasters = {};
    let lastHistoryData = {}; // 帳簿表示用などにデータを保持
    let lastRawData = {};    // サーバーから届いた生データを保持（マージ用）
    let ledgerYear = new Date().getFullYear();
    let ledgerPeriod = String(new Date().getMonth() + 1).padStart(2, '0'); // デフォルトは「今月」

    let isStocktakeMode = false;
    let stocktakeData = {}; // { itemName: { actual: number, diff: number } }
    const STOCKTAKE_CYCLE_DAYS = 30; // サイクルカウントの基準日数

    let pendingStockChanges = {
        thresholds: {}, // itemName -> newValue
        statuses: {}    // itemName -> newValue (0 or 1)
    };

    let html5QrCode = null; // Scanner instance

    // マスタ編集用スキーマ定義 (提案7・マスター管理強化)
    const MASTER_SCHEMAS = {
        'M_商品': {
            key: '品名',
            fields: [
                { name: '商品ID', type: 'text', visible: true, editable: false }, // 自動採番
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '品名', type: 'text', visible: true, editable: false }, // 編集時はReadOnly
                { name: 'カテゴリ', type: 'select', visible: true, editable: true, options: ['パーツ', '単体商品', '商品', '経費', '製造'] },
                { name: '説明', type: 'textarea', visible: true, editable: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true },
                { name: '画像URL', visible: false },
                { name: '保管場所', type: 'text', visible: true, editable: true },
                { name: 'QR/バーコード', type: 'text', visible: true, editable: true } // I列
            ]
        },
        'M_仕入先': {
            key: '仕入先',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '仕入先', type: 'text', visible: true, editable: false },
                { name: '用途区分', type: 'select', visible: true, editable: true, options: [{v:1, l:'1:仕入のみ'}, {v:2, l:'2:経費のみ'}, {v:3, l:'3:両方'}] },
                { name: '説明', type: 'textarea', visible: true, editable: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true }
            ]
        },
        'M_売先': {
            key: '売先',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '売先', type: 'text', visible: true, editable: false },
                { name: '手数料率', type: 'number', visible: true, editable: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true }
            ]
        },
        'M_発送': {
            key: '発送方法',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '発送方法', type: 'text', visible: true, editable: false },
                { name: '送料', type: 'number', visible: true, editable: true },
                { name: '使用FLG', type: 'switch', visible: true, editable: true }
            ]
        },
        'M_BOM': {
            key: '品名', // 実際は複合キーだがバックエンドで対応
            fields: [
                { name: '品名', type: 'select', visible: true, editable: false, refMaster: 'M_商品', filter: (r)=>r['カテゴリ']==='商品' },
                { name: '部品', type: 'select', visible: true, editable: false, refMaster: 'M_商品', filter: (r)=>['パーツ','単体商品'].includes(r['カテゴリ']) },
                { name: '数量', type: 'number', visible: true, editable: true },
                { name: '説明', type: 'textarea', visible: true, editable: true }
            ]
        },
        'M_経費品名': {
            key: '品名',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '品名', type: 'text', visible: true, editable: false },
                { name: 'デフォルト仕訳', type: 'select', visible: true, editable: true, refMaster: 'M_仕訳' },
                { name: '使用FLG', type: 'switch', visible: true, editable: true }
            ]
        },
        'T_在庫集計': {
            key: '品名',
            fields: [
                { name: '優先度', type: 'number', visible: true, editable: true },
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '品名', type: 'text', visible: true, editable: false },
                { name: 'カテゴリ', type: 'text', visible: true, editable: false },
                { name: '現在庫数', type: 'number', visible: true, editable: false },
                { name: '閾値', type: 'number', visible: true, editable: true },
                { name: '最終更新日', visible: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true },
                { name: '最終棚卸日', type: 'text', visible: true, editable: false },
                { name: '商品ID', type: 'text', visible: true, editable: false }, 
                { name: '保管場所', type: 'text', visible: true, editable: false },
                { name: 'QR/バーコード', type: 'text', visible: true, editable: false }
            ]
        },
        'M_支払': {
            key: '支払方法',
            fields: [
                { name: '表示順', type: 'number', visible: true, editable: true },
                { name: '支払方法', type: 'text', visible: true, editable: false },
                { name: '使用FLG', type: 'switch', visible: true, editable: true }
            ]
        }
    };

    // 棚卸管理用の状態
    let stocktakeSession = {
        verifiedItems: new Set(), // 現在のセッションで「済」にした品名
        uncheckedOnly: false      // フィルタ状態
    };

    try {
        // ---- 1. Event Listener Registrations (Essential UI) ----
        setupNavigation();
        setupToggleLogics();
        setupSettingsListeners();

        // ---- 2. System Initialization (Data Fetching) ----
        // 認証チェック (GAS環境以外の場合)
        if (typeof google === 'undefined' && !currentAuthKey) {
            setupLoginHandlers();
            showLoginModal();
            return; // ログイン完了まで中断
        }
        setupLoginHandlers(); // リトライ用などに常にセットアップ

        // 認証済みの場合、即座にコンテナを表示 (読み込み中アニメーションを表示させるため)
        const appContainer = document.querySelector('.app-container');
        if (appContainer) {
            appContainer.style.display = 'flex';
        }

        // 1. まずキャッシュからマスタを読み込んでUIを構築
        loadMastersFromCache();

        console.time('Essential Load');
        await initSystem('essential');
        console.timeEnd('Essential Load');


        // 4. 残りの詳細履歴データをバックグラウンドで非同期に取得 (マスタは取得済みなのでスキップ)
        initSystem('all', { skipMasters: true }).then(() => {
            console.log("Background data load completed.");
        }).catch(err => {
            console.warn("Background load failed:", err);
        });

        // ---- 3. Image Feature Initializations ----
        setupImagePreviewListeners();
        setupStockUpdateListeners();
        setupStockFilters();
        setupScannerListeners();

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

                // 最近のアクション（最新の活動日を基準にする）
                // 全ての日付列をスキャンして最大値を探す
                let latestActivityDate = d;
                for (const key in r) {
                    if (key.includes('日') && r[key]) {
                        const tempD = new Date(r[key]);
                        if (!isNaN(tempD.getTime()) && tempD > latestActivityDate) {
                            latestActivityDate = tempD;
                        }
                    }
                }

                let amount = 0;
                if (sheetName === 'T_仕入') amount = -price;
                else if (sheetName === 'T_経費') amount = -price;
                else if (sheetName === 'T_販売') amount = price;

                recentAll.push({
                    type: sheetName === 'T_仕入' ? 'purchase' : (sheetName === 'T_経費' ? 'expense' : (sheetName === 'T_販売' ? 'sales' : 'manufacturing')),
                    id: r[Object.keys(r)[0]], date: latestActivityDate, itemName: r['品名'],
                    amount: amount, status: status, quantity: r['数量'] || r['製造数量'],
                    dateStr: formatDate(latestActivityDate),
                    buyer: sheetName === 'T_販売' ? r['売先'] : null
                });
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

        // 販売・製造履歴の絞り込み用オプションを更新
        updateSalesFilterOptions(history['T_販売']);
        updateManufacturingFilterOptions(history['T_製造']);

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
    }

    /**
     * 販売履歴のフィルターオプションを更新する
     */
    function updateSalesFilterOptions(data) {
        if (!data) return;
        const statusSelect = document.getElementById('sales-history-status-filter');
        const buyerSelect = document.getElementById('sales-history-buyer-filter');
        if (!statusSelect || !buyerSelect) return;

        const currentStatus = statusSelect.value;
        const currentBuyer = buyerSelect.value;

        const statuses = new Set();
        const buyers = new Set();
        data.forEach(item => {
            if (item['ステータス']) statuses.add(item['ステータス'].trim());
            if (item['売先']) buyers.add(item['売先'].trim());
        });

        statusSelect.innerHTML = '<option value="">すべて</option>';
        [...statuses].sort().forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            statusSelect.appendChild(opt);
        });

        buyerSelect.innerHTML = '<option value="">すべて</option>';
        [...buyers].sort().forEach(b => {
            const opt = document.createElement('option');
            opt.value = b;
            opt.textContent = b;
            buyerSelect.appendChild(opt);
        });

        if (statuses.has(currentStatus)) statusSelect.value = currentStatus;
        if (buyers.has(currentBuyer)) buyerSelect.value = currentBuyer;
    }
    
    /**
     * 製造履歴のフィルターオプションを更新する
     */
    function updateManufacturingFilterOptions(data) {
        if (!data) return;
        const statusSelect = document.getElementById('manufacturing-history-status-filter');
        const itemSelect = document.getElementById('manufacturing-history-item-filter');
        if (!statusSelect || !itemSelect) return;

        const currentStatus = statusSelect.value;
        const currentItem = itemSelect.value;

        const statuses = new Set();
        const items = new Set();
        data.forEach(item => {
            if (item['ステータス']) statuses.add(item['ステータス'].trim());
            const itemName = item['品名'] || item['完成品名'];
            if (itemName) items.add(itemName.trim());
        });

        statusSelect.innerHTML = '<option value="">すべて</option>';
        [...statuses].sort().forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            statusSelect.appendChild(opt);
        });

        itemSelect.innerHTML = '<option value="">すべて</option>';
        [...items].sort().forEach(i => {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = i;
            itemSelect.appendChild(opt);
        });

        if (statuses.has(currentStatus)) statusSelect.value = currentStatus;
        if (items.has(currentItem)) itemSelect.value = currentItem;
    }

    /**
     * 製造履歴にフィルターを適用
     */
    window.applyManufacturingHistoryFilter = function() {
        if (!lastHistoryData || !lastHistoryData.history) return;
        renderHistoryCards('manufacturing', lastHistoryData.history['T_製造']);
    };

    window.applySalesHistoryFilter = function() {
        if (!lastHistoryData || !lastHistoryData.history) return;
        renderHistoryCards('sales', lastHistoryData.history['T_販売']);
    };

    /**
     * ナビゲーションバッジの更新
     */
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
        if (useMock) {
            if (window.MockAPI) {
                if (action === 'getMasters') return window.MockAPI.getMasterData();
                if (action === 'registerTransaction') return window.MockAPI.registerTransaction(payload.sheet, payload.data);
                if (action === 'updateTransaction') return window.MockAPI.updateTransaction(payload.id, payload.updates);
                if (action === 'verifyAndAddMaster') return window.MockAPI.verifyAndAddMaster(payload.sheet, payload.value);
            }
            return { status: 'error', message: 'Mock API not found' };
        }

        const bodyData = Object.assign({ action: action, key: currentAuthKey }, payload);
        
        if (!currentAuthKey && action !== 'getInitData') {
            console.warn(`[fetchAPI] Warning: No auth key for action: ${action}`);
        }

        // GAS本番環境 (google.script.run が存在する場合)
        if (typeof google !== 'undefined' && google.script && google.script.run) {
            return new Promise((resolve, reject) => {
                google.script.run
                    .withSuccessHandler(res => {
                        if (!res) reject(new Error("Server returned null response"));
                        else if (res.status === 'error' && res.message.includes('Unauthorized')) {
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
            if (result.status === 'error' && result.message.includes('Unauthorized')) {
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
    function handleUnauthorized() {
        localStorage.removeItem('inventory_auth_key');
        currentAuthKey = '';
        showLoginModal();
    }

    /**
     * パスワードをハッシュ化 (SHA-256)
     */
    async function hashPassword(password) {
        const encoder = new TextEncoder();
        const data = encoder.encode(password);
        const hash = await crypto.subtle.digest('SHA-256', data);
        return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
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
    function setupLoginHandlers() {
        const submitBtn = document.getElementById('login-submit-btn');
        const passwordInput = document.getElementById('login-password');
        const errorMsg = document.getElementById('login-error');

        if (!submitBtn || !passwordInput) return;

        const handleLogin = async () => {
            const password = passwordInput.value;
            if (!password) return;

            submitBtn.disabled = true;
            submitBtn.textContent = '認証中...';
            errorMsg.style.display = 'none';

            try {
                // バックエンドでハッシュ化を行うため、ここでは生のパスワードを送信
                const bodyData = { action: 'getInitData', key: password, scope: 'check' };
                const response = await fetch(GAS_URL, {
                    method: 'POST',
                    body: JSON.stringify(bodyData)
                });
                
                const result = await response.json();
                
                if (result.status === 'success') {
                    // 認証成功
                    currentAuthKey = password;
                    localStorage.setItem('inventory_auth_key', password);
                    document.getElementById('login-modal').style.display = 'none';
                    
                    // システムコンテナを表示（重要：ブランク画面回避）
                    const appContainer = document.querySelector('.app-container');
                    if (appContainer) {
                        appContainer.style.display = 'flex';
                    }

                    // システム初期化を再開
                    if (typeof showToast === 'function') showToast('システムを起動しています...', 'success');
                    initSystem('all');
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

        // Sync Button
        const syncBtn = document.getElementById('sync-btn');
        if (syncBtn) {
            syncBtn.addEventListener('click', () => {
                const icon = syncBtn.querySelector('ion-icon');
                if (icon) icon.style.animation = 'spin 1s linear infinite';
                setLoading(true, 'システムデータを同期中...');
                initSystem().then(() => {
                    if (icon) icon.style.animation = '';
                    setLoading(false);
                }).catch(e => {
                    if (icon) icon.style.animation = '';
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
            if (el) el.value = '';
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
        controls.forEach(ctrl => {
            const elmId = ctrl['要素ID'];
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
                            const excludeFields = ['表示順', '使用FLG', 'カテゴリ', '手数料率', '送料', '用途区分', '説明', 'デフォルト仕訳', '役割（タイプ）', '対象機能', '画面名称'];
                            const keyField = Object.keys(masterData[0]).find(k => !excludeFields.includes(k));
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

        const term = (stockSearchInput ? stockSearchInput.value : "").toLowerCase().trim();
        const showHidden = showHiddenChip ? showHiddenChip.classList.contains('active') : false;
        const thresholdZeroOnly = thresholdZeroChip ? thresholdZeroChip.classList.contains('active') : false;
        const uncheckedOnly = uncheckedOnlyChip ? uncheckedOnlyChip.classList.contains('active') : false;
        const inStockOnly = inStockOnlyChip ? inStockOnlyChip.classList.contains('active') : false;
        
        const allStockProducts = currentMasters['T_在庫集計'] || [];

        const filtered = allStockProducts.filter(p => {
            const name = (p['品名'] || "").toLowerCase();
            const category = (p['カテゴリ'] || "").toLowerCase();
            const id = (p['商品ID'] || "").toLowerCase();
            const barcode = (p['QR/バーコード'] || "").toLowerCase();
            const location = (p['保管場所'] || "").toLowerCase();
            const useFlag = parseInt(p['使用FLG']) !== 0;
            const threshold = parseFloat(p['閾値']) || 0;
            const stock = parseFloat(p['現在庫数']) || 0;
            const isUnchecked = !p['最終棚卸日'];

            const matchesSearch = name.includes(term) || category.includes(term) || id.includes(term) || barcode.includes(term) || location.includes(term);
            const matchesVisibility = showHidden || useFlag;
            const matchesThreshold = !thresholdZeroOnly || threshold === 0;
            const matchesUnchecked = !uncheckedOnly || isUnchecked;
            const matchesInStock = !inStockOnly || stock > 0;

            return matchesSearch && matchesVisibility && matchesThreshold && matchesUnchecked && matchesInStock;
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

    function toggleSaleItemInput(isPersonal) {
        let saleItemField = document.getElementById('sale-item');
        if (!saleItemField) return;

        // すでにラッパーの中にある場合は、ラッパーの親を取得
        const container = saleItemField.dataset.hasClear === "true" ?
            saleItemField.closest('.input-clear-wrapper').parentElement :
            saleItemField.parentElement;
        const currentVal = saleItemField.value;

        if (isPersonal && saleItemField.tagName === 'SELECT') {
            const input = document.createElement('input');
            input.type = 'text';
            input.id = 'sale-item';
            input.placeholder = '品名を自由入力';
            input.setAttribute('list', 'product-list');
            input.value = currentVal;

            // 古い要素を置換（ラッパーごと置換する必要がある場合を考慮）
            const targetToReplace = saleItemField.dataset.hasClear === "true" ?
                saleItemField.closest('.input-clear-wrapper') :
                saleItemField;
            container.replaceChild(input, targetToReplace);

            // 新規入力欄にクリアボタンを追加
            addClearButton(input);
        } else if (!isPersonal && saleItemField.tagName === 'INPUT') {
            // 事業用の場合は、M_画面制御の定義に従って select か suggest-strict かを判定
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

                const targetToReplace = saleItemField.dataset.hasClear === "true" ?
                    saleItemField.closest('.input-clear-wrapper') :
                    saleItemField;
                container.replaceChild(select, targetToReplace);
            } else {
                // suggest-strict の場合は INPUT のままでよいが、プレースホルダーなどを更新
                saleItemField.placeholder = '販売品名を検索...';
                const modeText = document.getElementById('sale-item-mode-text');
                if (modeText) modeText.textContent = '(事業用: リストから選択のみ)';
            }
        }

        // 差し替え後の要素にプレビューリスナーを再設定
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
                            if (el.type === 'number') {
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
    function validateData(sheet, data) {
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

        for (const rule of sheetRules) {
            const val = data[rule.key];

            // 必須チェック
            if (rule.required) {
                if (val === undefined || val === null || val === '' || (typeof val === 'number' && isNaN(val))) {
                    return `「${rule.label}」を入力してください。`;
                }
            }

            // 数値チェック（最小値）
            if (rule.min !== undefined && typeof val === 'number') {
                if (val < rule.min) {
                    return `「${rule.label}」は ${rule.min} 以上の数値を入力してください。`;
                }
            }

            // suggest-strict のバリデーション
            const el = document.getElementById(confFields[rule.key]);
            if (el && el.dataset.type === 'suggest-strict') {
                // 販売タブで「個人用」が選択されている場合は、マスタチェックをスキップする
                const isPersonalSales = (sheet === 'T_販売' && data.type === 'personal');
                if (!isPersonalSales) {
                    const masterName = el.dataset.master;
                    const masterData = currentMasters[masterName] || [];
                    
                    // 除外フィールドを考慮して品名などのキーフィールドを特定（buildDynamicUIと同様のロジック）
                    const excludeFields = ['表示順', '使用FLG', 'カテゴリ', '手数料率', '送料', '用途区分', '説明', 'デフォルト仕訳', '役割（タイプ）', '対象機能', '画面名称'];
                    const keyField = masterData.length > 0 ? Object.keys(masterData[0]).find(k => !excludeFields.includes(k)) : null;
                    
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
        if (sheet === 'T_販売') {
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
                    defaultQtys.forEach(el => { if (el.id.includes('quantity')) el.value = ''; });
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

        safeData.forEach(row => {
            const itemName = row['品名'];
            const category = row['カテゴリ'] || row['商品区分'] || ''; 
            const rawQty = parseFloat(row['現在庫数']);
            const stockQty = isNaN(rawQty) ? 0 : rawQty;

            const rawThreshold = parseFloat(row['閾値']);
            const threshold = isNaN(rawThreshold) ? 10 : rawThreshold;
            const useFlag = parseInt(row['使用FLG']) !== 0;

            const lastStocktakeDate = row['最終棚卸日'] ? new Date(row['最終棚卸日']) : null;
            const isRecent = lastStocktakeDate && (Math.abs(now - lastStocktakeDate) / (1000*60*60*24) <= 30);
            const isVerified = isRecent || stocktakeSession.verifiedItems.has(itemName);

            const card = document.createElement('div');
            card.className = 'stock-item-card';
            if (!useFlag) card.classList.add('inactive-row');
            if (isVerified && isStocktakeMode) card.classList.add('verified-row');

            const itemInMaster = (currentMasters['M_商品'] || []).find(m => m['品名'] === itemName);
            const imageUrl = itemInMaster ? itemInMaster['画像URL'] : null;
            const itemID = row['商品ID'] || '';
            const location = row['保管場所'] || '';
            const barcode = row['QR/バーコード'] || '';

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
                        <div class="product-thumb-container" onclick="showImageModal('${imageUrl}')">
                            ${imageUrl ? `<img src="${imageUrl}" loading="lazy">` : `<ion-icon name="image-outline"></ion-icon>`}
                        </div>
                        <div class="card-main-info">
                            <div class="card-product-name">${itemName}</div>
                            <div class="card-sub-info">
                                ${itemID ? `<span class="id-badge">${itemID}</span>` : ''}
                                ${barcode ? `<span class="barcode-badge"><ion-icon name="barcode-outline"></ion-icon>${barcode}</span>` : ''}
                                ${location ? `<span class="location-badge"><ion-icon name="location-outline"></ion-icon>${location}</span>` : ''}
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
                        <div class="product-thumb-container" onclick="showImageModal('${imageUrl}')">
                            ${imageUrl ? `<img src="${imageUrl}" loading="lazy">` : `<ion-icon name="image-outline"></ion-icon>`}
                        </div>
                        <div class="card-main-info">
                            <div class="card-product-name clickable" onclick="navigateToTransactionForm('${itemName}', '${category}')">${itemName}</div>
                            <div class="card-sub-info">
                                ${itemID ? `<span class="id-badge">${itemID}</span>` : ''}
                                ${barcode ? `<span class="barcode-badge">${barcode}</span>` : ''}
                                ${location ? `<span class="location-badge"><ion-icon name="location-outline"></ion-icon>${location}</span>` : ''}
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
                            <button class="camera-btn" onclick="triggerPhotoUpload('${itemName}')" title="写真を登録/変更" style="margin-left:0;">
                                <ion-icon name="camera-outline"></ion-icon>
                            </button>
                            <button class="update-mini-btn no-text btn-save-threshold ${isThresholdDirty ? 'is-dirty' : ''}" title="数値を確定待ちに追加">
                                <ion-icon name="time-outline"></ion-icon>
                            </button>
                            <button class="update-mini-btn no-text secondary-action btn-archive-item ${isStatusDirty ? 'is-dirty' : ''}" title="${currentUseFlag ? '廃盤にする' : '復活させる'}">
                                <ion-icon name="${currentUseFlag ? 'archive-outline' : 'refresh-outline'}"></ion-icon>
                            </button>
                        </div>
                    </div>
                `;

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

        if (!html5QrCode) {
            html5QrCode = new Html5Qrcode("reader");
        }

        const config = { fps: 10, qrbox: { width: 250, height: 250 } };

        html5QrCode.start({ facingMode: "environment" }, config, onScanSuccess)
            .catch(err => {
                console.error("Scanner start error:", err);
                alert("カメラの起動に失敗しました。カメラの使用を許可してください。");
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

        // 1. 場所QRコードの判定 (LOC-XXXX)
        if (decodedText.startsWith('LOC-')) {
            const loc = decodedText.replace('LOC-', '');
            const searchInput = document.getElementById('stock-search-input');
            if (searchInput) {
                searchInput.value = loc;
                searchInput.dispatchEvent(new Event('input')); // 検索実行
            }
            return;
        }

        // 2. 商品ID、品名、またはバーコードでの照合
        const cards = Array.from(document.querySelectorAll('.stock-item-card'));
        const matches = cards.filter(card => {
            const id = card.getAttribute('data-item-id');
            const name = card.getAttribute('data-item-name');
            const bc = card.getAttribute('data-barcode');
            return id === decodedText || bc === decodedText || name === decodedText;
        });

        if (matches.length === 1) {
            // 一意に決まる場合：スクロール、ハイライト、入力フォーカス
            const targetCard = matches[0];
            targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
            targetCard.classList.add('scan-highlight');
            setTimeout(() => targetCard.classList.remove('scan-highlight'), 1500);

            const input = targetCard.querySelector('.stepper-input');
            if (input) {
                input.focus();
                input.select();
            }
        } else if (matches.length > 1) {
            // 重複する場合（同一JANコードなど）：一覧をその値で絞り込む
            showToast(`${matches.length}件の商品がヒットしました。絞り込み表示します。`);
            const searchInput = document.getElementById('stock-search-input');
            if (searchInput) {
                searchInput.value = decodedText;
                searchInput.dispatchEvent(new Event('input')); // 検索実行
            }
        } else {
            alert(`スキャン結果: "${decodedText}" に一致する商品は見つかりませんでした。`);
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

        if (!data || data.length === 0) {
            container.innerHTML = '<p class="empty-msg">現在進行中の履歴はありません</p>';
            return;
        }

        // --- 絞り込みロジック (A案: サマリーは全体のまま、リストだけ絞る) ---
        let renderData = data;
        if (tab === 'sales') {
            const statusFilter = document.getElementById('sales-history-status-filter')?.value;
            const buyerFilter = document.getElementById('sales-history-buyer-filter')?.value;
            
            if (statusFilter) {
                renderData = renderData.filter(item => (item['ステータス'] || '').trim() === statusFilter);
            }
            if (buyerFilter) {
                renderData = renderData.filter(item => (item['売先'] || '').trim() === buyerFilter);
            }
        } else if (tab === 'manufacturing') {
            const statusFilter = document.getElementById('manufacturing-history-status-filter')?.value;
            const itemFilter = document.getElementById('manufacturing-history-item-filter')?.value;
            
            if (statusFilter) {
                renderData = renderData.filter(item => (item['ステータス'] || '').trim() === statusFilter);
            }
            if (itemFilter) {
                renderData = renderData.filter(item => {
                    const name = (item['品名'] || item['完成品名'] || '').trim();
                    return name === itemFilter;
                });
            }
        }

        if (!renderData || renderData.length === 0) {
            container.innerHTML = '<p class="empty-msg">条件に一致する履歴はありません</p>';
            return;
        }

        // --- 描画最適化ロジック ---
        const CHUNK_SIZE = 50; // 1回あたりに描画する件数
        let index = 0;

        function renderChunk() {
            const fragment = document.createDocumentFragment();
            const limit = Math.min(index + CHUNK_SIZE, renderData.length);

            for (; index < limit; index++) {
                const item = renderData[index];
                const card = createHistoryCardElement(tab, item);
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
        const price = item['合計金額'] || item['価格'] || item['販売価格'] || item['小計'] || item['売上'] || 0;
        const formattedPrice = typeof price === 'number' ? price.toLocaleString() : price;

        let innerHTML = `
            <div class="history-card-header">
                <div class="header-left">
                    <span class="history-id">${id}</span>
                    ${(tab === 'sales' && item['売先']) ? `<span class="badge badge-buyer">${item['売先']}</span>` : ''}
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
                        <button class="add-material-btn" data-id="${id}" title="部材を追加消費する">
                            <ion-icon name="add-circle-outline"></ion-icon>
                            <span>部材追加</span>
                        </button>
                    ` : ''}
                    <button class="update-mini-btn" data-id="${id}" title="変更を保存">
                        <span>保存</span>
                        <ion-icon name="save-outline"></ion-icon>
                    </button>
                </div>
            </div>
        `;

        const itemName = item['品名'] || item['完成品名'] || '';
        const itemInMaster = (currentMasters['M_商品'] || []).find(m => m['品名'] === itemName);
        const imageUrl = itemInMaster ? itemInMaster['画像URL'] : null;
        const thumbHtml = imageUrl ? `<div class="product-thumb-container" style="width:30px; height:30px; margin-right:8px; border-radius:4px;" onclick="showImageModal('${imageUrl}')"><img src="${imageUrl}" loading="lazy"></div>` : '';

        if (tab === 'purchase') {
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}${itemName} (¥${formattedPrice}) 数量:${item['数量']}</div>
                <div class="history-sub-info" style="margin-top: 4px; margin-bottom: 8px; font-size: 0.9em; color: var(--text-secondary);">
                    ${item['仕入先'] || ''} &nbsp;&nbsp; ${item['支払方法'] || ''} &nbsp;&nbsp; ${item['区分'] || ''}
                </div>
                <div class="history-inputs-grid">
                    <div class="input-group mini"><label>仕入日</label><input type="date" class="date-input" data-header="仕入日" value="${formatISODate(item['仕入日'] || item['発注日'])}"></div>
                    <div class="input-group mini"><label>入庫日</label><input type="date" class="date-input" data-header="入庫日" value="${formatISODate(item['入庫日'])}"></div>
                    <div class="input-group mini" style="grid-column: span 2;">${generateStatusSelect(id, '仕入', item['ステータス'])}</div>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.1); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${item['備考'] || item['note'] || ''}</textarea>
                </div>
            `;
        } else if (tab === 'expense') {
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}${itemName} (¥${formattedPrice}) 数量:${item['数量']}</div>
                <div class="history-sub-info" style="margin-top: 4px; margin-bottom: 8px; font-size: 0.9em; color: var(--text-secondary);">
                    ${item['購入先'] || ''} &nbsp;&nbsp; ${item['支払方法'] || ''} &nbsp;&nbsp; ${item['仕訳'] || ''}
                </div>
                <div class="history-inputs-grid">
                    <div class="input-group mini"><label>注文日</label><input type="date" class="date-input" data-header="注文日" value="${formatISODate(item['注文日'] || item['登録日'])}"></div>
                    <div class="input-group mini"><label>完了日</label><input type="date" class="date-input" data-header="完了日" value="${formatISODate(item['完了日'])}"></div>
                    <div class="input-group mini" style="grid-column: span 2;">${generateStatusSelect(id, '経費', item['ステータス'])}</div>
                </div>
                <div class="history-inputs-grid" style="grid-template-columns: 1fr; margin-top: 5px;">
                    <label style="display: flex; align-items: center; gap: 8px; font-size: 0.9em; color: var(--text-secondary); cursor: pointer;">
                        <input type="checkbox" class="chk-input" data-header="レシート" ${item['レシート'] == 1 ? 'checked' : ''} style="width:auto; transform:scale(1.1);">
                        レシート提出済みにする
                    </label>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.1); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${item['備考'] || item['note'] || ''}</textarea>
                </div>
            `;
        } else if (tab === 'manufacturing') {
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}${itemName} 数量:${item['数量'] || item['製造数量']} 単価:${item['単価'] || 0}</div>
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
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${item['備考'] || item['note'] || item['メモ'] || ''}</textarea>
                </div>
            `;
        } else if (tab === 'sales') {
            const unitPrice = item['合計単価'] || item['単価'] || 0;
            const totalPrice = item['販売価格'] || item['価格'] || item['合計金額'] || 0;
            const formattedUnit = typeof unitPrice === 'number' ? unitPrice.toLocaleString() : unitPrice;
            const formattedTotal = typeof totalPrice === 'number' ? totalPrice.toLocaleString() : totalPrice;
            const trackingUrl = getTrackingUrl(item['発送方法'], item['追跡番号']);
            innerHTML += `
                <div class="history-product-info" style="display: flex; align-items: center;">${thumbHtml}${itemName} 数量:${item['数量']} 価格:¥${formattedTotal}(単価:¥${formattedUnit}) 送料:¥${(item['送料'] || 0).toLocaleString()}</div>
                <div class="history-sub-info" style="margin-top: 4px; margin-bottom: 8px; font-size: 0.9em; color: var(--text-muted);">
                    注文日:${formatDate(item['販売開始日'])} &nbsp;&nbsp; ${item['発送方法'] || ''}
                </div>
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
                            <input type="text" class="note-input" data-header="追跡番号" value="${item['追跡番号'] || ''}" placeholder="追跡番号" style="flex: 1;">
                            ${trackingUrl ? `<a href="${trackingUrl}" target="_blank" class="mini-icon-link" title="配送状況を確認"><ion-icon name="navigate-circle-outline"></ion-icon></a>` : ''}
                        </div>
                    </div>
                </div>
                <div class="history-note" style="margin-top: 10px; padding-top: 8px; border-top: 1px dashed rgba(0,0,0,0.12); font-size: 0.85em; color: var(--text-muted);">
                    <textarea class="note-input" data-header="備考" rows="2" style="width:100%; background:rgba(0,0,0,0.05); border:1px solid rgba(0,0,0,0.1); border-radius:4px; padding:4px;">${item['備考'] || item['note'] || ''}</textarea>
                </div>
            `;
        }

        card.innerHTML = innerHTML;

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
            return `<option value="${s['ステータス名称']}" ${selected}>${s['ステータス名称']}</option>`;
        }).join('');

        return `<label>ステータス</label><select class="status-select">${options}</select>`;
    }

    async function handleHistorySave(btn) {
        const id = btn.getAttribute('data-id');
        const card = btn.closest('.history-card');
        const statusSelect = card.querySelector('.status-select');
        const newStatus = statusSelect ? statusSelect.value : '';

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
                dateUpdates[header] = input.value;
            }
        });

        // チェックボックスの取得
        card.querySelectorAll('.chk-input').forEach(chk => {
            const header = chk.getAttribute('data-header');
            if (header) {
                dateUpdates[header] = chk.checked ? 1 : 0;
            }
        });

        const originalHTML = btn.innerHTML;
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
                    dates: dateUpdates
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
                    const s = parseNumber(row['売上']);
                    const cost = parseNumber(row['仕入']) + parseNumber(row['通信費']) + parseNumber(row['修繕費']) +
                        parseNumber(row['消耗品費']) + parseNumber(row['諸会費']) + parseNumber(row['支払手数料']) + parseNumber(row['雑費']);
                    res.sales += s;
                    res.cost += cost;
                    res.profit += (s - cost);
                }
            });

            // 個人用売上の集計
            const psData = lastHistoryData.personalSalesByMonth || {};
            for (const ym in psData) {
                const [y, m] = ym.split('-');
                if (String(y) === String(targetY)) {
                    let match = false;
                    if (targetP === 'FY') match = true;
                    else if (targetP === 'Q1') match = (parseInt(m) >= 1 && parseInt(m) <= 3);
                    else if (targetP === 'Q2') match = (parseInt(m) >= 4 && parseInt(m) <= 6);
                    else if (targetP === 'Q3') match = (parseInt(m) >= 7 && parseInt(m) <= 9);
                    else if (targetP === 'Q4') match = (parseInt(m) >= 10 && parseInt(m) <= 12);
                    else if (m === targetP) match = true;

                    if (match) res.personalSales += psData[ym];
                }
            }
            return res;
        }

        // 現在・前月・前年の特定
        const current = getTotalsFor(ledgerYear, ledgerPeriod);

        let prevY = ledgerYear, prevP = ledgerPeriod;
        if (ledgerPeriod === 'FY') { prevY = parseInt(ledgerYear) - 1; prevP = 'FY'; }
        else if (ledgerPeriod === 'Q1') { prevY = parseInt(ledgerYear) - 1; prevP = 'Q4'; }
        else if (ledgerPeriod.startsWith('Q')) { prevP = 'Q' + (parseInt(ledgerPeriod[1]) - 1); }
        else {
            let m = parseInt(ledgerPeriod);
            if (m === 1) { prevY = parseInt(ledgerYear) - 1; prevP = '12'; }
            else { prevP = String(m - 1).padStart(2, '0'); }
        }
        const prev = getTotalsFor(prevY, prevP);
        const lastYear = getTotalsFor(parseInt(ledgerYear) - 1, ledgerPeriod);

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
            const isFY = (ledgerPeriod === 'FY');
            const prevLabel = isFY ? '前年比' : (ledgerPeriod.startsWith('Q') ? '前期比' : '前月比');
            const yearLabel = '前年比';

            const salesPrevHTML = getGrowthHTML(calcRate(current.sales, prev.sales), prevLabel);
            const salesYearHTML = isFY ? '' : getGrowthHTML(calcRate(current.sales, lastYear.sales), yearLabel);

            const profitPrevHTML = getGrowthHTML(calcRate(current.profit, prev.profit), prevLabel);
            const profitYearHTML = isFY ? '' : getGrowthHTML(calcRate(current.profit, lastYear.profit), yearLabel);

            const pSales = current.personalSales || 0;
            const salesTotalHTML = `¥${Math.round(current.sales).toLocaleString()}${pSales > 0 ? `<span style="font-size:0.75em;color:var(--text-muted);margin-left:4px;">(¥${Math.round(current.sales + pSales).toLocaleString()})</span>` : ''}`;
            const profitTotalHTML = `¥${Math.round(current.profit).toLocaleString()}${pSales > 0 ? `<span style="font-size:0.75em;color:var(--text-muted);margin-left:4px;">(¥${Math.round(current.profit + pSales).toLocaleString()})</span>` : ''}`;

            summaryCards.innerHTML = `
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
            const s = parseNumber(row['売上']);
            const cost = parseNumber(row['仕入']) + parseNumber(row['通信費']) + parseNumber(row['修繕費']) +
                parseNumber(row['消耗品費']) + parseNumber(row['諸会費']) + parseNumber(row['支払手数料']) + parseNumber(row['雑費']);
            monthlySum[ym].sales += s;
            monthlySum[ym].cost += cost;
        });

        // フィルタリング対象の明細
        const filteredLedger = [];
        ledger.forEach(row => {
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
        const personalSalesByMonth = (lastHistoryData && lastHistoryData.personalSalesByMonth) || {};
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
                const totalSalesHeight = ((businessSales + personalSales) / maxVal) * 100;
                const businessHeight = (businessSales / maxVal) * 100;
                const cHeight = (monthlySum[m].cost / maxVal) * 100;
                const label = m.split('-')[1] + '月';
                // 積層バー: 事業用(緑)の上に個人用(青緑)を積み上げ
                if (personalSales > 0) {
                    return `
                    <div class="bar-group">
                        <div class="bar-stack" style="height: ${totalSalesHeight}%; display:flex; flex-direction:column-reverse;">
                            <div class="bar-segment income" style="flex: 0 0 ${businessHeight > 0 ? (businessHeight / totalSalesHeight * 100) : 0}%;"></div>
                            <div class="bar-segment personal" style="flex: 1 1 auto;"></div>
                        </div>
                        <div class="bar expense" style="height: ${cHeight}%;"></div>
                        <span class="bar-label" style="font-size:0.75em;">${label}</span>
                    </div>
                `;
                } else {
                    return `
                    <div class="bar-group">
                        <div class="bar income" style="height: ${businessHeight}%;"></div>
                        <div class="bar expense" style="height: ${cHeight}%;"></div>
                        <span class="bar-label" style="font-size:0.75em;">${label}</span>
                    </div>
                `;
                }
            }).join('');

            // 凡例に「個人」を追加（個人用データがある場合）
            const hasPersonalData = monthsToShow.some(m => personalSalesByMonth[m] > 0);
            const legend = document.querySelector('.chart-legend');
            if (legend && hasPersonalData && !legend.querySelector('.dot.personal')) {
                const personalLegend = document.createElement('div');
                personalLegend.className = 'legend-item';
                personalLegend.innerHTML = '<div class="dot personal"></div>個人';
                legend.appendChild(personalLegend);
            }
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
            ledgerBody.innerHTML = filteredLedger.length > 0 ? filteredLedger.map(row => {
                const s = parseNumber(row['売上']);
                const p = parseNumber(row['仕入']);
                const c = parseNumber(row['通信費']);
                const r = parseNumber(row['修繕費']);
                const sp = parseNumber(row['消耗品費']);
                const d = parseNumber(row['諸会費']);
                const f = parseNumber(row['支払手数料']);
                const ms = parseNumber(row['雑費']);
                const rowProfit = s - (p + c + r + sp + d + f + ms);

                const itemName = row['品名'] || "-";

                return `
                    <tr>
                        <td class="sticky-col first">${formatDate(row['日付'], 'short')}</td>
                        <td class="sticky-col second" title="${itemName}">${itemName}</td>
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
            const totals = current;
            updateEl('ledger-total-sales', totals.sales);
            updateEl('ledger-total-purchase', totals.cost);
            updateEl('ledger-total-profit', totals.profit);
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
                </div>
            </div>
        `;

        container.innerHTML = html;

        // イベントバインド
        container.querySelectorAll('[data-year]').forEach(btn => {
            btn.onclick = () => {
                ledgerYear = parseInt(btn.dataset.year);
                renderLedger();
            };
        });
        container.querySelectorAll('[data-period]').forEach(btn => {
            btn.onclick = () => {
                ledgerPeriod = btn.dataset.period;
                renderLedger();
            };
        });
        const monthPicker = document.getElementById('ledger-month-picker');
        if (monthPicker) {
            monthPicker.onchange = () => {
                if (monthPicker.value) {
                    ledgerPeriod = monthPicker.value;
                    renderLedger();
                }
            };
        }
    }

    /**
     * 帳簿データをエクスポート
     */
    async function handleLedgerExport() {
        if (!confirm(`${ledgerYear}年${ledgerPeriod}の帳簿レポートをスプレッドシートに出力しますか？`)) return;

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
            statusArea.innerHTML = `<p>出力中... (${period})</p>`;

            const response = await fetchAPI('exportLedgerReport', { period: period });

            if (response.status === 'success') {
                showToast('レポート出力が完了しました');
                statusArea.className = 'export-status-area success';
                statusArea.innerHTML = `
                    <p>出力完了: ${response.fileName}</p>
                    <a href="${response.url}" target="_blank" class="download-link">
                        <ion-icon name="open-outline"></ion-icon> スプレッドシートを開く
                    </a>
                `;
            } else {
                throw new Error(response.message);
            }
        } catch (e) {
            showToast('エクスポート失敗: ' + e.message, 'error');
            statusArea.className = 'export-status-area error';
            statusArea.innerHTML = `<p>出力エラー: ${e.message}</p>`;
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
        const imageUrl = product ? product['画像URL'] : null;

        if (imageUrl) {
            previewElm.innerHTML = `<img src="${imageUrl}" onclick="showImageModal('${imageUrl}')">`;
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
    window.showImageModal = function (url) {
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
     * 在庫アラートダッシュボードの描画
     */
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

        const activeMfg = lastHistoryData.history['T_製造'] || [];
        const activePur = lastHistoryData.history['T_仕入'] || [];

        alerts.forEach(item => {
            const itemName = item['品名'];
            const qty = parseFloat(item['現在庫数']) || 0;
            const threshold = parseFloat(item['閾値']) || 0;
            const isCritical = qty === 0;
            const cardClass = isCritical ? 'critical' : 'warning';

            // 進行中の取引があるか確認 (件数ではなく合計数量を表示するように変更)
            const mfgTotal = activeMfg.filter(m => {
                const name = (m['完成品名'] || m['品名'] || m['商品名'] || '').toString().trim();
                return name === itemName.trim();
            }).reduce((sum, m) => sum + (parseFloat(m['数量'] || m['製造数量']) || 0), 0);
            
            const purTotal = activePur.filter(p => {
                const name = (p['品名'] || p['商品名'] || p['完成品名'] || '').toString().trim();
                return name === itemName.trim();
            }).reduce((sum, p) => sum + (parseFloat(p['数量']) || 0), 0);
            
            let statusBadge = '';
            if (mfgTotal > 0) {
                statusBadge = `<span class="alert-processing-badge mfg"><ion-icon name="hammer-outline"></ion-icon>製造中(${mfgTotal})</span>`;
            } else if (purTotal > 0) {
                statusBadge = `<span class="alert-processing-badge pur"><ion-icon name="cart-outline"></ion-icon>仕入中(${purTotal})</span>`;
            }

            // M_商品からカテゴリ情報を取得
            const product = (currentMasters['M_商品'] || []).find(m => {
                const name = (m['品名'] || '').toString().trim();
                return name === itemName.trim();
            });
            const category = product ? product['カテゴリ'] : "";

            const isMade = (category === '商品' || category === 'パーツ2');
            const actionBtn = isMade ?
                `<button class="alert-btn make" onclick="jumpToTab('manufacturing', '${itemName}')"><ion-icon name="hammer-outline"></ion-icon>製造へ</button>` :
                `<button class="alert-btn purchase" onclick="jumpToTab('purchase', '${itemName}')"><ion-icon name="cart-outline"></ion-icon>仕入へ</button>`;

            html += `
                <div class="alert-card ${cardClass}">
                    <div class="alert-info">
                        <h4>${itemName}</h4>
                        <div class="alert-status-row">
                            <span class="alert-current-qty">現在: <b>${qty}</b></span>
                            <span class="alert-threshold">閾値: ${threshold}</span>
                            <span class="alert-category">${category}</span>
                            ${statusBadge}
                        </div>
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
     * 特定のタブへ遷移し、品目を選択済みにする
     */
    window.jumpToTab = function (tabId, itemName) {
        const navItem = document.querySelector(`.nav-item[data-target="${tabId}"]`);
        if (navItem) {
            navItem.click(); // タブ切り替え

            // タブ切り替えのアニメーション待ち
            setTimeout(() => {
                const map = { purchase: 'buy-item', manufacturing: 'make-item' };
                const el = document.getElementById(map[tabId]);
                if (el) {
                    el.value = itemName;
                    // 入力イベントを発火させて連動するロジック（プレビュー表示など）を動かす
                    el.dispatchEvent(new Event('input'));
                    el.dispatchEvent(new Event('change'));

                    // スムーズにスクロールして目立たせる
                    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    el.focus();
                }
            }, 100);
        }
    };

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

    function renderMasterList() {
        const container = document.querySelector('.master-list-grid');
        if (!container) return;

        const masterConfig = {
            'M_商品': { title: '商品・パーツ', icon: 'cube-outline' },
            'M_仕入先': { title: '仕入先・購入先', icon: 'business-outline' },
            'M_売先': { title: '販売先(顧客)', icon: 'people-outline' },
            'M_BOM': { title: '製造レシピ(BOM)', icon: 'construct-outline' },
            'M_発送': { title: '配送・送料', icon: 'bus-outline' },
            'M_経費品名': { title: '経費科目名', icon: 'receipt-outline' },
            'M_支払': { title: '支払方法', icon: 'wallet-outline' },
            'T_在庫集計': { title: 'アラート設定(在庫閾値)', icon: 'notifications-outline' },
            'M_ステータス': { title: 'ステータス定義', icon: 'flag-outline' },
            'M_画面制御': { title: '画面入力制御', icon: 'options-outline' }
        };

        container.innerHTML = '';
        Object.keys(masterConfig).forEach(mkey => {
            const conf = masterConfig[mkey];
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

            head.innerHTML = `<tr>${keys.map(k => `<th>${k}</th>`).join('')}<th></th></tr>`;

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
            body.innerHTML = `<tr><td colspan="5" style="color:red; text-align:center; padding: 20px;">取得エラー: ${e.message}</td></tr>`;
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

            tdActions.appendChild(btnStatus);
            tdActions.appendChild(btnEdit);

            tr.innerHTML = html;
            tr.appendChild(tdActions);
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
                inputHtml = `<textarea name="${key}" ${!isEditable ? 'readonly class="readonly-field"' : ''}>${value !== undefined ? value : ''}</textarea>`;
            } else {
                inputHtml = `
                    <input type="${type}" 
                           name="${key}" 
                           value="${value !== undefined ? value : ''}"
                           ${!isEditable ? 'readonly class="readonly-field"' : ''}
                           ${key === '使用FLG' ? 'placeholder="1:有効, 0:無効"' : ''}>
                `;
            }

            group.innerHTML = `<label>${key}</label>${inputHtml}`;
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

                if (value === "" && type !== 'switch') {
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
        // カテゴリが「商品」の場合は選択モーダルを表示 (提案18/25 B案)
        if (category === '商品') {
            showActionSelectionModal(itemName);
            return;
        }

        // それ以外は従来通り直接遷移
        let targetTab = 'purchase';
        let inputId = 'buy-item';

        if (category === 'パーツ2') {
            targetTab = 'manufacturing';
            inputId = 'make-item';
        } else if (category === '経費') {
            targetTab = 'expense';
            inputId = 'exp-item';
        } else if (category === 'パーツ' || category === '単体商品') {
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
    function showActionSelectionModal(itemName) {
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
                        <button class="action-select-btn manufacturing" id="action-btn-manufacturing">
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

        const close = () => overlay.classList.remove('active');

        document.getElementById('action-btn-manufacturing').onclick = () => {
            close();
            executeNavigation('manufacturing', 'make-item', itemName);
        };
        document.getElementById('action-btn-sales').onclick = () => {
            close();
            executeNavigation('sales', 'sale-item', itemName);
        };
        document.getElementById('action-btn-cancel').onclick = close;
        overlay.onclick = (e) => { if (e.target === overlay) close(); };

        setTimeout(() => overlay.classList.add('active'), 10);
    }

});
