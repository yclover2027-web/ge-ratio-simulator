document.addEventListener('DOMContentLoaded', () => {
    const csvFileInput = document.getElementById('csvFileInput');
    const masterFileInput = document.getElementById('masterFileInput');
    const fetchLatestMasterButton = document.getElementById('fetchLatestMasterButton');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const masterFileNameDisplay = document.getElementById('masterFileNameDisplay');
    const simulationWorkspace = document.getElementById('simulationWorkspace');
    const dashboard = document.getElementById('dashboard');
    const drugListSection = document.getElementById('drugListSection');
    const drugTableBody = document.getElementById('drugTableBody');
    const searchInput = document.getElementById('searchInput');
    const filterSelect = document.getElementById('filterSelect');
    const sortAmountHeader = document.getElementById('sortAmountHeader');
    
    const currentRateEl = document.getElementById('currentRate');
    const simulatedRateEl = document.getElementById('simulatedRate');
    const rateDiffEl = document.getElementById('rateDiff');
    const progressFill = document.getElementById('progressFill');
    const progressMarker = document.getElementById('progressMarker');

    let globalDrugData = [];
    let baseGenericTotal = 0;
    let baseDenominatorTotal = 0;
    let sortDirection = 'none'; // 'none', 'desc', 'asc'

    const knownLatestMasterExcelUrl = 'https://www.mhlw.go.jp/topics/2026/04/xls/tp20260612-01_05.xlsx';
    
    // ひらがなをカタカナに変換する関数
    function hiraganaToKatakana(str) {
        return str.replace(/[\u3041-\u3096]/g, match => String.fromCharCode(match.charCodeAt(0) + 0x60));
    }

    csvFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        fileNameDisplay.textContent = `読み込み中: ${file.name}`;

        const reader = new FileReader();
        reader.readAsText(file, 'Shift_JIS');

        reader.onload = function(event) {
            const csvText = event.target.result;
            processCSV(csvText);
            fileNameDisplay.textContent = `読み込み完了: ${file.name}`;
        };

        reader.onerror = function() {
            fileNameDisplay.textContent = '読み込みに失敗しました。';
        };
    });

    fetchLatestMasterButton.addEventListener('click', async () => {
        if (globalDrugData.length === 0) {
            alert('先に使用薬剤一覧(CSV)を読み込んでください。');
            return;
        }

        fetchLatestMasterButton.disabled = true;
        masterFileNameDisplay.textContent = '厚生労働省の最新版データを確認中...';

        try {
            const latestMaster = await downloadLatestMasterViaLocalServer();
            parseMasterArrayBuffer(latestMaster.arrayBuffer, latestMaster.fileName, latestMaster.sourceUrl);
        } catch (err) {
            console.error(err);
            masterFileNameDisplay.innerHTML = `自動取得に失敗しました。<a href="${knownLatestMasterExcelUrl}" target="_blank" rel="noopener">最新版Excelを開く</a> から保存して、下の枠で読み込んでください。`;
            alert('最新版Excelの自動取得に失敗しました。ローカルサーバー経由で開いているか確認してください。難しい場合は、画面のリンクからExcelを保存して、下の枠で読み込んでください。');
        } finally {
            fetchLatestMasterButton.disabled = false;
        }
    });

    // Handle Master Data Upload
    masterFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        if (globalDrugData.length === 0) {
            alert('先に使用薬剤一覧(CSV)を読み込んでください。');
            masterFileInput.value = '';
            return;
        }

        masterFileNameDisplay.textContent = `読み込み中: ${file.name}`;
        
        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                parseMasterArrayBuffer(event.target.result, file.name, '');
            } catch (err) {
                console.error(err);
                masterFileNameDisplay.textContent = '読み込みに失敗しました。';
                alert('エクセルファイルの解析に失敗しました。');
            }
        };
        reader.readAsArrayBuffer(file);
    });

    async function downloadLatestMasterViaLocalServer() {
        // 厚労省サイトはブラウザから直接Excelの中身を読むための許可ヘッダーがありません。
        // そのため、同じローカルサーバー上のAPIを通して取得し、ブラウザ制限に引っかからない形にします。
        const response = await fetch('/api/latest-mhlw-master', { cache: 'no-store' });
        if (!response.ok) {
            throw new Error(`最新版Excelの取得に失敗しました: ${response.status}`);
        }
        return {
            arrayBuffer: await response.arrayBuffer(),
            fileName: response.headers.get('X-File-Name') || 'latest-master.xlsx',
            sourceUrl: response.headers.get('X-Source-Url') || ''
        };
    }

    function parseMasterArrayBuffer(arrayBuffer, sourceName, sourceUrl) {
        const data = new Uint8Array(arrayBuffer);
        // XLSX は CDNから読み込み済みです。Excelを2次元配列へ変換して、既存の更新処理に渡します。
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, {header: 1}); // 2次元配列

        const updatedCount = updateGenericsFromMaster(json, sourceName);
        const sourceLink = sourceUrl ? `（取得元: ${sourceUrl}）` : '';
        masterFileNameDisplay.textContent = `更新完了: ${sourceName} / ${updatedCount} 件 ${sourceLink}`;
    }

    function updateGenericsFromMaster(rows, sourceName = 'マスターデータ') {
        let headerRowIdx = -1;
        let yjColIdx = -1;
        let typeColIdx = -1;
        let nameColIdx = -1;

        // ヘッダー行を検索
        for (let i = 0; i < Math.min(rows.length, 50); i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            for (let j = 0; j < row.length; j++) {
                const cellStr = String(row[j] || '').replace(/\s+/g, '');
                if (cellStr.includes('YJ') || cellStr.includes('ＹＪ') || cellStr.includes('薬価基準収載医薬品コード')) yjColIdx = j;
                if ((cellStr.includes('後発') && cellStr.includes('区分')) || cellStr.includes('新区分') || cellStr.includes('各先発医薬品の後発医薬品の有無に関する情報')) typeColIdx = j;
                if (typeColIdx === -1 && cellStr === '区分') typeColIdx = j; 
                if (cellStr.includes('品名') || cellStr.includes('医薬品名')) nameColIdx = j;
            }
            if (yjColIdx !== -1 && typeColIdx !== -1) {
                headerRowIdx = i;
                break;
            }
        }

        if (headerRowIdx === -1) {
            alert('マスターデータ内に「YJコード」と「区分」を示す列が見つかりませんでした。');
            return;
        }

        const exactMap = new Map();
        const prefix9Map = new Map();

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length <= Math.max(yjColIdx, typeColIdx)) continue;
            
            let yjStr = String(row[yjColIdx] || '').replace(/[^a-zA-Z0-9Ａ-Ｚａ-ｚ０-９]/g, '');
            let typeStr = String(row[typeColIdx] || '').trim();
            if (!yjStr || yjStr.length < 9) continue;
            
            // 1. 12桁の完全一致用に登録
            exactMap.set(yjStr, typeStr);

            // 2. 前方9桁の共通マッチ用に登録
            const yj9 = yjStr.substring(0, 9);
            if (!prefix9Map.has(yj9)) {
                prefix9Map.set(yj9, []);
            }
            // 9桁マッチングでは重複を避けるため、YJコード自体を持たせておく
            prefix9Map.get(yj9).push({ yj: yjStr, type: typeStr });
        }

        let updatedCount = 0;
        globalDrugData.forEach(drug => {
            if (!drug.yjCode) return;
            
            const csvYj = drug.yjCode.replace(/[^a-zA-Z0-9Ａ-Ｚａ-ｚ０-９]/g, '');
            if (csvYj.length < 9) return;
            
            let matchedType = null;
            
            // ① 12桁の完全一致
            if (exactMap.has(csvYj)) {
                matchedType = exactMap.get(csvYj);
            } else {
                // ② 完全一致しない場合、前方9桁で照合
                const yj9 = csvYj.substring(0, 9);
                const candidates = prefix9Map.get(yj9) || [];
                
                if (candidates.length === 1) {
                    matchedType = candidates[0].type;
                } else if (candidates.length > 1) {
                    // ③ 複数の候補が出る場合は、前方11桁での照合を試みる
                    const yj11 = csvYj.substring(0, 11);
                    const match11 = candidates.find(c => c.yj.substring(0, 11) === yj11);
                    if (match11) {
                        matchedType = match11.type;
                    }
                }
            }
            
            if (matchedType !== null) {
                // マッチした区分で更新する
                // 空白の場合は分かりやすく表示する
                drug.genericType = matchedType === '' ? '空白 (Excel上書き)' : matchedType + ' (Excel上書き)';
                drug.isGeneric = false;
                drug.isOriginalWithGeneric = false;
                drug.badgeClass = 'badge-other';
                drug.typeLabel = 'その他';

                if (matchedType.includes('3') || matchedType.includes('３') || matchedType.includes('後発品')) {
                    drug.isGeneric = true;
                    drug.badgeClass = 'badge-generic';
                    drug.typeLabel = '後発品';
                } else if (matchedType.includes('2') || matchedType.includes('２') || matchedType.includes('先発品(後発あり)') || matchedType.includes('先発（後発あり）')) {
                    drug.isOriginalWithGeneric = true;
                    drug.badgeClass = 'badge-original';
                    drug.typeLabel = '先発品(後発あり)';
                } else if (matchedType.includes('★') || matchedType.includes('☆') || matchedType === '') {
                    drug.isGeneric = false;
                    drug.isOriginalWithGeneric = false;
                    drug.badgeClass = 'badge-other';
                    drug.typeLabel = '対象外(特例/空白)';
                } else if (matchedType.includes('1') || matchedType.includes('１') || matchedType.includes('先発')) {
                    drug.isGeneric = false;
                    drug.isOriginalWithGeneric = false;
                    drug.badgeClass = 'badge-other';
                    drug.typeLabel = '先発品(後発なし)';
                }
                updatedCount++;
            }
        });

        alert(`${sourceName} の読み込みが完了しました！\n${updatedCount} 件のお薬の「後発品区分」を最新の状態に更新しました。`);
        calculateBaseRate();
        renderTable();
        return updatedCount;
    }

    function parseCSVLine(text) {
        const result = [];
        let curVal = '';
        let inQuotes = false;
        
        for (let i = 0; i < text.length; i++) {
            const char = text[i];
            
            if (inQuotes) {
                if (char === '"') {
                    if (i + 1 < text.length && text[i + 1] === '"') {
                        curVal += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    curVal += char;
                }
            } else {
                if (char === '"') {
                    inQuotes = true;
                } else if (char === ',') {
                    result.push(curVal);
                    curVal = '';
                } else if (char === '\r' || char === '\n') {
                } else {
                    curVal += char;
                }
            }
        }
        result.push(curVal);
        return result;
    }

    function processCSV(csvText) {
        const lines = csvText.split('\n');
        
        let headerParsed = false;
        let colMap = {};
        let parsedData = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            
            const cols = parseCSVLine(line);
            
            if (!headerParsed && cols.includes('医薬品名') && cols.includes('使用量')) {
                cols.forEach((col, index) => {
                    colMap[col] = index;
                });
                headerParsed = true;
                continue;
            }

            if (headerParsed) {
                if (cols.length > colMap['医薬品名']) {
                    const name = cols[colMap['医薬品名']];
                    const amountStr = cols[colMap['使用量']] || '0';
                    const unit = cols[colMap['単位']] || '';
                    const genericType = cols[colMap['後発区分']] || '';
                    const yjCode = cols[colMap['YJコード']] || '';
                    // 先頭8桁を使用（9桁目を無視することで、同じ成分・規格の「普通錠とOD錠」を同グループとして扱います）
                    const yjBase = yjCode.substring(0, 8); 
                    
                    if (!name) continue;

                    const amount = parseFloat(amountStr.replace(/,/g, '')) || 0;

                    let isGeneric = false;
                    let isOriginalWithGeneric = false;
                    let badgeClass = 'badge-other';
                    let typeLabel = 'その他';

                    if (genericType.includes('3:後発品')) {
                        isGeneric = true;
                        badgeClass = 'badge-generic';
                        typeLabel = '後発品';
                    } else if (genericType.startsWith('2:')) {
                        isOriginalWithGeneric = true;
                        badgeClass = 'badge-original';
                        typeLabel = '先発品(後発あり)';
                    }

                    parsedData.push({
                        id: i,
                        name: name,
                        nameKana: hiraganaToKatakana(name), // カタカナ変換名も保持
                        amount: amount,
                        simulatedAmount: amount,
                        unit: unit,
                        genericType: genericType,
                        yjCode: yjCode,
                        yjBase: yjBase,
                        isGeneric: isGeneric,
                        isOriginalWithGeneric: isOriginalWithGeneric,
                        badgeClass: badgeClass,
                        typeLabel: typeLabel,
                        excludeFromCalc: false,
                        isReplacedWithGeneric: false
                    });
                }
            }
        }

        globalDrugData = parsedData;
        calculateBaseRate();
        renderTable();
        
        simulationWorkspace.classList.remove('hidden');
        dashboard.classList.remove('hidden');
        drugListSection.classList.remove('hidden');
    }

    function calculateBaseRate() {
        let genericTotal = 0;
        let denominatorTotal = 0;

        globalDrugData.forEach(drug => {
            if (drug.isGeneric) {
                genericTotal += drug.amount;
                denominatorTotal += drug.amount;
            } else if (drug.isOriginalWithGeneric) {
                denominatorTotal += drug.amount;
            }
        });

        baseGenericTotal = genericTotal;
        baseDenominatorTotal = denominatorTotal;

        const rate = denominatorTotal > 0 ? (genericTotal / denominatorTotal) * 100 : 0;
        
        currentRateEl.textContent = `${rate.toFixed(1)}%`;
        progressMarker.style.left = `${rate.toFixed(1)}%`;
        
        updateSimulationRate();
    }

    function updateSimulationRate() {
        let simGenericTotal = 0;
        let simDenominatorTotal = 0;

        globalDrugData.forEach(drug => {
            if (drug.excludeFromCalc) return; // 計算から除外

            const amt = drug.simulatedAmount;

            if (drug.isReplacedWithGeneric) {
                // ジェネリックに振り替えられた場合、分子にも分母にも加算
                simGenericTotal += amt;
                simDenominatorTotal += amt;
            } else {
                if (drug.isGeneric) {
                    simGenericTotal += amt;
                    simDenominatorTotal += amt;
                } else if (drug.isOriginalWithGeneric) {
                    simDenominatorTotal += amt;
                }
            }
        });

        const rate = simDenominatorTotal > 0 ? (simGenericTotal / simDenominatorTotal) * 100 : 0;
        simulatedRateEl.textContent = `${rate.toFixed(1)}%`;
        progressFill.style.width = `${rate.toFixed(1)}%`;

        const baseRate = baseDenominatorTotal > 0 ? (baseGenericTotal / baseDenominatorTotal) * 100 : 0;
        const diff = rate - baseRate;
        
        if (diff > 0) {
            rateDiffEl.textContent = `+${diff.toFixed(1)}% ↑`;
            rateDiffEl.className = 'stat-diff positive';
        } else if (diff < 0) {
            rateDiffEl.textContent = `${diff.toFixed(1)}% ↓`;
            rateDiffEl.className = 'stat-diff negative';
        } else {
            rateDiffEl.textContent = '±0.0%';
            rateDiffEl.className = 'stat-diff';
        }
    }

    function renderTable() {
        const rawSearchTerm = searchInput.value.toLowerCase();
        // 入力されたひらがなをカタカナに変換してマッチングしやすくする
        const searchTerm = hiraganaToKatakana(rawSearchTerm);
        const filterVal = filterSelect.value;
        
        // フィルタリング
        let displayData = globalDrugData.filter(drug => {
            if (filterVal === 'generic' && !drug.isGeneric) return false;
            if (filterVal === 'original' && !drug.isOriginalWithGeneric) return false;
            if (searchTerm && !(drug.name.toLowerCase().includes(searchTerm) || drug.nameKana.toLowerCase().includes(searchTerm))) return false;
            return true;
        });

        // ソート処理
        const sortIconEl = sortAmountHeader.querySelector('.sort-icon');
        if (sortDirection === 'desc') {
            displayData.sort((a, b) => b.amount - a.amount);
            sortIconEl.textContent = '↓';
        } else if (sortDirection === 'asc') {
            displayData.sort((a, b) => a.amount - b.amount);
            sortIconEl.textContent = '↑';
        } else {
            displayData.sort((a, b) => a.id - b.id);
            sortIconEl.textContent = '';
        }

        drugTableBody.innerHTML = '';

        displayData.forEach(drug => {

            const tr = document.createElement('tr');
            
            // 行のスタイル設定
            if (drug.excludeFromCalc) tr.classList.add('row-excluded');
            else if (drug.isReplacedWithGeneric) tr.classList.add('row-replaced');
            else if (!drug.isGeneric && !drug.isOriginalWithGeneric) tr.style.opacity = '0.5';

            let buttonsHtml = '';
            
            // ジェネリック振替ボタン (先発品(後発あり)のみ)
            if (drug.isOriginalWithGeneric) {
                buttonsHtml += `<button class="action-btn replace-btn ${drug.isReplacedWithGeneric ? 'active' : ''}" data-action="replace" data-id="${drug.id}">
                    🔄 ジェネリックへ振替
                </button>`;
            }

            // 関連まるごと除外ボタン
            // 先発品または後発品で、yjBase がある場合のみ表示
            if ((drug.isOriginalWithGeneric || drug.isGeneric) && drug.yjBase) {
                buttonsHtml += `<button class="action-btn exclude-all-btn" data-action="exclude-group" data-id="${drug.id}" title="同じ成分の医薬品（先発・後発）をまとめて除外します">
                    🗑 まとめて除外
                </button>`;
            }
            
            // 除外ボタン
            buttonsHtml += `<button class="action-btn exclude-btn ${drug.excludeFromCalc ? 'active' : ''}" data-action="exclude" data-id="${drug.id}">
                ⛔ 計算から除外
            </button>`;

            tr.innerHTML = `
                <td><strong>${drug.name}</strong></td>
                <td><span class="badge ${drug.badgeClass}">${drug.typeLabel}</span><br><small style="color:var(--text-secondary);font-size:0.7rem;">${drug.genericType}</small></td>
                <td>${drug.amount.toLocaleString(undefined, {minimumFractionDigits: 1, maximumFractionDigits: 2})}</td>
                <td>${drug.unit}</td>
                <td>
                    <input type="number" step="0.1" min="0" 
                           class="amount-input" 
                           data-id="${drug.id}" 
                           value="${drug.simulatedAmount}"
                           ${(!drug.isGeneric && !drug.isOriginalWithGeneric) || drug.excludeFromCalc ? 'disabled title="操作不可"' : ''}>
                </td>
                <td>${buttonsHtml}</td>
            `;
            drugTableBody.appendChild(tr);
        });

        // Event Listeners for new inputs and buttons
        document.querySelectorAll('.amount-input').forEach(input => {
            input.addEventListener('input', (e) => {
                const id = parseInt(e.target.getAttribute('data-id'));
                const newAmount = parseFloat(e.target.value) || 0;
                const drug = globalDrugData.find(d => d.id === id);
                if (drug) {
                    drug.simulatedAmount = newAmount;
                    updateSimulationRate();
                }
            });
        });

        document.querySelectorAll('.action-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = parseInt(e.currentTarget.getAttribute('data-id'));
                const action = e.currentTarget.getAttribute('data-action');
                const drug = globalDrugData.find(d => d.id === id);
                
                if (drug) {
                    if (action === 'replace') {
                        drug.isReplacedWithGeneric = !drug.isReplacedWithGeneric;
                    } else if (action === 'exclude') {
                        drug.excludeFromCalc = !drug.excludeFromCalc;
                    } else if (action === 'exclude-group') {
                        // yjBase が一致する医薬品をすべて抽出し、除外状態を同期（トグル）する
                        const groupDrugs = globalDrugData.filter(d => d.yjBase === drug.yjBase && d.yjBase !== '');
                        // グループの中に1つでも「除外されていない（計算に含まれる）」ものがあれば、すべて除外にする。
                        // すべて除外されていれば、すべて解除する。
                        const isAnyInclude = groupDrugs.some(d => !d.excludeFromCalc);
                        groupDrugs.forEach(d => {
                            d.excludeFromCalc = isAnyInclude;
                        });
                    }
                    updateSimulationRate();
                    renderTable(); // ボタンの見た目や行スタイルを更新するため再描画
                }
            });
        });
    }

    sortAmountHeader.addEventListener('click', () => {
        if (sortDirection === 'none') {
            sortDirection = 'desc';
        } else if (sortDirection === 'desc') {
            sortDirection = 'asc';
        } else {
            sortDirection = 'none';
        }
        renderTable();
    });

    searchInput.addEventListener('input', renderTable);
    filterSelect.addEventListener('change', renderTable);
});
