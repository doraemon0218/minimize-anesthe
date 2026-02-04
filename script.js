// グローバル変数
let surgeryData = [];
let timelineData = [];
let chart = null;
let assignmentRules = {
    hasSupervisor: false,
    allowSupervisorAnesthesia: false,
    strictSingleCaseRule: true,
    allowTraineeSolo: true  // 専攻医の単独麻酔を許容するか
};

// DOM要素の取得
const excelFileInput = document.getElementById('excelFile');
const residentInput = document.getElementById('resident');
const traineeInput = document.getElementById('trainee');
const fulltimeInput = document.getElementById('fulltime');
const parttimeInput = document.getElementById('parttime');
const nurseInput = document.getElementById('nurse');
const totalDoctorsSpan = document.getElementById('totalDoctors');
const analyzeBtn = document.getElementById('analyzeBtn');
const resultsSection = document.getElementById('resultsSection');

// モーダル関連
const rulesModal = document.getElementById('rulesModal');
const rulesBtn = document.getElementById('rulesBtn');
const closeModal = document.querySelector('.close');
const saveRulesBtn = document.getElementById('saveRulesBtn');
const cancelRulesBtn = document.getElementById('cancelRulesBtn');
const hasSupervisorCheck = document.getElementById('hasSupervisor');
const allowSupervisorAnesthesiaCheck = document.getElementById('allowSupervisorAnesthesia');
const strictSingleCaseRuleCheck = document.getElementById('strictSingleCaseRule');
const allowTraineeSoloCheck = document.getElementById('allowTraineeSolo');

// イベントリスナー
excelFileInput.addEventListener('change', handleFileUpload);
[residentInput, traineeInput, fulltimeInput, parttimeInput, nurseInput].forEach(input => {
    input.addEventListener('input', updateTotalDoctors);
    input.addEventListener('input', checkAnalyzeButton);
});
analyzeBtn.addEventListener('click', analyzeSchedule);

// モーダル関連イベント
rulesBtn.addEventListener('click', () => {
    rulesModal.style.display = 'block';
    // 現在の設定を反映
    hasSupervisorCheck.checked = assignmentRules.hasSupervisor;
    allowSupervisorAnesthesiaCheck.checked = assignmentRules.allowSupervisorAnesthesia;
    strictSingleCaseRuleCheck.checked = assignmentRules.strictSingleCaseRule;
    allowTraineeSoloCheck.checked = assignmentRules.allowTraineeSolo;
});

closeModal.addEventListener('click', () => {
    rulesModal.style.display = 'none';
});

cancelRulesBtn.addEventListener('click', () => {
    rulesModal.style.display = 'none';
});

saveRulesBtn.addEventListener('click', () => {
    assignmentRules.hasSupervisor = hasSupervisorCheck.checked;
    assignmentRules.allowSupervisorAnesthesia = allowSupervisorAnesthesiaCheck.checked;
    assignmentRules.strictSingleCaseRule = strictSingleCaseRuleCheck.checked;
    assignmentRules.allowTraineeSolo = allowTraineeSoloCheck.checked;
    rulesModal.style.display = 'none';
    alert('配置ルールを保存しました');
});

window.addEventListener('click', (event) => {
    if (event.target === rulesModal) {
        rulesModal.style.display = 'none';
    }
});

// 合計人数の更新
function updateTotalDoctors() {
    const total = parseInt(residentInput.value || 0) +
                  parseInt(traineeInput.value || 0) +
                  parseInt(fulltimeInput.value || 0) +
                  parseInt(parttimeInput.value || 0) +
                  parseInt(nurseInput.value || 0);
    totalDoctorsSpan.textContent = total;
}

// 分析ボタンの有効化チェック
function checkAnalyzeButton() {
    const hasFile = excelFileInput.files.length > 0;
    const hasDoctors = parseInt(totalDoctorsSpan.textContent) > 0;
    analyzeBtn.disabled = !(hasFile && hasDoctors);
}

// エクセルファイルの読み込み
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 最初のシートを読み込む
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            parseSurgeryData(jsonData);
            checkAnalyzeButton();
        } catch (error) {
            alert('エクセルファイルの読み込みに失敗しました: ' + error.message);
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

// 手術データのパース
function parseSurgeryData(jsonData) {
    surgeryData = [];
    
    // 列名のマッピング（様々な形式に対応）
    const columnMappings = {
        '入室時刻': ['入室時刻', '入室', '入室時間', '入室予定時刻'],
        '手術開始': ['手術開始予定時刻', '手術開始', '開始時刻', '開始時間', '手術開始時刻'],
        '手術終了': ['手術終了予定時刻', '手術終了', '終了時刻', '終了時間', '手術終了時刻'],
        '退室時刻': ['退室予定時刻', '退室', '退室時間', '退室時刻'],
        '手術内容': ['手術内容', '手術', '内容', '手術名', '術式']
    };
    
    jsonData.forEach((row, index) => {
        try {
            // 列名を自動検出
            const findColumn = (mappings) => {
                for (const key in row) {
                    const normalizedKey = key.trim();
                    if (mappings.some(m => normalizedKey.includes(m) || m.includes(normalizedKey))) {
                        return row[key];
                    }
                }
                return null;
            };
            
            const entryTime = findColumn(columnMappings['入室時刻']);
            const startTime = findColumn(columnMappings['手術開始']);
            const endTime = findColumn(columnMappings['手術終了']);
            const exitTime = findColumn(columnMappings['退室時刻']);
            const surgeryType = findColumn(columnMappings['手術内容']);
            
            if (!entryTime || !startTime || !endTime || !exitTime) {
                console.warn(`行 ${index + 2}: 必要な時刻データが不足しています`);
                return;
            }
            
            // 時刻のパース
            const entry = parseDateTime(entryTime);
            const start = parseDateTime(startTime);
            const end = parseDateTime(endTime);
            const exit = parseDateTime(exitTime);
            
            if (!entry || !start || !end || !exit) {
                console.warn(`行 ${index + 2}: 時刻のパースに失敗しました`);
                return;
            }
            
            surgeryData.push({
                entry: entry,
                start: start,
                end: end,
                exit: exit,
                surgeryType: surgeryType || '未指定',
                id: index
            });
        } catch (error) {
            console.warn(`行 ${index + 2}の処理中にエラー:`, error);
        }
    });
    
    console.log(`読み込んだ手術データ: ${surgeryData.length}件`);
}

// 日時文字列のパース
function parseDateTime(dateTimeStr) {
    if (!dateTimeStr) return null;
    
    // Excelの日付シリアル値の場合
    if (typeof dateTimeStr === 'number') {
        // Excelの日付は1900年1月1日からの日数
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + dateTimeStr * 24 * 60 * 60 * 1000);
        return date;
    }
    
    // 文字列の場合
    const str = String(dateTimeStr).trim();
    
    // 様々な形式を試す
    const formats = [
        /(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})\s+(\d{1,2}):(\d{1,2})/,
        /(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})\s+(\d{1,2}):(\d{1,2})/,
        /(\d{1,2}):(\d{1,2})/,
    ];
    
    for (const format of formats) {
        const match = str.match(format);
        if (match) {
            let year, month, day, hour, minute;
            
            if (format === formats[0]) {
                // YYYY-MM-DD HH:MM
                [, year, month, day, hour, minute] = match;
            } else if (format === formats[1]) {
                // MM-DD-YYYY HH:MM
                [, month, day, year, hour, minute] = match;
            } else {
                // HH:MM only (今日の日付を使用)
                const today = new Date();
                year = today.getFullYear();
                month = today.getMonth() + 1;
                day = today.getDate();
                [, hour, minute] = match;
            }
            
            const date = new Date(
                parseInt(year),
                parseInt(month) - 1,
                parseInt(day),
                parseInt(hour),
                parseInt(minute)
            );
            
            if (!isNaN(date.getTime())) {
                return date;
            }
        }
    }
    
    // 最後の手段: Dateコンストラクタ
    const date = new Date(str);
    if (!isNaN(date.getTime())) {
        return date;
    }
    
    return null;
}

// 配置ルールに基づく必要人数の計算（キャパシティベース）
function calculateRequiredStaff(timeSlot, activeSurgeries) {
    if (activeSurgeries.length === 0) {
        return {
            requiredCapacity: 0,
            suppliedCapacity: 0,
            requiredPeople: 0,
            assignments: [],
            unassignedSurgeries: []
        };
    }
    
    const staff = {
        resident: parseInt(residentInput.value || 0),
        trainee: parseInt(traineeInput.value || 0),
        fulltime: parseInt(fulltimeInput.value || 0),
        parttime: parseInt(parttimeInput.value || 0),
        nurse: parseInt(nurseInput.value || 0)
    };
    
    // 責任者がいる場合、常勤専門医から1人を引く
    if (assignmentRules.hasSupervisor && !assignmentRules.allowSupervisorAnesthesia) {
        staff.fulltime = Math.max(0, staff.fulltime - 1);
    }
    
    // 供給可能なキャパシティを計算
    // 専門医: 最大2症例まで（指導症例×2、または指導症例×1+単独×1）
    // 初期研修医・看護師: 単独ではキャパシティなし（専門医とペアで1症例、専門医のキャパシティに含まれる）
    // 専攻医: 単独で1症例担当可能（設定で許可されている場合のみ）
    let suppliedCapacity = (staff.fulltime + staff.parttime) * 2; // 専門医のキャパシティ
    if (assignmentRules.allowTraineeSolo) {
        suppliedCapacity += staff.trainee * 1; // 専攻医の単独キャパシティ
    }
    
    // 必要なキャパシティ（症例数）
    const requiredCapacity = activeSurgeries.length;
    
    // 各手術に必要な人員を割り当てる（貪欲法で最適化）
    let requiredPeople = 0;
    const assignments = [];
    const unassignedSurgeries = [];
    
    // 専門医の担当状況を追跡（各専門医が何症例を担当しているか）
    const specialistAssignments = {
        fulltime: new Array(staff.fulltime).fill(0), // 各常勤専門医の担当症例数
        parttime: new Array(staff.parttime).fill(0)   // 各非常勤専門医の担当症例数
    };
    
    // 各手術に対して人員を割り当て
    for (const surgery of activeSurgeries) {
        const assignment = assignStaffToSurgery(
            surgery, 
            staff, 
            assignments, 
            specialistAssignments,
            timeSlot
        );
        if (assignment) {
            assignments.push({
                ...assignment,
                surgery: surgery
            });
            requiredPeople += assignment.required;
            
            // 専門医の担当状況を更新
            if (assignment.fulltimeIndex !== undefined) {
                specialistAssignments.fulltime[assignment.fulltimeIndex]++;
            }
            if (assignment.parttimeIndex !== undefined) {
                specialistAssignments.parttime[assignment.parttimeIndex]++;
            }
        } else {
            // 割り当てできない場合
            unassignedSurgeries.push(surgery);
            requiredPeople += 1; // 最低1人は必要
        }
    }
    
    return {
        requiredCapacity: requiredCapacity,
        suppliedCapacity: suppliedCapacity,
        requiredPeople: requiredPeople,
        assignments: assignments,
        unassignedSurgeries: unassignedSurgeries,
        staff: staff
    };
}

// 手術への人員割り当て
function assignStaffToSurgery(surgery, availableStaff, existingAssignments, specialistAssignments, timeSlot) {
    // 既に割り当てられている人員を確認
    const usedStaff = {
        resident: 0,
        trainee: 0,
        nurse: 0
    };
    
    existingAssignments.forEach(assign => {
        if (assign.surgeryId !== surgery.id) {
            if (assign.resident) usedStaff.resident += assign.resident;
            if (assign.trainee) usedStaff.trainee += assign.trainee;
            if (assign.nurse) usedStaff.nurse += assign.nurse;
        }
    });
    
    const remainingStaff = {
        resident: Math.max(0, availableStaff.resident - usedStaff.resident),
        trainee: Math.max(0, availableStaff.trainee - usedStaff.trainee),
        nurse: Math.max(0, availableStaff.nurse - usedStaff.nurse)
    };
    
    // 利用可能な専門医を探す（2症例まで担当可能、1症例のみ担当中は原則追加担当しない）
    // ただし、初期研修医と一緒の場合は指導として1症例まで可能
    const findAvailableSpecialist = (forResident = false) => {
        // 常勤専門医から探す
        for (let i = 0; i < specialistAssignments.fulltime.length; i++) {
            const caseCount = specialistAssignments.fulltime[i];
            // 2症例まで担当可能
            if (caseCount < 2) {
                // 1症例のみ担当中の場合は、厳格ルール適用時は追加担当しない
                // ただし、初期研修医の指導の場合は可能
                if (caseCount === 1 && assignmentRules.strictSingleCaseRule && !forResident) {
                    continue;
                }
                return { type: 'fulltime', index: i };
            }
        }
        
        // 非常勤専門医から探す
        for (let i = 0; i < specialistAssignments.parttime.length; i++) {
            const caseCount = specialistAssignments.parttime[i];
            if (caseCount < 2) {
                if (caseCount === 1 && assignmentRules.strictSingleCaseRule && !forResident) {
                    continue;
                }
                return { type: 'parttime', index: i };
            }
        }
        
        return null;
    };
    
    // 優先順位の修正：
    // - 初期研修医: 必ず専門医とペア（単独不可）
    // - 麻酔看護師: 必ず専門医とペア（単独不可）
    // - 専攻医: 専門医とペア、または単独でも可能
    
    // 1. 初期研修医 + 専門医（必須ペア、専門医が1症例のみ担当中でも指導として可能）
    if (remainingStaff.resident > 0) {
        const specialist = findAvailableSpecialist(true); // 初期研修医の指導として
        if (specialist) {
            return {
                surgeryId: surgery.id,
                resident: 1,
                fulltime: specialist.type === 'fulltime' ? 1 : 0,
                parttime: specialist.type === 'parttime' ? 1 : 0,
                fulltimeIndex: specialist.type === 'fulltime' ? specialist.index : undefined,
                parttimeIndex: specialist.type === 'parttime' ? specialist.index : undefined,
                required: 2
            };
        }
        // 初期研修医は専門医とペアが必須なので、専門医がいない場合は割り当て不可
    }
    
    // 2. 専攻医 + 専門医（推奨）
    if (remainingStaff.trainee > 0) {
        const specialist = findAvailableSpecialist();
        if (specialist) {
            return {
                surgeryId: surgery.id,
                trainee: 1,
                fulltime: specialist.type === 'fulltime' ? 1 : 0,
                parttime: specialist.type === 'parttime' ? 1 : 0,
                fulltimeIndex: specialist.type === 'fulltime' ? specialist.index : undefined,
                parttimeIndex: specialist.type === 'parttime' ? specialist.index : undefined,
                required: 2
            };
        }
    }
    
    // 3. 看護師 + 専門医（必須ペア）
    if (remainingStaff.nurse > 0) {
        const specialist = findAvailableSpecialist();
        if (specialist) {
            return {
                surgeryId: surgery.id,
                nurse: 1,
                fulltime: specialist.type === 'fulltime' ? 1 : 0,
                parttime: specialist.type === 'parttime' ? 1 : 0,
                fulltimeIndex: specialist.type === 'fulltime' ? specialist.index : undefined,
                parttimeIndex: specialist.type === 'parttime' ? specialist.index : undefined,
                required: 2
            };
        }
        // 看護師は専門医とペアが必須なので、専門医がいない場合は割り当て不可
    }
    
    // 4. 専攻医単独（設定で許容されている場合のみ可能、ただし1症例のみ担当可能）
    // remainingStaffで既に他の手術に割り当てられている専攻医は除外されている
    if (remainingStaff.trainee > 0 && assignmentRules.allowTraineeSolo) {
        return {
            surgeryId: surgery.id,
            trainee: 1,
            required: 1
        };
    }
    
    // 5. 専門医単独
    const specialist = findAvailableSpecialist();
    if (specialist) {
        return {
            surgeryId: surgery.id,
            fulltime: specialist.type === 'fulltime' ? 1 : 0,
            parttime: specialist.type === 'parttime' ? 1 : 0,
            fulltimeIndex: specialist.type === 'fulltime' ? specialist.index : undefined,
            parttimeIndex: specialist.type === 'parttime' ? specialist.index : undefined,
            required: 1
        };
    }
    
    // 人員不足
    return null;
}

// 15分単位のタイムライン生成（配置ルール適用）
function generateTimeline() {
    if (surgeryData.length === 0) return [];
    
    // 全手術の開始時刻と終了時刻から範囲を決定
    const allTimes = [];
    surgeryData.forEach(surgery => {
        allTimes.push(surgery.entry, surgery.exit);
    });
    
    const minTime = new Date(Math.min(...allTimes));
    const maxTime = new Date(Math.max(...allTimes));
    
    // 15分単位で切り上げ・切り下げ
    minTime.setMinutes(Math.floor(minTime.getMinutes() / 15) * 15, 0, 0);
    maxTime.setMinutes(Math.ceil(maxTime.getMinutes() / 15) * 15, 0, 0);
    
    // 15分単位のタイムスロットを生成
    const timeline = [];
    const current = new Date(minTime);
    
    while (current <= maxTime) {
        const timeSlot = new Date(current);
        
        // この時間帯に進行中の手術を取得
        const activeSurgeries = surgeryData.filter(surgery => 
            timeSlot >= surgery.entry && timeSlot < surgery.exit
        );
        
        // 配置ルールに基づいて必要人数を計算（キャパシティベース）
        const result = calculateRequiredStaff(timeSlot, activeSurgeries);
        
        timeline.push({
            time: new Date(timeSlot),
            requiredCapacity: result.requiredCapacity,
            suppliedCapacity: result.suppliedCapacity,
            requiredPeople: result.requiredPeople,
            activeSurgeries: activeSurgeries.length,
            assignments: result.assignments,
            unassignedSurgeries: result.unassignedSurgeries,
            staff: result.staff
        });
        
        // 15分進める
        current.setMinutes(current.getMinutes() + 15);
    }
    
    return timeline;
}

// 分析実行
function analyzeSchedule() {
    if (surgeryData.length === 0) {
        alert('手術データが読み込まれていません');
        return;
    }
    
    const totalStaff = parseInt(residentInput.value || 0) +
                      parseInt(traineeInput.value || 0) +
                      parseInt(fulltimeInput.value || 0) +
                      parseInt(parttimeInput.value || 0) +
                      parseInt(nurseInput.value || 0);
    
    if (totalStaff === 0) {
        alert('人員の人数を入力してください');
        return;
    }
    
    // タイムライン生成
    timelineData = generateTimeline();
    
    // 最大必要キャパシティを計算
    const maxRequiredCapacity = Math.max(...timelineData.map(t => t.requiredCapacity), 0);
    const maxRequiredPeople = Math.max(...timelineData.map(t => t.requiredPeople), 0);
    
    // 結果表示
    displayResults(totalStaff, maxRequiredCapacity, maxRequiredPeople);
    drawChart(totalStaff);
    displayDetailTable(totalStaff);
    displayTimelineVisualization();
    displaySurgeryTable();
    
    // 結果セクションを表示
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

// 結果の表示（キャパシティベース）
function displayResults(totalStaff, maxRequiredCapacity, maxRequiredPeople) {
    // 供給可能なキャパシティを計算
    const staff = {
        resident: parseInt(residentInput.value || 0),
        trainee: parseInt(traineeInput.value || 0),
        fulltime: parseInt(fulltimeInput.value || 0),
        parttime: parseInt(parttimeInput.value || 0),
        nurse: parseInt(nurseInput.value || 0)
    };
    
    if (assignmentRules.hasSupervisor && !assignmentRules.allowSupervisorAnesthesia) {
        staff.fulltime = Math.max(0, staff.fulltime - 1);
    }
    
    // 供給可能なキャパシティを計算（上記と同じロジック）
    let suppliedCapacity = (staff.fulltime + staff.parttime) * 2; // 専門医のキャパシティ
    if (assignmentRules.allowTraineeSolo) {
        suppliedCapacity += staff.trainee * 1; // 専攻医の単独キャパシティ
    }
    
    document.getElementById('maxRequired').textContent = maxRequiredCapacity + '症例';
    document.getElementById('currentStaff').textContent = suppliedCapacity + '症例';
    
    const surplus = suppliedCapacity - maxRequiredCapacity;
    const surplusElement = document.getElementById('surplusDeficit');
    
    if (surplus >= 0) {
        surplusElement.textContent = `+${surplus}症例（余剰）`;
        surplusElement.parentElement.className = 'summary-card success';
    } else {
        surplusElement.textContent = `${surplus}症例（不足）`;
        surplusElement.parentElement.className = 'summary-card warning';
    }
    
    // 不足時間帯の計算
    const deficitPeriods = timelineData.filter(t => t.requiredCapacity > t.suppliedCapacity).length;
    document.getElementById('deficitPeriods').textContent = `${deficitPeriods}時間帯（15分×${deficitPeriods}）`;
}

// チャートの描画
function drawChart(totalStaff) {
    const ctx = document.getElementById('timelineChart').getContext('2d');
    
    // 既存のチャートを破棄
    if (chart) {
        chart.destroy();
    }
    
    const labels = timelineData.map(t => 
        t.time.toLocaleString('ja-JP', { 
            month: '2-digit', 
            day: '2-digit', 
            hour: '2-digit', 
            minute: '2-digit' 
        })
    );
    
    const requiredData = timelineData.map(t => t.requiredCapacity);
    const suppliedData = timelineData.map(t => t.suppliedCapacity);
    
    chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '必要キャパシティ（症例数）',
                    data: requiredData,
                    borderColor: 'rgb(255, 99, 132)',
                    backgroundColor: 'rgba(255, 99, 132, 0.1)',
                    tension: 0.1,
                    fill: true
                },
                {
                    label: '供給キャパシティ（症例数）',
                    data: suppliedData,
                    borderColor: 'rgb(54, 162, 235)',
                    backgroundColor: 'rgba(54, 162, 235, 0.1)',
                    borderDash: [5, 5],
                    fill: false
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    mode: 'index',
                    intersect: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    },
                    title: {
                        display: true,
                        text: '症例数（キャパシティ）'
                    }
                },
                x: {
                    ticks: {
                        maxRotation: 45,
                        minRotation: 45
                    }
                }
            }
        }
    });
}

// 詳細テーブルの表示（キャパシティベース）
function displayDetailTable(totalStaff) {
    const tbody = document.getElementById('detailTableBody');
    tbody.innerHTML = '';
    
    timelineData.forEach(slot => {
        const surplus = slot.suppliedCapacity - slot.requiredCapacity;
        const status = surplus >= 0 ? 'ok' : 'danger';
        const statusText = surplus >= 0 ? '充足' : '不足';
        
        // 必要な人員構成を計算
        const requiredComposition = calculateRequiredComposition(slot);
        
        const row = tbody.insertRow();
        row.innerHTML = `
            <td>${slot.time.toLocaleString('ja-JP', { 
                month: '2-digit', 
                day: '2-digit', 
                hour: '2-digit', 
                minute: '2-digit' 
            })}</td>
            <td>${slot.requiredCapacity}症例</td>
            <td>${slot.suppliedCapacity}症例</td>
            <td class="status-${status}">${surplus >= 0 ? '+' : ''}${surplus}症例</td>
            <td class="status-${status}">${statusText}</td>
            <td class="required-staff-composition">${requiredComposition}</td>
        `;
    });
}

// 必要な人員構成を計算
function calculateRequiredComposition(slot) {
    if (slot.requiredCapacity === 0) return '-';
    
    const required = slot.requiredCapacity;
    const assigned = slot.assignments.length;
    const unassigned = slot.unassignedSurgeries.length;
    
    if (unassigned > 0) {
        // 不足している場合、最小限の人員構成を提案
        let composition = [];
        let remaining = required;
        
        // 専門医でカバーできる分（専門医は最大2症例まで）
        const specialistCases = Math.min(remaining, (slot.staff.fulltime + slot.staff.parttime) * 2);
        if (specialistCases > 0) {
            const specialistsNeeded = Math.ceil(specialistCases / 2);
            composition.push(`専門医${specialistsNeeded}人`);
            remaining -= specialistCases;
        }
        
        // 残りを専攻医でカバー（設定で許可されている場合のみ）
        if (remaining > 0 && assignmentRules.allowTraineeSolo) {
            composition.push(`専攻医${remaining}人`);
        } else if (remaining > 0) {
            // 専攻医の単独が許可されていない場合、専門医が必要
            const additionalSpecialists = Math.ceil(remaining / 2);
            composition.push(`専門医${additionalSpecialists}人（追加）`);
        }
        
        return composition.join(' + ') + `（不足: ${unassigned}症例）`;
    } else {
        // 充足している場合、現在の構成を表示
        const staffCount = {
            resident: 0,
            trainee: 0,
            nurse: 0,
            fulltime: 0,
            parttime: 0
        };
        
        slot.assignments.forEach(assign => {
            if (assign.resident) staffCount.resident += assign.resident;
            if (assign.trainee) staffCount.trainee += assign.trainee;
            if (assign.nurse) staffCount.nurse += assign.nurse;
            if (assign.fulltime) staffCount.fulltime += assign.fulltime;
            if (assign.parttime) staffCount.parttime += assign.parttime;
        });
        
        const parts = [];
        if (staffCount.resident > 0) parts.push(`研修医${staffCount.resident}人`);
        if (staffCount.trainee > 0) parts.push(`専攻医${staffCount.trainee}人`);
        if (staffCount.nurse > 0) parts.push(`看護師${staffCount.nurse}人`);
        if (staffCount.fulltime > 0) parts.push(`常勤専門医${staffCount.fulltime}人`);
        if (staffCount.parttime > 0) parts.push(`非常勤専門医${staffCount.parttime}人`);
        
        return parts.join(' + ') || '-';
    }
}

// タイムライン可視化の表示
function displayTimelineVisualization() {
    const grid = document.getElementById('timelineGrid');
    grid.innerHTML = '';
    
    timelineData.forEach(slot => {
        const surplus = slot.suppliedCapacity - slot.requiredCapacity;
        const isAdequate = surplus >= 0;
        
        const slotDiv = document.createElement('div');
        slotDiv.className = `timeline-slot ${isAdequate ? 'adequate' : 'insufficient'}`;
        
        // 担当医の割り当て一覧を作成
        let assignmentsHtml = '';
        
        // 割り当て済みの手術
        slot.assignments.forEach(assign => {
            const staffBadges = [];
            if (assign.resident) staffBadges.push('<span class="staff-badge resident">研修医</span>');
            if (assign.trainee) staffBadges.push('<span class="staff-badge trainee">専攻医</span>');
            if (assign.nurse) staffBadges.push('<span class="staff-badge nurse">看護師</span>');
            if (assign.fulltime) staffBadges.push('<span class="staff-badge fulltime">常勤専門医</span>');
            if (assign.parttime) staffBadges.push('<span class="staff-badge parttime">非常勤専門医</span>');
            
            assignmentsHtml += `
                <div class="assignment-item">
                    <span class="assignment-surgery">${assign.surgery.surgeryType}</span>
                    <span class="assignment-staff">${staffBadges.join('')}</span>
                </div>
            `;
        });
        
        // 未割り当ての手術
        slot.unassignedSurgeries.forEach(surgery => {
            assignmentsHtml += `
                <div class="assignment-item unassigned">
                    <span class="assignment-surgery">${surgery.surgeryType}（未割り当て）</span>
                    <span class="assignment-staff"></span>
                </div>
            `;
        });
        
        if (slot.requiredCapacity === 0) {
            assignmentsHtml = '<div class="assignment-item"><span class="assignment-surgery">手術なし</span></div>';
        }
        
        slotDiv.innerHTML = `
            <div class="timeline-slot-header">
                <span class="timeline-slot-time">${slot.time.toLocaleString('ja-JP', { 
                    month: '2-digit', 
                    day: '2-digit', 
                    hour: '2-digit', 
                    minute: '2-digit' 
                })}</span>
                <span class="timeline-slot-status ${isAdequate ? 'adequate' : 'insufficient'}">
                    ${isAdequate ? '充足' : '不足'}
                </span>
            </div>
            <div class="timeline-slot-capacity">
                <div class="capacity-item">
                    <span class="capacity-label">必要:</span>
                    <span class="capacity-value required">${slot.requiredCapacity}症例</span>
                </div>
                <div class="capacity-item">
                    <span class="capacity-label">供給:</span>
                    <span class="capacity-value supplied">${slot.suppliedCapacity}症例</span>
                </div>
                <div class="capacity-item">
                    <span class="capacity-label">余剰/不足:</span>
                    <span class="capacity-value ${isAdequate ? 'supplied' : 'required'}">
                        ${surplus >= 0 ? '+' : ''}${surplus}症例
                    </span>
                </div>
            </div>
            <div class="timeline-slot-assignments">
                <div class="assignment-list">
                    ${assignmentsHtml}
                </div>
            </div>
        `;
        
        grid.appendChild(slotDiv);
    });
}

// 手術テーブルの表示
function displaySurgeryTable() {
    const tbody = document.getElementById('surgeryTableBody');
    tbody.innerHTML = '';
    
    surgeryData.forEach(surgery => {
        const duration = (surgery.end - surgery.start) / (1000 * 60); // 分
        const row = tbody.insertRow();
        row.innerHTML = `
            <td>${surgery.entry.toLocaleString('ja-JP', { 
                month: '2-digit', 
                day: '2-digit', 
                hour: '2-digit', 
                minute: '2-digit' 
            })}</td>
            <td>${surgery.start.toLocaleString('ja-JP', { 
                month: '2-digit', 
                day: '2-digit', 
                hour: '2-digit', 
                minute: '2-digit' 
            })}</td>
            <td>${surgery.end.toLocaleString('ja-JP', { 
                month: '2-digit', 
                day: '2-digit', 
                hour: '2-digit', 
                minute: '2-digit' 
            })}</td>
            <td>${surgery.exit.toLocaleString('ja-JP', { 
                month: '2-digit', 
                day: '2-digit', 
                hour: '2-digit', 
                minute: '2-digit' 
            })}</td>
            <td>${surgery.surgeryType}</td>
            <td>-</td>
        `;
    });
}

// 初期化
updateTotalDoctors();
