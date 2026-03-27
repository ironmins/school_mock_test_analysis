/**
 * 고등학교 모의고사 성적분석 V14
 * script.js — 영역별 선택과목 분석 반영
 */

Chart.register(ChartDataLabels);
Chart.defaults.plugins.datalabels.display = false;

const state = {
    gradeData: {},
    availableGrades: [],
    currentGradeTotal: null,
    currentGradeClass: null,
    currentGradeIndiv: null,
    metric: 'raw',
    classMetric: 'raw',
    classSort: 'no',
    charts: {},
    subjectCharts: [],
    uploadedFiles: []
};

Object.defineProperty(state, 'exams', {
    get() {
        const g = state.currentGradeIndiv || state.availableGrades[0];
        return state.gradeData[g] || [];
    }
});

/* ── 탐구 선택과목 분류 ── */
const socialSubjects = [
    '생활과 윤리', '윤리와 사상', '한국 지리', '세계 지리',
    '동아시아사', '세계사', '경제', '정치와 법', '사회·문화', '사회문화',
    '통합사회', '사회'
];
const scienceSubjects = [
    '물리학Ⅰ', '물리학Ⅱ', '물리학1', '물리학2',
    '화학Ⅰ', '화학Ⅱ', '화학1', '화학2',
    '생명 과학Ⅰ', '생명 과학Ⅱ', '생명과학Ⅰ', '생명과학Ⅱ',
    '생명 과학1', '생명 과학2', '생명과학1', '생명과학2',
    '지구 과학Ⅰ', '지구 과학Ⅱ', '지구과학Ⅰ', '지구과학Ⅱ',
    '지구 과학1', '지구 과학2', '지구과학1', '지구과학2',
    '통합과학', '과학'
];

function classifyInquiry(name) {
    if (!name) return 'unknown';
    const n = name.trim();
    if (socialSubjects.some(s => n.includes(s) || s.includes(n))) return 'social';
    if (scienceSubjects.some(s => n.includes(s) || s.includes(n))) return 'science';
    return 'unknown';
}

/* ── 영역 정의 (6개 영역) ── */
const areas = [
    { k: 'kor',  n: '국어 영역',  hasChoice: true  },
    { k: 'math', n: '수학 영역',  hasChoice: true  },
    { k: 'eng',  n: '영어 영역',  hasChoice: false },
    { k: 'hist', n: '한국사',     hasChoice: false },
    { k: 'inq1', n: '탐구영역1',  hasChoice: true  },
    { k: 'inq2', n: '탐구영역2',  hasChoice: true  },
];

const radarPointColors = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6'];

document.addEventListener('DOMContentLoaded', initializeEventListeners);

function initializeEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const tabBtns = document.querySelectorAll('.tab-btn');
    const uploadSection = document.querySelector('.upload-section');
    const fileLabel = document.querySelector('.file-input-label');

    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const newFiles = Array.from(e.target.files);
            if (newFiles.length > 0) {
                addFiles(newFiles);
                if (analyzeBtn) analyzeBtn.disabled = false;
            }
            // input 초기화 (같은 파일 재선택 가능하도록)
            fileInput.value = '';
        });
    }

    if (analyzeBtn) analyzeBtn.addEventListener('click', analyzeFiles);

    tabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            switchTab(e.target.closest('.tab-btn').dataset.tab);
        });
    });

    if (uploadSection) {
        const prevent = (ev) => { ev.preventDefault(); ev.stopPropagation(); };
        const setDragState = (on) => { if (fileLabel) fileLabel.classList.toggle('dragover', on); };
        ['dragover', 'drop'].forEach(evt => window.addEventListener(evt, prevent));
        ['dragenter', 'dragover'].forEach(evt => uploadSection.addEventListener(evt, (ev) => { prevent(ev); setDragState(true); }));
        ['dragleave', 'dragend'].forEach(evt => uploadSection.addEventListener(evt, (ev) => { prevent(ev); setDragState(false); }));
        uploadSection.addEventListener('drop', (ev) => {
            prevent(ev); setDragState(false);
            const dropped = Array.from(ev.dataTransfer?.files || []);
            const files = dropped.filter(f => /\.(xlsx|xls|csv|xlsm)$/i.test(f.name));
            if (files.length > 0) {
                addFiles(files);
                if (analyzeBtn) analyzeBtn.disabled = false;
            }
        });
    }

    document.getElementById('gradeSelectTotal')?.addEventListener('change', onGradeChangeTotal);
    document.getElementById('gradeSelectClass')?.addEventListener('change', onGradeChangeClass);
    document.getElementById('gradeSelectIndiv')?.addEventListener('change', onGradeChangeIndiv);
    document.getElementById('examSelectTotal')?.addEventListener('change', renderOverall);
    document.getElementById('examSelectClass')?.addEventListener('change', renderClass);
    document.getElementById('classSelect')?.addEventListener('change', renderClass);
    document.getElementById('indivClassSelect')?.addEventListener('change', updateIndivList);
    document.getElementById('indivStudentSelect')?.addEventListener('change', renderIndividual);
    // ★ 학생 검색 기능 이벤트 등록
    initStudentSearch();
    document.getElementById('indivExamSelect')?.addEventListener('change', renderIndividual);
    document.getElementById('pdfStudentBtn')?.addEventListener('click', generateStudentPDF);
    document.getElementById('pdfClassBtn')?.addEventListener('click', generateClassPDF);

    let resizeTimer = null;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            state.subjectCharts.forEach(chart => {
                if (chart && !chart.destroyed && typeof chart.resize === 'function') chart.resize();
            });
        }, 150);
    });

    const saveHtmlBtn = document.getElementById('saveHtmlBtn');
    if (saveHtmlBtn) {
        saveHtmlBtn.replaceWith(saveHtmlBtn.cloneNode(true));
        document.getElementById('saveHtmlBtn').addEventListener('click', saveHtmlFile);
    }
}

/* ── 헬퍼 함수들 ── */
function getExamsForGrade(grade) { return state.gradeData[grade] || []; }

function onGradeChangeTotal() {
    const grade = parseInt(document.getElementById('gradeSelectTotal').value);
    state.currentGradeTotal = grade;
    updateExamSelector('examSelectTotal', grade);
    renderOverall();
}
function onGradeChangeClass() {
    const grade = parseInt(document.getElementById('gradeSelectClass').value);
    state.currentGradeClass = grade;
    updateExamSelector('examSelectClass', grade);
    updateClassSelector('classSelect', grade);
    renderClass();
}
function onGradeChangeIndiv() {
    const grade = parseInt(document.getElementById('gradeSelectIndiv').value);
    state.currentGradeIndiv = grade;
    updateExamSelector('indivExamSelect', grade);
    updateClassSelector('indivClassSelect', grade);
    updateIndivList();
}
function updateExamSelector(selectId, grade) {
    const el = document.getElementById(selectId);
    if (!el) return;
    const exams = getExamsForGrade(grade);
    el.innerHTML = exams.map((e, i) => `<option value="${i}">${e.name}</option>`).join('');
}
function updateClassSelector(selectId, grade) {
    const el = document.getElementById(selectId);
    if (!el) return;
    const exams = getExamsForGrade(grade);
    if (!exams.length) { el.innerHTML = ''; return; }
    const classSet = new Set();
    exams.forEach(exam => exam.students.forEach(s => classSet.add(s.info.class)));
    const classes = [...classSet].sort((a, b) => a - b);
    el.innerHTML = classes.map(c => `<option value="${c}">${c}반</option>`).join('');
    el.value = classes[0];
}
function updateGradeSelectors() {
    const grades = state.availableGrades;
    const opts = grades.map(g => `<option value="${g}">${g}학년</option>`).join('');
    ['gradeSelectTotal', 'gradeSelectClass', 'gradeSelectIndiv'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = opts;
    });
    state.currentGradeTotal = grades[0];
    state.currentGradeClass = grades[0];
    state.currentGradeIndiv = grades[0];
}

/* ── 파일 누적 추가 (중복 방지) ── */
function addFiles(newFiles) {
    newFiles.forEach(file => {
        // 같은 이름 + 같은 크기의 파일은 중복으로 판단하여 스킵
        const isDuplicate = state.uploadedFiles.some(
            f => f.name === file.name && f.size === file.size
        );
        if (!isDuplicate) {
            state.uploadedFiles.push(file);
        }
    });
    renderFileList();
}

/* ── 개별 파일 삭제 ── */
window.removeFile = function(index) {
    state.uploadedFiles.splice(index, 1);
    renderFileList();
    const analyzeBtn = document.getElementById('analyzeBtn');
    if (analyzeBtn) {
        analyzeBtn.disabled = state.uploadedFiles.length === 0;
    }
};

/* ── 전체 파일 삭제 ── */
window.clearAllFiles = function() {
    state.uploadedFiles = [];
    renderFileList();
    const analyzeBtn = document.getElementById('analyzeBtn');
    if (analyzeBtn) analyzeBtn.disabled = true;
};

/* ── 파일 목록 렌더링 ── */
function renderFileList() {
    const fileList = document.getElementById('fileList');
    if (!fileList) return;

    if (state.uploadedFiles.length === 0) {
        fileList.style.display = 'none';
        fileList.innerHTML = '';
        return;
    }

    fileList.style.display = 'block';
    fileList.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
            <h4 style="margin:0;"><i class="fas fa-file-alt"></i> 업로드된 파일 (${state.uploadedFiles.length}개)</h4>
            <button onclick="clearAllFiles()" 
                style="background:#e74c3c; color:#fff; border:none; padding:4px 12px; 
                       border-radius:4px; cursor:pointer; font-size:0.8rem;">
                <i class="fas fa-trash-alt"></i> 전체 삭제
            </button>
        </div>
        <ul style="list-style:none; padding:0; margin:0;">
            ${state.uploadedFiles.map((file, idx) => `
                <li style="display:flex; justify-content:space-between; align-items:center; 
                           padding:6px 10px; margin-bottom:4px; background:rgba(0,0,0,0.03); 
                           border-radius:6px; font-size:0.88rem;">
                    <span>
                        <i class="fas fa-file-excel" style="color:#27ae60; margin-right:6px;"></i>
                        ${file.name}
                        <span style="color:#999; font-size:0.75rem; margin-left:8px;">
                            (${(file.size / 1024).toFixed(1)} KB)
                        </span>
                    </span>
                    <button onclick="removeFile(${idx})" 
                        style="background:none; border:none; color:#e74c3c; cursor:pointer; 
                               font-size:1rem; padding:2px 6px;" title="삭제">
                        <i class="fas fa-times-circle"></i>
                    </button>
                </li>
            `).join('')}
        </ul>
    `;
}

function showLoading(text = '분석 중...') {
    const loadingText = document.getElementById('loadingText');
    if (loadingText) loadingText.textContent = text;
    document.getElementById('loading').style.display = 'flex';
}
function hideLoading() { document.getElementById('loading').style.display = 'none'; }

/* ── 파일 분석 ── */
async function analyzeFiles() {
    const files = state.uploadedFiles;
    if (files.length === 0) return alert('파일을 선택해주세요.');

    showLoading();
    try {
        const promises = files.map(file => new Promise(resolve => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const wb = XLSX.read(evt.target.result, { type: 'array' });
                    let targetSheetName = wb.SheetNames.find(name => {
                        const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
                        for (let i = 0; i < Math.min(20, json.length); i++) {
                            const rowStr = (json[i] || []).join(' ');
                            if (rowStr.includes('이름') && (rowStr.includes('국어') || rowStr.includes('수학'))) return true;
                        }
                        return false;
                    }) || wb.SheetNames[0];
                    const jsonData = XLSX.utils.sheet_to_json(wb.Sheets[targetSheetName], { header: 1 });
                    resolve(parseExcel(jsonData, file.name));
                } catch (err) { console.error(err); resolve(null); }
            };
            reader.readAsArrayBuffer(file);
        }));

        const results = await Promise.all(promises);
        const validResults = results.filter(r => r && r.students.length > 0);
        if (!validResults.length) { hideLoading(); return alert('데이터를 찾을 수 없습니다.'); }

        state.gradeData = {};
        validResults.forEach(exam => {
            const gradeSet = new Set(exam.students.map(s => s.info.grade));
            gradeSet.forEach(grade => {
                if (!state.gradeData[grade]) state.gradeData[grade] = [];
                const gradeStudents = exam.students.filter(s => s.info.grade === grade);
                gradeStudents.sort((a, b) => b.totalRaw - a.totalRaw);
                gradeStudents.forEach((s, idx) => { s.totalRank = idx + 1; });
                state.gradeData[grade].push({ name: exam.name, students: gradeStudents });
            });
        });

        Object.keys(state.gradeData).forEach(grade => {
            state.gradeData[grade].sort((a, b) => b.name.localeCompare(a.name, undefined, { numeric: true }));
        });

        state.availableGrades = Object.keys(state.gradeData).map(Number).sort((a, b) => a - b);

        if (state.availableGrades.length) {
            hideLoading();
            document.getElementById('results').style.display = 'block';
            document.getElementById('saveHtmlBtn').style.display = 'inline-flex';
            updateLastUpdated();
            initSelectors();
        } else { hideLoading(); alert('데이터를 찾을 수 없습니다.'); }
    } catch (error) { hideLoading(); alert('파일 분석 중 오류가 발생했습니다: ' + error.message); }
}

/* ── 엑셀 파싱 (선택과목명 포함) ── */
function parseExcel(rows, fname) {
    let startRow = -1;
    for (let i = 0; i < rows.length; i++) {
        if (!rows[i]) continue;
        const rowStr = rows[i].map(c => String(c).replace(/\s/g, '')).join(',');
        if (rowStr.includes('이름') && rowStr.includes('번호')) { startRow = i; break; }
    }
    if (startRow === -1) return null;

    const students = [];
    const val = (r, i) => Number(r[i]) || 0;
    const grd = (r, i) => { const v = Number(r[i]); return (v > 0 && v < 10) ? v : 9; };
    const str = (r, i) => (r[i] && String(r[i]).trim() !== '' && String(r[i]).trim() !== 'nan') ? String(r[i]).trim() : '';

    for (let i = startRow + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || !r[4]) continue;

        const s = {
            info: { grade: parseInt(r[1]), class: parseInt(r[2]), no: parseInt(r[3]), name: r[4] },
            hist: { raw: val(r, 5), grd: grd(r, 6), std: 0, pct: 0, name: '' },
            kor:  { name: str(r, 7), raw: val(r, 8), std: val(r, 9), pct: val(r, 10), grd: grd(r, 11) },
            math: { name: str(r, 12), raw: val(r, 13), std: val(r, 14), pct: val(r, 15), grd: grd(r, 16) },
            eng:  { raw: val(r, 17), grd: grd(r, 18), std: 0, pct: 0, name: '' },
            inq1: { name: str(r, 19), raw: val(r, 20), std: val(r, 21), pct: val(r, 22), grd: grd(r, 23) },
            inq2: { name: str(r, 24), raw: val(r, 25), std: val(r, 26), pct: val(r, 27), grd: grd(r, 28) },
            uid: `${parseInt(r[2])}-${parseInt(r[3])}-${r[4]}`
        };

        s.totalRaw = s.kor.raw + s.math.raw + s.eng.raw + s.inq1.raw + s.inq2.raw + s.hist.raw;
        s.totalStd = s.kor.std + s.math.std + s.inq1.std + s.inq2.std;
        s.totalPct = parseFloat((s.kor.pct + s.math.pct + s.inq1.pct + s.inq2.pct).toFixed(2));
        students.push(s);
    }
    students.sort((a, b) => b.totalRaw - a.totalRaw);
    students.forEach((s, idx) => { s.totalRank = idx + 1; });
    return { name: fname.replace(/\.[^/.]+$/, ""), students };
}

function initSelectors() {
    if (!state.availableGrades.length) return;
    updateGradeSelectors();
    const defaultGrade = state.availableGrades[0];
    state.currentGradeTotal = defaultGrade;
    document.getElementById('gradeSelectTotal').value = defaultGrade;
    updateExamSelector('examSelectTotal', defaultGrade);
    state.currentGradeClass = defaultGrade;
    document.getElementById('gradeSelectClass').value = defaultGrade;
    updateExamSelector('examSelectClass', defaultGrade);
    updateClassSelector('classSelect', defaultGrade);
    state.currentGradeIndiv = defaultGrade;
    document.getElementById('gradeSelectIndiv').value = defaultGrade;
    updateExamSelector('indivExamSelect', defaultGrade);
    updateClassSelector('indivClassSelect', defaultGrade);
    switchTab('overall');
    renderOverall();
    renderClass();
    updateIndivList();
}

function switchTab(t) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById(t + '-tab').classList.add('active');
    document.querySelector(`.tab-btn[data-tab="${t}"]`).classList.add('active');
    if (t === 'overall' && state.charts.bubble) state.charts.bubble.resize();
}

/* ── 성적일람표에서 개인통계 탭으로 바로 이동 ── */
window.goToIndividual = function(uid, grade) {
    // 학년 동기화
    state.currentGradeIndiv = grade;
    const gradeSelectIndiv = document.getElementById('gradeSelectIndiv');
    if (gradeSelectIndiv) gradeSelectIndiv.value = grade;
    updateExamSelector('indivExamSelect', grade);

    // 반 추출 (uid 형식: "반-번호-이름")
    const cls = parseInt(uid.split('-')[0]);
    const indivClassSelect = document.getElementById('indivClassSelect');
    if (indivClassSelect) {
        indivClassSelect.value = cls;
    }

    // 학생 목록 갱신 후 해당 학생 선택
    updateIndivList();
    setTimeout(() => {
        const studentSelect = document.getElementById('indivStudentSelect');
        if (studentSelect) {
            studentSelect.value = uid;
            renderIndividual();
        }
        switchTab('individual');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }, 50);
};

/* ── 유틸: 학생 배열에서 영역별 선택과목 그룹 추출 ── */
function getChoiceGroups(students, areaKey) {
    // 해당 영역에서 어떤 선택과목들이 있는지 + 학생 분류
    const groups = {};
    students.forEach(s => {
        const subj = s[areaKey];
        const choiceName = subj.name || '(미분류)';
        if (!groups[choiceName]) groups[choiceName] = [];
        groups[choiceName].push(s);
    });
    return groups;
}

/* ── 과목 카드 HTML (학생 배열 기반) ── */
function buildSubjectCardHTML(title, studentsInGroup, areaKey, metric) {
    const isAbs = (areaKey === 'eng' || areaKey === 'hist');
    const scores = studentsInGroup.map(s => isAbs ? s[areaKey].raw : (s[areaKey][metric] || 0)).filter(v => v > 0);
    const avg = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : '-';
    const std = scores.length ? Math.sqrt(scores.reduce((a, b) => a + Math.pow(b - parseFloat(avg), 2), 0) / scores.length).toFixed(1) : '-';
    const maxScore = scores.length ? Math.max(...scores).toFixed(1) : '-';
    const counts = Array(9).fill(0);
    studentsInGroup.forEach(s => { if (s[areaKey].grd >= 1 && s[areaKey].grd <= 9) counts[s[areaKey].grd - 1]++; });
    return buildCardInnerHTML(title, scores.length, avg, std, maxScore, counts);
}

/* ── 과목 카드 HTML (entry 배열 기반 — 탐구 통합용) ── */
function buildEntryCardHTML(title, entries, metric) {
    const scores = entries.map(e => {
        if (metric === 'raw') return e.raw;
        if (metric === 'std') return e.std;
        return e.pct;
    }).filter(v => v > 0);
    const avg = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : '-';
    const std = scores.length ? Math.sqrt(scores.reduce((a, b) => a + Math.pow(b - parseFloat(avg), 2), 0) / scores.length).toFixed(1) : '-';
    const maxScore = scores.length ? Math.max(...scores).toFixed(1) : '-';
    const counts = Array(9).fill(0);
    entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
    return buildCardInnerHTML(title, entries.length, avg, std, maxScore, counts);
}

/* ── 카드 공통 내부 HTML ── */
function buildCardInnerHTML(title, count, avg, std, maxScore, counts) {
    return `
        <div class="subject-card">
            <div class="subject-card-header">
                <h4>${title}</h4>
                <span class="count-badge">응시 ${count}명</span>
            </div>
            <div class="subject-stats">
                <div class="stat-item"><div class="stat-label">평균</div><div class="stat-value-large">${avg}</div></div>
                <div class="stat-item"><div class="stat-label">표준편차</div><div class="stat-value-large">${std}</div></div>
                <div class="stat-item"><div class="stat-label">최고점</div><div class="stat-value-large">${maxScore}</div></div>
            </div>
            <div class="grade-distribution">
                ${counts.map((c, i) => {
                    const total = counts.reduce((a, b) => a + b, 0);
                    const pct = total ? ((c / total) * 100).toFixed(1) : 0;
                    return `<div class="grade-bar-item">
                        <div class="grade-label">${i + 1}등급</div>
                        <div class="grade-bar-container"><div class="grade-bar-fill g-${i + 1}" style="width: ${pct}%;"></div></div>
                        <div class="grade-count">${c}명</div>
                        <div class="grade-percentage">${pct}%</div>
                    </div>`;
                }).join('')}
            </div>
        </div>`;
}

/* ── 탐구영역 통합 그룹 추출 ── */
function getInquiryMergedGroups(students) {
    const allEntries = [];
    students.forEach(s => {
        if (s.inq1.name) allEntries.push({ ...s.inq1, student: s, source: 'inq1' });
        if (s.inq2.name) allEntries.push({ ...s.inq2, student: s, source: 'inq2' });
    });

    const socialGroup = {};
    const scienceGroup = {};
    const unknownGroup = {};

    allEntries.forEach(entry => {
        const cat = classifyInquiry(entry.name);
        const target = cat === 'social' ? socialGroup : (cat === 'science' ? scienceGroup : unknownGroup);
        if (!target[entry.name]) target[entry.name] = [];
        target[entry.name].push(entry);
    });

    return { socialGroup, scienceGroup, unknownGroup, allEntries };
}

/* ── 탐구 영역 통합 카드 HTML 생성 ── */
function buildInquiryAreaHTML(students, metric) {
    const { socialGroup, scienceGroup, unknownGroup } = getInquiryMergedGroups(students);
    let html = '';

    const socialNames = Object.keys(socialGroup).sort();
    const scienceNames = Object.keys(scienceGroup).sort();
    const unknownNames = Object.keys(unknownGroup).sort();

    if (socialNames.length > 0) {
        const allSocialEntries = socialNames.flatMap(n => socialGroup[n]);
        html += `<div class="area-section">
            <h4 class="area-title"><i class="fas fa-layer-group"></i> 사회탐구 영역</h4>
            <div class="subject-grid subject-grid-fixed">`;
        if (socialNames.length > 1) {
            html += buildEntryCardHTML('사회탐구 전체', allSocialEntries, metric);
        }
        socialNames.forEach(name => {
            html += buildEntryCardHTML(name, socialGroup[name], metric);
        });
        html += `</div></div>`;
    }

    if (scienceNames.length > 0) {
        const allScienceEntries = scienceNames.flatMap(n => scienceGroup[n]);
        html += `<div class="area-section">
            <h4 class="area-title"><i class="fas fa-layer-group"></i> 과학탐구 영역</h4>
            <div class="subject-grid subject-grid-fixed">`;
        if (scienceNames.length > 1) {
            html += buildEntryCardHTML('과학탐구 전체', allScienceEntries, metric);
        }
        scienceNames.forEach(name => {
            html += buildEntryCardHTML(name, scienceGroup[name], metric);
        });
        html += `</div></div>`;
    }

    if (unknownNames.length > 0) {
        html += `<div class="area-section">
            <h4 class="area-title"><i class="fas fa-layer-group"></i> 탐구 영역 (기타)</h4>
            <div class="subject-grid subject-grid-fixed">`;
        unknownNames.forEach(name => {
            html += buildEntryCardHTML(name, unknownGroup[name], metric);
        });
        html += `</div></div>`;
    }

    return html;
}

/* ── 모든 영역의 선택과목이 1개씩인지 판별 ── */
function checkSimpleMode(students) {
    // 국어, 수학: 선택과목 종류가 1개 이하
    for (const key of ['kor', 'math']) {
        const names = new Set();
        students.forEach(s => {
            const n = s[key].name;
            if (n && n.trim() !== '' && n !== '(미분류)') names.add(n.trim());
        });
        if (names.size > 1) return false;
    }

    // 탐구: 사회탐구 과목이 2개 이상이거나 과학탐구 과목이 2개 이상이면 복합모드
    const { socialGroup, scienceGroup, unknownGroup } = getInquiryMergedGroups(students);
    const socialNames = Object.keys(socialGroup);
    const scienceNames = Object.keys(scienceGroup);
    const unknownNames = Object.keys(unknownGroup);

    if (socialNames.length > 1 || scienceNames.length > 1 || unknownNames.length > 0) return false;

    return true;
}

/* ── 단일 영역 카드 생성 (선택과목 1개면 간소화) ── */
function buildSingleAreaHTML(students, areaKey, areaName, hasChoice, metric) {
    if (!hasChoice) {
        return `<div class="subject-grid subject-grid-fixed">${buildSubjectCardHTML(areaName, students, areaKey, metric)}</div>`;
    }

    const groups = getChoiceGroups(students, areaKey);
    const sortedNames = Object.keys(groups).filter(n => n !== '(미분류)' || groups[n].length > 0).sort();
    const realNames = sortedNames.filter(n => n !== '(미분류)');

    if (realNames.length <= 1) {
        return `<div class="subject-grid subject-grid-fixed">${buildSubjectCardHTML(realNames[0] || areaName, students, areaKey, metric)}</div>`;
    }

    let html = `<div class="subject-grid subject-grid-fixed">`;
    html += buildSubjectCardHTML(`${areaName} 전체`, students, areaKey, metric);
    sortedNames.forEach(choiceName => {
        html += buildSubjectCardHTML(choiceName, groups[choiceName], areaKey, metric);
    });
    html += `</div>`;
    return html;
}

/* ==============================
   전체통계 탭
   ============================== */
window.changeMetric = function (m) {
    state.metric = m;
    document.querySelectorAll('#overall-tab .opt-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + m).classList.add('active');
    renderOverall();
};

function renderOverall() {
    const grade = state.currentGradeTotal || state.availableGrades[0];
    const exams = getExamsForGrade(grade);
    const examSelect = document.getElementById('examSelectTotal');
    if (!examSelect || !exams.length) return;
    const examIdx = parseInt(examSelect.value) || 0;
    if (!exams[examIdx]) return;
    const students = exams[examIdx].students;
    const metric = state.metric;

    /* ── 버블 차트 (성적 분포) ── */
    const classes = [...new Set(students.map(s => s.info.class))].sort((a, b) => a - b);
    const maxClass = Math.max(...classes) || 12;
    const bubbleData = [];
    classes.forEach(c => {
        const cls = students.filter(s => s.info.class == c).sort((a, b) => b.totalRaw - a.totalRaw);
        cls.forEach((s, idx) => {
            const ratio = idx / (cls.length - 1 || 1);
            const r = ratio < 0.5 ? Math.floor(255 * (ratio * 2)) : 255;
            const g = ratio < 0.5 ? 255 : Math.floor(255 * (2 - ratio * 2));
            let score = metric === 'raw' ? s.totalRaw : (metric === 'std' ? s.totalStd : s.totalPct);
            bubbleData.push({ x: Number(c), y: score, r: 8, bg: `rgba(${r}, ${g}, 0, 0.8)`, name: s.info.name });
        });
    });
    const classScoreAvgData = classes.map(c => {
        const cls = students.filter(s => s.info.class == c);
        const avg = cls.reduce((sum, s) => sum + (metric === 'raw' ? s.totalRaw : (metric === 'std' ? s.totalStd : s.totalPct)), 0) / cls.length;
        return { x: Number(c), y: parseFloat(avg.toFixed(1)), r: 12, name: `${c}반 평균` };
    });

    if (state.charts.bubble) state.charts.bubble.destroy();
    state.charts.bubble = new Chart(document.getElementById('bubbleChart'), {
        type: 'bubble',
        data: {
            datasets: [
                { label: '학생', data: bubbleData, backgroundColor: bubbleData.map(d => d.bg), borderColor: 'transparent' },
                { label: '반 평균', data: classScoreAvgData, backgroundColor: 'rgba(80, 80, 220, 0.85)', borderColor: 'rgba(50, 50, 180, 1)', borderWidth: 1.5 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { min: 0, max: maxClass + 1, ticks: { stepSize: 1, callback: v => (Number.isInteger(v) && v > 0 && v <= maxClass) ? v + "반" : "" } },
                y: { title: { display: true, text: metric === 'raw' ? '원점수 합' : (metric === 'std' ? '표준점수 합' : '백분위 합') } }
            },
            plugins: {
                legend: { display: true, position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', font: { size: 11 } } },
                datalabels: { display: false },
                tooltip: { callbacks: { label: c => c.datasetIndex === 1 ? `${c.raw.name}: ${c.raw.y}` : `${c.raw.x}반 ${c.raw.name}: ${c.raw.y.toFixed(1)}` } }
            }
        }
    });

    /* ── 버블 차트 (평균등급) ── */
    const getAvgGradeVal = s => (s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4;
    const gradeBubbleData = [];
    classes.forEach(c => {
        const cls = students.filter(s => s.info.class == c).sort((a, b) => getAvgGradeVal(a) - getAvgGradeVal(b));
        cls.forEach((s, idx) => {
            const avgGrd = getAvgGradeVal(s);
            const ratio = idx / (cls.length - 1 || 1);
            const r = ratio < 0.5 ? Math.floor(255 * (ratio * 2)) : 255;
            const g = ratio < 0.5 ? 255 : Math.floor(255 * (2 - ratio * 2));
            gradeBubbleData.push({ x: Number(c), y: parseFloat(avgGrd.toFixed(2)), r: 8, bg: `rgba(${r}, ${g}, 0, 0.8)`, name: s.info.name });
        });
    });
    const classGradeAvgData = classes.map(c => {
        const cls = students.filter(s => s.info.class == c);
        const avg = cls.reduce((sum, s) => sum + getAvgGradeVal(s), 0) / cls.length;
        return { x: Number(c), y: parseFloat(avg.toFixed(2)), r: 12, name: `${c}반 평균` };
    });

    if (state.charts.gradesBubble) state.charts.gradesBubble.destroy();
    state.charts.gradesBubble = new Chart(document.getElementById('gradesBubbleChart'), {
        type: 'bubble',
        data: {
            datasets: [
                { label: '학생', data: gradeBubbleData, backgroundColor: gradeBubbleData.map(d => d.bg), borderColor: 'transparent' },
                { label: '반 평균', data: classGradeAvgData, backgroundColor: 'rgba(80, 80, 220, 0.85)', borderColor: 'rgba(50, 50, 180, 1)', borderWidth: 1.5 }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { min: 0, max: maxClass + 1, ticks: { stepSize: 1, callback: v => (Number.isInteger(v) && v > 0 && v <= maxClass) ? v + "반" : "" } },
                y: { reverse: true, min: 0, max: 9.5, ticks: { stepSize: 1, callback: v => (Number.isInteger(v) && v >= 1 && v <= 9) ? v + "등급" : "" }, title: { display: true, text: '평균등급' } }
            },
            plugins: {
                legend: { display: true, position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', font: { size: 11 } } },
                datalabels: { display: false },
                tooltip: { callbacks: { label: c => c.datasetIndex === 1 ? `${c.raw.name}: ${c.raw.y}등급` : `${c.raw.x}반 ${c.raw.name}: ${c.raw.y}등급` } }
            }
        }
    });

    /* ── 영역별 선택과목 종합 분석 카드 ── */
    const container = document.getElementById('combinedStatsContainer');
    container.innerHTML = '';

    // 선택과목이 모두 단일인지 판별
    const isSimpleMode = checkSimpleMode(students);

    if (isSimpleMode) {
        // V13 스타일: area-section 없이 6개 카드를 3\times2 그리드로 배치
        const simpleSubjects = [
            { k: 'kor',  n: '국어' },
            { k: 'math', n: '수학' },
            { k: 'eng',  n: '영어' },
            { k: 'hist', n: '한국사' },
            { k: 'inq1', n: '사회탐구' },
            { k: 'inq2', n: '과학탐구' }
        ];
        container.innerHTML = `<div class="subject-grid subject-grid-fixed">`;
        simpleSubjects.forEach(sub => {
            container.innerHTML += buildSubjectCardHTML(sub.n, students, sub.k, metric);
        });
        // innerHTML 연결 방식으로는 닫는 div가 안 들어가므로 별도 처리
        const wrapper = document.createElement('div');
        wrapper.className = 'subject-grid subject-grid-fixed';
        simpleSubjects.forEach(sub => {
            wrapper.innerHTML += buildSubjectCardHTML(sub.n, students, sub.k, metric);
        });
        container.innerHTML = '';
        container.appendChild(wrapper);
    } else {
        // 선택과목 다수: 영역별 area-section 분리
        ['kor', 'math', 'eng', 'hist'].forEach(key => {
            const area = areas.find(a => a.k === key);
            let areaHTML = `<div class="area-section">
                <h4 class="area-title"><i class="fas fa-layer-group"></i> ${area.n}</h4>`;
            areaHTML += buildSingleAreaHTML(students, area.k, area.n, area.hasChoice, metric);
            areaHTML += `</div>`;
            container.innerHTML += areaHTML;
        });

        // 탐구영역 통합 (사회탐구 / 과학탐구)
        container.innerHTML += buildInquiryAreaHTML(students, metric);
    }

    /* ── 전체 성적 일람표 (동적 헤더) ── */
    buildTotalTable(students, metric);

    /* ── 등급 필터 UI ── */
    initGradeFilter(students);
}

/* ── 전체 일람표 동적 헤더 ── */
function buildTotalTable(students, metric) {
    const thead = document.getElementById('totalTableHead');
    const tbody = document.getElementById('totalTableBody');

    // 영역별 선택과목 목록 수집
    const areaChoices = {};
    areas.forEach(area => {
        if (area.hasChoice) {
            const names = new Set();
            students.forEach(s => {
                const n = s[area.k].name || '(미분류)';
                names.add(n);
            });
            areaChoices[area.k] = [...names].sort();
        }
    });

    // 헤더 1행: 석차, 학번, 이름, 각 영역 (colspan 결정)
    let h1 = `<tr><th rowspan="2">석차</th><th rowspan="2">학번</th><th rowspan="2">이름</th>`;
    let h2 = `<tr>`;

    areas.forEach(area => {
        const isAbs = (area.k === 'eng' || area.k === 'hist');
        if (area.hasChoice) {
            // 선택과목 컬럼: 선택과목명, 점수, 등급 = 3열
            h1 += `<th colspan="3">${area.n}</th>`;
            h2 += `<th>선택과목</th><th>점수</th><th>등급</th>`;
        } else {
            h1 += `<th colspan="2">${area.n}</th>`;
            h2 += `<th>점수</th><th>등급</th>`;
        }
    });

    h1 += `<th rowspan="2" class="total-col">총점</th><th rowspan="2" class="total-col">평균등급</th></tr>`;
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    // 본문
    const getTot = s => metric === 'raw' ? s.totalRaw : (metric === 'std' ? s.totalStd.toFixed(1) : s.totalPct.toFixed(1));
    const getAvgGrade = s => ((s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4).toFixed(2);

    tbody.innerHTML = '';
    const grade = state.currentGradeTotal || state.availableGrades[0];
    students.slice(0, 500).forEach(s => {
        let row = `<tr>`;
        row += `<td style="font-weight:bold;color:var(--primary);">${s.totalRank}</td>`;
        row += `<td class="student-link-cell" onclick="goToIndividual('${s.uid}', ${grade})" title="개인통계 보기">${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}</td>`;
        row += `<td class="student-link-cell" onclick="goToIndividual('${s.uid}', ${grade})" title="개인통계 보기" style="font-weight:bold;">${s.info.name}</td>`;

        areas.forEach(area => {
            const isAbs = (area.k === 'eng' || area.k === 'hist');
            const subj = s[area.k];
            if (area.hasChoice) {
                const scoreVal = isAbs ? subj.raw : (subj[metric] || '-');
                row += `<td style="font-size:0.75rem;color:var(--text-secondary);">${subj.name || '-'}</td>`;
                row += `<td>${scoreVal}</td>`;
                row += `<td class="g-${subj.grd}">${subj.grd}</td>`;
            } else {
                row += `<td>${subj.raw}</td>`;
                row += `<td class="g-${subj.grd}">${subj.grd}</td>`;
            }
        });

        row += `<td class="total-col">${getTot(s)}</td>`;
        row += `<td class="total-col" style="font-weight:bold;color:var(--primary);">${getAvgGrade(s)}</td>`;
        row += `</tr>`;
        tbody.innerHTML += row;
    });
}

/* ── 등급 필터 (영역별 선택과목별) ── */
function initGradeFilter(students) {
    const group = document.getElementById('gradeFilterGroup');
    if (!group) return;

    const isSimple = checkSimpleMode(students);

    if (isSimple) {
        // V13 스타일: 단순 6과목 행
        const filterSubjects = [
            { k: 'kor',  n: '국어' },
            { k: 'math', n: '수학' },
            { k: 'eng',  n: '영어' },
            { k: 'hist', n: '한국사' },
            { k: 'inq1', n: '사회탐구' },
            { k: 'inq2', n: '과학탐구' }
        ];

        const gradeCounts = {};
        filterSubjects.forEach(sub => {
            gradeCounts[sub.k] = Array(9).fill(0);
            students.forEach(s => {
                const g = s[sub.k].grd;
                if (g >= 1 && g <= 9) gradeCounts[sub.k][g - 1]++;
            });
        });

        group.innerHTML = filterSubjects.map(sub => `
            <div class="grade-filter-subject">
                <div class="grade-filter-subject-label">${sub.n}</div>
                <div class="grade-filter-btns">
                    ${Array.from({length: 9}, (_, i) => i + 1).map(g => `
                        <button class="grade-filter-btn g-btn-${sub.k}-${g}"
                            onclick="renderGradeFilter('${sub.k}', ${g}, 'all')"
                            title="${g}등급: ${gradeCounts[sub.k][g-1]}명">
                            ${g}등급
                            <span class="grade-filter-count">${gradeCounts[sub.k][g-1]}</span>
                        </button>
                    `).join('')}
                </div>
            </div>
        `).join('');
    } else {
        // 복합모드: 영역별 헤더 + 선택과목별 행
        let html = '';

        // 국어, 수학, 영어, 한국사
        ['kor', 'math', 'eng', 'hist'].forEach(key => {
            const area = areas.find(a => a.k === key);
            html += `<div class="grade-filter-area-header">${area.n}</div>`;

            if (area.hasChoice) {
                const groups = getChoiceGroups(students, key);
                const sortedNames = Object.keys(groups).sort();
                const realNames = sortedNames.filter(n => n !== '(미분류)');

                if (realNames.length <= 1) {
                    const counts = Array(9).fill(0);
                    students.forEach(s => { const g = s[key].grd; if (g >= 1 && g <= 9) counts[g - 1]++; });
                    html += buildFilterSubjectRow(realNames[0] || area.n, key, 'all', counts);
                } else {
                    const allCounts = Array(9).fill(0);
                    students.forEach(s => { const g = s[key].grd; if (g >= 1 && g <= 9) allCounts[g - 1]++; });
                    html += buildFilterSubjectRow(`${area.n} 전체`, key, 'all', allCounts);

                    sortedNames.forEach(choiceName => {
                        const grp = groups[choiceName];
                        const counts = Array(9).fill(0);
                        grp.forEach(s => { const g = s[key].grd; if (g >= 1 && g <= 9) counts[g - 1]++; });
                        html += buildFilterSubjectRow(`└ ${choiceName}`, key, choiceName, counts, true);
                    });
                }
            } else {
                const counts = Array(9).fill(0);
                students.forEach(s => { const g = s[key].grd; if (g >= 1 && g <= 9) counts[g - 1]++; });
                html += buildFilterSubjectRow(area.n, key, 'all', counts);
            }
        });

        // 탐구영역 통합 필터
        const { socialGroup, scienceGroup, unknownGroup } = getInquiryMergedGroups(students);

        const socialNames = Object.keys(socialGroup).sort();
        if (socialNames.length > 0) {
            html += `<div class="grade-filter-area-header">사회탐구 영역</div>`;
            if (socialNames.length > 1) {
                const allEntries = socialNames.flatMap(n => socialGroup[n]);
                const allCounts = Array(9).fill(0);
                allEntries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) allCounts[e.grd - 1]++; });
                html += buildFilterSubjectRow('사회탐구 전체', 'inq_social', 'all', allCounts);
            }
            socialNames.forEach(name => {
                const entries = socialGroup[name];
                const counts = Array(9).fill(0);
                entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
                html += buildFilterSubjectRow(socialNames.length > 1 ? `└ ${name}` : name, 'inq_social', name, counts, socialNames.length > 1);
            });
        }

        const scienceNames = Object.keys(scienceGroup).sort();
        if (scienceNames.length > 0) {
            html += `<div class="grade-filter-area-header">과학탐구 영역</div>`;
            if (scienceNames.length > 1) {
                const allEntries = scienceNames.flatMap(n => scienceGroup[n]);
                const allCounts = Array(9).fill(0);
                allEntries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) allCounts[e.grd - 1]++; });
                html += buildFilterSubjectRow('과학탐구 전체', 'inq_science', 'all', allCounts);
            }
            scienceNames.forEach(name => {
                const entries = scienceGroup[name];
                const counts = Array(9).fill(0);
                entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
                html += buildFilterSubjectRow(scienceNames.length > 1 ? `└ ${name}` : name, 'inq_science', name, counts, scienceNames.length > 1);
            });
        }

        const unknownNames = Object.keys(unknownGroup).sort();
        if (unknownNames.length > 0) {
            html += `<div class="grade-filter-area-header">탐구 영역 (기타)</div>`;
            unknownNames.forEach(name => {
                const entries = unknownGroup[name];
                const counts = Array(9).fill(0);
                entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
                html += buildFilterSubjectRow(name, 'inq_unknown', name, counts, false);
            });
        }

        group.innerHTML = html;
    }

    document.getElementById('gradeFilterResult').style.display = 'none';
    document.getElementById('gradeFilterEmpty').style.display = 'none';
}

function buildFilterSubjectRow(label, filterArea, filterChoice, counts, indented) {
    const safeChoice = String(filterChoice).replace(/'/g, "\\'");
    return `<div class="grade-filter-subject">
        <div class="grade-filter-subject-label" ${indented ? 'style="padding-left:16px;"' : ''}>${label}</div>
        <div class="grade-filter-btns">
            ${Array.from({length: 9}, (_, i) => i + 1).map(g => `
                <button class="grade-filter-btn"
                    onclick="renderGradeFilter('${filterArea}', ${g}, '${safeChoice}')"
                    title="${g}등급: ${counts[g-1]}명">
                    ${g}등급 <span class="grade-filter-count">${counts[g-1]}</span>
                </button>`).join('')}
        </div>
    </div>`;
}

window.renderGradeFilter = function(filterArea, gradeLevel, choiceFilter) {
    // 클릭한 버튼 찾기
    let clickedBtn = null;
    const simpleBtnSelector = `.g-btn-${filterArea}-${gradeLevel}`;
    const simpleBtn = document.querySelector(simpleBtnSelector);

    if (simpleBtn) {
        clickedBtn = simpleBtn;
    } else {
        document.querySelectorAll('.grade-filter-btn').forEach(btn => {
            const oc = btn.getAttribute('onclick') || '';
            if (oc.includes(`'${filterArea}', ${gradeLevel}`) && oc.includes(`'${choiceFilter}'`)) {
                clickedBtn = btn;
            }
        });
    }

    // 토글: 이미 활성화된 버튼이면 비활성화
    if (clickedBtn && clickedBtn.classList.contains('active')) {
        clickedBtn.classList.remove('active');
    } else if (clickedBtn) {
        clickedBtn.classList.add('active');
    }

    // 현재 활성화된 모든 버튼 수집
    const activeButtons = document.querySelectorAll('.grade-filter-btn.active');

    if (activeButtons.length === 0) {
        // 활성 버튼이 없으면 결과 숨김
        document.getElementById('gradeFilterResult').style.display = 'none';
        document.getElementById('gradeFilterEmpty').style.display = 'none';
        return;
    }

    // 활성 버튼에서 필터 조건 추출
    const activeFilters = [];
    activeButtons.forEach(btn => {
        const oc = btn.getAttribute('onclick') || '';
        const match = oc.match(/renderGradeFilter\x28'([^']+)',\s*(\d+),\s*'([^']+)'\x29/);
        if (match) {
            activeFilters.push({ area: match[1], grade: parseInt(match[2]), choice: match[3] });
        }
    });

    if (activeFilters.length === 0) {
        document.getElementById('gradeFilterResult').style.display = 'none';
        document.getElementById('gradeFilterEmpty').style.display = 'none';
        return;
    }

    // 학생 데이터 가져오기
    const currentGrade = state.currentGradeTotal || state.availableGrades[0];
    const exams = getExamsForGrade(currentGrade);
    const examSelect = document.getElementById('examSelectTotal');
    const examIdx = parseInt(examSelect.value) || 0;
    const students = exams[examIdx]?.students || [];

    // 모든 활성 필터에 해당하는 결과 수집
    const allRows = [];
    const summaryParts = [];

    activeFilters.forEach(filter => {
        const isInquiry = filter.area.startsWith('inq_');

        if (isInquiry) {
            const { socialGroup, scienceGroup, unknownGroup } = getInquiryMergedGroups(students);
            let targetGroup = {};
            let categoryName = '';

            if (filter.area === 'inq_social') {
                targetGroup = socialGroup; categoryName = '사회탐구';
            } else if (filter.area === 'inq_science') {
                targetGroup = scienceGroup; categoryName = '과학탐구';
            } else {
                targetGroup = unknownGroup; categoryName = '탐구(기타)';
            }

            let entries = [];
            if (filter.choice === 'all') {
                entries = Object.values(targetGroup).flat();
            } else {
                entries = targetGroup[filter.choice] || [];
            }

            entries = entries.filter(e => e.grd === filter.grade);
            entries.sort((a, b) => b.raw - a.raw);

            const displayLabel = filter.choice === 'all' ? categoryName : `${categoryName}-${filter.choice}`;
            if (entries.length > 0) {
                summaryParts.push(`<strong>${displayLabel}</strong> <span class="grade-badge g-${filter.grade}">${filter.grade}등급</span> ${entries.length}명`);
            }

            entries.forEach(e => {
                const s = e.student;
                const id = `${s.info.grade}${String(s.info.class).padStart(2,'0')}${String(s.info.no).padStart(2,'0')}`;
                const srcLabel = e.source === 'inq1' ? '탐구1' : '탐구2';
                allRows.push({
                    type: 'inquiry',
                    id, name: s.info.name, subject: e.name || '-', src: srcLabel,
                    raw: e.raw, grd: e.grd, std: e.std || '-', pct: e.pct || '-',
                    filterLabel: displayLabel
                });
            });
        } else {
            const area = areas.find(a => a.k === filter.area);
            const isAbs = (filter.area === 'eng' || filter.area === 'hist');

            let filtered = students.filter(s => s[filter.area].grd === filter.grade);
            if (filter.choice !== 'all') {
                filtered = filtered.filter(s => (s[filter.area].name || '(미분류)') === filter.choice);
            }
            filtered.sort((a, b) => {
                const diff = b[filter.area].raw - a[filter.area].raw;
                return diff !== 0 ? diff : (a.info.class * 100 + a.info.no) - (b.info.class * 100 + b.info.no);
            });

            const displayLabel = filter.choice === 'all' ? area.n : `${area.n}-${filter.choice}`;
            if (filtered.length > 0) {
                summaryParts.push(`<strong>${displayLabel}</strong> <span class="grade-badge g-${filter.grade}">${filter.grade}등급</span> ${filtered.length}명`);
            }

            filtered.forEach(s => {
                const sub = s[filter.area];
                const id = `${s.info.grade}${String(s.info.class).padStart(2,'0')}${String(s.info.no).padStart(2,'0')}`;
                allRows.push({
                    type: 'normal',
                    id, name: s.info.name,
                    subject: area.hasChoice ? (sub.name || '-') : '',
                    hasChoice: area.hasChoice,
                    raw: sub.raw, grd: sub.grd,
                    std: isAbs ? null : (sub.std || '-'),
                    pct: isAbs ? null : (sub.pct || '-'),
                    isAbs,
                    filterLabel: displayLabel
                });
            });
        }
    });

    const resultEl = document.getElementById('gradeFilterResult');
    const emptyEl = document.getElementById('gradeFilterEmpty');
    const summaryEl = document.getElementById('gradeFilterSummary');
    const theadEl = document.getElementById('gradeFilterThead');
    const tbodyEl = document.getElementById('gradeFilterTbody');

    if (allRows.length === 0) {
        resultEl.style.display = 'none';
        emptyEl.style.display = 'flex';
        return;
    }
    emptyEl.style.display = 'none';
    resultEl.style.display = 'block';

    const totalCount = allRows.length;
    summaryEl.innerHTML = `<span class="grade-filter-summary-text">
        ${summaryParts.join(' &nbsp;|&nbsp; ')}
        &nbsp;&nbsp; 합계 <strong>${totalCount}명</strong>
    </span>`;

    // 통합 테이블 헤더 (모든 컬럼 표시)
    theadEl.innerHTML = `<tr>
        <th>조회조건</th><th>학번</th><th>이름</th><th>선택과목</th><th>원점수</th><th>등급</th><th>표준점수</th><th>백분위</th>
    </tr>`;

    tbodyEl.innerHTML = allRows.map(r => {
        return `<tr>
            <td style="font-size:0.75rem;color:var(--text-secondary);">${r.filterLabel}</td>
            <td>${r.id}</td>
            <td style="font-weight:bold;">${r.name}</td>
            <td style="font-size:0.8rem;color:var(--text-secondary);">${r.subject || (r.type === 'inquiry' ? r.subject : '-')}</td>
            <td>${r.raw}</td>
            <td class="g-${r.grd}">${r.grd}</td>
            <td>${r.std !== null ? r.std : '-'}</td>
            <td>${r.pct !== null ? r.pct : '-'}</td>
        </tr>`;
    }).join('');
};

/* ── 등급 필터 결과 엑셀 저장 ── */
window.exportGradeFilterToExcel = function() {
    const thead = document.getElementById('gradeFilterThead');
    const tbody = document.getElementById('gradeFilterTbody');
    if (!thead || !tbody || tbody.rows.length === 0) {
        return alert('저장할 데이터가 없습니다.');
    }

    const wb = XLSX.utils.book_new();
    const rows = [];

    // 헤더
    const headerRow = [];
    thead.querySelectorAll('tr').forEach(tr => {
        tr.querySelectorAll('th').forEach(th => {
            headerRow.push(th.textContent.trim());
        });
    });
    rows.push(headerRow);

    // 본문
    tbody.querySelectorAll('tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('td').forEach(td => {
            const val = td.textContent.trim();
            const num = Number(val);
            row.push(isNaN(num) || val === '' || val === '-' ? val : num);
        });
        rows.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // 열 너비 자동 조정
    const colWidths = rows[0].map((_, colIdx) => {
        let maxLen = 0;
        rows.forEach(row => {
            const cellLen = String(row[colIdx] || '').length;
            if (cellLen > maxLen) maxLen = cellLen;
        });
        return { wch: Math.max(maxLen + 2, 8) };
    });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, '등급구간조회');

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `등급구간_학생조회_${dateStr}.xlsx`);
};

/* ── 전체 성적 일람표 엑셀 저장 ── */
window.exportTotalTableToExcel = function() {
    const thead = document.getElementById('totalTableHead');
    const tbody = document.getElementById('totalTableBody');
    if (!thead || !tbody || tbody.rows.length === 0) {
        return alert('저장할 데이터가 없습니다.');
    }

    const wb = XLSX.utils.book_new();
    const rows = [];

    // 헤더 처리 (2행 병합 헤더 → 평탄화)
    const headerRows = thead.querySelectorAll('tr');
    if (headerRows.length === 2) {
        // 1행과 2행을 합쳐서 단일 헤더 생성
        const mergedHeader = [];
        const row1Cells = headerRows[0].querySelectorAll('th');
        const row2Cells = Array.from(headerRows[1].querySelectorAll('th'));
        let row2Idx = 0;

        row1Cells.forEach(th => {
            const colspan = parseInt(th.getAttribute('colspan')) || 1;
            const rowspan = parseInt(th.getAttribute('rowspan')) || 1;

            if (rowspan === 2) {
                // 2행 병합 셀: 그대로 추가
                mergedHeader.push(th.textContent.trim());
            } else if (colspan > 1) {
                // colspan 셀: 2행의 하위 셀과 결합
                const parentName = th.textContent.trim();
                for (let i = 0; i < colspan && row2Idx < row2Cells.length; i++) {
                    mergedHeader.push(`${parentName}_${row2Cells[row2Idx].textContent.trim()}`);
                    row2Idx++;
                }
            } else {
                mergedHeader.push(th.textContent.trim());
            }
        });
        rows.push(mergedHeader);
    } else {
        // 단일 행 헤더
        const headerRow = [];
        headerRows[0].querySelectorAll('th').forEach(th => {
            headerRow.push(th.textContent.trim());
        });
        rows.push(headerRow);
    }

    // 본문
    tbody.querySelectorAll('tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('td').forEach(td => {
            const val = td.textContent.trim();
            const num = Number(val);
            row.push(isNaN(num) || val === '' || val === '-' ? val : num);
        });
        rows.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // 열 너비 자동 조정
    const colWidths = rows[0].map((_, colIdx) => {
        let maxLen = 0;
        rows.forEach(row => {
            const cellLen = String(row[colIdx] || '').length;
            if (cellLen > maxLen) maxLen = cellLen;
        });
        return { wch: Math.max(maxLen + 2, 8) };
    });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, '성적일람표');

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `전체성적일람표_${dateStr}.xlsx`);
};

/* ==============================
   학급통계 탭
   ============================== */
window.changeClassMetric = function (m) {
    state.classMetric = m;
    document.querySelectorAll('#class-tab .opt-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-c-' + m).classList.add('active');
    renderClass();
};
window.sortClass = function (t) {
    state.classSort = t;
    document.getElementById('btn-sort-total').classList.remove('active');
    document.getElementById('btn-sort-no').classList.remove('active');
    document.getElementById('btn-sort-' + t).classList.add('active');
    renderClass();
};

function renderClass() {
    const grade = state.currentGradeClass || state.availableGrades[0];
    if (!grade) return;
    const exams = getExamsForGrade(grade);
    const examSelect = document.getElementById('examSelectClass');
    const classSelect = document.getElementById('classSelect');
    if (!examSelect || !classSelect || !exams.length) return;
    const examIdx = parseInt(examSelect.value) || 0;
    if (!exams[examIdx]) return;
    const studentsAll = exams[examIdx].students;
    const cls = parseInt(classSelect.value);
    if (isNaN(cls)) return;
    const metric = state.classMetric;

    let students = studentsAll.filter(s => s.info.class === cls);
    const getTot = s => metric === 'raw' ? s.totalRaw : (metric === 'std' ? s.totalStd.toFixed(1) : s.totalPct.toFixed(1));
    if (state.classSort === 'total') students.sort((a, b) => parseFloat(getTot(b)) - parseFloat(getTot(a)));
    else students.sort((a, b) => a.info.no - b.info.no);

    /* ── 학급 영역별 선택과목 카드 ── */
    const container = document.getElementById('classStatsContainer');
    if (container) {
        container.innerHTML = '';

        const isSimple = checkSimpleMode(students);

        if (isSimple) {
            const simpleSubjects = [
                { k: 'kor',  n: '국어' },
                { k: 'math', n: '수학' },
                { k: 'eng',  n: '영어' },
                { k: 'hist', n: '한국사' },
                { k: 'inq1', n: '사회탐구' },
                { k: 'inq2', n: '과학탐구' }
            ];
            const wrapper = document.createElement('div');
            wrapper.className = 'subject-grid subject-grid-fixed';
            simpleSubjects.forEach(sub => {
                wrapper.innerHTML += buildSubjectCardHTML(sub.n, students, sub.k, metric);
            });
            container.appendChild(wrapper);
        } else {
            ['kor', 'math', 'eng', 'hist'].forEach(key => {
                const area = areas.find(a => a.k === key);
                let areaHTML = `<div class="area-section">
                    <h4 class="area-title"><i class="fas fa-layer-group"></i> ${area.n}</h4>`;
                areaHTML += buildSingleAreaHTML(students, area.k, area.n, area.hasChoice, metric);
                areaHTML += `</div>`;
                container.innerHTML += areaHTML;
            });
            container.innerHTML += buildInquiryAreaHTML(students, metric);
        }
    }

    /* ── 학급 일람표 (동적 헤더) ── */
    const thead = document.getElementById('classTableHead');
    const tbody = document.getElementById('classTableBody');

    let h1 = `<tr><th rowspan="2">번호</th><th rowspan="2">이름</th>`;
    let h2 = `<tr>`;

    areas.forEach(area => {
        if (area.hasChoice) {
            h1 += `<th colspan="3">${area.n}</th>`;
            h2 += `<th>선택과목</th><th>점수</th><th>등급</th>`;
        } else {
            h1 += `<th colspan="2">${area.n}</th>`;
            h2 += `<th>점수</th><th>등급</th>`;
        }
    });
    h1 += `<th rowspan="2" class="total-col">총점</th><th rowspan="2" class="total-col">석차</th><th rowspan="2" class="total-col">평균등급</th></tr>`;
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    const getAvgGrade = s => ((s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4).toFixed(2);
    const gradeForClass = state.currentGradeClass || state.availableGrades[0];
    tbody.innerHTML = '';
    students.forEach(s => {
        const rank = students.filter(st => parseFloat(getTot(st)) > parseFloat(getTot(s))).length + 1;
        let row = `<tr><td class="student-link-cell" onclick="goToIndividual('${s.uid}', ${gradeForClass})" title="개인통계 보기">${s.info.no}</td><td class="student-link-cell" onclick="goToIndividual('${s.uid}', ${gradeForClass})" title="개인통계 보기" style="font-weight:bold;">${s.info.name}</td>`;

        areas.forEach(area => {
            const isAbs = (area.k === 'eng' || area.k === 'hist');
            const subj = s[area.k];
            if (area.hasChoice) {
                const scoreVal = isAbs ? subj.raw : (subj[metric] || '-');
                row += `<td style="font-size:0.75rem;color:var(--text-secondary);">${subj.name || '-'}</td>`;
                row += `<td>${scoreVal}</td>`;
                row += `<td class="g-${subj.grd}">${subj.grd}</td>`;
            } else {
                row += `<td>${subj.raw}</td>`;
                row += `<td class="g-${subj.grd}">${subj.grd}</td>`;
            }
        });

        row += `<td class="total-col">${getTot(s)}</td><td class="total-col">${rank}</td><td class="total-col" style="font-weight:bold;color:var(--primary);">${getAvgGrade(s)}</td></tr>`;
        tbody.innerHTML += row;
    });
}

/* ==============================
   개인통계 탭
   ============================== */
function updateIndivList() {
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return;
    const exams = getExamsForGrade(grade);
    if (!exams.length) return;
    const clsVal = document.getElementById('indivClassSelect')?.value;
    if (!clsVal) return;
    const cls = parseInt(clsVal);
    if (isNaN(cls)) return;

    // ★ 수정: 모든 시험에서 해당 반 학생을 수집 (중복 제거)
    const studentMap = new Map();
    exams.forEach(exam => {
        exam.students
            .filter(s => s.info.class === cls)
            .forEach(s => {
                if (!studentMap.has(s.uid)) {
                    studentMap.set(s.uid, s);
                }
            });
    });

    const list = Array.from(studentMap.values())
                      .sort((a, b) => a.info.no - b.info.no);

    document.getElementById('indivStudentSelect').innerHTML =
        list.map(s => `<option value="${s.uid}">${s.info.no}번 ${s.info.name}</option>`).join('');

    if (list.length > 0) renderIndividual();
}

/* ============================================================
   학생 검색 기능
   ============================================================ */
function initStudentSearch() {
    const searchInput = document.getElementById('studentSearchInput');
    const clearBtn = document.getElementById('studentSearchClear');
    const dropdown = document.getElementById('studentSearchDropdown');

    if (!searchInput) return;

    let highlightIdx = -1;

    // 입력 시 검색
    searchInput.addEventListener('input', () => {
        const query = searchInput.value.trim();
        clearBtn.style.display = query ? 'block' : 'none';

        if (query.length === 0) {
            dropdown.style.display = 'none';
            highlightIdx = -1;
            return;
        }

        const results = searchStudents(query);
        renderSearchDropdown(results, query);
        highlightIdx = -1;
    });

    // 포커스 시 이미 텍스트가 있으면 드롭다운 표시
    searchInput.addEventListener('focus', () => {
        const query = searchInput.value.trim();
        if (query.length > 0) {
            const results = searchStudents(query);
            renderSearchDropdown(results, query);
        }
    });

    // 키보드 탐색
    searchInput.addEventListener('keydown', (e) => {
        const items = dropdown.querySelectorAll('.student-search-item');
        if (!items.length || dropdown.style.display === 'none') return;

        if (e.key === 'ArrowDown') {
            e.preventDefault();
            highlightIdx = Math.min(highlightIdx + 1, items.length - 1);
            updateHighlight(items, highlightIdx);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            highlightIdx = Math.max(highlightIdx - 1, 0);
            updateHighlight(items, highlightIdx);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (highlightIdx >= 0 && highlightIdx < items.length) {
                // 화살표로 선택한 항목이 있으면 해당 항목 선택
                items[highlightIdx].click();
            } else if (items.length > 0) {
                // 화살표 선택 없이 Enter만 누르면 첫 번째 결과 자동 선택
                items[0].click();
            }
        } else if (e.key === 'Escape') {
            dropdown.style.display = 'none';
            highlightIdx = -1;
        }
    });

    // X 버튼 클릭
    clearBtn.addEventListener('click', () => {
        searchInput.value = '';
        clearBtn.style.display = 'none';
        dropdown.style.display = 'none';
        highlightIdx = -1;
        searchInput.focus();
    });

    // 외부 클릭 시 드롭다운 닫기
    document.addEventListener('click', (e) => {
        const wrapper = document.getElementById('studentSearchWrapper');
        if (wrapper && !wrapper.contains(e.target)) {
            dropdown.style.display = 'none';
            highlightIdx = -1;
        }
    });
}

function searchStudents(query) {
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return [];
    const exams = getExamsForGrade(grade);
    if (!exams.length) return [];

    // 모든 시험에서 학생 수집 (중복 제거)
    const studentMap = new Map();
    exams.forEach(exam => {
        exam.students.forEach(s => {
            if (!studentMap.has(s.uid)) {
                studentMap.set(s.uid, s);
            }
        });
    });

    const allStudents = Array.from(studentMap.values());
    const q = query.toLowerCase();

    return allStudents.filter(s => {
        const name = s.info.name.toLowerCase();
        const no = String(s.info.no);
        const classNo = String(s.info.class);
        const fullId = `${s.info.class}반 ${s.info.no}번`;
        return name.includes(q) || no === q || classNo === q || fullId.includes(q);
    }).sort((a, b) => {
        // 이름 매칭 우선, 그 다음 반-번호순
        const aNameMatch = a.info.name.toLowerCase().startsWith(q) ? 0 : 1;
        const bNameMatch = b.info.name.toLowerCase().startsWith(q) ? 0 : 1;
        if (aNameMatch !== bNameMatch) return aNameMatch - bNameMatch;
        if (a.info.class !== b.info.class) return a.info.class - b.info.class;
        return a.info.no - b.info.no;
    }).slice(0, 30); // 최대 30명
}

function renderSearchDropdown(results, query) {
    const dropdown = document.getElementById('studentSearchDropdown');
    if (!dropdown) return;

    if (results.length === 0) {
        dropdown.innerHTML = `<div class="student-search-no-result">
            <i class="fas fa-search" style="margin-right:6px;"></i>
            '${query}'에 해당하는 학생이 없습니다
        </div>`;
        dropdown.style.display = 'block';
        return;
    }

    let html = `<div class="student-search-count">검색 결과: ${results.length}명</div>`;

    results.forEach(s => {
        const highlightedName = highlightText(s.info.name, query);
        html += `<div class="student-search-item" data-uid="${s.uid}" data-class="${s.info.class}">
            <span class="student-search-item-name">${highlightedName}</span>
            <span class="student-search-item-info">${s.info.class}반 ${s.info.no}번</span>
        </div>`;
    });

    dropdown.innerHTML = html;
    dropdown.style.display = 'block';

    // 클릭 이벤트
    dropdown.querySelectorAll('.student-search-item').forEach(item => {
        item.addEventListener('click', () => {
            const uid = item.dataset.uid;
            const cls = item.dataset.class;

            // 반 선택 변경
            const classSelect = document.getElementById('indivClassSelect');
            if (classSelect) {
                classSelect.value = cls;
                updateIndivList();
            }

            // 학생 선택
            setTimeout(() => {
                const studentSelect = document.getElementById('indivStudentSelect');
                if (studentSelect) {
                    studentSelect.value = uid;
                    renderIndividual();
                }

                // 검색창 정리
                const searchInput = document.getElementById('studentSearchInput');
                const clearBtn = document.getElementById('studentSearchClear');
                if (searchInput) searchInput.value = '';
                if (clearBtn) clearBtn.style.display = 'none';
                dropdown.style.display = 'none';
            }, 100);
        });
    });
}

function highlightText(text, query) {
    if (!query) return text;
    const regex = new RegExp(`(${query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
    return text.replace(regex, '<mark>$1</mark>');
}

function updateHighlight(items, idx) {
    items.forEach((item, i) => {
        item.classList.toggle('highlighted', i === idx);
    });
    if (items[idx]) {
        items[idx].scrollIntoView({ block: 'nearest' });
    }
}

function renderIndividual() {
    const sel = document.getElementById('indivStudentSelect');
    if (!sel || !sel.value) return;
    const uid = sel.value;
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return;
    const exams = getExamsForGrade(grade);
    if (!exams.length) return;

    const history = [];
    for (let i = exams.length - 1; i >= 0; i--) {
        const ex = exams[i];
        const s = ex.students.find(st => st.uid === uid);
        if (s) history.push({ name: ex.name, data: s });
    }
    if (!history.length) return;

    const selectedExamIdx = parseInt(document.getElementById('indivExamSelect')?.value) || 0;
    let currentData = exams[selectedExamIdx]?.students.find(st => st.uid === uid);
    let selectedExamName = exams[selectedExamIdx]?.name;
    if (!currentData) {
        currentData = history[history.length - 1].data;
        selectedExamName = history[history.length - 1].name;
    }

    document.getElementById('indivName').innerText = currentData.info.name;
    document.getElementById('indivInfo').innerText = `${currentData.info.grade}학년 ${currentData.info.class}반 ${currentData.info.no}번`;
    document.getElementById('latestExamName').innerText = `선택 시험: ${selectedExamName}`;
    document.getElementById('indivTotalRaw').innerText = currentData.totalRaw;
    document.getElementById('indivTotalStd').innerText = currentData.totalStd.toFixed(1);
    document.getElementById('indivTotalPct').innerText = currentData.totalPct.toFixed(2);
    document.getElementById('indivRank').innerText = currentData.totalRank;

    const inqAvgGrade = (currentData.inq1.grd + currentData.inq2.grd) / 2;
    const avgGrade = ((currentData.kor.grd + currentData.math.grd + currentData.eng.grd + inqAvgGrade) / 4).toFixed(2);
    document.getElementById('indivAverageGrade').innerText = avgGrade;

    // 총점 추이 차트
    drawChart('totalTrendChart', 'line', {
        labels: history.map(h => h.name),
        datasets: [{
            label: '총점(백분위합)', data: history.map(h => h.data.totalPct),
            borderColor: '#8B5A8D', backgroundColor: '#8B5A8D', tension: 0.3, borderWidth: 2, pointRadius: 4
        }]
    }, {
        scales: { y: { min: 0, max: 450, ticks: { stepSize: 50, callback: v => v === 450 ? '' : v } } },
        plugins: { datalabels: { display: true, color: '#8B5A8D', align: 'top', font: { weight: 'bold' }, formatter: v => v.toFixed(1) } }
    });

    // 레이더 차트 — 선택과목명 표시
    const radarLabels = ['국어', '수학', '영어', '탐1', '탐2'];
    const radarGrades = [
        currentData.kor.grd,
        currentData.math.grd,
        currentData.eng.grd,
        currentData.inq1.grd,
        currentData.inq2.grd
    ];

    drawChart('radarChart', 'radar', {
        labels: radarLabels,
        datasets: [{
            label: '등급', data: radarGrades,
            backgroundColor: 'rgba(100, 150, 200, 0.2)', borderColor: 'rgba(100, 150, 200, 0.6)',
            borderWidth: 2, pointBackgroundColor: radarPointColors, pointBorderColor: radarPointColors,
            pointBorderWidth: 2, pointRadius: 6, pointHoverRadius: 8
        }]
    }, {
        layout: { padding: { top: 30, bottom: 30, left: 50, right: 50 } },
        scales: {
            r: {
                reverse: true, min: 1, max: 9,
                ticks: { stepSize: 1, display: true, font: { size: 10 }, color: '#999', backdropColor: 'transparent' },
                grid: { circular: false, color: 'rgba(0, 0, 0, 0.1)' },
                angleLines: { display: true, color: 'rgba(0, 0, 0, 0.1)' },
                pointLabels: { font: { size: 13, weight: 'bold', family: "'Pretendard', sans-serif" }, color: radarPointColors, padding: 12 }
            }
        },
        plugins: {
            legend: { display: false },
            datalabels: {
                display: true,
                backgroundColor: function (ctx) { return ctx.dataset.data[ctx.dataIndex] <= 1.5 ? 'rgba(255,255,255,0.85)' : '#ffffff'; },
                borderColor: function (ctx) { return radarPointColors[ctx.dataIndex]; },
                borderWidth: 2, color: function (ctx) { return radarPointColors[ctx.dataIndex]; },
                borderRadius: 4, padding: { top: 3, bottom: 3, left: 6, right: 6 },
                font: { weight: 'bold', size: 10 }, formatter: (v) => v.toFixed(2) + '등급',
                anchor: 'center',
                align: function (ctx) {
                    const idx = ctx.dataIndex;
                    const count = ctx.dataset.data.length;
                    const angleDeg = 90 - (360 / count) * idx;
                    return angleDeg * Math.PI / 180 + Math.PI;
                },
                offset: function (ctx) {
                    const v = ctx.dataset.data[ctx.dataIndex];
                    const normalized = (v - 1) / 8;
                    return Math.round(4 + (1 - normalized) * 10);
                },
                clip: false
            }
        }
    });

    /* ── 과목별 상세 (영역 단위) ── */
    const detailContainer = document.getElementById('subjectDetailContainer');
    if (detailContainer) {
        state.subjectCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
        state.subjectCharts = [];
        detailContainer.innerHTML = '';

        areas.forEach(area => {
            const k = area.k;
            const isAbs = (k === 'eng' || k === 'hist');

            // 영역명 + 선택과목명 결정
            let displayTitle = area.n;
            if (area.hasChoice && currentData[k].name) {
                displayTitle = `${area.n} (${currentData[k].name})`;
            }

            // 테이블 헤더: 각 회차
            let thead = `<tr><th>구분</th>` + history.map(h => `<th>${h.name}</th>`).join('') + `</tr>`;

            // 선택과목 행 (선택과목이 있는 영역만)
            let trChoice = '';
            if (area.hasChoice) {
                trChoice = `<tr><td style="font-weight:600;color:var(--text-secondary);">선택과목</td>` +
                    history.map(h => `<td style="font-size:0.78rem;color:var(--text-secondary);">${h.data[k].name || '-'}</td>`).join('') + `</tr>`;
            }

            let trGrd = `<tr><td>등급</td>` + history.map(h => `<td class="g-${h.data[k].grd}">${h.data[k].grd}</td>`).join('') + `</tr>`;
            let trRaw = `<tr><td>원점수</td>` + history.map(h => `<td>${h.data[k].raw}</td>`).join('') + `</tr>`;
            let trStd = `<tr><td>표준점수</td>` + history.map(h => `<td>${h.data[k].std || '-'}</td>`).join('') + `</tr>`;
            let trPct = `<tr><td>백분위</td>` + history.map(h => `<td>${h.data[k].pct || '-'}</td>`).join('') + `</tr>`;

            const chartId = `chart-${k}-${uid.replace(/[^a-zA-Z0-9]/g, '')}`;

            detailContainer.innerHTML += `
                <div class="chart-card subject-detail-card" data-subject="${k}">
                    <h3><i class="fas fa-book"></i> ${displayTitle} 성적 상세</h3>
                    <div class="subject-detail-grid">
                        <div class="chart-container" style="height: 200px; width: 100%;"><canvas id="${chartId}"></canvas></div>
                        <div class="table-wrapper">
                            <table class="data-table" style="font-size:0.8rem; min-width: 100%;">
                                <thead>${thead}</thead>
                                <tbody>${trChoice}${trGrd}${trRaw}${!isAbs ? trStd : ''}${!isAbs ? trPct : ''}</tbody>
                            </table>
                        </div>
                    </div>
                </div>`;

            setTimeout(() => {
                const ctx = document.getElementById(chartId);
                if (!ctx) return;
                const yVals = history.map(h => isAbs ? h.data[k].grd : h.data[k].pct);
                const chart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: history.map(h => h.name),
                        datasets: [{ label: isAbs ? '등급' : '백분위', data: yVals, borderColor: isAbs ? '#B8860B' : '#4A6B8A', tension: 0.1, borderWidth: 2, pointRadius: 4 }]
                    },
                    options: {
                        responsive: true, maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false },
                            datalabels: { display: true, color: isAbs ? '#B8860B' : '#4A6B8A', align: 'top', font: { weight: 'bold' }, formatter: v => v > 100 ? '' : v }
                        },
                        scales: {
                            y: {
                                reverse: isAbs, min: isAbs ? 0 : 0, max: isAbs ? 9 : 120,
                                ticks: { stepSize: isAbs ? 1 : 20, callback: function (value) { if (isAbs && value === 0) return ''; return value; } }
                            }
                        }
                    }
                });
                state.subjectCharts.push(chart);
            }, 0);
        });
    }
}

function drawChart(id, type, data, options) {
    const ctx = document.getElementById(id);
    if (!ctx) return;
    if (state.charts[id]) state.charts[id].destroy();
    state.charts[id] = new Chart(ctx.getContext('2d'), { type, data, options });
}

function updateLastUpdated() {
    const now = new Date();
    const el = document.getElementById('lastUpdated');
    if (el) el.textContent = `Last updated: ${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')} (KST)`;
}

/* ── HTML 저장 ── */
window.saveHtmlFile = function () {
    try {
        const htmlContent = document.documentElement.cloneNode(true);
        htmlContent.querySelector('#uploadSection')?.remove();
        htmlContent.querySelector('#loading')?.remove();
        htmlContent.querySelector('#saveHtmlBtn')?.remove();
        const stateToSave = {
            gradeData: state.gradeData, availableGrades: state.availableGrades,
            currentGradeTotal: state.currentGradeTotal, currentGradeClass: state.currentGradeClass,
            currentGradeIndiv: state.currentGradeIndiv, metric: state.metric,
            classMetric: state.classMetric, classSort: state.classSort, charts: {}
        };
        const scriptTag = document.createElement('script');
        scriptTag.textContent = `window.SAVED_STATE = ${JSON.stringify(stateToSave)}; window.addEventListener('DOMContentLoaded', function() { if (window.SAVED_STATE) { Object.assign(state, window.SAVED_STATE); state.charts = {}; document.getElementById('results').style.display = 'block'; initSelectors(); } });`;
        htmlContent.querySelector('head').appendChild(scriptTag);
        const blob = new Blob(['<!DOCTYPE html>\n' + htmlContent.outerHTML], { type: 'text/html;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        a.download = `모의고사_분석결과_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.html`;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (e) { alert('HTML 저장 중 오류가 발생했습니다: ' + e.message); }
}

/* ==============================
   PDF 생성
   ============================== */
async function captureTabPageV2(pdf, showSelectors, addNewPage) {
    const source = document.getElementById('individual-tab');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfPageHeight = pdf.internal.pageSize.getHeight();
    const marginX = 8, marginY = 8;
    const contentWidth = pdfWidth - marginX * 2;
    const cloneWidth = 1600;

    const canvasMap = new Map();
    source.querySelectorAll('canvas').forEach(c => {
        if (!c.id) return;
        try { canvasMap.set(c.id, { dataUrl: c.toDataURL('image/png'), w: c.offsetWidth || c.getBoundingClientRect().width, h: c.offsetHeight || c.getBoundingClientRect().height }); } catch (e) {}
    });

    const wrapper = document.createElement('div');
    wrapper.style.cssText = `position:fixed;top:-99999px;left:0;width:${cloneWidth}px;min-width:${cloneWidth}px;max-width:${cloneWidth}px;background:#FEFDFB;padding:20px 24px;box-sizing:border-box;display:block;overflow:visible;z-index:-9999;`;

    // chart-card 너비를 wrapper에 맞게 강제하는 헬퍼
    const fixCardWidths = (el) => {
        el.querySelectorAll('.chart-card, .student-profile-card').forEach(card => {
            card.style.setProperty('width', '100%', 'important');
            card.style.setProperty('box-sizing', 'border-box', 'important');
            card.style.setProperty('max-width', 'none', 'important');
        });
    };

    const allSubjects = ['kor', 'math', 'eng', 'hist', 'inq1', 'inq2'];

    if (showSelectors.includes('profile')) {
        const profileCard = source.querySelector('.student-profile-card');
        if (profileCard) {
            const pc = profileCard.cloneNode(true);
            pc.style.marginBottom = '16px';
            pc.style.width = '100%';
            pc.style.boxSizing = 'border-box';
            const statsGrid = pc.querySelector('.profile-stats');
            if (statsGrid) statsGrid.setAttribute('style', 'display:grid !important;grid-template-columns:repeat(5,1fr) !important;gap:15px !important;width:100% !important;');
            const profileHeader = pc.querySelector('.profile-header');
            if (profileHeader) profileHeader.setAttribute('style', 'display:flex !important;justify-content:space-between !important;align-items:flex-start !important;margin-bottom:20px !important;width:100% !important;');
            wrapper.appendChild(pc);
        }
    }

    if (showSelectors.includes('charts')) {
        const chartsRow = source.querySelector('.charts-row');
        if (chartsRow) {
            const cr = chartsRow.cloneNode(true);
            cr.setAttribute('style', 'display:grid !important;grid-template-columns:3fr 2fr !important;gap:20px !important;margin-bottom:20px !important;width:100% !important;box-sizing:border-box !important;');
            cr.querySelectorAll('.chart-half').forEach(ch => {
                ch.setAttribute('style', 'display:flex !important;flex-direction:column !important;overflow:visible !important;min-height:0 !important;height:auto !important;width:100% !important;box-sizing:border-box !important;');
            });
            cr.querySelectorAll('.chart-half .chart-container').forEach(cc => { cc.setAttribute('style', 'overflow:visible !important;height:360px !important;min-height:360px !important;width:100% !important;display:block !important;position:relative !important;'); });
            wrapper.appendChild(cr);
        }
    }

    allSubjects.forEach(subj => {
        if (!showSelectors.includes(subj)) return;
        const card = source.querySelector(`.subject-detail-card[data-subject="${subj}"]`);
        if (!card) return;
        const cc = card.cloneNode(true);
        cc.style.marginBottom = '16px';
        cc.querySelectorAll('.subject-detail-grid').forEach(grid => {
            grid.setAttribute('style', 'display:grid !important;grid-template-columns:2fr 3fr !important;gap:16px !important;align-items:start !important;width:100% !important;');
            grid.querySelectorAll('.chart-container').forEach(chartC => { chartC.setAttribute('style', 'height:auto !important;width:100% !important;position:relative !important;overflow:visible !important;'); });
            grid.querySelectorAll('.table-wrapper').forEach(tw => { tw.setAttribute('style', 'overflow:visible !important;overflow-x:visible !important;max-height:none !important;width:100% !important;'); });
            grid.querySelectorAll('table').forEach(tbl => {
                tbl.setAttribute('style', 'min-width:0 !important;width:100% !important;table-layout:auto !important;font-size:0.75rem !important;border-collapse:separate !important;border-spacing:0 !important;overflow:hidden !important;');
                // thead 행을 tbody 맨 위로 물리적으로 이동 (html2canvas thead/tbody display 버그 회피)
                const thead = tbl.querySelector('thead');
                const tbody = tbl.querySelector('tbody');
                if (thead && tbody) {
                    Array.from(thead.querySelectorAll('tr')).reverse().forEach(row => {
                        tbody.insertBefore(row, tbody.firstChild);
                    });
                    thead.remove();
                }
                // 이동 후: tbody 첫 번째 행(구분 헤더)에 상단 라운드 적용
                const firstRowCells = tbl.querySelectorAll('tbody tr:first-child th, tbody tr:first-child td');
                firstRowCells.forEach((cell, i) => {
                    let r = '';
                    if (i === 0) r += 'border-top-left-radius:12px !important;';
                    if (i === firstRowCells.length - 1) r += 'border-top-right-radius:12px !important;';
                    if (r) cell.setAttribute('style', (cell.getAttribute('style') || '') + r);
                });
                // tbody 마지막 행에 하단 라운드 적용
                const lastRowCells = tbl.querySelectorAll('tbody tr:last-child td');
                lastRowCells.forEach((td, i) => {
                    let r = '';
                    if (i === 0) r += 'border-bottom-left-radius:12px !important;';
                    if (i === lastRowCells.length - 1) r += 'border-bottom-right-radius:12px !important;';
                    if (r) td.setAttribute('style', (td.getAttribute('style') || '') + r);
                });
            });
            grid.querySelectorAll('th, td').forEach(cell => { cell.setAttribute('style', (cell.getAttribute('style')||'') + 'white-space:nowrap !important;padding:4px 6px !important;'); });
        });
        wrapper.appendChild(cc);
    });

    wrapper.querySelectorAll('canvas').forEach(cloneCanvas => {
        const snap = canvasMap.get(cloneCanvas.id);
        if (!snap) { cloneCanvas.style.display = 'none'; return; }
        const img = document.createElement('img');
        img.src = snap.dataUrl;
        const ar = snap.w > 0 && snap.h > 0 ? `${snap.w} / ${snap.h}` : 'auto';
        img.style.cssText = `display:block;width:100%;height:auto;aspect-ratio:${ar};object-fit:fill;`;
        cloneCanvas.parentNode.replaceChild(img, cloneCanvas);
    });

    // wrapper를 화면 밖 고정 위치에 추가 (레이아웃/스크롤에 영향 없음)
    wrapper.id = '__pdf_wrapper__';
    // body 스크롤/리플로우 방지 (overflow:hidden 임시 적용)
    const prevOverflow = document.documentElement.style.overflow;
    document.documentElement.style.overflow = 'hidden';
    document.body.appendChild(wrapper);
    fixCardWidths(wrapper);

    await new Promise(r => setTimeout(r, 300));

    const canvas = await html2canvas(wrapper, {
        scale: 2, useCORS: true, backgroundColor: '#FEFDFB',
        windowWidth: cloneWidth, logging: false, allowTaint: true,
        onclone: (clonedDoc, clonedEl) => {
            // html2canvas 내부 복제 DOM에서만 스타일 오버라이드 → 실제 화면 불변
            const overrideStyle = clonedDoc.createElement('style');
            overrideStyle.textContent = [
                `body { min-width: ${cloneWidth}px !important; overflow: visible !important; }`,
                `.container { max-width: none !important; width: 100% !important; }`,
                `#__pdf_wrapper__ { width: ${cloneWidth}px !important; min-width: ${cloneWidth}px !important; max-width: none !important; }`,
            ].join('\n');
            clonedDoc.head.appendChild(overrideStyle);
        },
    });

    document.body.removeChild(wrapper);
    document.documentElement.style.overflow = prevOverflow;

    const imgData = canvas.toDataURL('image/jpeg', 0.92);
    const imgProps = pdf.getImageProperties(imgData);
    const imgHeight = (imgProps.height * contentWidth) / imgProps.width;

    if (addNewPage) pdf.addPage();

    const maxH = pdfPageHeight - marginY * 2;
    if (imgHeight > maxH) {
        const scale = maxH / imgHeight;
        const scaledW = contentWidth * scale;
        const xOff = marginX + (contentWidth - scaledW) / 2;
        pdf.addImage(imgData, 'JPEG', xOff, marginY, scaledW, maxH);
    } else {
        pdf.addImage(imgData, 'JPEG', marginX, marginY, contentWidth, imgHeight);
    }
}

async function generateStudentPDF() {
    const btn = document.getElementById('pdfStudentBtn');
    if (!btn || btn.disabled) return;
    const uid = document.getElementById('indivStudentSelect')?.value;
    if (!uid) return alert('학생을 선택해주세요.');

    btn.disabled = true;
    const originalText = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> PDF 생성 중...';

    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');
        await new Promise(r => setTimeout(r, 1000));
        await captureTabPageV2(pdf, ['profile', 'charts'], false);
        await captureTabPageV2(pdf, ['kor', 'math', 'eng'], true);
        await captureTabPageV2(pdf, ['hist', 'inq1', 'inq2'], true);
        const studentName = document.getElementById('indivName').innerText;
        pdf.save(`모의고사_분석리포트_${studentName}.pdf`);
    } catch (error) { console.error('PDF 생성 오류:', error); alert('PDF 생성 중 오류가 발생했습니다.'); }
    finally { btn.disabled = false; btn.innerHTML = originalText; }
}

async function generateClassPDF() {
    const btn = document.getElementById('pdfClassBtn');
    if (!btn || btn.disabled) return;
    const classSelect = document.getElementById('indivClassSelect');
    if (!classSelect) return;
    const cls = parseInt(classSelect.value);
    const sel = document.getElementById('indivStudentSelect');
    const options = Array.from(sel.options);
    if (!options.length) return alert('해당 학급에 학생이 없습니다.');
    if (!confirm(`현재 선택된 ${cls}반 학생 ${options.length}명 전체 리포트를 PDF로 생성합니다.\n시간이 다소 소요될 수 있습니다. 진행하시겠습니까?`)) return;

    btn.disabled = true;
    const originalText = btn.innerHTML;

    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');
        for (let i = 0; i < options.length; i++) {
            btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> PDF 생성 중... (${i + 1}/${options.length})`;
            sel.value = options[i].value;
            renderIndividual();
            await new Promise(r => setTimeout(r, 1000));
            await captureTabPageV2(pdf, ['profile', 'charts'], i > 0);
            await captureTabPageV2(pdf, ['kor', 'math', 'eng'], true);
            await captureTabPageV2(pdf, ['hist', 'inq1', 'inq2'], true);
        }
        pdf.save(`모의고사_분석리포트_${cls}반_전체.pdf`);
    } catch (error) { console.error('PDF 생성 오류:', error); alert('학급 PDF 생성 중 오류가 발생했습니다.'); }
    finally { btn.disabled = false; btn.innerHTML = originalText; }
}
