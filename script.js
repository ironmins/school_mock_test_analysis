/**
 * 고등학교 모의고사 성적분석 V15.1
 * script.js — 상위권 Metric 선택 + 과목명 풀네임 + 테이블 라운드 개선
 */

Chart.register(ChartDataLabels);
Chart.defaults.plugins.datalabels.display = false;

/* ============================================================
   전역 상태
   ============================================================ */
const state = {
    gradeData: {},
    availableGrades: [],
    currentGradeTotal: null,
    currentGradeClass: null,
    currentGradeIndiv: null,
    currentGradeTop: null,
    metric: 'raw',
    classMetric: 'raw',
    classSort: 'no',
    topMetric: 'raw',
    topN: 20,
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

/* ============================================================
   과목 분류
   ============================================================ */
const socialSubjects = [
    '생활과 윤리', '윤리와 사상', '한국 지리', '세계 지리',
    '동아시아사', '세계사', '경제', '정치와 법',
    '사회·문화', '사회문화', '통합사회', '사회'
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

/* ============================================================
   영역 & 과목 정의
   ============================================================ */
const areas = [
    { k: 'kor',  n: '국어 영역', hasChoice: true },
    { k: 'math', n: '수학 영역', hasChoice: true },
    { k: 'eng',  n: '영어 영역', hasChoice: false },
    { k: 'hist', n: '한국사',     hasChoice: false },
    { k: 'inq1', n: '탐구영역1',  hasChoice: true },
    { k: 'inq2', n: '탐구영역2',  hasChoice: true },
];

const topSubjects = [
    { k: 'kor',  n: '국어' },
    { k: 'math', n: '수학' },
    { k: 'eng',  n: '영어' },
    { k: 'hist', n: '한국사' },
    { k: 'inq1', n: '탐구①' },
    { k: 'inq2', n: '탐구②' },
];

const radarPointColors = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6'];

/* ============================================================
   초기화
   ============================================================ */
document.addEventListener('DOMContentLoaded', initializeEventListeners);

function initializeEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const analyzeBtn = document.getElementById('analyzeBtn');
    const tabBtns = document.querySelectorAll('.tab-btn');
    const uploadSection = document.querySelector('.upload-section');
    const fileLabel = document.querySelector('.file-input-label');

    /* 파일 선택 */
    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const nf = Array.from(e.target.files);
            if (nf.length > 0) {
                addFiles(nf);
                if (analyzeBtn) analyzeBtn.disabled = false;
            }
            fileInput.value = '';
        });
    }

    /* 분석 버튼 */
    if (analyzeBtn) analyzeBtn.addEventListener('click', analyzeFiles);

    /* 탭 전환 */
    tabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            switchTab(e.target.closest('.tab-btn').dataset.tab);
        });
    });

    /* 드래그 앤 드롭 */
    if (uploadSection) {
        const prevent = (ev) => { ev.preventDefault(); ev.stopPropagation(); };
        const setDrag = (on) => { if (fileLabel) fileLabel.classList.toggle('dragover', on); };

        ['dragover', 'drop'].forEach(evt => window.addEventListener(evt, prevent));

        ['dragenter', 'dragover'].forEach(evt =>
            uploadSection.addEventListener(evt, (ev) => { prevent(ev); setDrag(true); })
        );
        ['dragleave', 'dragend'].forEach(evt =>
            uploadSection.addEventListener(evt, (ev) => { prevent(ev); setDrag(false); })
        );

        uploadSection.addEventListener('drop', (ev) => {
            prevent(ev);
            setDrag(false);
            const dr = Array.from(ev.dataTransfer?.files || []);
            const files = dr.filter(f => /\.(xlsx|xls|csv|xlsm)$/i.test(f.name));
            if (files.length > 0) {
                addFiles(files);
                if (analyzeBtn) analyzeBtn.disabled = false;
            }
        });
    }

    /* 셀렉터 이벤트 */
    document.getElementById('gradeSelectTotal')?.addEventListener('change', onGradeChangeTotal);
    document.getElementById('gradeSelectClass')?.addEventListener('change', onGradeChangeClass);
    document.getElementById('gradeSelectIndiv')?.addEventListener('change', onGradeChangeIndiv);
    document.getElementById('gradeSelectTop')?.addEventListener('change', onGradeChangeTop);

    document.getElementById('examSelectTotal')?.addEventListener('change', renderOverall);
    document.getElementById('examSelectClass')?.addEventListener('change', renderClass);
    document.getElementById('classSelect')?.addEventListener('change', renderClass);

    document.getElementById('indivClassSelect')?.addEventListener('change', updateIndivList);
    document.getElementById('indivStudentSelect')?.addEventListener('change', renderIndividual);
    initStudentSearch();
    document.getElementById('indivExamSelect')?.addEventListener('change', renderIndividual);

    document.getElementById('pdfStudentBtn')?.addEventListener('click', generateStudentPDF);
    document.getElementById('pdfClassBtn')?.addEventListener('click', generateClassPDF);

    document.getElementById('examSelectTop')?.addEventListener('change', renderTop);
    document.getElementById('examSelectTopPrev')?.addEventListener('change', renderTop);

    /* 상위 N명 슬라이더 */
    const topNSlider = document.getElementById('topNSlider');
    if (topNSlider) {
        topNSlider.addEventListener('input', () => {
            state.topN = parseInt(topNSlider.value);
            document.getElementById('topNValue').textContent = state.topN + '명';
            renderTop();
        });
    }

    /* 리사이즈 */
    let resizeTimer = null;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            state.subjectCharts.forEach(c => {
                if (c && !c.destroyed && typeof c.resize === 'function') c.resize();
            });
        }, 150);
    });

    /* HTML 저장 */
    const saveHtmlBtn = document.getElementById('saveHtmlBtn');
    if (saveHtmlBtn) {
        saveHtmlBtn.replaceWith(saveHtmlBtn.cloneNode(true));
        document.getElementById('saveHtmlBtn').addEventListener('click', saveHtmlFile);
    }
}

/* ============================================================
   헬퍼 — 학년/시험/반 셀렉터
   ============================================================ */
function getExamsForGrade(g) {
    return state.gradeData[g] || [];
}

function onGradeChangeTotal() {
    const g = parseInt(document.getElementById('gradeSelectTotal').value);
    state.currentGradeTotal = g;
    updateExamSelector('examSelectTotal', g);
    renderOverall();
}

function onGradeChangeClass() {
    const g = parseInt(document.getElementById('gradeSelectClass').value);
    state.currentGradeClass = g;
    updateExamSelector('examSelectClass', g);
    updateClassSelector('classSelect', g);
    renderClass();
}

function onGradeChangeIndiv() {
    const g = parseInt(document.getElementById('gradeSelectIndiv').value);
    state.currentGradeIndiv = g;
    updateExamSelector('indivExamSelect', g);
    updateClassSelector('indivClassSelect', g);
    updateIndivList();
}

function onGradeChangeTop() {
    const g = parseInt(document.getElementById('gradeSelectTop').value);
    state.currentGradeTop = g;
    updateTopExamSelectors(g);
    renderTop();
}

function updateExamSelector(id, grade) {
    const el = document.getElementById(id);
    if (!el) return;
    const exams = getExamsForGrade(grade);
    el.innerHTML = exams.map((e, i) =>
        `<option value="${i}">${e.name}</option>`
    ).join('');
}

function updateTopExamSelectors(grade) {
    const exams = getExamsForGrade(grade);
    const s1 = document.getElementById('examSelectTop');
    const s2 = document.getElementById('examSelectTopPrev');
    if (!s1 || !s2) return;

    s1.innerHTML = exams.map((e, i) =>
        `<option value="${i}">${e.name}</option>`
    ).join('');

    let po = `<option value="-1">(비교 없음)</option>`;
    exams.forEach((e, i) => {
        po += `<option value="${i}">${e.name}</option>`;
    });
    s2.innerHTML = po;

    if (exams.length >= 2) s2.value = '1';
    else s2.value = '-1';
}

function updateClassSelector(id, grade) {
    const el = document.getElementById(id);
    if (!el) return;
    const exams = getExamsForGrade(grade);
    if (!exams.length) { el.innerHTML = ''; return; }

    const cs = new Set();
    exams.forEach(ex => ex.students.forEach(s => cs.add(s.info.class)));
    const cls = [...cs].sort((a, b) => a - b);

    el.innerHTML = cls.map(c => `<option value="${c}">${c}반</option>`).join('');
    el.value = cls[0];
}

function updateGradeSelectors() {
    const gs = state.availableGrades;
    const opts = gs.map(g => `<option value="${g}">${g}학년</option>`).join('');

    ['gradeSelectTotal', 'gradeSelectClass', 'gradeSelectIndiv', 'gradeSelectTop']
        .forEach(id => {
            const el = document.getElementById(id);
            if (el) el.innerHTML = opts;
        });

    state.currentGradeTotal = gs[0];
    state.currentGradeClass = gs[0];
    state.currentGradeIndiv = gs[0];
    state.currentGradeTop = gs[0];
}

/* ============================================================
   파일 관리
   ============================================================ */
function addFiles(nf) {
    nf.forEach(f => {
        if (!state.uploadedFiles.some(x => x.name === f.name && x.size === f.size))
            state.uploadedFiles.push(f);
    });
    renderFileList();
}

window.removeFile = function(i) {
    state.uploadedFiles.splice(i, 1);
    renderFileList();
    const b = document.getElementById('analyzeBtn');
    if (b) b.disabled = state.uploadedFiles.length === 0;
};

window.clearAllFiles = function() {
    state.uploadedFiles = [];
    renderFileList();
    const b = document.getElementById('analyzeBtn');
    if (b) b.disabled = true;
};

function renderFileList() {
    const fl = document.getElementById('fileList');
    if (!fl) return;

    if (!state.uploadedFiles.length) {
        fl.style.display = 'none';
        fl.innerHTML = '';
        return;
    }

    fl.style.display = 'block';
    fl.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
            <h4 style="margin:0;">
                <i class="fas fa-file-alt"></i>
                업로드된 파일 (${state.uploadedFiles.length}개)
            </h4>
            <button onclick="clearAllFiles()"
                    style="background:#e74c3c;color:#fff;border:none;padding:4px 12px;
                           border-radius:4px;cursor:pointer;font-size:0.8rem;">
                <i class="fas fa-trash-alt"></i> 전체 삭제
            </button>
        </div>
        <ul style="list-style:none;padding:0;margin:0;">
            ${state.uploadedFiles.map((f, i) => `
                <li style="display:flex;justify-content:space-between;align-items:center;
                           padding:6px 10px;margin-bottom:4px;
                           background:rgba(0,0,0,0.03);border-radius:6px;font-size:0.88rem;">
                    <span>
                        <i class="fas fa-file-excel" style="color:#27ae60;margin-right:6px;"></i>
                        ${f.name}
                        <span style="color:#999;font-size:0.75rem;margin-left:8px;">
                            (${(f.size / 1024).toFixed(1)} KB)
                        </span>
                    </span>
                    <button onclick="removeFile(${i})"
                            style="background:none;border:none;color:#e74c3c;
                                   cursor:pointer;font-size:1rem;padding:2px 6px;"
                            title="삭제">
                        <i class="fas fa-times-circle"></i>
                    </button>
                </li>
            `).join('')}
        </ul>`;
}

function showLoading(t = '분석 중...') {
    const lt = document.getElementById('loadingText');
    if (lt) lt.textContent = t;
    document.getElementById('loading').style.display = 'flex';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

/* ============================================================
   분석 (파일 파싱 → state 구성)
   ============================================================ */
async function analyzeFiles() {
    const files = state.uploadedFiles;
    if (!files.length) return alert('파일을 선택해주세요.');

    showLoading();

    try {
        const promises = files.map(file => new Promise(resolve => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                try {
                    const wb = XLSX.read(evt.target.result, { type: 'array' });

                    let tsn = wb.SheetNames.find(name => {
                        const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
                        for (let i = 0; i < Math.min(20, json.length); i++) {
                            const rs = (json[i] || []).join(' ');
                            if (rs.includes('이름') && (rs.includes('국어') || rs.includes('수학')))
                                return true;
                        }
                        return false;
                    }) || wb.SheetNames[0];

                    const jd = XLSX.utils.sheet_to_json(wb.Sheets[tsn], { header: 1 });
                    resolve(parseExcel(jd, file.name));
                } catch (e) {
                    console.error(e);
                    resolve(null);
                }
            };
            reader.readAsArrayBuffer(file);
        }));

        const results = await Promise.all(promises);
        const vr = results.filter(r => r && r.students.length > 0);

        if (!vr.length) {
            hideLoading();
            return alert('데이터를 찾을 수 없습니다.');
        }

        /* state.gradeData 구성 */
        state.gradeData = {};
        vr.forEach(exam => {
            const gs = new Set(exam.students.map(s => s.info.grade));
            gs.forEach(grade => {
                if (!state.gradeData[grade]) state.gradeData[grade] = [];
                const gStudents = exam.students.filter(s => s.info.grade === grade);
                gStudents.sort((a, b) => b.totalRaw - a.totalRaw);
                gStudents.forEach((s, idx) => { s.totalRank = idx + 1; });
                state.gradeData[grade].push({ name: exam.name, students: gStudents });
            });
        });

        Object.keys(state.gradeData).forEach(g => {
            state.gradeData[g].sort((a, b) =>
                b.name.localeCompare(a.name, undefined, { numeric: true })
            );
        });

        state.availableGrades = Object.keys(state.gradeData)
            .map(Number)
            .sort((a, b) => a - b);

        if (state.availableGrades.length) {
            hideLoading();
            document.getElementById('results').style.display = 'block';
            document.getElementById('saveHtmlBtn').style.display = 'inline-flex';
            updateLastUpdated();
            initSelectors();
        } else {
            hideLoading();
            alert('데이터를 찾을 수 없습니다.');
        }
    } catch (e) {
        hideLoading();
        alert('파일 분석 중 오류: ' + e.message);
    }
}

/* ============================================================
   엑셀 파싱
   ============================================================ */
function parseExcel(rows, fname) {
    let startRow = -1;
    for (let i = 0; i < rows.length; i++) {
        if (!rows[i]) continue;
        const rs = rows[i].map(c => String(c).replace(/\s/g, '')).join(',');
        if (rs.includes('이름') && rs.includes('번호')) {
            startRow = i;
            break;
        }
    }
    if (startRow === -1) return null;

    const students = [];
    const v = (r, i) => Number(r[i]) || 0;
    const g = (r, i) => { const v = Number(r[i]); return (v > 0 && v < 10) ? v : 9; };
    const s = (r, i) =>
        (r[i] && String(r[i]).trim() !== '' && String(r[i]).trim() !== 'nan')
            ? String(r[i]).trim()
            : '';

    for (let i = startRow + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || !r[4]) continue;

        const st = {
            info: {
                grade: parseInt(r[1]),
                class: parseInt(r[2]),
                no: parseInt(r[3]),
                name: r[4]
            },
            hist: { raw: v(r, 5), grd: g(r, 6), std: 0, pct: 0, name: '' },
            kor:  { name: s(r, 7),  raw: v(r, 8),  std: v(r, 9),  pct: v(r, 10), grd: g(r, 11) },
            math: { name: s(r, 12), raw: v(r, 13), std: v(r, 14), pct: v(r, 15), grd: g(r, 16) },
            eng:  { raw: v(r, 17),  grd: g(r, 18), std: 0, pct: 0, name: '' },
            inq1: { name: s(r, 19), raw: v(r, 20), std: v(r, 21), pct: v(r, 22), grd: g(r, 23) },
            inq2: { name: s(r, 24), raw: v(r, 25), std: v(r, 26), pct: v(r, 27), grd: g(r, 28) },
            uid: `${parseInt(r[2])}-${parseInt(r[3])}-${r[4]}`
        };

        st.totalRaw = st.kor.raw + st.math.raw + st.eng.raw
                     + st.inq1.raw + st.inq2.raw + st.hist.raw;
        st.totalStd = st.kor.std + st.math.std + st.inq1.std + st.inq2.std;
        st.totalPct = parseFloat(
            (st.kor.pct + st.math.pct + st.inq1.pct + st.inq2.pct).toFixed(2)
        );

        students.push(st);
    }

    students.sort((a, b) => b.totalRaw - a.totalRaw);
    students.forEach((s, i) => { s.totalRank = i + 1; });

    return { name: fname.replace(/\.[^/.]+$/, ""), students };
}

/* ============================================================
   셀렉터 초기화 & 탭 전환
   ============================================================ */
function initSelectors() {
    if (!state.availableGrades.length) return;

    updateGradeSelectors();
    const dg = state.availableGrades[0];

    state.currentGradeTotal = dg;
    document.getElementById('gradeSelectTotal').value = dg;
    updateExamSelector('examSelectTotal', dg);

    state.currentGradeClass = dg;
    document.getElementById('gradeSelectClass').value = dg;
    updateExamSelector('examSelectClass', dg);
    updateClassSelector('classSelect', dg);

    state.currentGradeIndiv = dg;
    document.getElementById('gradeSelectIndiv').value = dg;
    updateExamSelector('indivExamSelect', dg);
    updateClassSelector('indivClassSelect', dg);

    state.currentGradeTop = dg;
    document.getElementById('gradeSelectTop').value = dg;
    updateTopExamSelectors(dg);

    switchTab('overall');
    renderOverall();
    renderClass();
    updateIndivList();
    renderTop();
}

function switchTab(t) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById(t + '-tab').classList.add('active');
    document.querySelector(`.tab-btn[data-tab="${t}"]`).classList.add('active');

    if (t === 'overall' && state.charts.bubble) state.charts.bubble.resize();
    if (t === 'top') renderTop();
}

window.goToIndividual = function(uid, grade) {
    state.currentGradeIndiv = grade;
    const gs = document.getElementById('gradeSelectIndiv');
    if (gs) gs.value = grade;
    updateExamSelector('indivExamSelect', grade);

    const cls = parseInt(uid.split('-')[0]);
    const ics = document.getElementById('indivClassSelect');
    if (ics) ics.value = cls;
    updateIndivList();

    setTimeout(() => {
        const ss = document.getElementById('indivStudentSelect');
        if (ss) { ss.value = uid; renderIndividual(); }
        switchTab('individual');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }, 50);
};

/* ============================================================
   과목 카드 유틸리티
   ============================================================ */
function getChoiceGroups(students, areaKey) {
    const groups = {};
    students.forEach(s => {
        const cn = s[areaKey].name || '(미분류)';
        if (!groups[cn]) groups[cn] = [];
        groups[cn].push(s);
    });
    return groups;
}

function buildSubjectCardHTML(title, studentsInGroup, areaKey, metric) {
    const isAbs = (areaKey === 'eng' || areaKey === 'hist');
    const scores = studentsInGroup
        .map(s => isAbs ? s[areaKey].raw : (s[areaKey][metric] || 0))
        .filter(v => v > 0);

    const avg = scores.length
        ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1)
        : '-';
    const std = scores.length
        ? Math.sqrt(
            scores.reduce((a, b) => a + Math.pow(b - parseFloat(avg), 2), 0) / scores.length
          ).toFixed(1)
        : '-';
    const mx = scores.length ? Math.max(...scores).toFixed(1) : '-';

    const counts = Array(9).fill(0);
    studentsInGroup.forEach(s => {
        if (s[areaKey].grd >= 1 && s[areaKey].grd <= 9)
            counts[s[areaKey].grd - 1]++;
    });

    return buildCardInnerHTML(title, scores.length, avg, std, mx, counts);
}

function buildEntryCardHTML(title, entries, metric) {
    const scores = entries.map(e => {
        if (metric === 'raw') return e.raw;
        if (metric === 'std') return e.std;
        return e.pct;
    }).filter(v => v > 0);

    const avg = scores.length
        ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1)
        : '-';
    const std = scores.length
        ? Math.sqrt(
            scores.reduce((a, b) => a + Math.pow(b - parseFloat(avg), 2), 0) / scores.length
          ).toFixed(1)
        : '-';
    const mx = scores.length ? Math.max(...scores).toFixed(1) : '-';

    const counts = Array(9).fill(0);
    entries.forEach(e => {
        if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++;
    });

    return buildCardInnerHTML(title, entries.length, avg, std, mx, counts);
}

function buildCardInnerHTML(title, count, avg, std, mx, counts) {
    return `
        <div class="subject-card">
            <div class="subject-card-header">
                <h4>${title}</h4>
                <span class="count-badge">응시 ${count}명</span>
            </div>
            <div class="subject-stats">
                <div class="stat-item">
                    <div class="stat-label">평균</div>
                    <div class="stat-value-large">${avg}</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">표준편차</div>
                    <div class="stat-value-large">${std}</div>
                </div>
                <div class="stat-item">
                    <div class="stat-label">최고점</div>
                    <div class="stat-value-large">${mx}</div>
                </div>
            </div>
            <div class="grade-distribution">
                ${counts.map((c, i) => {
                    const total = counts.reduce((a, b) => a + b, 0);
                    const pct = total ? ((c / total) * 100).toFixed(1) : 0;
                    return `
                        <div class="grade-bar-item">
                            <div class="grade-label">${i + 1}등급</div>
                            <div class="grade-bar-container">
                                <div class="grade-bar-fill g-${i + 1}"
                                     style="width:${pct}%;"></div>
                            </div>
                            <div class="grade-count">${c}명</div>
                            <div class="grade-percentage">${pct}%</div>
                        </div>`;
                }).join('')}
            </div>
        </div>`;
}

function getInquiryMergedGroups(students) {
    const ae = [];
    students.forEach(s => {
        if (s.inq1.name) ae.push({ ...s.inq1, student: s, source: 'inq1' });
        if (s.inq2.name) ae.push({ ...s.inq2, student: s, source: 'inq2' });
    });

    const sg = {}, scg = {}, ug = {};
    ae.forEach(e => {
        const cat = classifyInquiry(e.name);
        const t = cat === 'social' ? sg : (cat === 'science' ? scg : ug);
        if (!t[e.name]) t[e.name] = [];
        t[e.name].push(e);
    });

    return { socialGroup: sg, scienceGroup: scg, unknownGroup: ug, allEntries: ae };
}

function buildInquiryAreaHTML(students, metric) {
    const { socialGroup: sg, scienceGroup: scg, unknownGroup: ug }
        = getInquiryMergedGroups(students);
    let html = '';

    const sn = Object.keys(sg).sort();
    const scn = Object.keys(scg).sort();
    const un = Object.keys(ug).sort();

    if (sn.length > 0) {
        const ae = sn.flatMap(n => sg[n]);
        html += `
            <div class="area-section">
                <h4 class="area-title">
                    <i class="fas fa-layer-group"></i> 사회탐구 영역
                </h4>
                <div class="subject-grid subject-grid-fixed">`;
        if (sn.length > 1)
            html += buildEntryCardHTML('사회탐구 전체', ae, metric);
        sn.forEach(n => { html += buildEntryCardHTML(n, sg[n], metric); });
        html += `</div></div>`;
    }

    if (scn.length > 0) {
        const ae = scn.flatMap(n => scg[n]);
        html += `
            <div class="area-section">
                <h4 class="area-title">
                    <i class="fas fa-layer-group"></i> 과학탐구 영역
                </h4>
                <div class="subject-grid subject-grid-fixed">`;
        if (scn.length > 1)
            html += buildEntryCardHTML('과학탐구 전체', ae, metric);
        scn.forEach(n => { html += buildEntryCardHTML(n, scg[n], metric); });
        html += `</div></div>`;
    }

    if (un.length > 0) {
        html += `
            <div class="area-section">
                <h4 class="area-title">
                    <i class="fas fa-layer-group"></i> 탐구 영역 (기타)
                </h4>
                <div class="subject-grid subject-grid-fixed">`;
        un.forEach(n => { html += buildEntryCardHTML(n, ug[n], metric); });
        html += `</div></div>`;
    }

    return html;
}

function checkSimpleMode(students) {
    for (const key of ['kor', 'math']) {
        const ns = new Set();
        students.forEach(s => {
            const n = s[key].name;
            if (n && n.trim() !== '' && n !== '(미분류)') ns.add(n.trim());
        });
        if (ns.size > 1) return false;
    }

    const { socialGroup: sg, scienceGroup: scg, unknownGroup: ug }
        = getInquiryMergedGroups(students);
    if (Object.keys(sg).length > 1 || Object.keys(scg).length > 1 || Object.keys(ug).length > 0)
        return false;

    return true;
}

function buildSingleAreaHTML(students, areaKey, areaName, hasChoice, metric) {
    if (!hasChoice)
        return `<div class="subject-grid subject-grid-fixed">
                    ${buildSubjectCardHTML(areaName, students, areaKey, metric)}
                </div>`;

    const groups = getChoiceGroups(students, areaKey);
    const sn = Object.keys(groups)
        .filter(n => n !== '(미분류)' || groups[n].length > 0)
        .sort();
    const rn = sn.filter(n => n !== '(미분류)');

    if (rn.length <= 1)
        return `<div class="subject-grid subject-grid-fixed">
                    ${buildSubjectCardHTML(rn[0] || areaName, students, areaKey, metric)}
                </div>`;

    let html = `<div class="subject-grid subject-grid-fixed">`;
    html += buildSubjectCardHTML(`${areaName} 전체`, students, areaKey, metric);
    sn.forEach(cn => { html += buildSubjectCardHTML(cn, groups[cn], areaKey, metric); });
    html += `</div>`;
    return html;
}

/* ============================================================
   전체통계 탭
   ============================================================ */
window.changeMetric = function(m) {
    state.metric = m;
    document.querySelectorAll('#overall-tab .opt-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-' + m).classList.add('active');
    renderOverall();
};

function renderOverall() {
    const grade = state.currentGradeTotal || state.availableGrades[0];
    const exams = getExamsForGrade(grade);
    const es = document.getElementById('examSelectTotal');
    if (!es || !exams.length) return;

    const ei = parseInt(es.value) || 0;
    if (!exams[ei]) return;

    const students = exams[ei].students;
    const metric = state.metric;
    const classes = [...new Set(students.map(s => s.info.class))].sort((a, b) => a - b);
    const maxClass = Math.max(...classes) || 12;

    /* ── ★ 기준별 총점 계산 (한국사 제외) & 정렬 ── */
    const getTotalForMetric = (s) => {
        if (metric === 'raw')
            return s.kor.raw + s.math.raw + s.eng.raw + s.inq1.raw + s.inq2.raw;
        if (metric === 'std')
            return s.totalStd;  // 이미 한국사 미포함
        return s.totalPct;      // 이미 한국사 미포함
    };

    /* 정렬용 복사본 */
    const sorted = [...students].sort((a, b) =>
        getTotalForMetric(b) - getTotalForMetric(a)
    );
    sorted.forEach((s, i) => { s._metricRank = i + 1; });

    /* ── 반별 성적 분포 버블 차트 ── */
    const bd = [];
    classes.forEach(c => {
        const cls = sorted
            .filter(s => s.info.class == c)
            .sort((a, b) => getTotalForMetric(b) - getTotalForMetric(a));

        cls.forEach((s, idx) => {
            const ratio = idx / (cls.length - 1 || 1);
            const r = ratio < 0.5 ? Math.floor(255 * (ratio * 2)) : 255;
            const g = ratio < 0.5 ? 255 : Math.floor(255 * (2 - ratio * 2));
            let score = getTotalForMetric(s);
            bd.push({
                x: Number(c), y: score, r: 8,
                bg: `rgba(${r},${g},0,0.8)`,
                name: s.info.name
            });
        });
    });

    const csad = classes.map(c => {
        const cls = students.filter(s => s.info.class == c);
        const avg = cls.reduce((sum, s) => sum + getTotalForMetric(s), 0) / cls.length;
        return { x: Number(c), y: parseFloat(avg.toFixed(1)), r: 12, name: `${c}반 평균` };
    });

    if (state.charts.bubble) state.charts.bubble.destroy();
    state.charts.bubble = new Chart(document.getElementById('bubbleChart'), {
        type: 'bubble',
        data: {
            datasets: [
                {
                    label: '학생',
                    data: bd,
                    backgroundColor: bd.map(d => d.bg),
                    borderColor: 'transparent'
                },
                {
                    label: '반 평균',
                    data: csad,
                    backgroundColor: 'rgba(80,80,220,0.85)',
                    borderColor: 'rgba(50,50,180,1)',
                    borderWidth: 1.5
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    min: 0,
                    max: maxClass + 1,
                    ticks: {
                        stepSize: 1,
                        callback: v =>
                            (Number.isInteger(v) && v > 0 && v <= maxClass) ? v + "반" : ""
                    }
                },
                y: {
                    title: {
                        display: true,
                        text: metric === 'raw' ? '원점수 합 (한국사 제외)'
                            : (metric === 'std' ? '표준점수 합' : '백분위 합')
                    }
                }
            },
            plugins: {
                legend: {
                    display: true, position: 'top',
                    labels: { usePointStyle: true, pointStyle: 'circle', font: { size: 11 } }
                },
                datalabels: { display: false },
                tooltip: {
                    callbacks: {
                        label: c => c.datasetIndex === 1
                            ? `${c.raw.name}: ${c.raw.y}`
                            : `${c.raw.x}반 ${c.raw.name}: ${c.raw.y.toFixed(1)}`
                    }
                }
            }
        }
    });

    /* ── 반별 평균등급 분포 버블 차트 (기존과 동일) ── */
    const getAvgGradeVal = s =>
        (s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4;

    const gbd = [];
    classes.forEach(c => {
        const cls = students
            .filter(s => s.info.class == c)
            .sort((a, b) => getAvgGradeVal(a) - getAvgGradeVal(b));

        cls.forEach((s, idx) => {
            const avgGrd = getAvgGradeVal(s);
            const ratio = idx / (cls.length - 1 || 1);
            const r = ratio < 0.5 ? Math.floor(255 * (ratio * 2)) : 255;
            const g = ratio < 0.5 ? 255 : Math.floor(255 * (2 - ratio * 2));
            gbd.push({
                x: Number(c), y: parseFloat(avgGrd.toFixed(2)), r: 8,
                bg: `rgba(${r},${g},0,0.8)`,
                name: s.info.name
            });
        });
    });

    const cgad = classes.map(c => {
        const cls = students.filter(s => s.info.class == c);
        const avg = cls.reduce((sum, s) => sum + getAvgGradeVal(s), 0) / cls.length;
        return { x: Number(c), y: parseFloat(avg.toFixed(2)), r: 12, name: `${c}반 평균` };
    });

    if (state.charts.gradesBubble) state.charts.gradesBubble.destroy();
    state.charts.gradesBubble = new Chart(document.getElementById('gradesBubbleChart'), {
        type: 'bubble',
        data: {
            datasets: [
                {
                    label: '학생',
                    data: gbd,
                    backgroundColor: gbd.map(d => d.bg),
                    borderColor: 'transparent'
                },
                {
                    label: '반 평균',
                    data: cgad,
                    backgroundColor: 'rgba(80,80,220,0.85)',
                    borderColor: 'rgba(50,50,180,1)',
                    borderWidth: 1.5
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    min: 0,
                    max: maxClass + 1,
                    ticks: {
                        stepSize: 1,
                        callback: v =>
                            (Number.isInteger(v) && v > 0 && v <= maxClass) ? v + "반" : ""
                    }
                },
                y: {
                    reverse: true,
                    min: 0,
                    max: 9.5,
                    ticks: {
                        stepSize: 1,
                        callback: v =>
                            (Number.isInteger(v) && v >= 1 && v <= 9) ? v + "등급" : ""
                    },
                    title: { display: true, text: '평균등급' }
                }
            },
            plugins: {
                legend: {
                    display: true, position: 'top',
                    labels: { usePointStyle: true, pointStyle: 'circle', font: { size: 11 } }
                },
                datalabels: { display: false },
                tooltip: {
                    callbacks: {
                        label: c => c.datasetIndex === 1
                            ? `${c.raw.name}: ${c.raw.y}등급`
                            : `${c.raw.x}반 ${c.raw.name}: ${c.raw.y}등급`
                    }
                }
            }
        }
    });

    /* ── 과목별 종합 분석 ── */
    const container = document.getElementById('combinedStatsContainer');
    container.innerHTML = '';
    const isSimple = checkSimpleMode(students);

    if (isSimple) {
        const subs = [
            { k: 'kor',  n: '국어' },
            { k: 'math', n: '수학' },
            { k: 'eng',  n: '영어' },
            { k: 'hist', n: '한국사' },
            { k: 'inq1', n: '사회탐구' },
            { k: 'inq2', n: '과학탐구' }
        ];
        const wrapper = document.createElement('div');
        wrapper.className = 'subject-grid subject-grid-fixed';
        subs.forEach(sub => {
            wrapper.innerHTML += buildSubjectCardHTML(sub.n, students, sub.k, metric);
        });
        container.appendChild(wrapper);
    } else {
        ['kor', 'math', 'eng', 'hist'].forEach(key => {
            const area = areas.find(a => a.k === key);
            let ah = `
                <div class="area-section">
                    <h4 class="area-title">
                        <i class="fas fa-layer-group"></i> ${area.n}
                    </h4>`;
            ah += buildSingleAreaHTML(students, area.k, area.n, area.hasChoice, metric);
            ah += `</div>`;
            container.innerHTML += ah;
        });
        container.innerHTML += buildInquiryAreaHTML(students, metric);
    }

    /* ★ sorted (기준별 정렬된) 배열로 테이블 생성 */
    buildTotalTable(sorted, metric);
    initGradeFilter(students);
}

/* ── 전체 성적 일람표 ── */
function buildTotalTable(students, metric) {
    const thead = document.getElementById('totalTableHead');
    const tbody = document.getElementById('totalTableBody');

    let h1 = `<tr>
        <th rowspan="2">석차</th>
        <th rowspan="2">학번</th>
        <th rowspan="2">이름</th>`;
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

    const excludeLabel = metric === 'raw' ? '(한국사 제외)' : '(영어, 한국사 제외)';
    h1 += `<th rowspan="2" class="total-col">총점<br><span style="font-size:0.7rem;font-weight:400;">${excludeLabel}</span></th>
            <th rowspan="2" class="total-col">평균등급</th></tr>`;
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    /* ★ 한국사 제외 총점 계산 */
    const getTot = s => {
        if (metric === 'raw')
            return s.kor.raw + s.math.raw + s.eng.raw + s.inq1.raw + s.inq2.raw;
        if (metric === 'std')
            return s.totalStd;
        return s.totalPct;
    };

    const getAvgGrade = s =>
        ((s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4).toFixed(2);

    const grade = state.currentGradeTotal || state.availableGrades[0];
    tbody.innerHTML = '';

    /* ★ students는 이미 기준별 정렬된 상태 (_metricRank 부여됨) */
    students.slice(0, 500).forEach((s, idx) => {
        const rank = idx + 1;
        const total = getTot(s);
        const totalDisplay = typeof total === 'number'
            ? (Number.isInteger(total) ? total : total.toFixed(1))
            : total;

        let row = `<tr>
            <td style="font-weight:bold;color:var(--primary);">${rank}</td>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${grade})">
                ${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}
            </td>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${grade})"
                style="font-weight:bold;">
                ${s.info.name}
            </td>`;

        areas.forEach(area => {
            const isAbs = (area.k === 'eng' || area.k === 'hist');
            const subj = s[area.k];
            if (area.hasChoice) {
                row += `<td style="font-size:0.75rem;color:var(--text-secondary);">
                            ${subj.name || '-'}
                        </td>
                        <td>${isAbs ? subj.raw : (subj[metric] || '-')}</td>
                        <td class="g-${subj.grd}">${subj.grd}</td>`;
            } else {
                row += `<td>${subj.raw}</td>
                        <td class="g-${subj.grd}">${subj.grd}</td>`;
            }
        });

        row += `<td class="total-col">${totalDisplay}</td>
                <td class="total-col" style="font-weight:bold;color:var(--primary);">
                    ${getAvgGrade(s)}
                </td></tr>`;
        tbody.innerHTML += row;
    });
}

/* ============================================================
   등급 필터
   ============================================================ */
function initGradeFilter(students) {
    const group = document.getElementById('gradeFilterGroup');
    if (!group) return;

    const isSimple = checkSimpleMode(students);

    if (isSimple) {
        const fs = [
            { k: 'kor',  n: '국어' },
            { k: 'math', n: '수학' },
            { k: 'eng',  n: '영어' },
            { k: 'hist', n: '한국사' },
            { k: 'inq1', n: '사회탐구' },
            { k: 'inq2', n: '과학탐구' }
        ];

        const gc = {};
        fs.forEach(sub => {
            gc[sub.k] = Array(9).fill(0);
            students.forEach(s => {
                const g = s[sub.k].grd;
                if (g >= 1 && g <= 9) gc[sub.k][g - 1]++;
            });
        });

        group.innerHTML = fs.map(sub => `
            <div class="grade-filter-subject">
                <div class="grade-filter-subject-label">${sub.n}</div>
                <div class="grade-filter-btns">
                    ${Array.from({ length: 9 }, (_, i) => i + 1).map(g => `
                        <button class="grade-filter-btn g-btn-${sub.k}-${g}"
                                onclick="renderGradeFilter('${sub.k}',${g},'all')"
                                title="${g}등급: ${gc[sub.k][g - 1]}명">
                            ${g}등급
                            <span class="grade-filter-count">${gc[sub.k][g - 1]}</span>
                        </button>
                    `).join('')}
                </div>
            </div>
        `).join('');
    } else {
        let html = '';

        ['kor', 'math', 'eng', 'hist'].forEach(key => {
            const area = areas.find(a => a.k === key);
            html += `<div class="grade-filter-area-header">${area.n}</div>`;

            if (area.hasChoice) {
                const groups = getChoiceGroups(students, key);
                const sn = Object.keys(groups).sort();
                const rn = sn.filter(n => n !== '(미분류)');

                if (rn.length <= 1) {
                    const counts = Array(9).fill(0);
                    students.forEach(s => {
                        const g = s[key].grd;
                        if (g >= 1 && g <= 9) counts[g - 1]++;
                    });
                    html += buildFilterSubjectRow(rn[0] || area.n, key, 'all', counts);
                } else {
                    const ac = Array(9).fill(0);
                    students.forEach(s => {
                        const g = s[key].grd;
                        if (g >= 1 && g <= 9) ac[g - 1]++;
                    });
                    html += buildFilterSubjectRow(`${area.n} 전체`, key, 'all', ac);

                    sn.forEach(cn => {
                        const grp = groups[cn];
                        const counts = Array(9).fill(0);
                        grp.forEach(s => {
                            const g = s[key].grd;
                            if (g >= 1 && g <= 9) counts[g - 1]++;
                        });
                        html += buildFilterSubjectRow(`└ ${cn}`, key, cn, counts, true);
                    });
                }
            } else {
                const counts = Array(9).fill(0);
                students.forEach(s => {
                    const g = s[key].grd;
                    if (g >= 1 && g <= 9) counts[g - 1]++;
                });
                html += buildFilterSubjectRow(area.n, key, 'all', counts);
            }
        });

        /* 탐구 영역 */
        const { socialGroup: sg, scienceGroup: scg, unknownGroup: ug }
            = getInquiryMergedGroups(students);

        const sn = Object.keys(sg).sort();
        if (sn.length > 0) {
            html += `<div class="grade-filter-area-header">사회탐구 영역</div>`;
            if (sn.length > 1) {
                const ae = sn.flatMap(n => sg[n]);
                const ac = Array(9).fill(0);
                ae.forEach(e => { if (e.grd >= 1 && e.grd <= 9) ac[e.grd - 1]++; });
                html += buildFilterSubjectRow('사회탐구 전체', 'inq_social', 'all', ac);
            }
            sn.forEach(name => {
                const entries = sg[name];
                const counts = Array(9).fill(0);
                entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
                html += buildFilterSubjectRow(
                    sn.length > 1 ? `└ ${name}` : name,
                    'inq_social', name, counts, sn.length > 1
                );
            });
        }

        const scn = Object.keys(scg).sort();
        if (scn.length > 0) {
            html += `<div class="grade-filter-area-header">과학탐구 영역</div>`;
            if (scn.length > 1) {
                const ae = scn.flatMap(n => scg[n]);
                const ac = Array(9).fill(0);
                ae.forEach(e => { if (e.grd >= 1 && e.grd <= 9) ac[e.grd - 1]++; });
                html += buildFilterSubjectRow('과학탐구 전체', 'inq_science', 'all', ac);
            }
            scn.forEach(name => {
                const entries = scg[name];
                const counts = Array(9).fill(0);
                entries.forEach(e => { if (e.grd >= 1 && e.grd <= 9) counts[e.grd - 1]++; });
                html += buildFilterSubjectRow(
                    scn.length > 1 ? `└ ${name}` : name,
                    'inq_science', name, counts, scn.length > 1
                );
            });
        }

        const un = Object.keys(ug).sort();
        if (un.length > 0) {
            html += `<div class="grade-filter-area-header">탐구 영역 (기타)</div>`;
            un.forEach(name => {
                const entries = ug[name];
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
    const sc = String(filterChoice).replace(/'/g, "\\'");
    return `
        <div class="grade-filter-subject">
            <div class="grade-filter-subject-label"
                 ${indented ? 'style="padding-left:16px;"' : ''}>
                ${label}
            </div>
            <div class="grade-filter-btns">
                ${Array.from({ length: 9 }, (_, i) => i + 1).map(g => `
                    <button class="grade-filter-btn"
                            onclick="renderGradeFilter('${filterArea}',${g},'${sc}')"
                            title="${g}등급: ${counts[g - 1]}명">
                        ${g}등급
                        <span class="grade-filter-count">${counts[g - 1]}</span>
                    </button>
                `).join('')}
            </div>
        </div>`;
}

window.renderGradeFilter = function(filterArea, gradeLevel, choiceFilter) {
    /* 버튼 토글 */
    let cb = null;
    const sb = document.querySelector(`.g-btn-${filterArea}-${gradeLevel}`);
    if (sb) {
        cb = sb;
    } else {
        document.querySelectorAll('.grade-filter-btn').forEach(btn => {
            const oc = btn.getAttribute('onclick') || '';
            if (oc.includes(`'${filterArea}', ${gradeLevel}`) &&
                oc.includes(`'${choiceFilter}'`))
                cb = btn;
        });
    }

    if (cb && cb.classList.contains('active'))
        cb.classList.remove('active');
    else if (cb)
        cb.classList.add('active');

    /* 활성 필터 수집 */
    const activeButtons = document.querySelectorAll('.grade-filter-btn.active');
    if (!activeButtons.length) {
        document.getElementById('gradeFilterResult').style.display = 'none';
        document.getElementById('gradeFilterEmpty').style.display = 'none';
        return;
    }

    const activeFilters = [];
    activeButtons.forEach(btn => {
        const oc = btn.getAttribute('onclick') || '';
        const match = oc.match(
            /renderGradeFilter\x28'([^']+)',\s*(\d+),\s*'([^']+)'\x29/
        );
        if (match)
            activeFilters.push({
                area: match[1],
                grade: parseInt(match[2]),
                choice: match[3]
            });
    });

    if (!activeFilters.length) {
        document.getElementById('gradeFilterResult').style.display = 'none';
        document.getElementById('gradeFilterEmpty').style.display = 'none';
        return;
    }

    /* 학생 데이터 */
    const cg = state.currentGradeTotal || state.availableGrades[0];
    const exams = getExamsForGrade(cg);
    const es = document.getElementById('examSelectTotal');
    const ei = parseInt(es.value) || 0;
    const students = exams[ei]?.students || [];

    const allRows = [];
    const summaryParts = [];

    activeFilters.forEach(filter => {
        const isInq = filter.area.startsWith('inq_');

        if (isInq) {
            const { socialGroup: sg, scienceGroup: scg, unknownGroup: ug }
                = getInquiryMergedGroups(students);

            let tg = {}, cn = '';
            if (filter.area === 'inq_social')  { tg = sg;  cn = '사회탐구'; }
            else if (filter.area === 'inq_science') { tg = scg; cn = '과학탐구'; }
            else { tg = ug; cn = '탐구(기타)'; }

            let entries = [];
            if (filter.choice === 'all')
                entries = Object.values(tg).flat();
            else
                entries = tg[filter.choice] || [];

            entries = entries.filter(e => e.grd === filter.grade);
            entries.sort((a, b) => b.raw - a.raw);

            const dl = filter.choice === 'all' ? cn : `${cn}-${filter.choice}`;
            if (entries.length > 0)
                summaryParts.push(
                    `<strong>${dl}</strong>
                     <span class="grade-badge g-${filter.grade}">${filter.grade}등급</span>
                     ${entries.length}명`
                );

            entries.forEach(e => {
                const s = e.student;
                const id = `${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}`;
                allRows.push({
                    type: 'inquiry',
                    id, name: s.info.name,
                    subject: e.name || '-',
                    raw: e.raw, grd: e.grd,
                    std: e.std || '-', pct: e.pct || '-',
                    filterLabel: dl
                });
            });
        } else {
            const area = areas.find(a => a.k === filter.area);
            const isAbs = (filter.area === 'eng' || filter.area === 'hist');

            let filtered = students.filter(s => s[filter.area].grd === filter.grade);
            if (filter.choice !== 'all')
                filtered = filtered.filter(s =>
                    (s[filter.area].name || '(미분류)') === filter.choice
                );

            filtered.sort((a, b) => {
                const diff = b[filter.area].raw - a[filter.area].raw;
                return diff !== 0
                    ? diff
                    : (a.info.class * 100 + a.info.no) - (b.info.class * 100 + b.info.no);
            });

            const dl = filter.choice === 'all' ? area.n : `${area.n}-${filter.choice}`;
            if (filtered.length > 0)
                summaryParts.push(
                    `<strong>${dl}</strong>
                     <span class="grade-badge g-${filter.grade}">${filter.grade}등급</span>
                     ${filtered.length}명`
                );

            filtered.forEach(s => {
                const sub = s[filter.area];
                const id = `${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}`;
                allRows.push({
                    type: 'normal',
                    id, name: s.info.name,
                    subject: area.hasChoice ? (sub.name || '-') : '',
                    raw: sub.raw, grd: sub.grd,
                    std: isAbs ? null : (sub.std || '-'),
                    pct: isAbs ? null : (sub.pct || '-'),
                    filterLabel: dl
                });
            });
        }
    });

    /* 결과 렌더링 */
    const rEl = document.getElementById('gradeFilterResult');
    const eEl = document.getElementById('gradeFilterEmpty');
    const sEl = document.getElementById('gradeFilterSummary');
    const thEl = document.getElementById('gradeFilterThead');
    const tbEl = document.getElementById('gradeFilterTbody');

    if (!allRows.length) {
        rEl.style.display = 'none';
        eEl.style.display = 'flex';
        return;
    }

    eEl.style.display = 'none';
    rEl.style.display = 'block';

    sEl.innerHTML = `
        <span class="grade-filter-summary-text">
            ${summaryParts.join(' &nbsp;|&nbsp; ')}
            &nbsp;&nbsp; 합계 <strong>${allRows.length}명</strong>
        </span>`;

    thEl.innerHTML = `<tr>
        <th>조회조건</th><th>학번</th><th>이름</th><th>선택과목</th>
        <th>원점수</th><th>등급</th><th>표준점수</th><th>백분위</th>
    </tr>`;

    tbEl.innerHTML = allRows.map(r => `
        <tr>
            <td style="font-size:0.75rem;color:var(--text-secondary);">${r.filterLabel}</td>
            <td>${r.id}</td>
            <td style="font-weight:bold;">${r.name}</td>
            <td style="font-size:0.8rem;color:var(--text-secondary);">${r.subject || '-'}</td>
            <td>${r.raw}</td>
            <td class="g-${r.grd}">${r.grd}</td>
            <td>${r.std !== null ? r.std : '-'}</td>
            <td>${r.pct !== null ? r.pct : '-'}</td>
        </tr>
    `).join('');
};

/* ── 등급 필터 엑셀 저장 ── */
window.exportGradeFilterToExcel = function() {
    const thead = document.getElementById('gradeFilterThead');
    const tbody = document.getElementById('gradeFilterTbody');
    if (!thead || !tbody || tbody.rows.length === 0)
        return alert('저장할 데이터가 없습니다.');

    const wb = XLSX.utils.book_new();
    const rows = [];

    const hr = [];
    thead.querySelectorAll('tr').forEach(tr => {
        tr.querySelectorAll('th').forEach(th => { hr.push(th.textContent.trim()); });
    });
    rows.push(hr);

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
    ws['!cols'] = rows[0].map((_, ci) => {
        let mx = 0;
        rows.forEach(r => {
            const cl = String(r[ci] || '').length;
            if (cl > mx) mx = cl;
        });
        return { wch: Math.max(mx + 2, 8) };
    });

    XLSX.utils.book_append_sheet(wb, ws, '등급구간조회');
    const now = new Date();
    XLSX.writeFile(wb,
        `등급구간_학생조회_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.xlsx`
    );
};

/* ── 전체 일람표 엑셀 저장 ── */
window.exportTotalTableToExcel = function() {
    const thead = document.getElementById('totalTableHead');
    const tbody = document.getElementById('totalTableBody');
    if (!thead || !tbody || tbody.rows.length === 0)
        return alert('저장할 데이터가 없습니다.');

    const wb = XLSX.utils.book_new();
    const rows = [];

    const headerRows = thead.querySelectorAll('tr');
    if (headerRows.length === 2) {
        const mh = [];
        const r1 = headerRows[0].querySelectorAll('th');
        const r2 = Array.from(headerRows[1].querySelectorAll('th'));
        let r2i = 0;

        r1.forEach(th => {
            const cs = parseInt(th.getAttribute('colspan')) || 1;
            const rs = parseInt(th.getAttribute('rowspan')) || 1;

            if (rs === 2) {
                mh.push(th.textContent.trim());
            } else if (cs > 1) {
                const pn = th.textContent.trim();
                for (let i = 0; i < cs && r2i < r2.length; i++) {
                    mh.push(`${pn}_${r2[r2i].textContent.trim()}`);
                    r2i++;
                }
            } else {
                mh.push(th.textContent.trim());
            }
        });
        rows.push(mh);
    } else {
        const hr = [];
        headerRows[0].querySelectorAll('th').forEach(th => {
            hr.push(th.textContent.trim());
        });
        rows.push(hr);
    }

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
    ws['!cols'] = rows[0].map((_, ci) => {
        let mx = 0;
        rows.forEach(r => {
            const cl = String(r[ci] || '').length;
            if (cl > mx) mx = cl;
        });
        return { wch: Math.max(mx + 2, 8) };
    });

    XLSX.utils.book_append_sheet(wb, ws, '성적일람표');
    const now = new Date();
    XLSX.writeFile(wb,
        `전체성적일람표_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.xlsx`
    );
};

/* ============================================================
   학급통계 탭
   ============================================================ */
window.changeClassMetric = function(m) {
    state.classMetric = m;
    document.querySelectorAll('#class-tab .opt-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-c-' + m).classList.add('active');
    renderClass();
};

window.sortClass = function(t) {
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
    const es = document.getElementById('examSelectClass');
    const cs = document.getElementById('classSelect');
    if (!es || !cs || !exams.length) return;

    const ei = parseInt(es.value) || 0;
    if (!exams[ei]) return;

    const sa = exams[ei].students;
    const cls = parseInt(cs.value);
    if (isNaN(cls)) return;

    const metric = state.classMetric;
    let students = sa.filter(s => s.info.class === cls);

    const getTot = s => {
        if (metric === 'raw')
            return s.kor.raw + s.math.raw + s.eng.raw + s.inq1.raw + s.inq2.raw;
        if (metric === 'std')
            return parseFloat(s.totalStd.toFixed(1));
        return parseFloat(s.totalPct.toFixed(1));
    };

    if (state.classSort === 'total')
        students.sort((a, b) => parseFloat(getTot(b)) - parseFloat(getTot(a)));
    else
        students.sort((a, b) => a.info.no - b.info.no);

    /* ── 과목별 종합 분석 ── */
    const container = document.getElementById('classStatsContainer');
    if (container) {
        container.innerHTML = '';
        const isSimple = checkSimpleMode(students);

        if (isSimple) {
            const subs = [
                { k: 'kor',  n: '국어' },
                { k: 'math', n: '수학' },
                { k: 'eng',  n: '영어' },
                { k: 'hist', n: '한국사' },
                { k: 'inq1', n: '사회탐구' },
                { k: 'inq2', n: '과학탐구' }
            ];
            const wrapper = document.createElement('div');
            wrapper.className = 'subject-grid subject-grid-fixed';
            subs.forEach(sub => {
                wrapper.innerHTML += buildSubjectCardHTML(sub.n, students, sub.k, metric);
            });
            container.appendChild(wrapper);
        } else {
            ['kor', 'math', 'eng', 'hist'].forEach(key => {
                const area = areas.find(a => a.k === key);
                let ah = `
                    <div class="area-section">
                        <h4 class="area-title">
                            <i class="fas fa-layer-group"></i> ${area.n}
                        </h4>`;
                ah += buildSingleAreaHTML(students, area.k, area.n, area.hasChoice, metric);
                ah += `</div>`;
                container.innerHTML += ah;
            });
            container.innerHTML += buildInquiryAreaHTML(students, metric);
        }
    }

    /* ── 학급 일람표 ── */
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

    const excludeLabel = metric === 'raw' ? '(한국사 제외)' : '(영어, 한국사 제외)';
    h1 += `<th rowspan="2" class="total-col">총점<br><span style="font-size:0.7rem;font-weight:400;">${excludeLabel}</span></th>
            <th rowspan="2" class="total-col">석차</th>
            <th rowspan="2" class="total-col">평균등급</th></tr>`;
    h2 += `</tr>`;
    thead.innerHTML = h1 + h2;

    const getAvgGrade = s =>
        ((s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4).toFixed(2);

    const gfc = state.currentGradeClass || state.availableGrades[0];
    tbody.innerHTML = '';

    students.forEach(s => {
        const rank = students.filter(st =>
            parseFloat(getTot(st)) > parseFloat(getTot(s))
        ).length + 1;

        let row = `<tr>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${gfc})">${s.info.no}</td>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${gfc})"
                style="font-weight:bold;">${s.info.name}</td>`;

        areas.forEach(area => {
            const isAbs = (area.k === 'eng' || area.k === 'hist');
            const subj = s[area.k];
            if (area.hasChoice) {
                row += `<td style="font-size:0.75rem;color:var(--text-secondary);">
                            ${subj.name || '-'}
                        </td>
                        <td>${isAbs ? subj.raw : (subj[metric] || '-')}</td>
                        <td class="g-${subj.grd}">${subj.grd}</td>`;
            } else {
                row += `<td>${subj.raw}</td>
                        <td class="g-${subj.grd}">${subj.grd}</td>`;
            }
        });

        row += `<td class="total-col">${getTot(s)}</td>
                <td class="total-col">${rank}</td>
                <td class="total-col" style="font-weight:bold;color:var(--primary);">
                    ${getAvgGrade(s)}
                </td></tr>`;
        tbody.innerHTML += row;
    });
}

/* ============================================================
   개인통계 탭
   ============================================================ */
function updateIndivList() {
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return;

    const exams = getExamsForGrade(grade);
    if (!exams.length) return;

    const cv = document.getElementById('indivClassSelect')?.value;
    if (!cv) return;

    const cls = parseInt(cv);
    if (isNaN(cls)) return;

    const sm = new Map();
    exams.forEach(ex => {
        ex.students
            .filter(s => s.info.class === cls)
            .forEach(s => { if (!sm.has(s.uid)) sm.set(s.uid, s); });
    });

    const list = Array.from(sm.values()).sort((a, b) => a.info.no - b.info.no);
    document.getElementById('indivStudentSelect').innerHTML = list.map(s =>
        `<option value="${s.uid}">${s.info.no}번 ${s.info.name}</option>`
    ).join('');

    if (list.length > 0) renderIndividual();
}

/* ── 학생 검색 ── */
function initStudentSearch() {
    const si = document.getElementById('studentSearchInput');
    const cb = document.getElementById('studentSearchClear');
    const dd = document.getElementById('studentSearchDropdown');
    if (!si) return;

    let hi = -1;

    si.addEventListener('input', () => {
        const q = si.value.trim();
        cb.style.display = q ? 'block' : 'none';
        if (!q.length) { dd.style.display = 'none'; hi = -1; return; }
        const r = searchStudents(q);
        renderSearchDropdown(r, q);
        hi = -1;
    });

    si.addEventListener('focus', () => {
        const q = si.value.trim();
        if (q.length > 0) {
            const r = searchStudents(q);
            renderSearchDropdown(r, q);
        }
    });

    si.addEventListener('keydown', (e) => {
        const items = dd.querySelectorAll('.student-search-item');
        if (!items.length || dd.style.display === 'none') return;

        if (e.key === 'ArrowDown') {
            e.preventDefault();
            hi = Math.min(hi + 1, items.length - 1);
            updateHighlight(items, hi);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            hi = Math.max(hi - 1, 0);
            updateHighlight(items, hi);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (hi >= 0 && hi < items.length) items[hi].click();
            else if (items.length > 0) items[0].click();
        } else if (e.key === 'Escape') {
            dd.style.display = 'none';
            hi = -1;
        }
    });

    cb.addEventListener('click', () => {
        si.value = '';
        cb.style.display = 'none';
        dd.style.display = 'none';
        hi = -1;
        si.focus();
    });

    document.addEventListener('click', (e) => {
        const w = document.getElementById('studentSearchWrapper');
        if (w && !w.contains(e.target)) {
            dd.style.display = 'none';
            hi = -1;
        }
    });
}

function searchStudents(query) {
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return [];

    const exams = getExamsForGrade(grade);
    if (!exams.length) return [];

    const sm = new Map();
    exams.forEach(ex => {
        ex.students.forEach(s => { if (!sm.has(s.uid)) sm.set(s.uid, s); });
    });

    const all = Array.from(sm.values());
    const q = query.toLowerCase();

    return all.filter(s => {
        const name = s.info.name.toLowerCase();
        const no = String(s.info.no);
        const cn = String(s.info.class);
        const fi = `${s.info.class}반 ${s.info.no}번`;
        return name.includes(q) || no === q || cn === q || fi.includes(q);
    }).sort((a, b) => {
        const am = a.info.name.toLowerCase().startsWith(q) ? 0 : 1;
        const bm = b.info.name.toLowerCase().startsWith(q) ? 0 : 1;
        if (am !== bm) return am - bm;
        if (a.info.class !== b.info.class) return a.info.class - b.info.class;
        return a.info.no - b.info.no;
    }).slice(0, 30);
}

function renderSearchDropdown(results, query) {
    const dd = document.getElementById('studentSearchDropdown');
    if (!dd) return;

    if (!results.length) {
        dd.innerHTML = `
            <div class="student-search-no-result">
                <i class="fas fa-search" style="margin-right:6px;"></i>
                '${query}'에 해당하는 학생이 없습니다
            </div>`;
        dd.style.display = 'block';
        return;
    }

    let html = `<div class="student-search-count">검색 결과: ${results.length}명</div>`;
    results.forEach(s => {
        const hn = highlightText(s.info.name, query);
        html += `
            <div class="student-search-item"
                 data-uid="${s.uid}" data-class="${s.info.class}">
                <span class="student-search-item-name">${hn}</span>
                <span class="student-search-item-info">${s.info.class}반 ${s.info.no}번</span>
            </div>`;
    });

    dd.innerHTML = html;
    dd.style.display = 'block';

    dd.querySelectorAll('.student-search-item').forEach(item => {
        item.addEventListener('click', () => {
            const uid = item.dataset.uid;
            const cls = item.dataset.class;

            const cs = document.getElementById('indivClassSelect');
            if (cs) { cs.value = cls; updateIndivList(); }

            setTimeout(() => {
                const ss = document.getElementById('indivStudentSelect');
                if (ss) { ss.value = uid; renderIndividual(); }

                const si = document.getElementById('studentSearchInput');
                const cb = document.getElementById('studentSearchClear');
                if (si) si.value = '';
                if (cb) cb.style.display = 'none';
                dd.style.display = 'none';
            }, 100);
        });
    });
}

function highlightText(t, q) {
    if (!q) return t;
    const re = new RegExp(
        `(${q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi'
    );
    return t.replace(re, '<mark>$1</mark>');
}

function updateHighlight(items, idx) {
    items.forEach((item, i) => {
        item.classList.toggle('highlighted', i === idx);
    });
    if (items[idx]) items[idx].scrollIntoView({ block: 'nearest' });
}

/* ── 개인통계 렌더링 ── */
function renderIndividual() {
    const sel = document.getElementById('indivStudentSelect');
    if (!sel || !sel.value) return;

    const uid = sel.value;
    const grade = state.currentGradeIndiv || state.availableGrades[0];
    if (!grade) return;

    const exams = getExamsForGrade(grade);
    if (!exams.length) return;

    /* 히스토리 수집 */
    const history = [];
    for (let i = exams.length - 1; i >= 0; i--) {
        const ex = exams[i];
        const s = ex.students.find(st => st.uid === uid);
        if (s) history.push({ name: ex.name, data: s });
    }
    if (!history.length) return;

    /* 선택 시험 */
    const sei = parseInt(document.getElementById('indivExamSelect')?.value) || 0;
    let cd = exams[sei]?.students.find(st => st.uid === uid);
    let sen = exams[sei]?.name;
    if (!cd) {
        cd = history[history.length - 1].data;
        sen = history[history.length - 1].name;
    }

    /* 프로필 카드 */
    document.getElementById('indivName').innerText = cd.info.name;
    document.getElementById('indivInfo').innerText =
        `${cd.info.grade}학년 ${cd.info.class}반 ${cd.info.no}번`;
    document.getElementById('latestExamName').innerText = `선택 시험: ${sen}`;
    document.getElementById('indivTotalRaw').innerText = cd.totalRaw;
    document.getElementById('indivTotalStd').innerText = cd.totalStd.toFixed(1);
    document.getElementById('indivTotalPct').innerText = cd.totalPct.toFixed(2);
    document.getElementById('indivRank').innerText = cd.totalRank;

    const iag = (cd.inq1.grd + cd.inq2.grd) / 2;
    const avgGrade = ((cd.kor.grd + cd.math.grd + cd.eng.grd + iag) / 4).toFixed(2);
    document.getElementById('indivAverageGrade').innerText = avgGrade;

    /* 총점 추이 차트 */
    drawChart('totalTrendChart', 'line', {
        labels: history.map(h => h.name),
        datasets: [{
            label: '총점(백분위합)',
            data: history.map(h => h.data.totalPct),
            borderColor: '#8B5A8D',
            backgroundColor: '#8B5A8D',
            tension: 0.3,
            borderWidth: 2,
            pointRadius: 4
        }]
    }, {
        scales: {
            y: {
                min: 0, max: 450,
                ticks: { stepSize: 50, callback: v => v === 450 ? '' : v }
            }
        },
        plugins: {
            datalabels: {
                display: true,
                color: '#8B5A8D',
                align: 'top',
                font: { weight: 'bold' },
                formatter: v => v.toFixed(1)
            }
        }
    });

    /* 레이더 차트 */
    const rl = ['국어', '수학', '영어', '탐1', '탐2'];
    const rg = [cd.kor.grd, cd.math.grd, cd.eng.grd, cd.inq1.grd, cd.inq2.grd];

    drawChart('radarChart', 'radar', {
        labels: rl,
        datasets: [{
            label: '등급',
            data: rg,
            backgroundColor: 'rgba(100,150,200,0.2)',
            borderColor: 'rgba(100,150,200,0.6)',
            borderWidth: 2,
            pointBackgroundColor: radarPointColors,
            pointBorderColor: radarPointColors,
            pointBorderWidth: 2,
            pointRadius: 6,
            pointHoverRadius: 8
        }]
    }, {
        layout: { padding: { top: 30, bottom: 30, left: 50, right: 50 } },
        scales: {
            r: {
                reverse: true, min: 1, max: 9,
                ticks: {
                    stepSize: 1, display: true,
                    font: { size: 10 }, color: '#999',
                    backdropColor: 'transparent'
                },
                grid: { circular: false, color: 'rgba(0,0,0,0.1)' },
                angleLines: { display: true, color: 'rgba(0,0,0,0.1)' },
                pointLabels: {
                    font: { size: 13, weight: 'bold', family: "'Pretendard',sans-serif" },
                    color: radarPointColors,
                    padding: 12
                }
            }
        },
        plugins: {
            legend: { display: false },
            datalabels: {
                display: true,
                backgroundColor: function(ctx) {
                    return ctx.dataset.data[ctx.dataIndex] <= 1.5
                        ? 'rgba(255,255,255,0.85)' : '#ffffff';
                },
                borderColor: function(ctx) { return radarPointColors[ctx.dataIndex]; },
                borderWidth: 2,
                color: function(ctx) { return radarPointColors[ctx.dataIndex]; },
                borderRadius: 4,
                padding: { top: 3, bottom: 3, left: 6, right: 6 },
                font: { weight: 'bold', size: 10 },
                formatter: v => v.toFixed(2) + '등급',
                anchor: 'center',
                align: function(ctx) {
                    const idx = ctx.dataIndex;
                    const count = ctx.dataset.data.length;
                    const ad = 90 - (360 / count) * idx;
                    return ad * Math.PI / 180 + Math.PI;
                },
                offset: function(ctx) {
                    const v = ctx.dataset.data[ctx.dataIndex];
                    const n = (v - 1) / 8;
                    return Math.round(4 + (1 - n) * 10);
                },
                clip: false
            }
        }
    });

    /* ── 과목 상세 카드 ── */
    const dc = document.getElementById('subjectDetailContainer');
    if (dc) {
        state.subjectCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
        state.subjectCharts = [];
        dc.innerHTML = '';

        areas.forEach(area => {
            const k = area.k;
            const isAbs = (k === 'eng' || k === 'hist');

            let dt = area.n;
            if (area.hasChoice && cd[k].name)
                dt = `${area.n} (${cd[k].name})`;

            let th = `<tr><th>구분</th>`
                + history.map(h => `<th>${h.name}</th>`).join('')
                + `</tr>`;

            let tc = '';
            if (area.hasChoice)
                tc = `<tr><td style="font-weight:600;color:var(--text-secondary);">선택과목</td>`
                    + history.map(h =>
                        `<td style="font-size:0.78rem;color:var(--text-secondary);">${h.data[k].name || '-'}</td>`
                      ).join('')
                    + `</tr>`;

            let tg = `<tr><td>등급</td>`
                + history.map(h => `<td class="g-${h.data[k].grd}">${h.data[k].grd}</td>`).join('')
                + `</tr>`;

            let tr = `<tr><td>원점수</td>`
                + history.map(h => `<td>${h.data[k].raw}</td>`).join('')
                + `</tr>`;

            let ts = `<tr><td>표준점수</td>`
                + history.map(h => `<td>${h.data[k].std || '-'}</td>`).join('')
                + `</tr>`;

            let tp = `<tr><td>백분위</td>`
                + history.map(h => `<td>${h.data[k].pct || '-'}</td>`).join('')
                + `</tr>`;

            const cid = `chart-${k}-${uid.replace(/[^a-zA-Z0-9]/g, '')}`;

            dc.innerHTML += `
                <div class="chart-card subject-detail-card" data-subject="${k}">
                    <h3><i class="fas fa-book"></i> ${dt} 성적 상세</h3>
                    <div class="subject-detail-grid">
                        <div class="chart-container" style="height:200px;width:100%;">
                            <canvas id="${cid}"></canvas>
                        </div>
                        <div class="table-wrapper">
                            <table class="data-table" style="font-size:0.8rem;min-width:100%;">
                                <thead>${th}</thead>
                                <tbody>
                                    ${tc}${tg}${tr}
                                    ${!isAbs ? ts : ''}
                                    ${!isAbs ? tp : ''}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>`;

            setTimeout(() => {
                const ctx = document.getElementById(cid);
                if (!ctx) return;

                const yv = history.map(h => isAbs ? h.data[k].grd : h.data[k].pct);
                const chart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: history.map(h => h.name),
                        datasets: [{
                            label: isAbs ? '등급' : '백분위',
                            data: yv,
                            borderColor: isAbs ? '#B8860B' : '#4A6B8A',
                            tension: 0.1,
                            borderWidth: 2,
                            pointRadius: 4
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false },
                            datalabels: {
                                display: true,
                                color: isAbs ? '#B8860B' : '#4A6B8A',
                                align: 'top',
                                font: { weight: 'bold' },
                                formatter: v => v > 100 ? '' : v
                            }
                        },
                        scales: {
                            y: {
                                reverse: isAbs,
                                min: isAbs ? 0 : 0,
                                max: isAbs ? 9 : 120,
                                ticks: {
                                    stepSize: isAbs ? 1 : 20,
                                    callback: function(value) {
                                        if (isAbs && value === 0) return '';
                                        return value;
                                    }
                                }
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

/* ============================================================
   상위권 분석 탭
   ============================================================ */
window.changeTopMetric = function(m) {
    state.topMetric = m;
    document.querySelectorAll('#top-tab .opt-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-t-' + m).classList.add('active');
    renderTop();
};

/* ── 점수 추출 헬퍼 (Metric 반영) ── */
function getSubjScore(s, subKey, metric) {
    const isAbs = (subKey === 'eng' || subKey === 'hist');
    if (isAbs) return s[subKey].raw; // 영어/한국사는 항상 원점수
    return s[subKey][metric] || 0;
}

function getTotalByMetric(s, metric) {
    if (metric === 'raw')
        return s.kor.raw + s.math.raw + s.eng.raw + s.inq1.raw + s.inq2.raw;
    if (metric === 'std')
        return s.totalStd;
    return s.totalPct;
}

function getMetricLabel(metric) {
    if (metric === 'raw') return '원점수';
    if (metric === 'std') return '표준점수';
    return '백분위';
}

/* ── 메인 렌더 ── */
function renderTop() {
    const grade = state.currentGradeTop || state.availableGrades[0];
    if (!grade) return;

    const exams = getExamsForGrade(grade);
    const curIdx = parseInt(document.getElementById('examSelectTop')?.value) || 0;
    const prevIdx = parseInt(document.getElementById('examSelectTopPrev')?.value);
    if (!exams[curIdx]) return;

    const curStudents = exams[curIdx].students;
    const prevStudents = (prevIdx >= 0 && exams[prevIdx])
        ? exams[prevIdx].students
        : null;
    const topN = state.topN;
    const metric = state.topMetric;

    /* 정렬 & 상위 N명 */
    const sorted = [...curStudents].sort((a, b) =>
        getTotalByMetric(b, metric) - getTotalByMetric(a, metric)
    );
    const topCur = sorted.slice(0, topN);
    topCur.forEach((s, i) => { s._topRank = i + 1; });

    /* 직전 시험 맵 */
    const prevMap = new Map();
    const prevRankMap = new Map();
    if (prevStudents) {
        const prevSorted = [...prevStudents].sort((a, b) =>
            getTotalByMetric(b, metric) - getTotalByMetric(a, metric)
        );
        prevSorted.forEach((s, i) => {
            prevMap.set(s.uid, s);
            prevRankMap.set(s.uid, i + 1);
        });
    }

    /* 분석 데이터 구성 */
    const analysisData = topCur.map((s, idx) => {
        const prev = prevMap.get(s.uid);
        const rank = idx + 1;
        const prevRank = prev ? prevRankMap.get(s.uid) : null;
        const isNew = !prev;

        const deltas = {};
        topSubjects.forEach(sub => {
            if (prev) {
                deltas[sub.k] = {
                    scoreDelta: parseFloat(
                        (getSubjScore(s, sub.k, metric) - getSubjScore(prev, sub.k, metric)).toFixed(2)
                    ),
                    grdDelta: prev[sub.k].grd - s[sub.k].grd
                };
            } else {
                deltas[sub.k] = { scoreDelta: null, grdDelta: null };
            }
        });

        const totalDelta = prev
            ? parseFloat((getTotalByMetric(s, metric) - getTotalByMetric(prev, metric)).toFixed(2))
            : null;

        let strongSubj = null, weakSubj = null;
        if (prev) {
            let maxD = -Infinity, minD = Infinity;
            topSubjects.forEach(sub => {
                const d = deltas[sub.k].scoreDelta;
                if (d !== null) {
                    if (d > maxD) { maxD = d; strongSubj = sub; }
                    if (d < minD) { minD = d; weakSubj = sub; }
                }
            });
            if (maxD <= 0) strongSubj = null;
            if (minD >= 0) weakSubj = null;
        }

        return { student: s, rank, prevRank, isNew, prev, deltas, totalDelta, strongSubj, weakSubj };
    });

    /* 각 섹션 렌더 */
    renderTopSummary(analysisData, topN, curStudents.length, prevStudents, metric);
    renderTopTable(analysisData, grade, prevStudents !== null, metric);
    renderTopHeatmap(analysisData, prevStudents !== null, metric);
    renderTopGrade1Chart(analysisData, prevStudents !== null);
    renderTopStrengthCards(analysisData, metric);
}

/* ── 요약 배너 ── */
function renderTopSummary(data, topN, totalStudents, hasPrev, metric) {
    const banner = document.getElementById('topSummaryBanner');
    if (!banner) return;

    const avgTotal = data.reduce((sum, d) =>
        sum + getTotalByMetric(d.student, metric), 0
    ) / data.length;

    const avgGrade = data.reduce((sum, d) => {
        const s = d.student;
        return sum + (s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4;
    }, 0) / data.length;

    let avgDelta = '-';
    let upCount = 0, downCount = 0, newCount = 0;
    if (hasPrev) {
        const deltas = data.filter(d => d.totalDelta !== null).map(d => d.totalDelta);
        avgDelta = deltas.length
            ? (deltas.reduce((a, b) => a + b, 0) / deltas.length).toFixed(1)
            : '-';
        upCount = data.filter(d => d.totalDelta !== null && d.totalDelta > 0).length;
        downCount = data.filter(d => d.totalDelta !== null && d.totalDelta < 0).length;
        newCount = data.filter(d => d.isNew).length;
    }

    let g1c = 0, tsc = 0;
    data.forEach(d => {
        topSubjects.forEach(sub => {
            if (d.student[sub.k].grd === 1) g1c++;
            tsc++;
        });
    });
    const g1p = tsc ? ((g1c / tsc) * 100).toFixed(1) : '0';

    banner.innerHTML = `
        <div class="top-summary-item">
            <div class="summary-label">분석 대상</div>
            <div class="summary-value" style="color:var(--primary);">${topN}명</div>
            <div class="summary-sub">전체 ${totalStudents}명 중</div>
        </div>
        <div class="top-summary-item">
            <div class="summary-label">평균 총점 (${getMetricLabel(metric)})</div>
            <div class="summary-value" style="color:var(--accent);">${avgTotal.toFixed(1)}</div>
            <div class="summary-sub">상위 ${topN}명 평균</div>
        </div>
        <div class="top-summary-item">
            <div class="summary-label">평균 등급</div>
            <div class="summary-value" style="color:var(--info);">${avgGrade.toFixed(2)}</div>
            <div class="summary-sub">4과목 평균</div>
        </div>
        <div class="top-summary-item">
            <div class="summary-label">1등급 비율</div>
            <div class="summary-value" style="color:var(--success);">${g1p}%</div>
            <div class="summary-sub">${g1c}/${tsc} (전과목)</div>
        </div>
        ${hasPrev ? `
            <div class="top-summary-item">
                <div class="summary-label">평균 증감</div>
                <div class="summary-value ${parseFloat(avgDelta) > 0 ? 'delta-up' : (parseFloat(avgDelta) < 0 ? 'delta-down' : '')}">
                    ${parseFloat(avgDelta) > 0 ? '+' : ''}${avgDelta}pt
                </div>
                <div class="summary-sub">
                    ▲${upCount}명 ▼${downCount}명
                    ${newCount > 0 ? `NEW ${newCount}명` : ''}
                </div>
            </div>` : ''}`;
}

/* ── 상위권 일람표 ── */
function renderTopTable(data, grade, hasPrev, metric) {
    const thead = document.getElementById('topTableHead');
    const tbody = document.getElementById('topTableBody');
    if (!thead || !tbody) return;

    let h1 = `<tr>
        <th rowspan="2">순위</th>
        <th rowspan="2">학번</th>
        <th rowspan="2">이름</th>`;

    topSubjects.forEach(sub => {
        if (hasPrev)
            h1 += `<th colspan="3">${sub.n}</th>`;
        else
            h1 += `<th colspan="2">${sub.n}</th>`;
    });

    h1 += `<th rowspan="2" class="total-col">합산</th>`;
    if (hasPrev)
        h1 += `<th rowspan="2" class="total-col">증감</th>
                <th rowspan="2" class="total-col">순위변동</th>`;
    h1 += `</tr>`;

    let h2 = `<tr>`;
    topSubjects.forEach(() => {
        h2 += `<th>점수</th><th>등급</th>`;
        if (hasPrev) h2 += `<th>증감</th>`;
    });
    h2 += `</tr>`;

    thead.innerHTML = h1 + h2;
    tbody.innerHTML = '';

    data.forEach(d => {
        const s = d.student;
        const id = `${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}`;

        let row = `<tr>
            <td style="font-weight:bold;color:var(--primary);">${d.rank}</td>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${grade})">${id}</td>
            <td class="student-link-cell"
                onclick="goToIndividual('${s.uid}',${grade})"
                style="font-weight:bold;">${s.info.name}</td>`;

        topSubjects.forEach(sub => {
            const score = getSubjScore(s, sub.k, metric);
            row += `<td>${typeof score === 'number' ? score : '-'}</td>`;
            row += `<td class="g-${s[sub.k].grd}">${s[sub.k].grd}</td>`;

            if (hasPrev) {
                const delta = d.deltas[sub.k];
                if (delta.scoreDelta !== null) {
                    const cls = delta.scoreDelta > 0 ? 'delta-up'
                              : (delta.scoreDelta < 0 ? 'delta-down' : 'delta-same');
                    const prefix = delta.scoreDelta > 0 ? '+' : '';

                    let gt = '';
                    if (delta.grdDelta > 0)
                        gt = ` <small style="color:#2D5A3D;">▲${delta.grdDelta}</small>`;
                    else if (delta.grdDelta < 0)
                        gt = ` <small style="color:#c0392b;">▼${Math.abs(delta.grdDelta)}</small>`;

                    row += `<td class="${cls}">${prefix}${delta.scoreDelta}${gt}</td>`;
                } else {
                    row += `<td class="delta-new">NEW</td>`;
                }
            }
        });

        row += `<td class="total-col">${getTotalByMetric(s, metric)}</td>`;

        if (hasPrev) {
            if (d.totalDelta !== null) {
                const cls = d.totalDelta > 0 ? 'delta-up'
                          : (d.totalDelta < 0 ? 'delta-down' : 'delta-same');
                const prefix = d.totalDelta > 0 ? '+' : '';
                row += `<td class="total-col ${cls}">${prefix}${d.totalDelta}</td>`;
            } else {
                row += `<td class="total-col delta-new">NEW</td>`;
            }

            if (d.isNew) {
                row += `<td class="total-col">
                            <span class="rank-change new-entry">NEW</span>
                        </td>`;
            } else if (d.prevRank !== null) {
                const rd = d.prevRank - d.rank;
                if (rd > 0)
                    row += `<td class="total-col">
                                <span class="rank-change up">▲${rd}</span>
                            </td>`;
                else if (rd < 0)
                    row += `<td class="total-col">
                                <span class="rank-change down">▼${Math.abs(rd)}</span>
                            </td>`;
                else
                    row += `<td class="total-col">
                                <span class="rank-change same">-</span>
                            </td>`;
            } else {
                row += `<td class="total-col">-</td>`;
            }
        }

        row += `</tr>`;
        tbody.innerHTML += row;
    });
}

/* ── 히트맵 ── */
function renderTopHeatmap(data, hasPrev, metric) {
    const wrapper = document.getElementById('topHeatmapWrapper');
    if (!wrapper) return;

    if (!hasPrev) {
        wrapper.innerHTML = `
            <div style="padding:30px;text-align:center;color:var(--text-muted);">
                비교 시험을 선택하면 증감 히트맵이 표시됩니다.
            </div>`;
        return;
    }

    let html = `<table class="heatmap-table"><thead><tr>
        <th>순위</th><th>학번</th><th>이름</th>`;
    topSubjects.forEach(sub => { html += `<th>${sub.n}</th>`; });
    html += `<th>합산</th></tr></thead><tbody>`;

    data.forEach(d => {
        const s = d.student;
        const id = `${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}`;

        html += `<tr>
            <td style="font-weight:bold;color:var(--primary);">${d.rank}</td>
            <td style="font-size:0.78rem;color:var(--text-secondary);">${id}</td>
            <td class="hm-name">${s.info.name}</td>`;

        topSubjects.forEach(sub => {
            const delta = d.deltas[sub.k];
            if (delta.scoreDelta !== null) {
                const color = getHeatmapColor(delta.scoreDelta);
                const prefix = delta.scoreDelta > 0 ? '+' : '';
                html += `<td style="background:${color};
                                    color:${Math.abs(delta.scoreDelta) > 10 ? 'white' : 'var(--text-primary)'};
                                    font-weight:700;">
                            ${prefix}${delta.scoreDelta}
                         </td>`;
            } else {
                html += `<td style="background:var(--info-bg);color:var(--info);font-style:italic;">
                            NEW
                         </td>`;
            }
        });

        if (d.totalDelta !== null) {
            const color = getHeatmapColor(d.totalDelta, 50);
            const prefix = d.totalDelta > 0 ? '+' : '';
            html += `<td style="background:${color};
                                color:${Math.abs(d.totalDelta) > 20 ? 'white' : 'var(--text-primary)'};
                                font-weight:700;">
                        ${prefix}${d.totalDelta}
                     </td>`;
        } else {
            html += `<td style="background:var(--info-bg);color:var(--info);">NEW</td>`;
        }

        html += `</tr>`;
    });

    html += `</tbody></table>`;
    wrapper.innerHTML = html;
}

function getHeatmapColor(value, maxAbs = 30) {
    if (value === null) return 'transparent';
    const clamped = Math.max(-maxAbs, Math.min(maxAbs, value));
    const ratio = clamped / maxAbs;

    if (ratio > 0) return `rgba(45,90,61,${0.15 + ratio * 0.65})`;
    else if (ratio < 0) return `rgba(192,57,43,${0.15 + Math.abs(ratio) * 0.65})`;
    return 'rgba(0,0,0,0.03)';
}

/* ── 1등급 인원 차트 ── */
function renderTopGrade1Chart(data, hasPrev) {
    const ctx = document.getElementById('topGrade1Chart');
    if (!ctx) return;

    const curCounts = {};
    const prevCounts = {};
    topSubjects.forEach(sub => { curCounts[sub.k] = 0; prevCounts[sub.k] = 0; });

    data.forEach(d => {
        topSubjects.forEach(sub => {
            if (d.student[sub.k].grd === 1) curCounts[sub.k]++;
            if (hasPrev && d.prev && d.prev[sub.k].grd === 1) prevCounts[sub.k]++;
        });
    });

    const labels = topSubjects.map(sub => sub.n);
    const datasets = [];

    if (hasPrev) {
        datasets.push({
            label: '직전 시험',
            data: topSubjects.map(sub => prevCounts[sub.k]),
            backgroundColor: 'rgba(156,148,136,0.6)',
            borderColor: 'rgba(156,148,136,1)',
            borderWidth: 1
        });
    }

    datasets.push({
        label: '최신 시험',
        data: topSubjects.map(sub => curCounts[sub.k]),
        backgroundColor: 'rgba(139,41,66,0.7)',
        borderColor: 'rgba(139,41,66,1)',
        borderWidth: 1
    });

    /* ★ 최대값 계산 후 여유 확보 */
    const allValues = [
        ...topSubjects.map(sub => curCounts[sub.k]),
        ...topSubjects.map(sub => prevCounts[sub.k])
    ];
    const maxVal = Math.max(...allValues, 1);
    const yMax = Math.ceil(maxVal * 1.10) + 2;

    if (state.charts.topGrade1) state.charts.topGrade1.destroy();

    state.charts.topGrade1 = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    max: yMax,
                    ticks: { stepSize: 1 },
                    title: { display: true, text: '1등급 인원(명)' }
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: { padding: 20 }
                },
                datalabels: {
                    display: true,
                    anchor: 'end',
                    align: 'top',
                    font: { weight: 'bold', size: 11 },
                    color: function(ctx) {
                        return ctx.datasetIndex === 0 ? '#6B6560' : '#8B2942';
                    },
                    formatter: v => v > 0 ? v + '명' : ''
                }
            }
        }
    });
}

/* ── 강약점 카드 ── */
function renderTopStrengthCards(data, metric) {
    const container = document.getElementById('topStrengthContainer');
    if (!container) return;
    container.innerHTML = '';

    data.forEach(d => {
        const s = d.student;
        const id = `${s.info.grade}${String(s.info.class).padStart(2, '0')}${String(s.info.no).padStart(2, '0')}`;

        /* 트렌드 클래스 */
        let trendClass = 'trend-same';
        if (d.isNew) trendClass = 'trend-new';
        else if (d.totalDelta > 0) trendClass = 'trend-up';
        else if (d.totalDelta < 0) trendClass = 'trend-down';

        /* 순위 클래스 */
        let rankClass = 'rank-other';
        if (d.rank === 1) rankClass = 'rank-1';
        else if (d.rank === 2) rankClass = 'rank-2';
        else if (d.rank === 3) rankClass = 'rank-3';

        /* 증감 뱃지 */
        let deltaBadge = '';
        if (d.isNew) {
            deltaBadge = `<span class="strength-delta-badge new-entry">
                              <i class="fas fa-star"></i> NEW
                          </span>`;
        } else if (d.totalDelta !== null) {
            const prefix = d.totalDelta > 0 ? '+' : '';
            const cls = d.totalDelta > 0 ? 'up' : (d.totalDelta < 0 ? 'down' : 'same');
            const icon = d.totalDelta > 0 ? 'fa-arrow-up'
                       : (d.totalDelta < 0 ? 'fa-arrow-down' : 'fa-minus');
            deltaBadge = `<span class="strength-delta-badge ${cls}">
                              <i class="fas ${icon}"></i> ${prefix}${d.totalDelta}pt
                          </span>`;
        }

        /* 태그 */
        let tags = '';
        if (d.strongSubj)
            tags += `<span class="strength-tag strong">
                         <i class="fas fa-arrow-up"></i> 강점: ${d.strongSubj.n}
                     </span>`;
        if (d.weakSubj)
            tags += `<span class="strength-tag weak">
                         <i class="fas fa-arrow-down"></i> 약점: ${d.weakSubj.n}
                     </span>`;

        /* 코멘트 */
        let comment = '';
        const ml = getMetricLabel(metric);

        if (d.isNew) {
            comment = '이번 시험에 새롭게 상위권에 진입한 학생입니다. 지속 관리가 필요합니다.';
        } else if (d.totalDelta !== null) {
            const abs = Math.abs(d.totalDelta);

            if (d.totalDelta >= 10)
                comment = `직전 대비 큰 폭 상승 (+${d.totalDelta}pt, ${ml}). 우수한 성적 향상을 보이고 있습니다.`;
            else if (d.totalDelta > 0)
                comment = `직전 대비 소폭 상승 (+${d.totalDelta}pt, ${ml}). 안정적으로 유지되고 있습니다.`;
            else if (d.totalDelta === 0)
                comment = `직전 대비 동일한 성적입니다 (${ml}).`;
            else if (abs <= 10)
                comment = `직전 대비 소폭 하락 (${d.totalDelta}pt, ${ml}). 취약과목 보완이 필요합니다.`;
            else if (abs <= 25)
                comment = `직전 대비 큰 폭 하락 (${d.totalDelta}pt, ${ml}). 집중 관리가 권장됩니다.`;
            else
                comment = `직전 대비 매우 큰 폭 하락 (${d.totalDelta}pt, ${ml}). 긴급 면담 및 학습 점검이 필요합니다.`;

            if (d.strongSubj && d.weakSubj) {
                comment += ` ${d.strongSubj.n} ▲${d.deltas[d.strongSubj.k].scoreDelta}, `
                         + `${d.weakSubj.n} ▼${Math.abs(d.deltas[d.weakSubj.k].scoreDelta)}.`;
            }
        }

        /* 순위 변동 */
        let rankChangeText = '';
        if (d.isNew) {
            rankChangeText = '<span class="rank-change new-entry">NEW</span>';
        } else if (d.prevRank !== null) {
            const rd = d.prevRank - d.rank;
            if (rd > 0)
                rankChangeText = `<span class="rank-change up">▲${rd}등</span>`;
            else if (rd < 0)
                rankChangeText = `<span class="rank-change down">▼${Math.abs(rd)}등</span>`;
            else
                rankChangeText = '<span class="rank-change same">-</span>';
        }

        const totalScore = getTotalByMetric(s, metric);
        const prevTotalScore = d.prev ? getTotalByMetric(d.prev, metric) : null;

        container.innerHTML += `
            <div class="strength-card ${trendClass}">
                <div class="strength-header">
                    <div class="strength-header-left">
                        <span class="strength-rank ${rankClass}">${d.rank}</span>
                        <div>
                            <div class="strength-name">${s.info.name}</div>
                            <div class="strength-id">${id} ${rankChangeText}</div>
                        </div>
                    </div>
                    ${deltaBadge}
                </div>
                <div class="strength-body">
                    <div class="strength-row">
                        <span class="label">총점 (${ml})</span>
                        <span class="value">
                            ${totalScore}${prevTotalScore !== null ? ` ← ${prevTotalScore}` : ''}
                        </span>
                    </div>
                    <div class="strength-row">
                        <span class="label">평균등급</span>
                        <span class="value">
                            ${((s.kor.grd + s.math.grd + s.eng.grd + (s.inq1.grd + s.inq2.grd) / 2) / 4).toFixed(2)}
                        </span>
                    </div>
                </div>
                ${tags ? `<div class="strength-tags">${tags}</div>` : ''}
                ${comment ? `<div class="strength-comment">${comment}</div>` : ''}
            </div>`;
    });
}

/* ── 상위권 엑셀 저장 ── */
window.exportTopTableToExcel = function() {
    const thead = document.getElementById('topTableHead');
    const tbody = document.getElementById('topTableBody');
    if (!thead || !tbody || tbody.rows.length === 0)
        return alert('저장할 데이터가 없습니다.');

    const wb = XLSX.utils.book_new();
    const rows = [];

    const headerRows = thead.querySelectorAll('tr');
    if (headerRows.length === 2) {
        const mh = [];
        const r1 = headerRows[0].querySelectorAll('th');
        const r2 = Array.from(headerRows[1].querySelectorAll('th'));
        let r2i = 0;

        r1.forEach(th => {
            const cs = parseInt(th.getAttribute('colspan')) || 1;
            const rs = parseInt(th.getAttribute('rowspan')) || 1;

            if (rs === 2) {
                mh.push(th.textContent.trim());
            } else if (cs > 1) {
                const pn = th.textContent.trim();
                for (let i = 0; i < cs && r2i < r2.length; i++) {
                    mh.push(`${pn}_${r2[r2i].textContent.trim()}`);
                    r2i++;
                }
            } else {
                mh.push(th.textContent.trim());
            }
        });
        rows.push(mh);
    }

    tbody.querySelectorAll('tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('td').forEach(td => {
            const val = td.textContent.trim();
            const num = Number(val);
            row.push(isNaN(num) || val === '' || val === '-' || val === 'NEW' ? val : num);
        });
        rows.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = rows[0].map((_, ci) => {
        let mx = 0;
        rows.forEach(r => {
            const cl = String(r[ci] || '').length;
            if (cl > mx) mx = cl;
        });
        return { wch: Math.max(mx + 2, 8) };
    });

    XLSX.utils.book_append_sheet(wb, ws, '상위권분석');
    const now = new Date();
    XLSX.writeFile(wb,
        `상위권_분석_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.xlsx`
    );
};

/* ============================================================
   기타 유틸
   ============================================================ */
function updateLastUpdated() {
    const now = new Date();
    const el = document.getElementById('lastUpdated');
    if (el) {
        el.textContent = `Last updated: ${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')} (KST)`;
    }
}

/* ── HTML 저장 ── */
window.saveHtmlFile = function() {
    try {
        const hc = document.documentElement.cloneNode(true);
        hc.querySelector('#uploadSection')?.remove();
        hc.querySelector('#loading')?.remove();
        hc.querySelector('#saveHtmlBtn')?.remove();

        const sts = {
            gradeData: state.gradeData,
            availableGrades: state.availableGrades,
            currentGradeTotal: state.currentGradeTotal,
            currentGradeClass: state.currentGradeClass,
            currentGradeIndiv: state.currentGradeIndiv,
            currentGradeTop: state.currentGradeTop,
            metric: state.metric,
            classMetric: state.classMetric,
            classSort: state.classSort,
            topMetric: state.topMetric,
            topN: state.topN,
            charts: {}
        };

        const st = document.createElement('script');
        st.textContent = `
            window.SAVED_STATE=${JSON.stringify(sts)};
            window.addEventListener('DOMContentLoaded', function() {
                if (window.SAVED_STATE) {
                    Object.assign(state, window.SAVED_STATE);
                    state.charts = {};
                    document.getElementById('results').style.display = 'block';
                    initSelectors();
                }
            });`;
        hc.querySelector('head').appendChild(st);

        const blob = new Blob(
            ['<!DOCTYPE html>\n' + hc.outerHTML],
            { type: 'text/html;charset=utf-8' }
        );
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        const now = new Date();
        a.download = `모의고사_분석결과_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.html`;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (e) {
        alert('HTML 저장 중 오류: ' + e.message);
    }
};

/* ============================================================
   PDF 생성
   ============================================================ */
async function captureTabPageV2(pdf, showSelectors, addNewPage) {
    const source = document.getElementById('individual-tab');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfPageHeight = pdf.internal.pageSize.getHeight();
    const marginX = 8, marginY = 8;
    const contentWidth = pdfWidth - marginX * 2;
    const cloneWidth = 1600;

    /* 캔버스 스냅샷 */
    const canvasMap = new Map();
    source.querySelectorAll('canvas').forEach(c => {
        if (!c.id) return;
        try {
            canvasMap.set(c.id, {
                dataUrl: c.toDataURL('image/png'),
                w: c.offsetWidth || c.getBoundingClientRect().width,
                h: c.offsetHeight || c.getBoundingClientRect().height
            });
        } catch (e) {}
    });

    /* 클론 래퍼 */
    const wrapper = document.createElement('div');
    wrapper.style.cssText = `
        position:fixed; top:-99999px; left:0;
        width:${cloneWidth}px; min-width:${cloneWidth}px; max-width:${cloneWidth}px;
        background:#FEFDFB; padding:20px 24px;
        box-sizing:border-box; display:block; overflow:visible; z-index:-9999;`;

    const fixCardWidths = (el) => {
        el.querySelectorAll('.chart-card,.student-profile-card').forEach(card => {
            card.style.setProperty('width', '100%', 'important');
            card.style.setProperty('box-sizing', 'border-box', 'important');
            card.style.setProperty('max-width', 'none', 'important');
        });
    };

    const allSubjects = ['kor', 'math', 'eng', 'hist', 'inq1', 'inq2'];

    /* 프로필 */
    if (showSelectors.includes('profile')) {
        const pc = source.querySelector('.student-profile-card');
        if (pc) {
            const c = pc.cloneNode(true);
            c.style.marginBottom = '16px';
            c.style.width = '100%';
            c.style.boxSizing = 'border-box';

            const sg = c.querySelector('.profile-stats');
            if (sg) sg.setAttribute('style',
                'display:grid !important;grid-template-columns:repeat(5,1fr) !important;gap:15px !important;width:100% !important;'
            );

            const ph = c.querySelector('.profile-header');
            if (ph) ph.setAttribute('style',
                'display:flex !important;justify-content:space-between !important;align-items:flex-start !important;margin-bottom:20px !important;width:100% !important;'
            );

            wrapper.appendChild(c);
        }
    }

    /* 차트 */
    if (showSelectors.includes('charts')) {
        const cr = source.querySelector('.charts-row');
        if (cr) {
            const c = cr.cloneNode(true);
            c.setAttribute('style',
                'display:grid !important;grid-template-columns:3fr 2fr !important;gap:20px !important;margin-bottom:20px !important;width:100% !important;box-sizing:border-box !important;'
            );
            c.querySelectorAll('.chart-half').forEach(ch => {
                ch.setAttribute('style',
                    'display:flex !important;flex-direction:column !important;overflow:visible !important;min-height:0 !important;height:auto !important;width:100% !important;box-sizing:border-box !important;'
                );
            });
            c.querySelectorAll('.chart-half .chart-container').forEach(cc => {
                cc.setAttribute('style',
                    'overflow:visible !important;height:360px !important;min-height:360px !important;width:100% !important;display:block !important;position:relative !important;'
                );
            });
            wrapper.appendChild(c);
        }
    }

    /* 과목 상세 */
    allSubjects.forEach(subj => {
        if (!showSelectors.includes(subj)) return;
        const card = source.querySelector(`.subject-detail-card[data-subject="${subj}"]`);
        if (!card) return;

        const cc = card.cloneNode(true);
        cc.style.marginBottom = '16px';

        cc.querySelectorAll('.subject-detail-grid').forEach(grid => {
            grid.setAttribute('style',
                'display:grid !important;grid-template-columns:2fr 3fr !important;gap:16px !important;align-items:start !important;width:100% !important;'
            );
            grid.querySelectorAll('.chart-container').forEach(chartC => {
                chartC.setAttribute('style',
                    'height:auto !important;width:100% !important;position:relative !important;overflow:visible !important;'
                );
            });
            grid.querySelectorAll('.table-wrapper').forEach(tw => {
                tw.setAttribute('style',
                    'overflow:visible !important;overflow-x:visible !important;max-height:none !important;width:100% !important;'
                );
            });
            grid.querySelectorAll('table').forEach(tbl => {
                tbl.setAttribute('style',
                    'min-width:0 !important;width:100% !important;table-layout:auto !important;font-size:0.75rem !important;border-collapse:separate !important;border-spacing:0 !important;overflow:hidden !important;'
                );
                const thead = tbl.querySelector('thead');
                const tbody = tbl.querySelector('tbody');
                if (thead && tbody) {
                    Array.from(thead.querySelectorAll('tr')).reverse().forEach(row => {
                        tbody.insertBefore(row, tbody.firstChild);
                    });
                    thead.remove();
                }
            });
            grid.querySelectorAll('th,td').forEach(cell => {
                cell.setAttribute('style',
                    (cell.getAttribute('style') || '') +
                    'white-space:nowrap !important;padding:4px 6px !important;'
                );
            });
        });

        wrapper.appendChild(cc);
    });

    /* 캔버스 → 이미지 교체 */
    wrapper.querySelectorAll('canvas').forEach(cloneCanvas => {
        const snap = canvasMap.get(cloneCanvas.id);
        if (!snap) { cloneCanvas.style.display = 'none'; return; }

        const img = document.createElement('img');
        img.src = snap.dataUrl;
        const ar = snap.w > 0 && snap.h > 0 ? `${snap.w}/${snap.h}` : 'auto';
        img.style.cssText = `display:block;width:100%;height:auto;aspect-ratio:${ar};object-fit:fill;`;
        cloneCanvas.parentNode.replaceChild(img, cloneCanvas);
    });

    wrapper.id = '__pdf_wrapper__';
    const prevOverflow = document.documentElement.style.overflow;
    document.documentElement.style.overflow = 'hidden';
    document.body.appendChild(wrapper);
    fixCardWidths(wrapper);

    await new Promise(r => setTimeout(r, 300));

    const canvas = await html2canvas(wrapper, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#FEFDFB',
        windowWidth: cloneWidth,
        logging: false,
        allowTaint: true,
        onclone: (clonedDoc) => {
            const os = clonedDoc.createElement('style');
            os.textContent = [
                `body{min-width:${cloneWidth}px !important;overflow:visible !important;}`,
                `.container{max-width:none !important;width:100% !important;}`,
                `#__pdf_wrapper__{width:${cloneWidth}px !important;min-width:${cloneWidth}px !important;max-width:none !important;}`
            ].join('\n');
            clonedDoc.head.appendChild(os);
        }
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
    const ot = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> PDF 생성 중...';

    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');

        await new Promise(r => setTimeout(r, 1000));
        await captureTabPageV2(pdf, ['profile', 'charts'], false);
        await captureTabPageV2(pdf, ['kor', 'math', 'eng'], true);
        await captureTabPageV2(pdf, ['hist', 'inq1', 'inq2'], true);

        const sn = document.getElementById('indivName').innerText;
        pdf.save(`모의고사_분석리포트_${sn}.pdf`);
    } catch (e) {
        console.error('PDF 생성 오류:', e);
        alert('PDF 생성 중 오류가 발생했습니다.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = ot;
    }
}

async function generateClassPDF() {
    const btn = document.getElementById('pdfClassBtn');
    if (!btn || btn.disabled) return;

    const cs = document.getElementById('indivClassSelect');
    if (!cs) return;

    const cls = parseInt(cs.value);
    const sel = document.getElementById('indivStudentSelect');
    const options = Array.from(sel.options);

    if (!options.length) return alert('해당 학급에 학생이 없습니다.');

    if (!confirm(
        `현재 선택된 ${cls}반 학생 ${options.length}명 전체 리포트를 PDF로 생성합니다.\n` +
        `시간이 다소 소요될 수 있습니다. 진행하시겠습니까?`
    )) return;

    btn.disabled = true;
    const ot = btn.innerHTML;

    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('l', 'mm', 'a4');

        for (let i = 0; i < options.length; i++) {
            btn.innerHTML =
                `<i class="fas fa-spinner fa-spin"></i> PDF 생성 중... (${i + 1}/${options.length})`;

            sel.value = options[i].value;
            renderIndividual();

            await new Promise(r => setTimeout(r, 1000));
            await captureTabPageV2(pdf, ['profile', 'charts'], i > 0);
            await captureTabPageV2(pdf, ['kor', 'math', 'eng'], true);
            await captureTabPageV2(pdf, ['hist', 'inq1', 'inq2'], true);
        }

        pdf.save(`모의고사_분석리포트_${cls}반_전체.pdf`);
    } catch (e) {
        console.error('PDF 생성 오류:', e);
        alert('학급 PDF 생성 중 오류가 발생했습니다.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = ot;
    }
}
