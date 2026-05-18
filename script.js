/* 예산서 검토 도구 v1
 * - 모든 처리는 브라우저 내부에서 수행됩니다.
 * - CDN: SheetJS, pdf.js
 */

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

const TOLERANCE = 1000;
const SHEET_KEYWORDS = {
  retirement: ['퇴직금적립현황', '퇴직적립금', '퇴직금', '퇴직충당금', '적립현황'],
  payroll: ['교직원보수일람표', '교직원보수', '보수일람표', '보수기준', '급여기준', '급여']
};

const EXPECTED_BUCKETS = {
  teacherSalary: ['교원급여', '원장', '교사', '방과후교사', '교원 기본급', '기본급'],
  teacherAllowance: ['교원수당', '정액급식비', '급식비', '연구수당', '연구활동비', '시간외수당', '직급보조비', '관리업무수당', '기타수당', '자가운전보조금', '명절휴가비', '상여금', '성과상여금', '처우개선비', '교통보조비'],
  staffSalary: ['직원급여', '사무직원', '조리직원', '조리사', '영양사', '보조교사', '방과후보조', '차량기사', '차량보조', '환경미화원', '청소직', '관리직'],
  staffAllowance: ['직원수당', '직원정액급식비', '직원급식비', '연장근로수당', '명절휴가비', '상여금', '성과상여금', '방학휴가비']
};

const PDF_TARGET_MOKS = ['교원급여', '교원수당', '직원급여', '직원수당', '교원퇴직금및퇴직적립금', '교원퇴직적립금', '직원퇴직금및퇴직적립금', '직원퇴직적립금'];

const $ = (id) => document.getElementById(id);

$('runBtn').addEventListener('click', runReview);
$('copyBtn').addEventListener('click', () => {
  const text = Array.from(document.querySelectorAll('.issue')).map(el => el.innerText).join('\n\n');
  navigator.clipboard.writeText(text || '지적사항 없음');
  setStatus('검토 결과를 클립보드에 복사했습니다.');
});

async function runReview() {
  clearUI();
  const formFile = $('formFile').files[0];
  const budgetFile = $('budgetFile').files[0];
  if (!formFile || !budgetFile) return setStatus('예산서 작성서식과 세출예산명세서 파일을 모두 선택해 주세요.', true);

  try {
    setStatus('파일을 읽는 중입니다...');
    const formData = await parseFormFile(formFile);
    const budgetData = await parseBudgetFile(budgetFile);

    setStatus('검토 기준을 적용하는 중입니다...');
    const result = review(formData, budgetData);
    renderResult(result, formData, budgetData);
    setStatus('검토가 완료되었습니다.');
  } catch (err) {
    console.error(err);
    setStatus(`처리 중 오류가 발생했습니다: ${err.message || err}`, true);
  }
}

function clearUI() {
  ['summaryPanel', 'issuesPanel', 'debugPanel'].forEach(id => $(id).classList.add('hidden'));
  $('summary').innerHTML = '';
  $('issues').innerHTML = '';
  $('debug').textContent = '';
}

function setStatus(msg, isError=false) {
  const el = $('status');
  el.textContent = msg;
  el.classList.remove('hidden');
  el.style.background = isError ? '#fff1f3' : '#eef4ff';
  el.style.color = isError ? '#b42318' : '#1849a9';
}

async function parseFormFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'pdf') return parseFormPdf(await readArrayBuffer(file));
  return parseFormExcel(await readArrayBuffer(file));
}

async function parseBudgetFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'pdf') return parseBudgetPdf(await readArrayBuffer(file));
  const parsed = parseWorkbookBudget(await readArrayBuffer(file));
  return parsed;
}

function readArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseFormExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false, cellNF: false, cellText: false, WTF: false });
  const retirementSheetName = findSheetName(wb, SHEET_KEYWORDS.retirement);
  const payrollSheetName = findSheetName(wb, SHEET_KEYWORDS.payroll);

  const retirementRows = retirementSheetName ? sheetToRows(wb.Sheets[retirementSheetName]) : [];
  const payrollRows = payrollSheetName ? sheetToRows(wb.Sheets[payrollSheetName]) : [];

  return {
    sourceType: 'excel',
    selectedSheets: { retirement: retirementSheetName || '미인식', payroll: payrollSheetName || '미인식' },
    retirement: parseRetirementRows(retirementRows),
    payroll: parsePayrollRows(payrollRows),
    raw: { retirementRows, payrollRows }
  };
}

function findSheetName(wbOrNames, keywords) {
  const names = Array.isArray(wbOrNames) ? wbOrNames : (wbOrNames?.SheetNames || []);
  const sheets = Array.isArray(wbOrNames) ? null : (wbOrNames?.Sheets || {});
  const norm = s => normalizeText(s);
  let best = null, score = 0;

  // 1순위: 시트명으로 찾기
  for (const name of names) {
    const n = norm(name);
    let s = 0;
    keywords.forEach(k => { if (n.includes(norm(k))) s += norm(k).length + 10; });
    if (s > score) { best = name; score = s; }
  }

  // 2순위: 시트명은 다르지만 내용 첫 부분에 관련 제목이 있는 경우
  if (sheets) {
    for (const name of names) {
      const sheet = sheets[name];
      if (!sheet || !sheet['!ref']) continue;
      const preview = sheetToRows(sheet, 25).flat().join(' ');
      const n = norm(preview);
      let s = 0;
      keywords.forEach(k => { if (n.includes(norm(k))) s += norm(k).length; });
      if (s > score) { best = name; score = s; }
    }
  }
  return best;
}

function sheetToRows(sheet, maxRows = Infinity) {
  if (!sheet || !sheet['!ref']) return [];
  let range;
  try {
    range = XLSX.utils.decode_range(sheet['!ref']);
  } catch (e) {
    return [];
  }
  if (Number.isFinite(maxRows)) range.e.r = Math.min(range.e.r, range.s.r + maxRows - 1);
  const merges = sheet['!merges'] || [];
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      let cell = sheet[addr];
      if (!cell) {
        const merge = merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
        if (merge) cell = sheet[XLSX.utils.encode_cell(merge.s)];
      }
      row.push(cell ? (cell.w ?? cell.v ?? '') : '');
    }
    rows.push(row);
  }
  return rows;
}

function parseRetirementRows(rows) {
  let carryoverTotal = 0;
  const flatRows = rows.map(r => r.map(v => String(v ?? '').trim()));
  for (let i = 0; i < flatRows.length; i++) {
    const rowText = flatRows[i].join(' ');
    if (rowText.includes('적립금') && rowText.includes('이월') && rowText.includes('계')) {
      carryoverTotal = Math.max(carryoverTotal, ...flatRows[i].map(parseMoney), ...(flatRows[i+1] || []).map(parseMoney));
    }
  }
  if (!carryoverTotal) {
    const suspicious = flatRows.filter(r => r.join(' ').includes('이월') || r.join(' ').includes('퇴직'));
    suspicious.forEach(r => { carryoverTotal = Math.max(carryoverTotal, ...r.map(parseMoney)); });
  }
  return { carryoverTotal };
}

function parsePayrollRows(rows) {
  const items = [];
  rows.forEach((row, idx) => {
    const cells = row.map(v => String(v ?? '').trim()).filter(Boolean);
    if (!cells.length) return;
    const nums = cells.map(parseMoney).filter(n => n > 0);
    if (!nums.length) return;
    const label = cells.filter(c => !/^[-\d,\.\s원]+$/.test(c)).join(' ').replace(/\s+/g, ' ').trim();
    const amount = Math.max(...nums);
    if (label && amount > 0) {
      const bucket = guessBucket(label);
      if (bucket) items.push({ label, amount, bucket, row: idx + 1 });
    }
  });
  return normalizePayrollItems(items);
}

function normalizePayrollItems(items) {
  const seen = new Set();
  return items.filter(it => {
    const key = `${it.bucket}|${normalizeText(it.label)}|${it.amount}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function guessBucket(label) {
  const n = normalizeText(label);
  let best = null, score = 0;
  for (const [bucket, keys] of Object.entries(EXPECTED_BUCKETS)) {
    let s = 0;
    keys.forEach(k => { if (n.includes(normalizeText(k))) s += k.length; });
    if (s > score) { best = bucket; score = s; }
  }
  return score ? best : null;
}

async function parseFormPdf(buffer) {
  const pages = await extractPdfPages(buffer);
  const allText = pages.map(p => p.text).join('\n');
  const retirementText = findSectionText(allText, SHEET_KEYWORDS.retirement);
  const payrollText = findSectionText(allText, SHEET_KEYWORDS.payroll);
  return {
    sourceType: 'pdf',
    selectedSheets: { retirement: retirementText ? 'PDF 내 퇴직금 관련 구간' : '미인식', payroll: payrollText ? 'PDF 내 보수 관련 구간' : '미인식' },
    retirement: parseRetirementText(retirementText || allText),
    payroll: parsePayrollText(payrollText || allText),
    raw: { pages }
  };
}

function parseRetirementText(text) {
  const lines = text.split(/\n/).map(s => s.trim()).filter(Boolean);
  let carryoverTotal = 0;
  lines.forEach((line, i) => {
    const context = [lines[i-1], line, lines[i+1]].filter(Boolean).join(' ');
    if ((context.includes('이월') || context.includes('적립')) && context.includes('계')) {
      carryoverTotal = Math.max(carryoverTotal, ...extractAmounts(context));
    }
  });
  return { carryoverTotal };
}

function parsePayrollText(text) {
  const lines = text.split(/\n/).map(s => s.trim()).filter(Boolean);
  const items = [];
  lines.forEach((line, idx) => {
    const amounts = extractAmounts(line);
    if (!amounts.length) return;
    const label = line.replace(/[\d,]+\s*원?/g, ' ').replace(/\s+/g, ' ').trim();
    const bucket = guessBucket(label);
    if (bucket) items.push({ label, amount: Math.max(...amounts), bucket, row: idx + 1 });
  });
  return normalizePayrollItems(items);
}

function findSectionText(text, keywords) {
  const lines = text.split(/\n/);
  const start = lines.findIndex(l => keywords.some(k => normalizeText(l).includes(normalizeText(k))));
  if (start < 0) return '';
  return lines.slice(start, Math.min(lines.length, start + 160)).join('\n');
}

async function parseBudgetPdf(buffer) {
  const pages = await extractPdfPages(buffer);
  const text = pages.map(p => `\n[PAGE ${p.page}]\n${p.text}`).join('\n');
  const entries = parseBudgetEntries(text);
  return { sourceType: 'pdf', pages, entries, text };
}

function parseWorkbookBudget(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false, cellNF: false, cellText: false, WTF: false });
  let text = '';
  (wb.SheetNames || []).forEach(name => {
    const sheet = wb.Sheets?.[name];
    if (!sheet || !sheet['!ref']) return;
    let csv = '';
    try {
      csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
    } catch (e) {
      csv = sheetToRows(sheet).map(r => r.join(',')).join('\n');
    }
    text += `\n[SHEET ${name}]\n${csv}\n`;
  });
  return { sourceType: 'excel', entries: parseBudgetEntries(text), text };
}

async function extractPdfPages(buffer) {
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const items = content.items.map(it => ({ str: it.str, x: it.transform[4], y: Math.round(it.transform[5]) }));
    const grouped = new Map();
    for (const it of items) {
      if (!grouped.has(it.y)) grouped.set(it.y, []);
      grouped.get(it.y).push(it);
    }
    const lines = Array.from(grouped.entries())
      .sort((a,b) => b[0] - a[0])
      .map(([, arr]) => arr.sort((a,b) => a.x - b.x).map(it => it.str).join(' ').replace(/\s+/g, ' ').trim())
      .filter(Boolean);
    pages.push({ page: i, text: lines.join('\n'), lines });
  }
  return pages;
}

function parseBudgetEntries(text) {
  const lines = text.split(/\n/).map(s => s.trim()).filter(Boolean);
  const entries = [];
  let currentMok = '';
  let currentGwan = '';
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].replace(/\s+/g, ' ').trim();
    const clean = normalizeText(line);
    for (const mok of PDF_TARGET_MOKS) if (clean.includes(normalizeText(mok))) currentMok = mok;
    const possibleMoks = ['인건비','교원인건비','직원인건비','운영비','일반교육활동비','선택적교육활동비','통학차량이용비','적립금','시설설비비품비','일반급식비간식비'];
    for (const g of possibleMoks) if (clean === normalizeText(g) || clean.startsWith(normalizeText(g))) currentGwan = g;

    if (line.includes('본예산') || line.includes('=') || /[\d,]+원/.test(line)) {
      let block = line;
      let j = i + 1;
      while (j < lines.length && !/[\d,]+\s*$/.test(block) && j < i + 5) {
        block += ' ' + lines[j];
        j++;
      }
      const amount = extractFinalAmount(block);
      if (amount > 0) {
        const detail = block.replace(/\d[\d,]*\s*원\s*[*xX×]?.*$/g, '').trim();
        entries.push({ mok: currentMok || inferMokFromContext(block), parent: currentGwan, detail: normalizeBudgetDetail(detail || block), formula: block, amount, pageHint: findPageHint(lines, i) });
      }
    }
  }
  return mergeSplitBudgetEntries(entries);
}

function mergeSplitBudgetEntries(entries) {
  return entries.map(e => ({ ...e, detail: e.detail.replace(/\s+/g, ' ').trim(), formula: e.formula.replace(/\s+/g, ' ').trim() }));
}

function inferMokFromContext(text) {
  const n = normalizeText(text);
  for (const mok of PDF_TARGET_MOKS) if (n.includes(normalizeText(mok))) return mok;
  return '';
}

function findPageHint(lines, index) {
  for (let i = index; i >= 0; i--) if (/^\[PAGE \d+\]/.test(lines[i])) return lines[i].replace(/\D/g, '');
  return '';
}

function normalizeBudgetDetail(s) {
  return s.replace(/\(본예산\)/g, '').replace(/\s+/g, ' ').trim();
}

function review(formData, budgetData) {
  const issues = [];
  const entries = budgetData.entries || [];
  const payroll = formData.payroll || [];

  const hasRetirementCarryover = (formData.retirement?.carryoverTotal || 0) > 0;
  const hasTeacherRetirement = entries.some(e => normalizeText(e.mok + e.detail).includes('교원퇴직'));
  const hasStaffRetirement = entries.some(e => normalizeText(e.mok + e.detail).includes('직원퇴직'));
  if (hasRetirementCarryover && !(hasTeacherRetirement || hasStaffRetirement)) {
    issues.push({ type: 'danger', title: '퇴직적립금 미편성', text: `퇴직금 적립금 이월액 계가 ${fmt(formData.retirement.carryoverTotal)}원으로 인식되었으나, 세출예산명세서에 교원퇴직적립금 또는 직원퇴직적립금 편성 항목을 찾지 못했습니다.` });
  }

  for (const item of payroll) {
    const normalMok = bucketToMok(item.bucket);
    const direct = findMatchingEntry(entries, item, normalMok);
    if (direct) continue;

    const integrated = findIntegrated(entries, payroll, item.bucket, normalMok);
    if (integrated) {
      if (!integrated.reported) {
        issues.push({ type: 'warn', title: `${normalMok} 통합편성`, text: `${normalMok} 세부 항목이 PDF에 개별 편성되어 있지 않지만, 엑셀 항목 합계(${integrated.labels}) = ${fmt(integrated.sum)}원과 PDF '${integrated.entry.detail}' ${fmt(integrated.entry.amount)}원이 허용범위 내에서 일치합니다.` });
        integrated.reported = true;
      }
      continue;
    }

    const misplaced = findMisplaced(entries, item, normalMok);
    if (misplaced.length) {
      const locations = misplaced.map(e => `${e.parent || e.mok || '다른 목'}의 '${e.detail}' ${fmt(e.amount)}원`).join(', ');
      issues.push({ type: 'danger', title: '오편성 의심', text: `${item.label} ${fmt(item.amount)}원이 정상 목(${normalMok})이 아닌 ${locations}에 편성된 것으로 보입니다.` });
      continue;
    }

    issues.push({ type: 'danger', title: '미편성 또는 금액 불일치', text: `${item.label} ${fmt(item.amount)}원을 세출예산명세서의 ${normalMok} 산출내역에서 확인하지 못했습니다.` });
  }

  if (!issues.length) issues.push({ type: 'ok', title: '지적사항 없음', text: '현재 1차 검토 기준에서 지적사항을 찾지 못했습니다.' });
  return { issues, counts: { payrollItems: payroll.length, budgetEntries: entries.length, issues: issues.filter(i => i.type !== 'ok').length } };
}

function bucketToMok(bucket) {
  return ({ teacherSalary: '교원급여', teacherAllowance: '교원수당', staffSalary: '직원급여', staffAllowance: '직원수당' })[bucket] || '';
}

function findMatchingEntry(entries, item, mok) {
  const labelN = normalizeText(item.label);
  return entries.find(e => normalizeText(e.mok).includes(normalizeText(mok)) && isSameMoney(e.amount, item.amount) && similarLabel(labelN, normalizeText(e.detail)));
}

function similarLabel(a, b) {
  if (!a || !b) return false;
  if (a.includes(b) || b.includes(a)) return true;
  const words = a.match(/[가-힣A-Za-z0-9]+/g) || [];
  return words.some(w => w.length >= 2 && b.includes(w));
}

function findIntegrated(entries, payroll, bucket, mok) {
  const missingGroup = payroll.filter(p => p.bucket === bucket);
  const sum = missingGroup.reduce((acc, p) => acc + p.amount, 0);
  const entry = entries.find(e => normalizeText(e.mok).includes(normalizeText(mok)) && isSameMoney(e.amount, sum));
  if (!entry) return null;
  return { entry, sum, labels: missingGroup.map(p => `${p.label} ${fmt(p.amount)}원`).join(' + ') };
}

function findMisplaced(entries, item, expectedMok) {
  return entries.filter(e => !normalizeText(e.mok).includes(normalizeText(expectedMok)) && (isSameMoney(e.amount, item.amount) || similarLabel(normalizeText(item.label), normalizeText(e.detail))));
}

function isSameMoney(a, b) {
  if (!a || !b) return false;
  if (Math.abs(a - b) <= TOLERANCE) return true;
  const ak = Math.round(a / 1000) * 1000;
  const bk = Math.round(b / 1000) * 1000;
  const af = Math.floor(a / 1000) * 1000;
  const bf = Math.floor(b / 1000) * 1000;
  const ac = Math.ceil(a / 1000) * 1000;
  const bc = Math.ceil(b / 1000) * 1000;
  return ak === bk || af === bf || ac === bc || Math.abs(ak - bk) <= TOLERANCE;
}

function parseMoney(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return Math.round(v);
  const s = String(v).replace(/\s/g, '');
  const matches = s.match(/-?\d{1,3}(?:,\d{3})+|-?\d+/g);
  if (!matches) return 0;
  return Math.max(...matches.map(m => Number(m.replace(/,/g, ''))).filter(Number.isFinite));
}

function extractAmounts(text) {
  const matches = String(text).match(/\d{1,3}(?:,\d{3})+(?=\s*원?)|\d+(?=\s*원)/g) || [];
  return matches.map(m => Number(m.replace(/,/g, ''))).filter(n => Number.isFinite(n) && n > 0);
}

function extractFinalAmount(text) {
  const amounts = extractAmounts(text);
  if (!amounts.length) return 0;
  return amounts[amounts.length - 1];
}

function normalizeText(s) {
  return String(s ?? '').replace(/[\s\n\r\t,()（）·ㆍ_\-]/g, '').toLowerCase();
}

function fmt(n) { return Number(n || 0).toLocaleString('ko-KR'); }

function renderResult(result, formData, budgetData) {
  $('summaryPanel').classList.remove('hidden');
  $('issuesPanel').classList.remove('hidden');
  $('debugPanel').classList.remove('hidden');
  $('summary').innerHTML = `
    <div class="stat"><strong>${result.counts.payrollItems}</strong><span>보수 항목 인식</span></div>
    <div class="stat"><strong>${result.counts.budgetEntries}</strong><span>예산 산출내역 인식</span></div>
    <div class="stat"><strong>${result.counts.issues}</strong><span>지적사항</span></div>
  `;
  $('issues').innerHTML = result.issues.map(i => `<div class="issue ${i.type === 'warn' ? 'warn' : i.type === 'ok' ? 'ok' : ''}"><div class="issue-title">${escapeHtml(i.title)}</div><div>${escapeHtml(i.text)}</div></div>`).join('');
  $('debug').textContent = JSON.stringify({ formData, budgetEntries: budgetData.entries }, null, 2);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
}

/* ===== v1.2 보강 패치: 보수일람표 열 기반 파싱 + PDF 목 추적 + 오편성 합산 탐색 ===== */

const ALL_BUDGET_MOK_HEADINGS = [
  '교원급여','교원수당','교원법정부담금','교원퇴직금및퇴직적립금','교원퇴직적립금',
  '직원급여','직원수당','직원법정부담금','직원퇴직금및퇴직적립금','직원퇴직적립금',
  '그밖의인건비','통학차량이용비','일반급식비간식비','방과후특성화비','방과후교육돌봄비',
  '수용비','수수료및제세공과금','연료비','여비','일반업무추진비','직책급업무추진비','적립금','시설비','취득비','유지비'
];
const REVIEW_TARGET_MOKS = ['교원급여','교원수당','직원급여','직원수당'];
const DEDUCT_WORDS = ['소득세','주민세','본인부담금','공제','실수령','건강보험','장기요양','고용보험','사학연금','국민연금'];
const ALLOWANCE_WORDS = ['수당','보조','급식','식대','연구','교통','명절','상여','휴가','성과','처우','운전'];

// 기존 sheetToRows는 서식 때문에 16,384열까지 도는 파일에서 느리고 불안정해서 실제 값이 있는 범위 중심으로 재정의
function sheetToRows(sheet, maxRows = Infinity) {
  if (!sheet) return [];
  const cellKeys = Object.keys(sheet).filter(k => /^[A-Z]+\d+$/.test(k));
  if (!cellKeys.length) return [];
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
  for (const key of cellKeys) {
    const p = XLSX.utils.decode_cell(key);
    const cell = sheet[key];
    const value = cell ? (cell.w ?? cell.v ?? '') : '';
    if (String(value ?? '').trim() === '') continue;
    minR = Math.min(minR, p.r); minC = Math.min(minC, p.c);
    maxR = Math.max(maxR, p.r); maxC = Math.max(maxC, p.c);
  }
  if (maxR < 0) return [];
  if (Number.isFinite(maxRows)) maxR = Math.min(maxR, minR + maxRows - 1);
  const merges = sheet['!merges'] || [];
  const rows = [];
  for (let r = minR; r <= maxR; r++) {
    const row = [];
    for (let c = minC; c <= maxC; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      let cell = sheet[addr];
      if (!cell) {
        const merge = merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
        if (merge) cell = sheet[XLSX.utils.encode_cell(merge.s)];
      }
      row.push(cell ? (cell.w ?? cell.v ?? '') : '');
    }
    rows.push(row);
  }
  return rows;
}

function parsePayrollRows(rows) {
  const cleaned = rows.map(r => r.map(v => String(v ?? '').replace(/\n/g, ' ').trim()));
  const headerRow = findPayrollHeaderRow(cleaned);
  if (headerRow < 0) return parsePayrollRowsFallback(cleaned);
  const headers = buildPayrollHeaders(cleaned, headerRow);
  const roleCol = findHeaderCol(headers, ['직명','직위','직종']);
  const salaryCols = headers.map((h, i) => isSalaryHeader(h) ? i : -1).filter(i => i >= 0);
  const allowanceCols = headers.map((h, i) => isAllowanceHeader(h) ? i : -1).filter(i => i >= 0);

  const teacherSubtotal = findSubtotalRow(cleaned, headerRow, ['소계교원','교원소계']);
  const staffSubtotal = findSubtotalRow(cleaned, headerRow, ['소계일반직','일반직소계','소계직원','직원소계']);
  const items = [];

  // 1) 기본급은 직명별로 묶어서 오편성 추적이 가능하게 함(차량기사 등)
  const roleSalaryMap = new Map();
  for (let r = headerRow + 1; r < cleaned.length; r++) {
    const row = cleaned[r];
    const rowText = normalizeText(row.join(' '));
    if (!rowText || rowText.includes('소계') || rowText.includes('합계')) continue;
    const role = row[roleCol] || '';
    if (!role) continue;
    const group = staffSubtotal >= 0 && r > teacherSubtotal ? 'staffSalary' : 'teacherSalary';
    const roleKey = normalizeRole(role, group);
    const salary = salaryCols.reduce((acc, c) => acc + parseMoney(row[c]), 0);
    if (salary > 0) {
      const key = `${group}|${roleKey}`;
      roleSalaryMap.set(key, (roleSalaryMap.get(key) || 0) + salary);
    }
  }
  for (const [key, amount] of roleSalaryMap.entries()) {
    const [bucket, label] = key.split('|');
    items.push({ label: `${label} 급여`, amount, bucket, source: '직명별 기본급 합계' });
  }

  // 2) 수당은 소계 행의 열별 합계를 읽음. PDF에서 세부 목이 없으면 통합편성 판단에 사용
  addSubtotalAllowanceItems(items, cleaned, headers, teacherSubtotal, 'teacherAllowance');
  addSubtotalAllowanceItems(items, cleaned, headers, staffSubtotal, 'staffAllowance');

  // 너무 세분화된 급여가 PDF와 합산 비교될 수 있도록 bucket 총액도 보관
  return normalizePayrollItems(items.filter(i => i.amount > 0));
}

function parsePayrollRowsFallback(rows) {
  const items = [];
  rows.forEach((row, idx) => {
    const cells = row.map(v => String(v ?? '').trim()).filter(Boolean);
    if (!cells.length) return;
    const nums = cells.map(parseMoney).filter(n => n > 0);
    if (!nums.length) return;
    const label = cells.filter(c => !/^[-\d,\.\s원]+$/.test(c)).join(' ').replace(/\s+/g, ' ').trim();
    const amount = Math.max(...nums);
    const bucket = guessBucket(label);
    if (label && bucket) items.push({ label, amount, bucket, row: idx + 1, source: 'fallback' });
  });
  return normalizePayrollItems(items);
}

function findPayrollHeaderRow(rows) {
  let best = -1, score = 0;
  rows.forEach((row, i) => {
    const t = normalizeText(row.join(' '));
    let s = 0;
    if (t.includes('직명')) s += 3;
    if (t.includes('기본급') || t.includes('본봉')) s += 4;
    if (t.includes('지급액계')) s += 2;
    if (t.includes('급식') || t.includes('식대')) s += 1;
    if (s > score) { score = s; best = i; }
  });
  return score >= 5 ? best : -1;
}

function buildPayrollHeaders(rows, headerRow) {
  const maxLen = Math.max(...rows.map(r => r.length));
  const headers = [];
  for (let c = 0; c < maxLen; c++) {
    const parts = [];
    for (let r = Math.max(0, headerRow - 1); r <= Math.min(rows.length - 1, headerRow + 1); r++) {
      const v = rows[r]?.[c] || '';
      if (v) parts.push(v);
    }
    headers[c] = parts.join(' ').replace(/\s+/g, ' ').trim();
  }
  return headers;
}

function findHeaderCol(headers, keys) {
  let idx = headers.findIndex(h => keys.some(k => normalizeText(h).includes(normalizeText(k))));
  return idx >= 0 ? idx : 1;
}

function isSalaryHeader(h) {
  const n = normalizeText(h);
  return (n.includes('기본급') || n.includes('본봉')) && !DEDUCT_WORDS.some(w => n.includes(normalizeText(w)));
}

function isAllowanceHeader(h) {
  const n = normalizeText(h);
  if (!n || DEDUCT_WORDS.some(w => n.includes(normalizeText(w)))) return false;
  if (n.includes('지급액계')) return false;
  if (isSalaryHeader(h)) return false;
  return ALLOWANCE_WORDS.some(w => n.includes(normalizeText(w)));
}

function findSubtotalRow(rows, start, keys) {
  for (let i = start + 1; i < rows.length; i++) {
    const n = normalizeText(rows[i].join(' '));
    if (keys.some(k => n.includes(k))) return i;
  }
  return -1;
}

function addSubtotalAllowanceItems(items, rows, headers, subtotalRow, bucket) {
  if (subtotalRow < 0) return;
  const row = rows[subtotalRow];
  headers.forEach((h, c) => {
    if (!isAllowanceHeader(h)) return;
    const amount = parseMoney(row[c]);
    if (amount > 0) items.push({ label: cleanHeaderLabel(h), amount, bucket, source: '소계 행 열별 합계' });
  });
}

function cleanHeaderLabel(h) {
  return h.replace(/[A-Z]\)?/g, '').replace(/일련 번호|직명|성 명|호봉|경력/g, '').replace(/\s+/g, ' ').trim() || h;
}

function normalizeRole(role, group) {
  const n = normalizeText(role);
  if (n.includes('원장')) return '원장';
  if (n.includes('방과후교사')) return '방과후교사';
  if (n.includes('정교사') || n.includes('교사') && group === 'teacherSalary') return '교사';
  if (n.includes('교원')) return '교원';
  if (n.includes('차량기사')) return '차량기사';
  if (n.includes('차량보조')) return '차량보조';
  if (n.includes('방과후보조')) return '방과후보조';
  if (n.includes('보조교사')) return '보조교사';
  if (n.includes('조리')) return '조리직원';
  if (n.includes('영양')) return '영양사';
  if (n.includes('환경') || n.includes('청소')) return '환경미화원';
  if (n.includes('사무') || n.includes('행정')) return '사무직원';
  if (n.includes('관리')) return '관리직';
  return role.replace(/\d+/g, '').trim();
}

function parseBudgetEntries(text) {
  const lines = text.split(/\n/).map(s => s.trim()).filter(Boolean);
  const entries = [];
  let currentMok = '';
  let currentParent = '';
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/\s+/g, ' ').trim();
    const clean = normalizeText(line);
    const heading = detectBudgetHeading(clean);
    if (heading) {
      currentMok = heading;
      if (!REVIEW_TARGET_MOKS.includes(heading) && !heading.includes('퇴직')) currentParent = heading;
      continue;
    }
    const parent = detectParentHeading(clean);
    if (parent) currentParent = parent;

    if (line.includes('본예산') || line.includes('=') || /\d[\d,]*\s*원/.test(line)) {
      let block = line;
      let j = i + 1;
      while (j < lines.length && !/=[^=]*\d[\d,]*\s*$/.test(block) && !/\d{1,3}(?:,\d{3})+\s*$/.test(block) && j < i + 6) {
        const nextClean = normalizeText(lines[j]);
        if (detectBudgetHeading(nextClean)) break;
        block += ' ' + lines[j].replace(/\s+/g, ' ').trim();
        j++;
      }
      const amount = extractFinalAmount(block);
      if (amount > 0) {
        const detail = extractBudgetDetail(block);
        entries.push({ mok: currentMok, parent: currentParent || currentMok, detail: normalizeBudgetDetail(detail || block), formula: block.replace(/\s+/g, ' ').trim(), amount, pageHint: findPageHint(lines, i) });
      }
    }
  }
  return mergeSplitBudgetEntries(entries);
}

function detectBudgetHeading(clean) {
  // 숫자가 섞인 산출내역 줄은 제목으로 보지 않음
  if (/\d/.test(clean)) return '';
  let found = '';
  for (const h of ALL_BUDGET_MOK_HEADINGS) {
    const nh = normalizeText(h);
    if (clean === nh || clean.startsWith(nh)) {
      if (!found || nh.length > normalizeText(found).length) found = h;
    }
  }
  return found;
}

function detectParentHeading(clean) {
  const parents = ['인건비','교원인건비','직원인건비','운영비','관리운영비','일반교육활동비','선택적교육활동비','그밖의교육활동비','시설설비비품비','잡지출','예비비및반환금'];
  return parents.find(p => clean === normalizeText(p) || clean.startsWith(normalizeText(p))) || '';
}

function extractBudgetDetail(block) {
  let left = block;
  const wonIdx = left.search(/\d[\d,]*\s*원/);
  if (wonIdx >= 0) left = left.slice(0, wonIdx);
  left = left.replace(/\(본예산\)/g, '').replace(/\(보조금및지원금\)|\(수익자부담금\)|\(그밖의수입\)/g, '').trim();
  return left;
}

function review(formData, budgetData) {
  const issues = [];
  const notices = [];
  const entries = budgetData.entries || [];
  const payroll = formData.payroll || [];

  const hasRetirementCarryover = (formData.retirement?.carryoverTotal || 0) > 0;
  const hasTeacherRetirement = entries.some(e => normalizeText(e.mok + e.detail).includes('교원퇴직'));
  const hasStaffRetirement = entries.some(e => normalizeText(e.mok + e.detail).includes('직원퇴직'));
  if (hasRetirementCarryover && !(hasTeacherRetirement || hasStaffRetirement)) {
    issues.push({ type: 'danger', title: '퇴직적립금 미편성', text: `퇴직금 적립금 이월액 계가 ${fmt(formData.retirement.carryoverTotal)}원으로 인식되었으나, 세출예산명세서에 교원퇴직적립금 또는 직원퇴직적립금 편성 항목을 찾지 못했습니다.` });
  }

  for (const bucket of ['teacherSalary','teacherAllowance','staffSalary','staffAllowance']) {
    const mok = bucketToMok(bucket);
    const group = payroll.filter(p => p.bucket === bucket);
    const expectedTotal = group.reduce((a, p) => a + p.amount, 0);
    if (!expectedTotal) continue;
    const normalEntries = entries.filter(e => normalizeText(e.mok).includes(normalizeText(mok)));
    const normalTotal = normalEntries.reduce((a, e) => a + e.amount, 0);

    // 통합편성 안내: 일반 산출내역명(교원수당/직원수당 등) 하나가 엑셀의 여러 열 합계와 일치하는 경우
    const generic = findGenericIntegratedEntry(normalEntries, group, mok);
    if (generic) notices.push({ type: 'warn', title: `${mok} 통합편성`, text: `${generic.labels} = ${fmt(generic.sum)}원 → PDF '${generic.entry.detail}' ${fmt(generic.entry.amount)}원으로 통합편성된 것으로 보입니다.` });

    if (isSameMoney(expectedTotal, normalTotal)) continue;

    const diff = Math.abs(expectedTotal - normalTotal);
    const misplacedCombo = findMisplacedCombo(entries, mok, diff, group);
    if (misplacedCombo.length) {
      const labelGuess = guessIssueLabel(group, diff, bucket);
      const locations = misplacedCombo.map(e => `${e.parent || e.mok || '다른 목'}의 '${e.detail}' ${fmt(e.amount)}원`).join(' + ');
      issues.push({ type: 'danger', title: '오편성 의심', text: `${labelGuess} ${fmt(diff)}원이 ${mok}이 아닌 ${locations}에 편성된 것으로 보입니다.` });
    } else {
      issues.push({ type: 'danger', title: `${mok} 금액 불일치`, text: `엑셀 보수일람표 기준 ${mok} 합계는 ${fmt(expectedTotal)}원이나, PDF ${mok} 산출내역 합계는 ${fmt(normalTotal)}원으로 인식되었습니다. 차액 ${fmt(diff)}원을 확인해야 합니다.` });
    }
  }

  const allIssues = [...issues, ...dedupeNotices(notices)];
  if (!allIssues.length) allIssues.push({ type: 'ok', title: '지적사항 없음', text: '현재 1차 검토 기준에서 지적사항을 찾지 못했습니다.' });
  return { issues: allIssues, counts: { payrollItems: payroll.length, budgetEntries: entries.length, issues: issues.length } };
}

function findGenericIntegratedEntry(normalEntries, group, mok) {
  const genericEntries = normalEntries.filter(e => {
    const d = normalizeText(e.detail);
    const m = normalizeText(mok);
    return d === m || d.endsWith(m) || d.includes(m);
  });
  if (!genericEntries.length || group.length < 2) return null;
  // 명절/급식처럼 PDF에 개별 항목이 있는 것은 제외하고, 남은 항목 합계가 generic과 맞는지 확인
  for (const entry of genericEntries) {
    const explicitLabels = new Set(normalEntries.filter(e => e !== entry).map(e => normalizeText(e.detail)));
    const candidates = group.filter(p => !Array.from(explicitLabels).some(d => similarLabel(normalizeText(p.label), d)));
    const sum = candidates.reduce((a, p) => a + p.amount, 0);
    if (candidates.length >= 2 && isSameMoney(sum, entry.amount)) {
      return { entry, sum, labels: candidates.map(p => `${p.label} ${fmt(p.amount)}원`).join(' + ') };
    }
  }
  return null;
}

function dedupeNotices(notices) {
  const seen = new Set();
  return notices.filter(n => { const k = n.title + n.text; if (seen.has(k)) return false; seen.add(k); return true; });
}

function findMisplacedCombo(entries, expectedMok, diff, group) {
  const nExpected = normalizeText(expectedMok);
  const keyWords = group.map(g => normalizeText(g.label)).join(' ');
  let candidates = entries.filter(e => !normalizeText(e.mok).includes(nExpected) && !REVIEW_TARGET_MOKS.some(m => normalizeText(e.mok).includes(normalizeText(m))));
  candidates = candidates.filter(e => isSameMoney(e.amount, diff) || similarLabel(keyWords, normalizeText(e.detail)) || normalizeText(e.detail + e.parent).includes('차량'));
  const direct = candidates.find(e => isSameMoney(e.amount, diff));
  if (direct) return [direct];
  // 최대 4개까지 합산 탐색
  const arr = candidates.slice(0, 30);
  for (let size = 2; size <= 4; size++) {
    const combo = findComboSum(arr, diff, size, 0, []);
    if (combo) return combo;
  }
  return [];
}

function findComboSum(arr, target, size, start, picked) {
  if (picked.length === size) {
    const sum = picked.reduce((a, e) => a + e.amount, 0);
    return isSameMoney(sum, target) ? picked : null;
  }
  for (let i = start; i < arr.length; i++) {
    const res = findComboSum(arr, target, size, i + 1, [...picked, arr[i]]);
    if (res) return res;
  }
  return null;
}

function guessIssueLabel(group, diff, bucket) {
  const exact = group.find(p => isSameMoney(p.amount, diff));
  if (exact) return exact.label;
  if (bucket === 'staffSalary') {
    const vehicle = group.filter(p => normalizeText(p.label).includes('차량'));
    const vSum = vehicle.reduce((a, p) => a + p.amount, 0);
    if (vehicle.length && isSameMoney(vSum, diff)) return '차량기사 급여';
  }
  return bucketToMok(bucket);
}
