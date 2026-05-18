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
