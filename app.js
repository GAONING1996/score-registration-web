(() => {
  const $ = (id) => document.getElementById(id);
  const state = {
    workbook: null, fileName: '', sheetName: '', scoreCol: 4,
    headers: [], rows: [], className: '', candidates: [], history: [], orphan: [],
    recognitionRunning: false, recognition: null, recognitionStartedAt: 0,
    voiceBridge: false, voiceCursor: 0, voicePollTimer: null,
    voiceHelperOpened: false, voiceHelperEverConnected: false,
    lastSpeechAt: 0, silenceTimer: null, sessionTimer: null, pendingVoice: false,
  };
  const el = {
    file: $('fileInput'), fileName: $('fileName'), sheet: $('sheetSelect'),
    cls: $('classSelect'), col: $('scoreColSelect'), status: $('fileStatus'),
    name: $('nameInput'), score: $('scoreInput'), clear: $('clearBtn'),
    mic: $('micBtn'), transcript: $('transcript'), candidates: $('candidateList'),
    scope: $('scopeText'), undo: $('undoBtn'), total: $('totalN'), done: $('doneN'),
    todo: $('todoN'), pending: $('pendingN'), summary: $('classSummary'),
    export: $('exportBtn'), unreg: $('unregisteredWrap'), registered: $('registeredWrap'), orphan: $('orphanWrap'),
    backdrop: $('modalBackdrop'), modalTitle: $('modalTitle'), modalText: $('modalText'),
    modalBody: $('modalBody'), modalCancel: $('modalCancel'), modalConfirm: $('modalConfirm'),
    toast: $('toast'),
  };
  const text = (v) => String(v ?? '').trim();
  const normalize = (v) => text(v).toLowerCase().replace(/[\s·•・,，。．、'"“”‘’（）()\[\]【】]/g, '');
  const escapeHtml = (v) => text(v).replace(/[&<>"']/g, (c) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
  const colName = (i) => { let out = ''; let n = i + 1; while (n) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); } return out; };
  const pinyin = (v, opts) => { try { return window.pinyinPro?.pinyin(v, opts) || ''; } catch { return ''; } };
  const fullPinyin = (v) => pinyin(v, { toneType: 'none' }).replace(/\s/g, '').toLowerCase();
  const initials = (v) => pinyin(v, { pattern: 'first', toneType: 'none', type: 'array' }).join('').toLowerCase().replace(/[^a-z]/g, '');
  const distance = (a, b) => {
    a = normalize(a); b = normalize(b); if (!a || !b) return 99;
    const row = Array.from({ length: a.length + 1 }, (_, i) => i);
    for (let j = 1; j <= b.length; j += 1) { let prev = row[0]; row[0] = j; for (let i = 1; i <= a.length; i += 1) { const cur = row[i]; row[i] = Math.min(row[i] + 1, row[i - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1)); prev = cur; } }
    return row[a.length];
  };
  const chineseNumber = (v) => {
    const map = { 零:0, 〇:0, 一:1, 二:2, 两:2, 三:3, 四:4, 五:5, 六:6, 七:7, 八:8, 九:9, 十:10 };
    if (!v) return null; if (/^\d+$/.test(v)) return Number(v);
    let total = 0; let current = 0;
    for (const ch of v) { if (map[ch] === undefined) return null; if (map[ch] === 10) { total += (current || 1) * 10; current = 0; } else current = map[ch]; }
    return total + current;
  };
  const scoreValue = (v) => {
    if (v === null || v === undefined || text(v) === '') return null;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const s = text(v).replace(/分$/, '').replace(/点/g, '.').replace(/，/g, '.');
    const match = s.match(/-?\d+(?:\.\d+)?/); if (match) return Number(match[0]);
    if (s.includes('.')) { const [whole, dec] = s.split('.'); const w = chineseNumber(whole); const d = [...dec].map((ch) => chineseNumber(ch)).join(''); return w === null || !d ? null : Number(`${w}.${d}`); }
    return chineseNumber(s);
  };
  const parseSpeech = (v) => {
    const raw = text(v).replace(/(这个是|这是|同学|的成绩|的分数|成绩是|分数是|分数为|得分|分数|成绩|考了|考得|分$)/g, ' ').replace(/\s+/g, ' ').trim();
    const nums = raw.match(/\d+(?:\.\d+)?|[零〇一二两三四五六七八九十百点]+/g);
    const last = nums?.[nums.length - 1]; const index = last ? raw.lastIndexOf(last) : -1;
    return { name: index >= 0 ? raw.slice(0, index).replace(/[，,：:；;\s]+$/, '').trim() : raw, score: last ? scoreValue(last) : null };
  };
  const toast = (msg) => { el.toast.textContent = msg; el.toast.classList.add('show'); clearTimeout(toast.timer); toast.timer = setTimeout(() => el.toast.classList.remove('show'), 2600); };
  const currentRows = () => state.rows.filter((r) => r.className === state.className);
  const isDone = (r) => text(r.values[state.scoreCol]) !== '';
  const setEnabled = (yes) => [el.name, el.score, el.clear, el.mic, el.export].forEach((node) => { node.disabled = !yes; });
  const scoreColumns = () => state.headers.map((h, i) => ({ index: i, label: `${colName(i)}列${text(h) ? `（${text(h)}）` : ''}`, hasData: state.rows.some((r) => text(r.values[i]) !== ''), empty: !state.rows.some((r) => text(r.values[i]) !== '') }));
  const defaultColumn = () => { const after = Math.max(state.nameCol || 1, state.classCol || 2, state.examCol || 3) + 1; const cols = scoreColumns(); return cols.find((c) => c.index === after && c.empty)?.index ?? cols.find((c) => c.empty)?.index ?? after; };
  const buildRows = (sheetName) => {
    const matrix = XLSX.utils.sheet_to_json(state.workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
    const headerRow = matrix.findIndex((r) => r.some((x) => text(x) === '姓名') && r.some((x) => text(x) === '班级'));
    if (headerRow < 0) throw new Error(`工作表“${sheetName}”没有找到姓名和班级表头`);
    state.headers = matrix[headerRow] || []; state.headerRow = headerRow;
    state.nameCol = state.headers.findIndex((x) => text(x) === '姓名'); state.classCol = state.headers.findIndex((x) => text(x) === '班级'); state.examCol = state.headers.findIndex((x) => text(x) === '考试号');
    state.rows = matrix.slice(headerRow + 1).map((values, offset) => ({ values, rowIndex: headerRow + 1 + offset, name: text(values[state.nameCol]), className: text(values[state.classCol]), exam: text(values[state.examCol]), pinyin: fullPinyin(text(values[state.nameCol])), initials: initials(text(values[state.nameCol])) })).filter((r) => r.name || r.className);
  };
  const refreshStats = () => {
    const rows = currentRows(); const doneCount = rows.filter(isDone).length;
    el.total.textContent = rows.length; el.done.textContent = doneCount; el.todo.textContent = rows.length - doneCount; el.pending.textContent = doneCount;
    el.summary.textContent = `${state.className}班 · ${rows.length}人`; el.scope.textContent = `${state.sheetName} · ${state.className}班`;
    const last = state.history[state.history.length - 1]; el.undo.disabled = !last; el.undo.textContent = last ? `撤销上一条：${last.row.name}（${last.row.className}班）→ ${last.newValue}` : '撤销上一条：暂无';
    const pending = rows.filter((r) => !isDone(r));
    el.unreg.innerHTML = pending.length ? `<table><thead><tr><th>姓名</th><th>考试号</th><th>状态</th></tr></thead><tbody>${pending.map((r) => `<tr><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.exam)}</td><td><span class="tag">未登记</span></td></tr>`).join('')}</tbody></table>` : '<div class="empty">本班学生已全部登记</div>';
    const registered = rows.filter(isDone);
    el.registered.innerHTML = registered.length ? `<table><thead><tr><th>姓名</th><th>考试号</th><th>成绩</th></tr></thead><tbody>${registered.map((r) => `<tr><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.exam || '—')}</td><td>${escapeHtml(r.values[state.scoreCol])}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">暂无已登记学生</div>';
    el.orphan.innerHTML = state.orphan.length ? `<table><thead><tr><th>成绩</th><th>原始输入</th><th>原因</th></tr></thead><tbody>${state.orphan.map((o) => `<tr><td>${escapeHtml(o.value ?? '—')}</td><td>${escapeHtml(o.raw)}</td><td>${escapeHtml(o.reason)}</td></tr>`).join('')}</tbody></table>` : '<div class="empty">没有无法匹配记录</div>';
  };
  const findCandidates = (query) => {
    const q = normalize(query); const qp = fullPinyin(query); const qi = initials(query); if (!q) return [];
    return currentRows().map((r) => {
      const name = normalize(r.name); let score = 0; let reason = '';
      if (name === q) { score = 100; reason = '姓名完全匹配'; } else if (name.includes(q)) { score = 92 - (name.length - q.length); reason = '姓名包含'; } else if ([...q].every((ch) => name.includes(ch))) { score = 78; reason = '姓名分字匹配'; } else if (qp && r.pinyin.includes(qp)) { score = 86; reason = '拼音匹配'; } else if (qi && r.initials.includes(qi)) { score = 82 - Math.max(0, r.initials.length - qi.length); reason = '拼音首字母'; } else { const d = distance(q, name); const di = qi ? distance(qi, r.initials) : 99; if (d <= 2) { score = 72 - d * 8; reason = '姓名相似'; } else if (di <= 1) { score = 68 - di * 8; reason = '首字母相似'; } }
      const rowDone = isDone(r); if (score > 0 && !rowDone) score += 2; return { ...r, score, reason, isDone: rowDone };
    }).filter((r) => r.score > 0).sort((a, b) => b.score - a.score || Number(a.isDone) - Number(b.isDone)).slice(0, 5);
  };
  const renderCandidates = (list) => {
    state.candidates = list;
    if (!list.length) { el.candidates.innerHTML = '<div class="empty">没有找到可靠候选。可以改用姓名中的某个字、拼音首字母，或选择“不登记”。</div>'; return; }
    el.candidates.innerHTML = list.map((r, i) => `<div class="candidate ${i === 0 ? 'top' : ''}"><div><div class="name">${escapeHtml(r.name)}</div><div class="meta">${escapeHtml(r.className)}班 · 考试号 ${escapeHtml(r.exam || '—')} · ${escapeHtml(r.reason)} · 匹配 ${r.score}</div></div><div class="right"><span class="mini-status ${r.isDone ? 'done' : 'todo'}">${r.isDone ? '已登记' : '未登记'}</span>${r.isDone ? `<span class="score">${escapeHtml(r.values[state.scoreCol])}</span><button class="btn light duplicate" data-index="${i}">处理重复</button>` : `<button class="btn primary choose" data-index="${i}">登记</button>`}</div></div>`).join('') + '<div style="text-align:right"><button class="btn light" id="noRegister">不登记</button></div>';
    document.querySelectorAll('.choose').forEach((button) => button.addEventListener('click', () => commit(state.candidates[Number(button.dataset.index)])));
    document.querySelectorAll('.duplicate').forEach((button) => button.addEventListener('click', () => duplicate(state.candidates[Number(button.dataset.index)], scoreValue(el.score.value))));
    $('noRegister')?.addEventListener('click', () => orphan('用户选择不登记'));
  };
  const clearInputs = () => { el.name.value = ''; el.score.value = ''; state.candidates = []; el.candidates.innerHTML = '<div class="empty">等待新的姓名或语音输入</div>'; state.pendingVoice = false; el.name.focus(); el.transcript.textContent = state.recognitionRunning ? '语音状态：正在监听下一张试卷' : '语音转写：尚未开始'; };
  const voiceToken = new URLSearchParams(location.search).get('token') || '';
  const voiceApi = (path) => `${path}${path.includes('?') ? '&' : '?'}token=${encodeURIComponent(voiceToken)}`;
  const sendVoiceControl = (action) => fetch(voiceApi('/voice/control'), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action }) }).catch(() => {});
  const handleSpeechFinal = (finalText) => {
    const parsed = parseSpeech(finalText);
    if (!parsed.name || parsed.score === null) {
      el.transcript.textContent = state.pendingVoice ? '语音状态：本句不完整，保留当前待确认内容' : '语音状态：姓名或成绩不完整，请补充或选择不登记';
      return;
    }
    const replaced = state.pendingVoice;
    state.pendingVoice = true;
    el.name.value = parsed.name;
    el.score.value = parsed.score;
    renderCandidates(findCandidates(parsed.name));
    el.transcript.textContent = replaced ? `语音状态：已覆盖上一条待确认内容：${parsed.name} → ${parsed.score}` : `语音状态：等待确认：${parsed.name} → ${parsed.score}`;
  };
  const startBrowserVoiceBridge = async () => {
    state.voiceBridge = true;
    state.recognitionRunning = true;
    state.voiceCursor = 0;
    el.mic.textContent = '停止语音';
    el.transcript.textContent = state.voiceHelperOpened ? '语音状态：正在连接已打开的浏览器语音窗口' : '语音状态：首次使用将打开一次浏览器语音窗口';
    await sendVoiceControl('start');
    if (!state.voiceHelperOpened) {
      state.voiceHelperEverConnected = false;
      try {
        await window.desktopAPI.openVoiceHelper();
        state.voiceHelperOpened = true;
      } catch (error) {
        state.voiceBridge = false;
        state.recognitionRunning = false;
        el.mic.textContent = '开始连续语音';
        el.transcript.textContent = '语音状态：无法打开浏览器语音窗口';
        toast(error?.message || '无法打开浏览器语音窗口');
        return;
      }
    }
    if (state.voicePollTimer) clearInterval(state.voicePollTimer);
    state.voicePollTimer = setInterval(async () => {
      try {
        const response = await fetch(voiceApi(`/voice/events?after=${state.voiceCursor}`), { cache: 'no-store' });
        const data = await response.json();
        if (data.helperConnected) state.voiceHelperEverConnected = true;
        if (state.voiceHelperEverConnected && !data.helperConnected && state.voiceBridge) {
          state.voiceHelperOpened = false;
          state.voiceHelperEverConnected = false;
          stopRecognition('语音状态：浏览器语音窗口已关闭，下次开始语音时会重新打开');
          return;
        }
        for (const event of data.events || []) {
          state.voiceCursor = Math.max(state.voiceCursor, event.id || 0);
          if (event.type === 'result') {
            const display = event.finalText || event.interim;
            if (display) el.transcript.textContent = `语音转写：${display}`;
            if (event.finalText) handleSpeechFinal(event.finalText);
          } else if (event.type === 'error') {
            el.transcript.textContent = `浏览器语音识别失败：${event.error || '未知错误'}`;
            if (event.error === 'network') toast('浏览器语音服务网络连接失败，请检查网络或改用文字输入');
          } else if (event.type === 'status' && event.message) {
            el.transcript.textContent = `语音状态：${event.message}`;
          }
        }
        if (!data.active && state.voiceBridge) stopRecognition('语音状态：已停止监听，可继续录入下一张试卷');
      } catch {}
    }, 180);
  };
  const openModal = (confirm) => { el.backdrop.style.display = 'flex'; el.modalConfirm.onclick = () => { el.backdrop.style.display = 'none'; confirm(); }; el.modalCancel.onclick = () => { el.backdrop.style.display = 'none'; }; };
  const orphan = (reason) => { const value = scoreValue(el.score.value); if (value === null) { toast('请先填写有效成绩'); return; } state.orphan.push({ value, raw: el.name.value || '无姓名', reason, timestamp: Date.now() }); refreshStats(); clearInputs(); toast(`已记录无法匹配内容：${value}`); };
  const duplicate = (row, value) => { if (!row) return; if (value === null) { toast('请先填写本次新成绩'); return; } el.modalTitle.textContent = '发现重复登记'; el.modalText.textContent = `${row.name}（${row.className}班，考试号 ${row.exam || '—'}）已有成绩：${row.values[state.scoreCol]}。本次识别为：${value}。`; el.modalBody.innerHTML = '<div class="field"><label>处理方式</label><select id="duplicateChoice"><option value="keep">保留原成绩</option><option value="replace">使用新成绩</option><option value="other">重新选择候选人</option><option value="skip">不登记本次记录</option></select></div>'; openModal(() => { const choice = $('duplicateChoice').value; if (choice === 'replace') { state.history.push({ row, oldValue: row.values[state.scoreCol], newValue: value, column: state.scoreCol, timestamp: Date.now() }); row.values[state.scoreCol] = value; refreshStats(); clearInputs(); resumeRecognition(); toast(`已替换 ${row.name} 的成绩`); } else if (choice === 'other') renderCandidates(findCandidates(el.name.value).filter((candidate) => candidate.name !== row.name)); else if (choice === 'skip') { orphan('重复记录未登记'); resumeRecognition(); } else resumeRecognition(); }); };
  const commit = (row) => { const value = scoreValue(el.score.value); if (!row || value === null) { toast('请先填写有效成绩'); return; } if (isDone(row)) { duplicate(row, value); return; } state.history.push({ row, oldValue: row.values[state.scoreCol], newValue: value, column: state.scoreCol, timestamp: Date.now() }); row.values[state.scoreCol] = value; el.transcript.textContent = `已登记：${row.name}（${row.className}班）→ ${value}`; refreshStats(); clearInputs(); resumeRecognition(); toast(`已登记 ${row.name}：${value}`); };
  el.file.addEventListener('change', async (event) => { const file = event.target.files?.[0]; if (!file) return; try { state.workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellStyles: true }); state.fileName = file.name; el.fileName.textContent = file.name; el.sheet.innerHTML = state.workbook.SheetNames.map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join(''); el.sheet.disabled = false; el.sheet.dispatchEvent(new Event('change')); el.status.innerHTML = `<span class="tag ok">已读取 ${state.workbook.SheetNames.length} 张工作表</span>`; toast('Excel 已读取，未修改原文件'); } catch (error) { console.error(error); toast(`读取 Excel 失败：${error.message}`); } });
  el.sheet.addEventListener('change', () => { state.sheetName = el.sheet.value; state.scoreCol = 4; buildRows(state.sheetName); state.scoreCol = defaultColumn(); el.col.innerHTML = scoreColumns().map((c) => `<option value="${c.index}">${escapeHtml(c.label)}${c.hasData ? ' · 已有数据' : ''}</option>`).join(''); el.col.value = String(state.scoreCol); el.col.disabled = false; const classes = [...new Set(state.rows.map((r) => r.className).filter(Boolean))]; el.cls.innerHTML = classes.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}班</option>`).join(''); el.cls.disabled = !classes.length; if (classes.length) el.cls.dispatchEvent(new Event('change')); });
  el.col.addEventListener('change', () => { state.scoreCol = Number(el.col.value); refreshStats(); });
  el.cls.addEventListener('change', () => { state.className = el.cls.value; state.history = []; state.orphan = []; setEnabled(true); refreshStats(); clearInputs(); });
  let searchTimer = null;
  const liveSearch = () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => { const query = text(el.name.value); renderCandidates(query ? findCandidates(query) : []); }, 120); };
  el.clear.addEventListener('click', () => { clearInputs(); toast('已清除当前录入内容'); });
  el.name.addEventListener('input', liveSearch);
  [el.name, el.score].forEach((input) => input.addEventListener('keydown', (event) => { if (event.key === 'Enter') { event.preventDefault(); liveSearch(); } }));
  el.undo.addEventListener('click', () => { const item = state.history.pop(); if (!item) return; item.row.values[item.column ?? state.scoreCol] = item.oldValue; refreshStats(); toast(`已撤销 ${item.row.name} 的登记：恢复为 ${item.oldValue || '空白'}`); });
  el.export.addEventListener('click', () => { if (!state.workbook) return; const sheet = state.workbook.Sheets[state.sheetName]; state.rows.forEach((row) => { const address = XLSX.utils.encode_cell({ r: row.rowIndex, c: state.scoreCol }); sheet[address] = { t: typeof row.values[state.scoreCol] === 'number' ? 'n' : 's', v: row.values[state.scoreCol] }; }); const output = XLSX.write(state.workbook, { bookType: 'xlsx', type: 'array', cellStyles: true }); const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); link.download = state.fileName.replace(/\.(xlsx|xls)$/i, '') + '_登记完成.xlsx'; link.click(); setTimeout(() => URL.revokeObjectURL(link.href), 1000); toast('已下载登记后的 Excel'); });
  const resumeRecognition = () => { if (!state.recognitionRunning) return; state.pendingVoice = false; el.transcript.textContent = '语音状态：正在监听下一张试卷'; if (state.recognition) { try { state.recognition.start(); } catch {} } };
  const stopRecognition = (message = '语音状态：已停止') => {
    clearTimeout(state.silenceTimer); clearTimeout(state.sessionTimer);
    if (state.voicePollTimer) { clearInterval(state.voicePollTimer); state.voicePollTimer = null; }
    if (state.voiceBridge) sendVoiceControl('stop');
    state.voiceBridge = false;
    if (state.recognition) { try { state.recognition.onend = null; state.recognition.stop(); } catch {} }
    state.recognition = null; state.recognitionRunning = false; state.pendingVoice = false;
    el.mic.textContent = '开始连续语音'; el.transcript.textContent = message;
  };
  const startRecognition = () => {
    if (state.recognitionRunning) { stopRecognition(); return; }
    if (window.desktopAPI?.isElectron) { startBrowserVoiceBridge(); return; }
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) { toast('当前浏览器不支持语音识别，请使用最新版 Chrome 或 Edge'); return; }
    const recognition = new SpeechRecognition();
    state.recognition = recognition;
    recognition.lang = 'zh-CN';
    recognition.continuous = true;
    recognition.interimResults = true;
    state.recognitionRunning = true;
    state.recognitionStartedAt = Date.now();
    state.lastSpeechAt = Date.now();
    el.mic.textContent = '停止语音';
    el.transcript.textContent = '语音状态：正在监听（首次使用请允许麦克风）';
    const resetSessionTimer = () => {
      clearTimeout(state.sessionTimer);
      state.sessionTimer = setTimeout(() => {
        if (state.recognitionRunning) stopRecognition('语音状态：30 秒没有收到语音识别结果，已自动停止');
      }, 30000);
    };
    const resetSilence = () => {
      state.lastSpeechAt = Date.now();
      clearTimeout(state.silenceTimer);
      state.silenceTimer = setTimeout(() => {
        if (state.recognitionRunning) el.transcript.textContent = '语音状态：尚未收到转写，仍在等待；如果已说话，请检查输入设备或语音服务网络';
      }, 10000);
    };
    recognition.onstart = resetSilence;
    recognition.onresult = (event) => {
      resetSilence();
      resetSessionTimer();
      let finalText = '';
      let interim = '';
      for (let i = event.resultIndex; i < event.results.length; i += 1) {
        const result = event.results[i];
        if (result.isFinal) finalText += result[0].transcript;
        else interim += result[0].transcript;
      }
      el.transcript.textContent = `语音转写：${finalText || interim}`;
      if (!finalText) return;
      const parsed = parseSpeech(finalText);
      if (!parsed.name || parsed.score === null) {
        el.transcript.textContent = state.pendingVoice ? '语音状态：本句不完整，保留当前待确认内容' : '语音状态：姓名或成绩不完整，请补充或选择不登记';
        return;
      }
      const replaced = state.pendingVoice;
      state.pendingVoice = true;
      el.name.value = parsed.name;
      el.score.value = parsed.score;
      renderCandidates(findCandidates(parsed.name));
      el.transcript.textContent = replaced ? `语音状态：已覆盖上一条待确认内容：${parsed.name} → ${parsed.score}` : `语音状态：等待确认：${parsed.name} → ${parsed.score}`;
    };
    recognition.onerror = (event) => {
      clearTimeout(state.silenceTimer);
      el.transcript.textContent = `语音识别失败：${event.error}`;
      if (event.error === 'not-allowed') stopRecognition('语音状态：麦克风权限未允许，请检查地址栏权限');
      else if (event.error === 'service-not-allowed') stopRecognition('语音状态：浏览器或网络不允许使用语音识别，请检查浏览器策略和语音服务网络');
      else if (event.error === 'network') stopRecognition('语音状态：浏览器语音服务网络连接失败，已停止自动重试；请检查网络后重新开始，或改用文字输入');
      else if (event.error === 'audio-capture') stopRecognition('语音状态：麦克风设备不可用');
      else if (event.error !== 'no-speech') toast(`语音识别失败：${event.error}`);
    };
    recognition.onend = () => {
      if (state.recognitionRunning) setTimeout(() => { if (state.recognitionRunning && state.recognition === recognition) { try { recognition.start(); } catch {} } }, 150);
    };
    try { recognition.start(); } catch (error) { stopRecognition('语音状态：无法启动识别，请稍后重试'); toast(error.message || '语音识别启动失败'); return; }
    resetSessionTimer();
  };
  el.mic.addEventListener('click', startRecognition);
  renderCandidates([]);
})();
