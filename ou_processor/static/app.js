'use strict';
const $ = id => document.getElementById(id);
let files = [], config = null, busy = false, currentJob = null, timer = null, registryTimer = null, licensesTimer = null;
const sheets = new Map();
let manualPending = 0;
function showError(message) { $('error').textContent = message; $('error').hidden = !message; }
function setBusy(value) { busy = value; $('controls').disabled = value; updateButton(); }
function updateButton() {
  $('start').disabled = busy || !files.length || !config;
  $('check-dadata').disabled = busy || !files.length || !config;
  $('manual-file').disabled = busy || !currentJob || !manualPending;
  $('start').firstChild.textContent = !files.length ? 'Сначала загрузите Excel ' : 'Обработать по ТЗ ';
}
function updateConnection() {
  $('connection').textContent = $('token').value.trim() ? 'Ключ для каскада введён'
    : config?.token_saved ? 'Ключ сохранён в системе'
    : config?.token_configured ? 'Ключ для каскада задан на сервере' : 'Без ключа: только локальные источники';
  $('cascade-note').textContent = $('token').value.trim() || config?.token_configured
    ? 'DaData автоматически уточнит реквизиты по ИНН или предложит кандидатов для ручной проверки. ФИО и номера дел остаются на компьютере.'
    : 'Без ключа DaData проверка ограничена локальными источниками. Неразрешённые случаи попадут на ручную проверку.';
}
$('token').addEventListener('input', updateConnection);
async function addFiles(incoming) {
  if (busy) return;
  showError('');
  const next = [...files];
  for (const file of incoming) {
    if (!file.name.toLowerCase().endsWith('.xlsx') || file.name.startsWith('~$')) return showError('Выберите книгу Excel в формате .xlsx.');
    if (!next.some(x => x.name === file.name && x.size === file.size && x.lastModified === file.lastModified)) next.push(file);
  }
  if (next.length > 10 || next.reduce((n, file) => n + file.size, 0) > 100 * 1024 * 1024) return showError('Можно загрузить до 10 файлов общим размером не больше 100 МБ.');
  files = next; renderFiles();
  if (!config) return;
  setBusy(true);
  try {
    const pending = files.filter(file => !sheets.has(file));
    if (pending.length) {
      const info = await request('/api/sheets', {method:'POST', headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf}, body:JSON.stringify({files:await Promise.all(pending.map(encodeFile))})});
      for (const [index, file] of pending.entries()) {
        const entry = info[index];
        if (entry) sheets.set(file, {...entry, selected:''});
      }
    }
    renderFiles();
  } catch (error) { showError(error.message); }
  finally { setBusy(false); }
}
function renderFiles() {
  $('file-list').replaceChildren();
  files.forEach((file, index) => {
    const li = document.createElement('li'), name = document.createElement('span'), size = document.createElement('small'), remove = document.createElement('button');
    name.textContent = file.name; size.textContent = Math.max(1, Math.round(file.size / 1024)) + ' КБ';
    remove.type = 'button'; remove.className = 'remove'; remove.textContent = '×'; remove.setAttribute('aria-label', 'Убрать ' + file.name);
    remove.onclick = () => { sheets.delete(file); files.splice(index, 1); renderFiles(); };
    li.append(name, size, remove);
    const info = sheets.get(file);
    if (info) {
      const label = document.createElement('label'), select = document.createElement('select');
      label.textContent = 'Рабочий лист'; select.setAttribute('aria-label', 'Рабочий лист: ' + file.name);
      for (const title of ['', ...info.sheets]) {
        const option = document.createElement('option'); option.value = title; option.textContent = title || 'Определить автоматически'; select.append(option);
      }
      select.value = info.selected; select.onchange = () => { info.selected = select.value; };
      label.append(select); li.append(label);
    }
    $('file-list').append(li);
  });
  updateButton();
}
$('files').onchange = e => { addFiles(e.target.files); e.target.value = ''; };
$('drop').onkeydown = e => { if (!busy && ['Enter', ' '].includes(e.key)) { e.preventDefault(); $('files').click(); } };
['dragenter', 'dragover'].forEach(event => $('drop').addEventListener(event, e => { e.preventDefault(); if (!busy) $('drop').classList.add('drag'); }));
['dragleave', 'drop'].forEach(event => $('drop').addEventListener(event, e => { e.preventDefault(); $('drop').classList.remove('drag'); }));
$('drop').addEventListener('drop', e => addFiles(e.dataTransfer.files));
async function request(path, options) {
  const response = await fetch(path, options);
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || 'Не удалось выполнить запрос.');
  return data;
}
function encodeFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({name: file.name, data: reader.result.split(',')[1]});
    reader.onerror = () => reject(new Error('Не удалось прочитать ' + file.name));
    reader.readAsDataURL(file);
  });
}
$('form').onsubmit = async e => {
  e.preventDefault();
  if (busy || !files.length || !config) return;
  const checking = e.submitter?.id === 'check-dadata';
  if (checking && !config.token_configured && !$('token').value.trim()) {
    $('settings').open = true; $('token').focus(); return showError('Введите API-ключ DaData в настройках подключения.');
  }
  showError(''); setBusy(true); currentJob = null; $('cancel').hidden = true;
  $('result').hidden = true; $('progress-card').hidden = false;
  $('phase').textContent = 'Загружаем файлы'; $('progress').value = 0; $('percentage').textContent = '0%'; $('progress-note').textContent = 'Подготавливаем таблицы…';
  try {
    const encoded = await Promise.all(files.map(encodeFile));
    const selected = files.map(file => sheets.get(file)?.selected || '');
    const payload = {files:encoded, sheets:selected, mode:checking ? 'dadata_check' : 'cascade',
      token:$('token').value, save_token:$('save-token').checked};
    const job = await request('/api/jobs', {method:'POST', headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf}, body:JSON.stringify(payload)});
    if ($('token').value.trim()) { config.token_configured = true; config.token_saved ||= $('save-token').checked; }
    $('token').value = ''; updateConnection(); updateTokenNote(); currentJob = job.id; sessionStorage.setItem('schools-job', currentJob); poll();
  } catch (error) { showError(error.message); $('progress-card').hidden = true; $('cancel').hidden = true; setBusy(false); }
};
async function cancelJob() {
  if (!currentJob || !busy) return;
  $('cancel').disabled = true;
  try {
    await request('/api/jobs/' + currentJob + '/cancel', {method:'POST', headers:{'X-CSRF-Token':config.csrf}});
    poll();
  } catch (error) { showError('Не удалось отменить обработку: ' + error.message); $('cancel').disabled = false; }
}
$('cancel').onclick = cancelJob;
async function poll() {
  clearTimeout(timer); $('retry').hidden = true;
  try {
    const job = await request('/api/jobs/' + currentJob);
    showError(''); $('progress-card').hidden = false;
    $('phase').textContent = job.phase; $('progress').value = job.progress; $('percentage').textContent = job.progress + '%';
    $('progress-note').textContent = job.total ? `Шагов поиска и сверки: ${job.done} из ${job.total}` : 'Подготавливаем данные…';
    if (job.status === 'running' || job.status === 'cancelling') {
      setBusy(true); $('cancel').hidden = false; $('cancel').disabled = job.status === 'cancelling';
      $('cancel').textContent = job.status === 'cancelling' ? 'Отменяем…' : 'Отменить обработку';
      timer = setTimeout(poll, 1000);
    } else {
      setBusy(false); $('cancel').hidden = true;
      if (job.status === 'cancelled') $('progress-note').textContent = 'Можно запустить новую обработку.';
      else if (job.status === 'error') { showError(job.error); $('progress-card').hidden = true; }
      else { $('progress-card').hidden = true; renderResult(job); }
    }
  } catch (error) { showError('Не удалось получить статус: ' + error.message); $('retry').hidden = false; }
}
$('retry').onclick = poll;
function updateTokenNote() {
  $('key-note').textContent = config?.token_saved
    ? 'Ключ сохранён в системном хранилище этого компьютера. Можно указать другой.'
    : 'Новый ключ можно сохранить в системном хранилище этого компьютера.';
  $('forget-token').hidden = !config?.token_saved;
  $('token').placeholder = config?.token_configured ? 'Ключ уже настроен' : 'Введите ключ';
}
function renderRegistry() {
  const state = config.registry_update;
  const date = config.registry_date;
  const dateText = date ? `${date.slice(6,8)}.${date.slice(4,6)}.${date.slice(0,4)}` : 'без даты';
  $('registry').textContent = config.registry_ready ? `Резервный снимок: ${dateText}` : 'Резервного снимка пока нет';
  $('registry').classList.toggle('bad', !config.registry_ready && state.phase === 'error');
  $('registry-progress').textContent = state.message + (config.registry_ready && ['checking','waiting','error'].includes(state.phase)
    ? ' Пока используется локальный снимок.' : '');
  $('registry-retry').hidden = !state.retry_available || state.phase === 'ready';
  clearTimeout(registryTimer);
  if (['checking','waiting'].includes(state.phase)) registryTimer = setTimeout(refreshRegistry, 1500);
}
function renderLicenses() {
  const state = config.licenses_update || {};
  const date = config.licenses_date;
  const dateText = date ? `${date.slice(6,8)}.${date.slice(4,6)}.${date.slice(0,4)}` : 'без даты';
  $('licenses').textContent = config.licenses_ready ? `Основной снимок: ${dateText}` : 'Основного снимка пока нет';
  $('licenses').classList.toggle('bad', !config.licenses_ready && state.phase === 'error');
  $('licenses-progress').textContent = state.message + (config.licenses_ready && ['checking','waiting','error'].includes(state.phase)
    ? ' Пока используется локальный снимок.' : '');
  $('licenses-retry').hidden = !state.retry_available || state.phase === 'ready';
  clearTimeout(licensesTimer);
  if (['checking','waiting'].includes(state.phase)) licensesTimer = setTimeout(refreshRegistry, 1500);
}
async function refreshRegistry() {
  try {
    const next = await request('/api/config');
    config = next;
    updateConnection(); updateTokenNote(); renderRegistry(); renderLicenses(); updateButton();
  } catch (error) {
    $('registry-progress').textContent = 'Нет связи с приложением. Обновите страницу, когда сервер снова заработает.';
  }
}
$('registry-retry').onclick = async () => {
  $('registry-retry').disabled = true;
  try {
    await request('/api/registry/retry', {method:'POST', headers:{'X-CSRF-Token':config.csrf}});
    await refreshRegistry();
  } catch (error) { $('registry-progress').textContent = 'Не удалось запустить проверку: ' + error.message; }
  finally { $('registry-retry').disabled = false; }
};
$('licenses-retry').onclick = async () => {
  $('licenses-retry').disabled = true;
  try {
    await request('/api/licenses/retry', {method:'POST', headers:{'X-CSRF-Token':config.csrf}});
    await refreshRegistry();
  } catch (error) { $('licenses-progress').textContent = 'Не удалось запустить проверку: ' + error.message; }
  finally { $('licenses-retry').disabled = false; }
};
$('forget-token').onclick = async () => {
  try {
    await request('/api/token/forget', {method:'POST', headers:{'X-CSRF-Token':config.csrf}});
    $('token').value = ''; await refreshRegistry();
  } catch (error) { showError(error.message); }
};
async function saveDecision(groupId, inn, comment, button, card) {
  button.disabled = true;
  try {
    const result = await request('/api/jobs/' + currentJob + '/decisions', {
      method:'POST', headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf},
      body:JSON.stringify({group_id:groupId, inn, comment})
    });
    renderResult(result.job);
  } catch (error) { showError(error.message); button.disabled = false; }
}
function renderManualReview(items, selectionItems=[]) {
  const section = $('manual-review'), list = $('review-list');
  list.replaceChildren();
  const pending = (items || []).filter(item => item.status !== 'ПРОВЕРЕНО');
  section.hidden = false; // Очередь видна даже при недоступном реестре и пустом поиске.
  if (!pending.length && !selectionItems.length) {
    const note=document.createElement('p'); note.className='muted';
    note.textContent='Нет записей для ручной проверки в этом запуске.'; list.append(note);
  }
  selectionItems.forEach(item => {
    const card=document.createElement('article'); card.className='review-card';
    const heading=document.createElement('h4'); heading.textContent=item.name || 'Без названия'; card.append(heading);
    const explanation=document.createElement('p'); explanation.textContent='Проверка исходных данных: '+(item.reason || 'Недостаточно данных')+'. Вид образования: '+(item.education || 'не определён')+'. Уточните вид образования в исходном Excel и повторите обработку. Выбор ИНН недоступен до отбора записи.'; card.append(explanation);
    list.append(card);
  });
  pending.forEach(item => {
    const card = document.createElement('article'); card.className = 'review-card';
    const title = document.createElement('h4'); title.textContent = item.source_name || 'Учреждение без названия'; card.append(title);
    const meta = document.createElement('p'); meta.className = 'muted'; meta.textContent = [item.region,item.city].filter(Boolean).join(' · '); card.append(meta);
    const reason = document.createElement('p'); reason.textContent = 'Причина проверки: ' + (item.reason || 'Требуется ручное решение'); card.append(reason);
    const select = document.createElement('select'); select.setAttribute('aria-label','Кандидат для подтверждения');
    const placeholder = document.createElement('option'); placeholder.value=''; placeholder.textContent='Выберите кандидата'; select.append(placeholder);
    (item.options || []).slice(0,3).forEach(c => {
      const option = document.createElement('option'); option.value=String(c.candidate_inn || '');
      option.textContent = `${c.candidate_name || 'Без названия'} — ИНН ${c.candidate_inn || 'нет'} — ${c.address || 'адрес не указан'} — балл ${c.score ?? 'н/д'} — источник ${c.source || 'не указан'} — ${(c.reasons || []).join(', ') || 'требует подтверждения'}`;
      select.append(option);
    });
    card.append(select);
    if (!(item.options || []).length) {
      const note=document.createElement('p'); note.className='muted';
      note.textContent='Кандидатов пока нет: локальный реестр недоступен или поиск не дал результатов. Дождитесь загрузки реестра и повторите обработку либо нажмите «Проверить в DaData». Без подтверждённого кандидата ИНН выбрать нельзя.';
      card.append(note); select.disabled=true;
    }
    const dadata = document.createElement('button'); dadata.type='button'; dadata.className='secondary'; dadata.textContent='Проверить в DaData';
    dadata.onclick = async () => {
      const token = prompt('Введите API-ключ DaData. Запрос отправит название и географию учреждения, но не ФИО и номера личных дел.');
      if (!token) return;
      dadata.disabled = true;
      try {
        const result = await request('/api/jobs/' + currentJob + '/dadata', {method:'POST',
          headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf},
          body:JSON.stringify({group_id:item.group_id, token})});
        renderResult(result.job);
      } catch (error) { showError(error.message); dadata.disabled = false; }
    };
    card.append(dadata);
    const comment = document.createElement('textarea'); comment.maxLength=1000; comment.placeholder='Комментарий (необязательно)'; comment.setAttribute('aria-label','Комментарий к решению'); card.append(comment);
    const save = document.createElement('button'); save.type='button'; save.className='secondary'; save.textContent='Подтвердить ИНН';
    const status = document.createElement('p'); status.className='decision-status muted';
    save.onclick = () => { if (!select.value) return showError('Выберите ИНН из списка кандидатов.'); saveDecision(item.group_id, select.value, comment.value, save, card); };
    if (!(item.options || []).length) save.disabled=true;
    card.append(save,status); list.append(card);
  });
}
function renderResult(job) {
  renderManualReview(job.review || [], job.selection_review || []);
  $('result').hidden = false; $('download').href = job.download;
  $('filled-downloads').replaceChildren();
  for (const source of job.filled_sources || []) {
    const link = document.createElement('a');
    link.href = source.url;
    link.download = source.source;
    link.textContent = `Скачать заполненную копию: ${source.source}`;
    $('filled-downloads').append(link, document.createElement('br'));
  }
  $('journal').href = job.journal; $('journal').hidden = !job.journal;
  $('manual-review').hidden = !job.manual_template;
  $('manual-template').href = job.manual_template;
  manualPending = job.manual_pending || 0;
  $('manual-status').textContent = manualPending
    ? `Групп для ручного решения: ${manualPending}. В шаблоне указаны до трёх кандидатов и причины проверки.`
    : 'Неподтверждённых групп сопоставления нет. Поля писем и причины отбора проверяйте в итоговом Excel.';
  updateButton();
  $('result-title').textContent = job.partial ? 'Таблица готова частично' : 'Таблица готова';
  $('metrics').replaceChildren();
  for (const [value, label] of [[job.summary.rows,'строк обработано'],[job.summary.schools,'учреждений объединено'],[job.summary.students,'пар ФИО и номер дела'],[job.summary.review,'строк требуют проверки']]) {
    const box = document.createElement('div'), number = document.createElement('strong'), text = document.createElement('span');
    box.className = 'metric'; number.textContent = value; text.textContent = label; box.append(number, text); $('metrics').append(box);
  }
  $('warnings').replaceChildren();
  const warnings = [...job.warnings];
  if (job.summary.review) warnings.push('Неоднозначные записи сохранены отдельно. Посмотрите лист «Ручная проверка» и причины отбора на листе «Исходные записи».');
  if (job.summary.excluded) warnings.push(`Исключено по правилам отбора: ${job.summary.excluded}. Причины сохранены на листе «Исходные записи».`);
  for (const warning of warnings) { const p = document.createElement('p'); p.className = 'notice'; p.textContent = warning; $('warnings').append(p); }
  $('preview').replaceChildren();
  $('preview-caption').textContent = job.preview.length ? `Показано ${job.preview.length} из ${job.summary.schools}. Все записи — в скачиваемом Excel.` : 'Учреждения не сопоставлены. Скачайте Excel, чтобы посмотреть причины и кандидатов.';
  for (const school of job.preview) {
    const tr = document.createElement('tr');
    for (const key of ['school_name','address','students','files']) { const td = document.createElement('td'); td.textContent = school[key] || '—'; tr.append(td); }
    $('preview').append(tr);
  }
}
$('manual-file').onchange = async e => {
  const file = e.target.files[0];
  if (!file || busy || !currentJob) return;
  if (!file.name.toLowerCase().endsWith('.xlsx') || file.size > 6 * 1024 * 1024) {
    e.target.value = ''; return showError('Выберите заполненный шаблон .xlsx размером до 6 МБ.');
  }
  showError(''); setBusy(true); $('manual-status').textContent = 'Проверяем решения и обновляем Excel…';
  try {
    const encoded = await encodeFile(file);
    const response = await request('/api/jobs/' + currentJob + '/decisions-import', {
      method:'POST', headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf},
      body:JSON.stringify({data:encoded.data})
    });
    renderResult(response.job);
    $('manual-status').textContent = `Применено решений: ${response.imported}. Статус — ПРОВЕРЕНО. Отчёт, журнал и копии исходных книг обновлены.`;
  } catch (error) {
    showError(error.message); $('manual-status').textContent = 'Импорт не завершён. Проверьте сообщение об ошибке.';
  } finally { e.target.value = ''; setBusy(false); }
};
async function init() {
  try {
    config = await request('/api/config');
    updateConnection();
    updateTokenNote(); renderRegistry(); renderLicenses();
    if (files.length) await addFiles([]);
    updateButton();
    currentJob = sessionStorage.getItem('schools-job');
    if (currentJob) { const previous = await request('/api/jobs/' + currentJob).catch(() => null); if (previous) { setBusy(['running','cancelling'].includes(previous.status)); poll(); } else sessionStorage.removeItem('schools-job'); }
  } catch (error) { showError('Сервер недоступен. Запустите приложение по инструкции и обновите страницу.'); }
}
init();
