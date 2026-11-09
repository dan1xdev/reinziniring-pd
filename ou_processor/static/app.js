'use strict';
const $ = id => document.getElementById(id);
let files = [], config = null, busy = false, currentJob = null, timer = null, registryTimer = null, licensesTimer = null;
const sheets = new Map();
function showError(message) { $('error').textContent = message; $('error').hidden = !message; }
function setBusy(value) { busy = value; $('controls').disabled = value; updateButton(); }
function updateButton() {
  $('start').disabled = busy || !files.length || !config;
  $('check-dadata').disabled = busy || !files.length || !config;
  $('start').firstChild.textContent = !files.length ? 'Сначала загрузите Excel ' : 'Обработать таблицы ';
}
function updateConnection() {
  $('connection').textContent = $('token').value.trim() ? 'Ключ для отдельной сверки введён'
    : config?.token_saved ? 'Ключ сохранён в системе'
    : config?.token_configured ? 'Ключ для сверки задан на сервере' : 'Для обычной обработки ключ не нужен';
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
    const payload = {files:encoded, sheets:selected, mode:checking ? 'dadata_check' : 'registry'};
    if (checking) Object.assign(payload, {token:$('token').value, save_token:$('save-token').checked});
    const job = await request('/api/jobs', {method:'POST', headers:{'Content-Type':'application/json','X-CSRF-Token':config.csrf}, body:JSON.stringify(payload)});
    if (checking && $('token').value.trim()) { config.token_configured = true; config.token_saved ||= $('save-token').checked; }
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
function renderResult(job) {
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
