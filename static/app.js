/* ===================== ESTADO E CONFIGURAÇÃO ===================== */
let PFX_LIST = ['REPROGRAMAÇÃO:', 'CANCELAMENTO:'];
let BASE_SUBJECTS = [
  'INFORMATIVO DE ATIVIDADE PROGRAMADA - AMERICAN TOWER',
  'MANUTENÇÃO PROGRAMADA'
];
let SCENARIO_NOTICES = {
  repro: 'ATENÇÃO: A atividade informada abaixo foi REPROGRAMADA. Seguem as novas datas de execução.',
  cancel: 'ATENÇÃO: A atividade informada abaixo foi CANCELADA.'
};
const PFX_RE = /^(REPROGRAMAÇÃO|REPROGRAMACAO|CANCELAMENTO|ATUALIZAÇÃO|ATUALIZACAO|ALTERAÇÃO|ALTERACAO):\s*/i;
const DATE_RE = /^(\d{2})\/(\d{2})\/(\d{4})\s*[–\-]\s*(\d{2})h(\d{2})$/;

const S = {
  config: {},
  stores: [],
  pickedFolder: null,
  isDemo: false,
  emails: [],
  fields: [],
  subj: { prefix: '', fixed: '', mode: 'set', bases: [] },
  scenario: null,
  useNotice: false,
  noticeText: '',
  over: {},      // email_id -> custom full subject
  perEmail: {},  // email_id -> { label: custom_value }
  reps: [{ find: '', rep: '' }],
  sel: new Set(),
  cur: null,
  pv: 'new',
  logsCount: 0,
  queue: [],
  qi: 0
};

/* ===================== UTILITÁRIOS ===================== */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

let toastT;
function toast(m) {
  const t = $('toast');
  t.textContent = m;
  t.classList.add('show');
  clearTimeout(toastT);
  toastT = setTimeout(() => t.classList.remove('show'), 3500);
}

function showBanner(msg, isWarn = false) {
  const b = $('topBanner');
  if (!msg) {
    b.className = 'banner';
    document.body.classList.remove('has-banner');
    return;
  }
  b.textContent = msg;
  b.className = 'banner show' + (isWarn ? ' warn' : '');
  document.body.classList.add('has-banner');
}

function openModal(id) { $(id).classList.add('show'); }
function closeModal(id) { $(id).classList.remove('show'); }

let pvT;
function refreshSoon() {
  clearTimeout(pvT);
  pvT = setTimeout(() => {
    renderPreview();
    updateBadges();
    updateFooter();
  }, 160);
}

async function api(url, opts = {}) {
  const res = await fetch(url, {
    headers: { 'Content-Type': 'application/json' },
    ...opts
  });
  const data = await res.json();
  if (!res.ok || data.ok === false) {
    throw new Error(data.error || `Erro HTTP ${res.status}`);
  }
  return data;
}

/* ===================== INICIALIZAÇÃO ===================== */
async function initApp() {
  try {
    const res = await api('/api/config');
    S.config = res.config || {};
    if (Array.isArray(S.config.prefixes) && S.config.prefixes.length) {
      PFX_LIST = S.config.prefixes;
    }
    if (Array.isArray(S.config.subject_bases) && S.config.subject_bases.length) {
      BASE_SUBJECTS = S.config.subject_bases;
    }
    if (S.config.scenario_notices) {
      SCENARIO_NOTICES = { ...SCENARIO_NOTICES, ...S.config.scenario_notices };
    }
    if (S.config.shared_mailbox) {
      $('sharedMailboxInput').value = S.config.shared_mailbox;
    }
    if (S.config.default_sender_mode) {
      $('senderMode').value = S.config.default_sender_mode;
      $('sharedMailboxInput').disabled = $('senderMode').value !== 'shared';
    }
    if (S.config.last_folder_entry_id && S.config.last_folder_path) {
      const btnLast = $('btnLastFolder');
      btnLast.hidden = false;
      btnLast.textContent = `↺ Reabrir: ${S.config.last_folder_path}`;
      btnLast.onclick = () => loadOutlookFolder(S.config.last_folder_entry_id, S.config.last_folder_store_id);
    }
    if (!res.outlook_ok) {
      showBanner(`Aviso: Não foi possível conectar ao Outlook clássico (${res.outlook_error}). Verifique se o Outlook está aberto ou use o lote de demonstração.`, true);
    }
  } catch (err) {
    toast('Erro ao carregar configuração inicial: ' + err.message);
  }

  // Heartbeat periódico
  setInterval(() => {
    fetch('/api/heartbeat', { method: 'POST' }).catch(() => {});
  }, 5000);
}

/* ===================== SELEÇÃO DE PASTAS DO OUTLOOK ===================== */
async function openFolderModal() {
  openModal('folderModal');
  $('tree').innerHTML = '<div class="muted">Consultando caixas de correio e pastas no Outlook…</div>';
  $('btnLoadFolder').disabled = !S.pickedFolder;

  try {
    const res = await api('/api/folders');
    S.stores = res.stores || [];
    // Abre o primeiro nível por padrão
    S.stores.forEach(st => { st.open = true; });
    renderTree();
  } catch (err) {
    $('tree').innerHTML = `<div style="color:var(--danger);padding:8px">
      <b>Não foi possível listar as pastas do Outlook:</b><br>${esc(err.message)}<br><br>
      Certifique-se de que o <b>Outlook clássico (desktop)</b> está aberto, ou clique em <b>🧪 Lote de demonstração</b> no rodapé desta janela para testar.
    </div>`;
  }
}

function renderTree() {
  if (!S.stores.length) {
    $('tree').innerHTML = '<div class="muted">Nenhuma caixa encontrada no Outlook.</div>';
    return;
  }
  const build = nodes => '<ul>' + nodes.map(n => {
    const isSel = S.pickedFolder && S.pickedFolder.entry_id === n.entry_id;
    const icon = n.is_store ? '📫' : '📁';
    const tg = (n.has_subfolders || (n.children && n.children.length)) ? (n.open ? '▾' : '▸') : '';
    const kids = (n.open && n.children && n.children.length) ? build(n.children) : '';
    return `<li>
      <div data-eid="${esc(n.entry_id)}" class="${isSel ? 'sel' : ''}">
        <span class="tg">${tg}</span>${icon} ${esc(n.name)}
        ${n.count != null && !n.is_store ? `<span class="cnt">${n.count.toLocaleString('pt-BR')}</span>` : ''}
      </div>${kids}
    </li>`;
  }).join('') + '</ul>';

  $('tree').innerHTML = build(S.stores);
}

function findFolderNode(entryId, nodes = S.stores) {
  for (const n of nodes) {
    if (n.entry_id === entryId) return n;
    if (n.children && n.children.length) {
      const found = findFolderNode(entryId, n.children);
      if (found) return found;
    }
  }
  return null;
}

$('tree').addEventListener('click', async ev => {
  const row = ev.target.closest('div[data-eid]');
  if (!row) return;
  const node = findFolderNode(row.dataset.eid);
  if (!node) return;

  const clickedArrow = ev.target.classList.contains('tg');
  S.pickedFolder = node;
  $('btnLoadFolder').disabled = false;

  if (node.has_subfolders && (!node.children || !node.children.length)) {
    try {
      const sub = await api('/api/subfolders', {
        method: 'POST',
        body: JSON.stringify({ entry_id: node.entry_id, store_id: node.store_id })
      });
      node.children = sub.children || [];
      node.open = true;
    } catch (e) {
      toast('Erro ao expandir subpastas: ' + e.message);
    }
  } else if (clickedArrow || node.is_store) {
    node.open = !node.open;
  }
  renderTree();
});

$('tree').addEventListener('dblclick', ev => {
  const row = ev.target.closest('div[data-eid]');
  if (!row) return;
  const node = findFolderNode(row.dataset.eid);
  if (node && !node.is_store) {
    S.pickedFolder = node;
    closeModal('folderModal');
    loadOutlookFolder(node.entry_id, node.store_id);
  }
});

$('btnFolder').onclick = openFolderModal;

$('btnLoadFolder').onclick = () => {
  if (!S.pickedFolder) return;
  closeModal('folderModal');
  loadOutlookFolder(S.pickedFolder.entry_id, S.pickedFolder.store_id);
};

$('btnNativePick').onclick = async () => {
  try {
    toast('Verifique a janela de seleção de pastas aberta pelo Outlook…');
    const res = await api('/api/pick-folder', { method: 'POST' });
    if (res.cancelled || !res.folder) return;
    closeModal('folderModal');
    loadOutlookFolder(res.folder.entry_id, res.folder.store_id);
  } catch (err) {
    toast('Erro ao abrir seletor do Outlook: ' + err.message);
  }
};

async function loadOutlookFolder(entryId, storeId) {
  const subjFilter = $('subjFilterInput') ? $('subjFilterInput').value.trim() : '';
  const maxItems = $('maxItemsInput') ? (parseInt($('maxItemsInput').value, 10) || 200) : 200;
  await doLoadFolder({ demo: false, entry_id: entryId, store_id: storeId, subject_filter: subjFilter, max_items: maxItems });
}

async function loadDemoFolder() {
  await doLoadFolder({ demo: true });
}

async function doLoadFolder(payload) {
  $('empty').hidden = true;
  $('workspace').hidden = true;
  $('loading').hidden = false;
  $('loadingMsg').textContent = payload.demo ? 'Carregando lote de demonstração…' : 'Lendo e-mails da pasta no Outlook e extraindo campos…';

  try {
    const res = await api('/api/load-folder', {
      method: 'POST',
      body: JSON.stringify(payload)
    });

    if (!res.emails || !res.emails.length) {
      $('loading').hidden = true;
      $('empty').hidden = false;
      toast('Nenhum e-mail encontrado na pasta selecionada (verifique o filtro).');
      return;
    }

    S.isDemo = !!res.is_demo;
    if (S.isDemo) {
      showBanner('MODO DEMONSTRAÇÃO ATIVO: Exibindo 6 informativos de exemplo. Se o Outlook estiver aberto, o botão Confirmar abrirá rascunhos reais nele; caso contrário, abrirá a janela simulada.', true);
    } else {
      showBanner('');
    }

    $('folderPath').textContent = `${res.folder_path} (${res.emails.length} e-mails)`;
    S.emails = res.emails;
    S.fields = res.fields || [];
    S.subj = res.subject || { prefix: '', fixed: '', mode: 'set', bases: [] };
    S.scenario = null;
    S.useNotice = false;
    S.noticeText = '';
    S.over = {};
    S.perEmail = {};
    S.reps = [{ find: '', rep: '' }];
    S.sel = new Set(S.emails.map(e => e.id));
    S.cur = S.emails[0].id;

    $('loading').hidden = true;
    $('workspace').hidden = false;

    const nv = S.fields.filter(f => f.varied).length;
    $('detSummary').textContent = `· ${S.fields.length} tópicos + assunto detectados (${nv} variável)`;

    renderAll();
    toast(`${S.emails.length} e-mail(s) carregado(s) · ${S.fields.length} tópico(s) detectado(s)`);
  } catch (err) {
    $('loading').hidden = true;
    $('empty').hidden = false;
    toast('Erro ao carregar pasta: ' + err.message);
  }
}

/* ===================== CÁLCULO DE ALTERAÇÕES POR E-MAIL ===================== */
function getFieldUpdatesForEmail(e) {
  const updates = {};
  (e.fields || []).forEach(f => {
    const pe = S.perEmail[e.id] && S.perEmail[e.id][f.label];
    if (pe != null && pe !== f.value) {
      updates[f.label] = pe;
      return;
    }
    const g = S.fields.find(x => x.label === f.label);
    if (g && g.apply && !g.varied && g.nv !== '' && g.nv !== f.value) {
      updates[f.label] = g.nv;
    }
  });
  return updates;
}

function stripPrefix(s) {
  return String(s || '').replace(PFX_RE, '').trim();
}

function finalSubject(e) {
  if (S.over[e.id] != null) return S.over[e.id];
  const base = (S.subj.mode === 'keep' && !S.subj.fixed)
    ? stripPrefix(e.subject)
    : S.subj.fixed;
  return (S.subj.prefix ? S.subj.prefix + ' ' : '') + base;
}

function activeNotice() {
  return (S.useNotice && S.noticeText.trim()) ? S.noticeText.trim() : '';
}

function changesFor(e) {
  const ch = [];
  if (finalSubject(e) !== e.subject) ch.push('Assunto');
  if (activeNotice()) ch.push('Aviso no topo');
  const fUpdates = getFieldUpdatesForEmail(e);
  Object.keys(fUpdates).forEach(k => ch.push(k));
  const bodyTxt = (e.text || '') + ' ' + (e.html || '');
  if (S.reps.some(r => r.find && bodyTxt.includes(r.find))) ch.push('Localizar/Substituir');
  return ch;
}

/* ===================== RENDERIZAÇÃO DA INTERFACE ===================== */
function renderAll() {
  renderList();
  renderSubject();
  renderFields();
  renderReps();
  renderPreview();
  updateFooter();
  setScenario(null);
}

function renderList() {
  $('emailList').innerHTML = S.emails.map(e => `
    <div class="mail ${S.cur === e.id ? 'active' : ''} ${S.sel.has(e.id) ? '' : 'off'}" data-id="${e.id}">
      <div class="top">
        <input type="checkbox" class="pick" ${S.sel.has(e.id) ? 'checked' : ''}>
        <span class="to" title="${esc(e.to)}">${esc(e.to || '(sem destinatário)')}</span>
        <span class="date">${esc(e.sent_on)}</span>
      </div>
      <input type="text" class="subj ${S.over[e.id] != null ? 'over' : ''}" value="${esc(finalSubject(e))}" title="Assunto final deste e-mail">
      <div class="badges" data-b="${e.id}"></div>
    </div>`).join('');
  updateBadges();
  updateCount();
}

function updateBadges() {
  S.emails.forEach(e => {
    const b = document.querySelector(`[data-b="${e.id}"]`);
    if (!b) return;
    const ch = changesFor(e);
    b.innerHTML = (ch.length
      ? `<span class="badge chg" title="${esc(ch.join(', '))}">${ch.length} alteração(ões)</span>`
      : '<span class="badge">sem alterações</span>')
      + (S.over[e.id] != null ? `<span class="badge ov" data-reset="${e.id}" title="Voltar ao assunto padrão">assunto individual ✕</span>` : '');

    const inp = document.querySelector(`.mail[data-id="${e.id}"] input.subj`);
    if (inp && document.activeElement !== inp) {
      inp.value = finalSubject(e);
      inp.classList.toggle('over', S.over[e.id] != null);
    }
  });
}

function updateCount() {
  $('countSel').textContent = `· ${S.sel.size} de ${S.emails.length} selecionados`;
  $('chkAll').checked = S.emails.length > 0 && S.sel.size === S.emails.length;
}

$('emailList').addEventListener('click', ev => {
  const card = ev.target.closest('.mail');
  if (!card) return;
  const id = card.dataset.id;
  if (ev.target.dataset.reset) {
    delete S.over[ev.target.dataset.reset];
    updateBadges();
    renderPreview();
    return;
  }
  if (ev.target.classList.contains('pick')) {
    ev.target.checked ? S.sel.add(id) : S.sel.delete(id);
    card.classList.toggle('off', !ev.target.checked);
    updateCount();
    updateFooter();
    return;
  }
  if (ev.target.classList.contains('subj')) {
    if (S.cur !== id) { S.cur = id; markActive(); renderPreview(); }
    return;
  }
  S.cur = id;
  markActive();
  renderPreview();
});

$('emailList').addEventListener('input', ev => {
  if (!ev.target.classList.contains('subj')) return;
  const id = ev.target.closest('.mail').dataset.id;
  S.over[id] = ev.target.value;
  ev.target.classList.add('over');
  refreshSoon();
});

function markActive() {
  document.querySelectorAll('.mail').forEach(m => m.classList.toggle('active', m.dataset.id === S.cur));
}

$('chkAll').onchange = ev => {
  S.sel = ev.target.checked ? new Set(S.emails.map(e => e.id)) : new Set();
  renderList();
  updateFooter();
};

/* Assunto */
function renderSubject() {
  const opts = [''].concat(PFX_LIST);
  const custom = S.subj.prefix && !PFX_LIST.includes(S.subj.prefix);
  $('pfxChips').innerHTML = '<span class="muted">Prefixo:</span>' + opts.map(p =>
    `<button class="chip ${S.subj.prefix === p ? 'on' : ''}" data-p="${esc(p)}">${p || 'Nenhum'}</button>`).join('')
    + `<input type="text" id="pfxCustom" placeholder="Outro…" style="width:120px;padding:4px 8px" value="${custom ? esc(S.subj.prefix) : ''}">`;

  const baseCandidates = [...new Set([...BASE_SUBJECTS, ...(S.subj.bases || [])])].filter(Boolean);
  $('baseChips').innerHTML = '<span class="muted">Assunto base:</span>' + baseCandidates.map(b =>
    `<button class="chip ${S.subj.fixed === b ? 'on' : ''}" data-base="${esc(b)}">${esc(b)}</button>`
  ).join('');

  $('pfxShow').textContent = S.subj.prefix;
  $('subjFixed').value = S.subj.fixed;
}

$('pfxChips').addEventListener('click', ev => {
  const b = ev.target.closest('button[data-p]');
  if (!b) return;
  S.subj.prefix = b.dataset.p;
  renderSubject();
  refreshSoon();
});

$('baseChips').addEventListener('click', ev => {
  const b = ev.target.closest('button[data-base]');
  if (!b) return;
  S.subj.fixed = b.dataset.base;
  S.subj.mode = 'set';
  renderSubject();
  refreshSoon();
});

$('pfxChips').addEventListener('input', ev => {
  if (ev.target.id !== 'pfxCustom') return;
  S.subj.prefix = ev.target.value.trim();
  $('pfxShow').textContent = S.subj.prefix;
  document.querySelectorAll('#pfxChips .chip').forEach(c => c.classList.toggle('on', c.dataset.p === S.subj.prefix));
  refreshSoon();
});

$('subjFixed').addEventListener('input', ev => {
  S.subj.fixed = ev.target.value;
  S.subj.mode = 'set';
  document.querySelectorAll('#baseChips .chip').forEach(c => c.classList.toggle('on', c.dataset.base === S.subj.fixed));
  refreshSoon();
});

/* Cenários rápidos */
function setScenario(s) {
  S.scenario = s;
  document.querySelectorAll('#scen button').forEach(b => b.classList.toggle('on', b.dataset.s === s));
  if (!s) {
    $('noticeBox').hidden = true;
    return;
  }
  S.subj.prefix = s === 'repro' ? 'REPROGRAMAÇÃO:' : s === 'cancel' ? 'CANCELAMENTO:' : '';
  if (s === 'repro' || s === 'cancel') {
    $('noticeBox').hidden = false;
    S.noticeText = SCENARIO_NOTICES[s] || '';
    $('noticeInput').value = S.noticeText;
    $('chkNotice').checked = S.useNotice;
  } else {
    $('noticeBox').hidden = true;
    S.useNotice = false;
    $('chkNotice').checked = false;
  }
  renderSubject();
  refreshSoon();
  if (s === 'repro') {
    document.querySelectorAll('#fieldsBody tr').forEach(tr => {
      const f = S.fields[tr.dataset.i];
      tr.classList.remove('hl');
      if (f && f.type === 'date') {
        void tr.offsetWidth;
        tr.classList.add('hl');
      }
    });
    const d = document.querySelector('#fieldsBody input[type=date]');
    if (d) d.focus();
    toast('Reprogramação: prefixo REPROGRAMAÇÃO: aplicado. Ajuste as datas destacadas.');
  } else if (s === 'cancel') {
    toast('Cancelamento: prefixo CANCELAMENTO: aplicado.');
  }
}

$('scen').addEventListener('click', ev => {
  const b = ev.target.closest('button');
  if (b) setScenario(b.dataset.s);
});

$('chkNotice').addEventListener('change', ev => {
  S.useNotice = ev.target.checked;
  refreshSoon();
});

$('noticeInput').addEventListener('input', ev => {
  S.noticeText = ev.target.value;
  if (S.useNotice) refreshSoon();
});

/* Tópicos do corpo */
function dateParts(v) {
  const m = (v || '').match(DATE_RE);
  return m ? { d: `${m[3]}-${m[2]}-${m[1]}`, t: `${m[4]}:${m[5]}` } : { d: '', t: '' };
}

function dateFmt(d, t) {
  if (!d) return '';
  const [y, mo, da] = d.split('-');
  const [h, mi] = (t || '00:00').split(':');
  return `${da}/${mo}/${y} – ${h}h${mi}`;
}

function renderFields() {
  if (!S.fields.length) {
    $('fieldsBody').innerHTML = `<tr><td colspan="4" class="muted" style="text-align:center;padding:16px">
      Nenhum tópico no formato "Rótulo: valor" encontrado. Você ainda pode editar o Assunto ou usar Localizar e Substituir abaixo.
    </td></tr>`;
    return;
  }
  $('fieldsBody').innerHTML = S.fields.map((f, i) => {
    let cur, edit;
    if (f.varied) {
      cur = `<span class="lock">🔒 ${f.uniq.length} valores diferentes</span>`;
      const n = S.emails.filter(e => S.perEmail[e.id] && S.perEmail[e.id][f.label] != null).length;
      edit = `<button class="small" data-pe="${i}">Editar por e-mail…</button>`
        + (n ? `<div class="hint">${n} e-mail(s) com valor individual</div>` : '<div class="hint">Diferente em cada e-mail, editável individualmente</div>');
    } else {
      cur = esc(f.cur);
      if (f.type === 'date') {
        const p = dateParts(f.nv);
        edit = `<div class="dt"><input type="date" data-d="${i}" value="${p.d}"><input type="time" data-t="${i}" value="${p.t}"></div><div class="hint" data-fmt="${i}">${esc(f.nv)}</div>`;
      } else if (f.type === 'long') {
        edit = `<textarea data-v="${i}">${esc(f.nv)}</textarea>`;
      } else {
        edit = `<input type="text" data-v="${i}" value="${esc(f.nv)}">`;
      }
    }
    return `<tr data-i="${i}" class="${!f.varied && f.nv !== f.cur ? 'changed' : ''}">
      <td class="lbl">${esc(f.label)}</td>
      <td class="cur">${cur}</td>
      <td>${edit}</td>
      <td class="ap">${f.varied ? '—' : `<input type="checkbox" data-a="${i}" ${f.apply ? 'checked' : ''}>`}</td>
    </tr>`;
  }).join('');
}

$('fieldsBody').addEventListener('input', ev => {
  const t = ev.target, ds = t.dataset;
  const i = ds.v ?? ds.d ?? ds.t;
  if (i == null) return;
  const f = S.fields[i];
  if (ds.v != null) {
    f.nv = t.value;
  } else {
    const row = t.closest('tr');
    f.nv = dateFmt(row.querySelector('[data-d]').value, row.querySelector('[data-t]').value);
    row.querySelector('[data-fmt]').textContent = f.nv;
  }
  t.closest('tr').classList.toggle('changed', f.nv !== f.cur);
  refreshSoon();
});

$('fieldsBody').addEventListener('change', ev => {
  const a = ev.target.dataset.a;
  if (a != null) {
    S.fields[a].apply = ev.target.checked;
    refreshSoon();
  }
});

$('fieldsBody').addEventListener('click', ev => {
  const p = ev.target.dataset.pe;
  if (p != null) openPerEmail(+p);
});

/* Edição por e-mail */
let peField = null;
function openPerEmail(i) {
  peField = S.fields[i];
  $('peTitle').textContent = 'Editar por e-mail: ' + peField.label;
  $('peBody').innerHTML = S.emails.map(e => {
    const found = (e.fields || []).find(f => f.label === peField.label);
    const orig = found ? found.value : '';
    const v = (S.perEmail[e.id] && S.perEmail[e.id][peField.label]) ?? orig;
    return `<div style="margin-bottom:12px">
      <div style="font-weight:600;font-size:13px;margin-bottom:4px">${esc(e.to)} <span class="muted">(${esc(e.subject)})</span></div>
      <textarea data-e="${e.id}" rows="2">${esc(v)}</textarea>
    </div>`;
  }).join('');
  openModal('peModal');
}

$('peSave').onclick = () => {
  document.querySelectorAll('#peBody textarea').forEach(t => {
    const e = S.emails.find(x => x.id === t.dataset.e);
    const found = (e.fields || []).find(f => f.label === peField.label);
    const orig = found ? found.value : '';
    S.perEmail[e.id] = S.perEmail[e.id] || {};
    if (t.value.trim() === orig) {
      delete S.perEmail[e.id][peField.label];
    } else {
      S.perEmail[e.id][peField.label] = t.value.trim();
    }
  });
  closeModal('peModal');
  renderFields();
  refreshSoon();
};

/* Localizar e substituir */
function renderReps() {
  $('repList').innerHTML = S.reps.map((r, i) => `<div class="rep">
    <input type="text" placeholder="Localizar…" data-f="${i}" value="${esc(r.find)}">
    <input type="text" placeholder="Substituir por…" data-r="${i}" value="${esc(r.rep)}">
    <button class="small" data-x="${i}" title="Remover">✕</button>
  </div>`).join('');
}

$('repList').addEventListener('input', ev => {
  const d = ev.target.dataset;
  if (d.f != null) S.reps[d.f].find = ev.target.value;
  if (d.r != null) S.reps[d.r].rep = ev.target.value;
  refreshSoon();
});

$('repList').addEventListener('click', ev => {
  const x = ev.target.dataset.x;
  if (x == null) return;
  S.reps.splice(x, 1);
  if (!S.reps.length) S.reps.push({ find: '', rep: '' });
  renderReps();
  refreshSoon();
});

$('addRep').onclick = () => {
  S.reps.push({ find: '', rep: '' });
  renderReps();
};

/* Pré-visualização via Backend */
document.querySelectorAll('.seg button').forEach(b => b.onclick = () => {
  S.pv = b.dataset.m;
  document.querySelectorAll('.seg button').forEach(x => x.classList.toggle('on', x === b));
  renderPreview();
});

async function renderPreview() {
  const e = S.emails.find(x => x.id === S.cur);
  if (!e) return;
  const isNew = S.pv === 'new';
  const subj = isNew ? finalSubject(e) : e.subject;
  const senderMode = $('senderMode').value;
  const sharedName = $('sharedMailboxInput').value.trim() || 'sir.fibra';
  const fromHtml = senderMode === 'shared'
    ? esc(sharedName)
    : '<span style="color:#c42b1c">(escolher manualmente no Outlook)</span>';
  const ch = changesFor(e);

  $('meta').innerHTML = `
    <div><b>De</b>${isNew ? fromHtml : esc(e.sender_name || sharedName)}</div>
    <div><b>Para</b>${esc(e.to)}</div>
    <div><b>Assunto</b>${isNew && subj !== e.subject ? `<mark style="background:#fff1a8">${esc(subj)}</mark>` : esc(subj)}</div>
    ${isNew ? `<div><b>Alterado</b>${ch.length ? esc(ch.join(' · ')) : '<span class="muted">nada ainda</span>'}</div>` : ''}`;

  if (!isNew) {
    $('preview').srcdoc = e.preview_html || e.html || '';
    return;
  }

  try {
    const res = await api('/api/preview', {
      method: 'POST',
      body: JSON.stringify({
        email_id: e.id,
        field_updates: getFieldUpdatesForEmail(e),
        replacements: S.reps.filter(r => r.find),
        notice: activeNotice(),
        highlight: true
      })
    });
    $('preview').srcdoc = res.html;
  } catch (err) {
    $('preview').srcdoc = e.preview_html || e.html || '';
  }
}

$('senderMode').onchange = () => {
  $('sharedMailboxInput').disabled = $('senderMode').value !== 'shared';
  renderPreview();
};
$('sharedMailboxInput').oninput = () => refreshSoon();

/* ===================== CONFIRMAR -> ABRIR NO OUTLOOK ===================== */
function updateFooter() {
  const n = S.sel.size;
  $('btnConfirm').disabled = !n;
  $('btnConfirm').textContent = `Confirmar e abrir ${n} rascunho(s) no Outlook`;
  $('btnLog').disabled = !S.logsCount;
  $('btnLog').textContent = `📄 Baixar log (${S.logsCount})`;
}

$('btnConfirm').onclick = async () => {
  const selectedEmails = S.emails.filter(e => S.sel.has(e.id));
  if (!selectedEmails.length) return;

  const noChange = selectedEmails.filter(e => !changesFor(e).length).length;
  if (noChange && !confirm(`${noChange} e-mail(s) selecionado(s) não têm alterações. Deseja abrir mesmo assim?`)) {
    return;
  }

  const mode = document.querySelector('input[name=mode]:checked').value;
  if (mode === 'all' && selectedEmails.length > 10) {
    if (!confirm(`Você selecionou abrir ${selectedEmails.length} janelas do Outlook de uma só vez. Continuar?`)) {
      return;
    }
  }

  S.queue = selectedEmails.map(e => ({
    e,
    subj: finalSubject(e),
    field_updates: getFieldUpdatesForEmail(e),
    changes: changesFor(e),
    status: 'pendente'
  }));
  S.qi = 0;

  openModal('queueModal');
  renderQueueModal();

  if (mode === 'all') {
    $('btnQueueNext').disabled = true;
    $('btnQueueSkip').disabled = true;
    for (let i = 0; i < S.queue.length; i++) {
      S.qi = i;
      renderQueueModal();
      await openQueueItem(i);
    }
    renderQueueModal();
    toast(`Concluído: ${S.queue.length} rascunho(s) processado(s).`);
  } else {
    // Modo um por vez: já abre o primeiro imediatamente e aguarda clique para o próximo
    await openQueueItem(0);
    renderQueueModal();
  }
};

async function openQueueItem(idx) {
  const item = S.queue[idx];
  if (!item) return;
  item.status = 'abrindo…';
  renderQueueModal();

  try {
    const res = await api('/api/open-draft', {
      method: 'POST',
      body: JSON.stringify({
        email_id: item.e.id,
        subject: item.subj,
        field_updates: item.field_updates,
        replacements: S.reps.filter(r => r.find),
        notice: activeNotice(),
        sender_mode: $('senderMode').value,
        shared_mailbox: $('sharedMailboxInput').value.trim(),
        changes: item.changes
      })
    });

    S.logsCount = res.logs_count || (S.logsCount + 1);
    updateFooter();

    if (res.opened_in_outlook) {
      item.status = 'aberto no Outlook ✓';
    } else {
      item.status = 'simulado ✓';
      showSimulatedOutlookModal(item, res.simulated_html);
    }
  } catch (err) {
    item.status = 'erro: ' + err.message;
    toast('Erro ao abrir rascunho no Outlook: ' + err.message);
  }
}

function renderQueueModal() {
  const doneCount = S.queue.filter(q => q.status.includes('✓') || q.status === 'pulado').length;
  $('queueStatusMsg').textContent = `Progresso: ${doneCount} de ${S.queue.length} rascunho(s)`;

  $('queueList').innerHTML = S.queue.map((q, i) => {
    const isDone = q.status.includes('✓');
    const isSkip = q.status === 'pulado';
    const isCur = i === S.qi && !isDone && !isSkip;
    return `<div class="queue-item ${isCur ? 'current' : ''} ${isDone ? 'done' : ''} ${isSkip ? 'skipped' : ''}">
      <span class="q-to">${esc(q.e.to)}</span>
      <span class="q-subj">${esc(q.subj)}</span>
      <span class="q-st">${esc(q.status)}</span>
    </div>`;
  }).join('');

  const nextIdx = S.queue.findIndex(q => q.status === 'pendente');
  if (nextIdx === -1) {
    $('btnQueueNext').disabled = true;
    $('btnQueueNext').textContent = 'Todos abertos ✓';
    $('btnQueueSkip').disabled = true;
  } else {
    $('btnQueueNext').disabled = false;
    $('btnQueueNext').textContent = `Abrir próximo no Outlook (${nextIdx + 1}/${S.queue.length}) ▶`;
    $('btnQueueSkip').disabled = false;
  }
}

$('btnQueueNext').onclick = async () => {
  const nextIdx = S.queue.findIndex(q => q.status === 'pendente');
  if (nextIdx === -1) return;
  S.qi = nextIdx;
  await openQueueItem(nextIdx);
  renderQueueModal();
};

$('btnQueueSkip').onclick = () => {
  const nextIdx = S.queue.findIndex(q => q.status === 'pendente');
  if (nextIdx === -1) return;
  S.queue[nextIdx].status = 'pulado';
  renderQueueModal();
};

$('btnQueueClose').onclick = () => closeModal('queueModal');

function showSimulatedOutlookModal(item, html) {
  $('olTitle').textContent = item.subj.slice(0, 65);
  $('olFrom').value = $('senderMode').value === 'shared' ? $('sharedMailboxInput').value.trim() : '';
  $('olTo').value = item.e.to;
  $('olCc').value = item.e.cc || '';
  $('olSubj').value = item.subj;
  const f = $('olBody');
  f.onload = () => { try { f.contentDocument.designMode = 'on'; } catch (e) {} };
  f.srcdoc = html || '';
  $('olCounter').textContent = `Simulação (Outlook não detectado) · Destinatário: ${item.e.to}`;
  openModal('olModal');
}

$('olSend').onclick = () => {
  toast('Modo simulação: quando usado com o Outlook aberto, você envia pela própria janela do Outlook.');
  closeModal('olModal');
};

/* Log e Encerramento */
$('btnLog').onclick = () => {
  window.location.href = '/api/logs/csv';
};

$('btnExit').onclick = async () => {
  if (!confirm('Deseja encerrar o Editor de Informativos?')) return;
  try {
    await api('/api/shutdown', { method: 'POST' });
    document.body.innerHTML = '<div style="display:grid;place-items:center;height:100vh;font-family:sans-serif;color:#555"><h2>Aplicação encerrada. Você já pode fechar esta aba.</h2></div>';
  } catch (e) {
    window.close();
  }
};

document.querySelectorAll('.modal').forEach(m => m.addEventListener('click', ev => {
  if (ev.target === m && m.id !== 'queueModal' && m.id !== 'olModal') closeModal(m.id);
}));

initApp();
