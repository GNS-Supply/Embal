// ══════════════════════════════════════════════════════════════════════
// planejamento.js
// Módulo de Planejamento de Embalagens — NOVO e ISOLADO do restante do
// sistema (app.js nunca é importado aqui). Reaproveita apenas:
//   - firebase-config.js (mesma conexão Firebase/Firestore de sempre)
//   - as coleções já existentes 'clientes', 'embalagensCat', 'plantas' e
//     a subcoleção 'embalagensCat/{id}/saldos/{plantaId}' (SÓ LEITURA —
//     este módulo nunca escreve nelas)
// Coleções PRÓPRIAS deste módulo (não usadas em nenhum outro lugar):
//   - planejamento_relacoes  (Cliente + Item + Embalagem + Multiplicador)
//   - planejamento_demanda   (projeção de demanda importada por planilha)
// Nenhuma função/variável daqui tem o mesmo nome do app.js original —
// os dois nunca rodam na mesma página, mas a separação de nomes evita
// qualquer confusão futura caso isso mude.
// ══════════════════════════════════════════════════════════════════════
import {
  auth, db, onAuthStateChanged, signOut,
  collection, addDoc, getDocs, getDoc, setDoc, deleteDoc, doc, updateDoc,
  query, orderBy, where, serverTimestamp, writeBatch, collectionGroup
} from './firebase-config.js';

// ── ESTADO ───────────────────────────────────────────────────────────
window._plCurrentUser = null;
window._plUserRole = 'visualizador';
window._plIsAdmMaster = false;
window._plPodeEscrever = false;

window._plClientes = [];
window._plEmbCat = [];
window._plPlantas = [];
window._plRelacoes = [];
window._plDemanda = [];
window._plSaldos = {};      // { [embCatId]: {vazias, cheias} } — relativo à planta escolhida na aba Comparação
window._plPlantaEscolhida = 'matriz';

window._plImportRelacaoRows = null;
window._plImportDemandaRows = null;
window._plUltimoCalculoPeriodo = null; // guarda o resultado do cálculo por período para a aba de Comparação reaproveitar

// ── UTILITÁRIOS (equivalentes pequenos aos do app.js — não importamos
//    app.js para manter os dois módulos totalmente independentes) ─────
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function p2(n){ return String(n).padStart(2,'0'); }
function formatDt(d){ return `${p2(d.getDate())}/${p2(d.getMonth()+1)}/${d.getFullYear()} ${p2(d.getHours())}:${p2(d.getMinutes())}`; }
function ds(){ const d=new Date(); return `${d.getFullYear()}${p2(d.getMonth()+1)}${p2(d.getDate())}`; }
function qtdInteiraValida(n){ return Number.isInteger(n) && n >= 0; }
function isoDate(d){ return `${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())}`; }

window.showToast = (msg, error=false) => {
  const t = document.getElementById('pl-toast');
  t.textContent = msg;
  t.className = 'toast show' + (error ? ' error' : '');
  clearTimeout(window._plToastTimer);
  window._plToastTimer = setTimeout(()=>{ t.className = 'toast'; }, 3500);
};
function showErr(el, msg){ el.style.display = 'block'; el.textContent = msg; }
window.closeModal = (id) => document.getElementById(id).classList.remove('open');

// Converte um valor de célula de planilha (Date, número serial do Excel, ou texto) em
// 'YYYY-MM-DD' — formato que permite comparar/filtrar datas como string, sem precisar de
// biblioteca de datas. Retorna null se não conseguir interpretar a célula como uma data válida.
function parseDataFlexivel(v){
  if (v instanceof Date && !isNaN(v)) return isoDate(v);
  if (typeof v === 'number' && v > 0) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000)); // serial de data do Excel
    if (!isNaN(d)) return isoDate(d);
  }
  const s = String(v ?? '').trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);           // DD/MM/AAAA
  if (m) return `${m[3]}-${p2(m[2])}-${p2(m[1])}`;
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);                  // AAAA-MM-DD
  if (m) return `${m[1]}-${p2(m[2])}-${p2(m[3])}`;
  return null;
}

// Acha, na linha de cabeçalho da planilha, a coluna cujo texto contém um dos nomes candidatos
// (usado na importação de demanda, cuja ordem de colunas não é fixa — ver instrução do usuário).
function acharColuna(headerRow, candidatos){
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] ?? '').trim().toLowerCase();
    if (candidatos.some(c => h.includes(c))) return i;
  }
  return -1;
}

// ── TEMA (aplica o mesmo tema já escolhido no sistema principal — só leitura) ──
async function aplicarTemaAtual(){
  try{
    const snap = await getDoc(doc(db,'configuracoes','tema_sistema'));
    const tema = snap.exists() ? (snap.data().tema || 'magius') : 'magius';
    const linkYMS = document.getElementById('theme-yms');
    const linkMagius = document.getElementById('theme-magius');
    if (tema === 'yms') {
      if (linkYMS) linkYMS.disabled = false;
      if (linkMagius) linkMagius.disabled = true;
    } else {
      if (linkYMS) linkYMS.disabled = true;
      if (linkMagius) linkMagius.disabled = false;
    }
  }catch(e){ console.error('aplicarTemaAtual:', e); }
}

// ── AUTENTICAÇÃO / GATE DE ACESSO ───────────────────────────────────
onAuthStateChanged(auth, async (user) => {
 try {
  if (!user) { window.location.href = 'index.html'; return; }
  window._plCurrentUser = user;

  let ud = null;
  try {
    const udSnap = await getDoc(doc(db,'usuarios', user.uid));
    ud = udSnap.exists() ? udSnap.data() : null;
  } catch(e) { console.error('erro ao carregar usuário:', e); }

  if (!ud || ud.ativo === false) {
    alert('Sua conta está bloqueada ou não foi encontrada. Contate o administrador.');
    window.location.href = 'index.html';
    return;
  }

  window._plUserRole = ud.perfil || 'operador';
  window._plIsAdmMaster = ud.admMaster === true;
  window._plPodeEscrever = ['operador','administrador'].includes(window._plUserRole) || window._plIsAdmMaster;
  const podeAdministrar = window._plUserRole === 'administrador' || window._plIsAdmMaster;

  const elNome = document.getElementById('pl-topbar-name');
  const elRole = document.getElementById('pl-topbar-role');
  if (elNome) elNome.textContent = user.displayName || user.email || '–';
  if (elRole) elRole.textContent = (window._plUserRole || '').toLowerCase();

  await aplicarTemaAtual();

  // módulo habilitado? (toggle fica no painel Admin do sistema principal — ver index.html)
  let moduloAtivo = false;
  try {
    const cfgSnap = await getDoc(doc(db,'configuracoes','modulo_planejamento'));
    moduloAtivo = cfgSnap.exists() ? !!cfgSnap.data().ativo : false;
  } catch(e) { console.error('erro ao checar módulo:', e); }

  if (!moduloAtivo && !podeAdministrar) {
    document.getElementById('pl-loading').style.display = 'none';
    document.getElementById('pl-modulo-desabilitado').style.display = 'flex';
    return;
  }

  // visualizador não vê nenhum controle de escrita (formulários, importação, edição)
  document.querySelectorAll('.pl-somente-escrita').forEach(el => { el.style.display = window._plPodeEscrever ? '' : 'none'; });
  if (!podeAdministrar) {
    const avisoModulo = document.getElementById('pl-aviso-modulo-desabilitado-admin');
    if (avisoModulo) avisoModulo.style.display = 'none';
  }

  const carregamentos = [
    ['Clientes', loadClientes()],
    ['Catálogo de Embalagens', loadEmbCat()],
    ['Plantas', loadPlantas()],
    ['Relações', loadRelacoes()],
    ['Demanda', loadDemanda()]
  ];
  const resultados = await Promise.allSettled(carregamentos.map(c => c[1]));
  const falhas = resultados
    .map((r, i) => ({ nome: carregamentos[i][0], r }))
    .filter(x => x.r.status === 'rejected');
  if (falhas.length) {
    falhas.forEach(f => console.error(`Erro ao carregar ${f.nome}:`, f.r.reason));
    window.showToast(`Erro ao carregar ${falhas.map(f=>f.nome).join(', ')}: ${falhas[0].r.reason?.message || falhas[0].r.reason}`, true);
  }

  popularSelects();
  renderRelacoes();
  renderResumoDemanda();
  if (!moduloAtivo && podeAdministrar) {
    const avisoModulo = document.getElementById('pl-aviso-modulo-desabilitado-admin');
    if (avisoModulo) avisoModulo.style.display = 'block';
  }

  document.getElementById('pl-loading').style.display = 'none';
  document.getElementById('pl-app').style.display = 'block';
  switchPlTab('relacao');
 } catch(e) {
  // rede de segurança: qualquer erro não previsto acima (ex: elemento inesperado, versão de
  // arquivo desatualizada em cache) agora aparece de forma visível, em vez de deixar a tela
  // travada em "Carregando" para sempre sem nenhuma pista do que aconteceu.
  console.error('[planejamento] erro fatal ao inicializar a página:', e);
  const loadingEl = document.getElementById('pl-loading');
  if (loadingEl) loadingEl.innerHTML = `⚠ Ocorreu um erro ao carregar esta página.<br><span style="font-size:12px;color:var(--text3);">Detalhes no console (F12): ${esc(e.message||e)}</span><br><a href="index.html" style="color:var(--accent);margin-top:10px;display:inline-block;">← Voltar ao sistema</a>`;
 }
});

window.doLogoutPl = async () => {
  try { await signOut(auth); } catch(e) { /* ignora */ }
  window.location.href = 'index.html';
};

// ── CARREGAMENTO DE DADOS (leitura de coleções já existentes) ───────
async function loadClientes(){
  const snap = await getDocs(query(collection(db,'clientes'), orderBy('nome')));
  window._plClientes = snap.docs.map(d => ({ id: d.id, ...d.data() }));
}
async function loadEmbCat(){
  const snap = await getDocs(query(collection(db,'embalagensCat'), orderBy('codigo')));
  window._plEmbCat = snap.docs.map(d => ({ id: d.id, ...d.data() }));
}
async function loadPlantas(){
  const snap = await getDocs(query(collection(db,'plantas'), orderBy('nome')));
  window._plPlantas = snap.docs.map(d => ({ id: d.id, ...d.data() }));
  if (!window._plPlantas.length) window._plPlantas = [{ id: 'matriz', nome: 'Matriz' }];
  window._plPlantaEscolhida = window._plPlantas[0].id === 'matriz' ? 'matriz' : window._plPlantas[0].id;
}
async function loadRelacoes(){
  const snap = await getDocs(collection(db,'planejamento_relacoes'));
  window._plRelacoes = snap.docs.map(d => ({ id: d.id, ...d.data() }));
}
async function loadDemanda(){
  const snap = await getDocs(collection(db,'planejamento_demanda'));
  window._plDemanda = snap.docs.map(d => ({ id: d.id, ...d.data() }));
}
// Busca (só leitura) o saldo atual (vazias/cheias) de todas as embalagens, na planta informada —
// mesma subcoleção 'saldos' que o sistema principal já usa; nunca escrevemos aqui.
async function loadSaldosDaPlanta(plantaId){
  const snap = await getDocs(query(collectionGroup(db,'saldos'), where('plantaId','==', plantaId)));
  const saldos = {};
  snap.docs.forEach(d => {
    const embCatId = d.ref.parent.parent.id;
    saldos[embCatId] = { vazias: Number(d.data().vazias)||0, cheias: Number(d.data().cheias)||0 };
  });
  return saldos;
}

function popularSelects(){
  const opts = window._plClientes.map(c => `<option value="${c.id}">${esc(c.nome)}</option>`).join('');
  ['rel-cliente','import-relacao-cliente-info'].forEach(id => {
    const sel = document.getElementById(id);
    if (sel && sel.tagName === 'SELECT') sel.innerHTML = `<option value="">— Selecione —</option>` + opts;
  });
  const plantaOpts = window._plPlantas.map(p => `<option value="${p.id}">${esc(p.nome)}</option>`).join('');
  const selPlanta = document.getElementById('comp-planta');
  if (selPlanta) { selPlanta.innerHTML = plantaOpts; selPlanta.value = window._plPlantaEscolhida; }
}

window.onRelClienteChange = () => {
  const clienteId = document.getElementById('rel-cliente').value;
  const embs = clienteId ? window._plEmbCat.filter(e => e.clienteId === clienteId) : [];
  const sel = document.getElementById('rel-embalagem');
  sel.innerHTML = `<option value="">${clienteId ? '— Selecione —' : '— Selecione o cliente acima —'}</option>`
    + embs.map(e => `<option value="${e.id}">${esc(e.codigo)} – ${esc(e.descricao||'')}</option>`).join('');
  sel.disabled = !clienteId;
};

// ── ABAS ─────────────────────────────────────────────────────────────
window.switchPlTab = (tab) => {
  document.querySelectorAll('.pl-page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.pl-nav-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('pl-tab-' + tab).classList.add('active');
  document.getElementById('pl-nav-' + tab).classList.add('active');
  if (tab === 'relacao') renderRelacoes();
  if (tab === 'demanda') renderResumoDemanda();
};

// ══════════════════════════════════════════════════════════════════════
// PARTE 1 — RELAÇÃO CLIENTE × ITEM × EMBALAGEM × MULTIPLICADOR
// ══════════════════════════════════════════════════════════════════════

// Todas as relações já cadastradas para uma combinação específica de cliente+item —
// usada para decidir automaticamente qual é a padrão (a primeira cadastrada) e para
// impedir mais de uma padrão ao mesmo tempo.
function relacoesDoItemCliente(clienteId, itemCodigo){
  return window._plRelacoes.filter(r => r.clienteId === clienteId && r.itemCodigo === itemCodigo);
}

window.salvarRelacao = async () => {
  if (!window._plPodeEscrever) { window.showToast('Sem permissão.', true); return; }
  const errEl = document.getElementById('relacao-form-error'); errEl.style.display = 'none';

  const clienteId = document.getElementById('rel-cliente').value;
  const cliente = window._plClientes.find(c => c.id === clienteId);
  const embCatId = document.getElementById('rel-embalagem').value;
  const emb = window._plEmbCat.find(e => e.id === embCatId);
  const itemCodigo = document.getElementById('rel-item').value.trim().toUpperCase();
  const multiplicadorRaw = document.getElementById('rel-multiplicador').value;
  const multiplicador = Number(multiplicadorRaw);

  if (!clienteId) { showErr(errEl,'Selecione o cliente.'); return; }
  if (!embCatId) { showErr(errEl,'Selecione a embalagem (precisa já estar cadastrada no Catálogo para este cliente).'); return; }
  if (!itemCodigo) { showErr(errEl,'Informe o código do item.'); return; }
  if (!multiplicadorRaw || !qtdInteiraValida(multiplicador) || multiplicador <= 0) { showErr(errEl,'Informe um multiplicador inteiro maior que zero.'); return; }

  const duplicada = window._plRelacoes.some(r => r.clienteId === clienteId && r.itemCodigo === itemCodigo && r.embCatId === embCatId);
  if (duplicada) { showErr(errEl,'Essa combinação de cliente, item e embalagem já está cadastrada.'); return; }

  const ehPrimeiraDoItem = relacoesDoItemCliente(clienteId, itemCodigo).length === 0;

  const btn = document.getElementById('btn-salvar-relacao');
  btn.disabled = true; btn.textContent = 'Salvando...';
  try {
    const data = {
      clienteId, clienteNome: cliente?.nome || '',
      embCatId, codigoEmbalagem: emb?.codigo || '',
      itemCodigo, multiplicador,
      padrao: ehPrimeiraDoItem,
      criadoEm: serverTimestamp(),
      uid: window._plCurrentUser.uid
    };
    const ref = await addDoc(collection(db,'planejamento_relacoes'), data);
    window._plRelacoes.push({ id: ref.id, ...data });
    window.showToast('✓ Relação salva.' + (ehPrimeiraDoItem ? ' Definida como padrão (primeira para este item/cliente).' : ''));
    document.getElementById('rel-item').value = '';
    document.getElementById('rel-multiplicador').value = '';
    renderRelacoes();
  } catch(e) {
    showErr(errEl, 'Erro: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '✓ Salvar Relação';
  }
};

window.marcarComoPadrao = async (id) => {
  if (!window._plPodeEscrever) { window.showToast('Sem permissão.', true); return; }
  const rel = window._plRelacoes.find(r => r.id === id); if (!rel) return;
  if (rel.padrao) return;
  const outrasPadrao = window._plRelacoes.filter(r => r.clienteId === rel.clienteId && r.itemCodigo === rel.itemCodigo && r.id !== id && r.padrao);
  try {
    for (const outra of outrasPadrao) {
      await updateDoc(doc(db,'planejamento_relacoes', outra.id), { padrao: false });
      outra.padrao = false;
    }
    await updateDoc(doc(db,'planejamento_relacoes', id), { padrao: true });
    rel.padrao = true;
    window.showToast('✓ Embalagem padrão atualizada para ' + rel.itemCodigo + '.');
    renderRelacoes();
  } catch(e) { window.showToast('Erro: ' + e.message, true); }
};

window.excluirRelacao = async (id) => {
  if (!window._plPodeEscrever) { window.showToast('Sem permissão.', true); return; }
  const rel = window._plRelacoes.find(r => r.id === id); if (!rel) return;
  if (!confirm(`Excluir a relação ${rel.itemCodigo} × ${rel.codigoEmbalagem} (${rel.clienteNome})?`)) return;
  try {
    await deleteDoc(doc(db,'planejamento_relacoes', id));
    window._plRelacoes = window._plRelacoes.filter(r => r.id !== id);
    // se a excluída era a padrão e ainda sobraram alternativas, promove a primeira restante
    if (rel.padrao) {
      const restantes = relacoesDoItemCliente(rel.clienteId, rel.itemCodigo);
      if (restantes.length && !restantes.some(r => r.padrao)) {
        await updateDoc(doc(db,'planejamento_relacoes', restantes[0].id), { padrao: true });
        restantes[0].padrao = true;
      }
    }
    window.showToast('✓ Relação excluída.');
    renderRelacoes();
  } catch(e) { window.showToast('Erro: ' + e.message, true); }
};

function badgePadraoAlternativa(padrao){
  return padrao
    ? `<span class="badge-status atendido">Padrão</span>`
    : `<span class="badge-status pendente">Alternativa</span>`;
}

function renderRelacoes(){
  const filtroCliente = document.getElementById('filter-rel-cliente')?.value || '';
  const tbody = document.getElementById('pl-relacoes-grid');
  if (!tbody) return;
  let dados = window._plRelacoes.filter(r => !filtroCliente || r.clienteId === filtroCliente);
  dados = [...dados].sort((a,b) => (a.clienteNome||'').localeCompare(b.clienteNome||'') || a.itemCodigo.localeCompare(b.itemCodigo) || (b.padrao?1:0)-(a.padrao?1:0));

  const empty = document.getElementById('pl-relacoes-empty');
  if (!dados.length) { tbody.innerHTML=''; empty.style.display='block'; return; }
  empty.style.display = 'none';

  tbody.innerHTML = dados.map(r => `
    <tr>
      <td data-label="Cliente">${esc(r.clienteNome)}</td>
      <td data-label="Item" style="font-family:var(--font-mono)">${esc(r.itemCodigo)}</td>
      <td data-label="Embalagem" style="font-family:var(--font-mono)">${esc(r.codigoEmbalagem)}</td>
      <td data-label="Multiplicador" style="font-family:var(--font-mono)">${r.multiplicador}</td>
      <td data-label="Situação">${badgePadraoAlternativa(r.padrao)}</td>
      <td data-label="Ações">
        ${window._plPodeEscrever && !r.padrao ? `<button class="btn btn-secondary btn-xs" style="margin:0 4px 4px 0" onclick="marcarComoPadrao('${r.id}')">Tornar padrão</button>` : ''}
        ${window._plPodeEscrever ? `<button class="btn btn-danger btn-xs" onclick="excluirRelacao('${r.id}')">Excluir</button>` : ''}
      </td>
    </tr>`).join('');
};
window.renderRelacoes = renderRelacoes;

window.buscarRelacoes = () => {
  const modo = document.getElementById('busca-relacao-modo').value || 'item';
  const termo = document.getElementById('busca-relacao-termo').value.trim().toUpperCase();
  const resultsEl = document.getElementById('pl-consulta-results');
  if (!termo) { resultsEl.innerHTML = '<div class="empty-state"><p>Digite um código para buscar.</p></div>'; return; }

  let linhas;
  if (modo === 'item') {
    linhas = window._plRelacoes.filter(r => r.itemCodigo.includes(termo));
    if (!linhas.length) { resultsEl.innerHTML = `<div class="empty-state"><p>Nenhuma embalagem encontrada para o item "${esc(termo)}".</p></div>`; return; }
    linhas.sort((a,b) => (b.padrao?1:0)-(a.padrao?1:0));
    resultsEl.innerHTML = `
      <div class="table-wrap"><div style="overflow-x:auto;"><table>
        <thead><tr><th>Cliente</th><th>Embalagem</th><th>Multiplicador</th><th>Situação</th></tr></thead>
        <tbody>${linhas.map(r => `<tr>
          <td data-label="Cliente">${esc(r.clienteNome)}</td>
          <td data-label="Embalagem" style="font-family:var(--font-mono)">${esc(r.codigoEmbalagem)}</td>
          <td data-label="Multiplicador" style="font-family:var(--font-mono)">${r.multiplicador}</td>
          <td data-label="Situação">${badgePadraoAlternativa(r.padrao)}</td>
        </tr>`).join('')}</tbody>
      </table></div></div>`;
  } else {
    linhas = window._plRelacoes.filter(r => r.codigoEmbalagem.includes(termo));
    if (!linhas.length) { resultsEl.innerHTML = `<div class="empty-state"><p>Nenhum item encontrado para a embalagem "${esc(termo)}".</p></div>`; return; }
    linhas.sort((a,b) => a.itemCodigo.localeCompare(b.itemCodigo));
    resultsEl.innerHTML = `
      <div class="table-wrap"><div style="overflow-x:auto;"><table>
        <thead><tr><th>Item</th><th>Cliente</th><th>Multiplicador</th><th>Situação</th></tr></thead>
        <tbody>${linhas.map(r => `<tr>
          <td data-label="Item" style="font-family:var(--font-mono)">${esc(r.itemCodigo)}</td>
          <td data-label="Cliente">${esc(r.clienteNome)}</td>
          <td data-label="Multiplicador" style="font-family:var(--font-mono)">${r.multiplicador}</td>
          <td data-label="Situação">${badgePadraoAlternativa(r.padrao)}</td>
        </tr>`).join('')}</tbody>
      </table></div></div>`;
  }
};

// ── IMPORTAÇÃO DA RELAÇÃO (planilha: Cliente | Código Embalagem | Código Item | Multiplicador) ──
window.baixarModeloRelacao = () => {
  const ws = XLSX.utils.aoa_to_sheet([['Cliente','Código Embalagem','Código Item','Multiplicador'],['Cliente Exemplo LTDA','EMB001','ITEM001',20]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Modelo');
  XLSX.writeFile(wb, 'modelo_relacao_item_embalagem.xlsx');
};

window.handleImportRelacaoFile = (input) => {
  const errEl = document.getElementById('modal-import-relacao-error'); errEl.style.display = 'none';
  document.getElementById('modal-import-relacao-success').style.display = 'none';
  document.getElementById('import-relacao-summary').style.display = 'none';
  const file = input.files?.[0];
  window._plImportRelacaoRows = null;
  document.getElementById('btn-processar-import-relacao').disabled = true;
  const label = document.getElementById('import-relacao-file-label');
  if (!file) { label.textContent = '📄 Selecionar arquivo…'; return; }
  label.textContent = '📄 ' + file.name;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const dataRows = rows.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
      if (!dataRows.length) { showErr(errEl, 'Nenhuma linha de dados encontrada na planilha.'); return; }
      window._plImportRelacaoRows = dataRows;
      document.getElementById('btn-processar-import-relacao').disabled = false;
      document.getElementById('import-relacao-summary').style.display = 'block';
      document.getElementById('import-relacao-summary').textContent = `${dataRows.length} linha(s) encontrada(s), prontas para processar.`;
    } catch(e) { showErr(errEl, 'Erro ao ler o arquivo: ' + e.message); }
  };
  reader.onerror = () => showErr(errEl, 'Erro ao ler o arquivo.');
  reader.readAsArrayBuffer(file);
};

window.processarImportRelacao = async () => {
  const errEl = document.getElementById('modal-import-relacao-error'); errEl.style.display = 'none';
  const okEl = document.getElementById('modal-import-relacao-success'); okEl.style.display = 'none';
  const rows = window._plImportRelacaoRows;
  if (!rows?.length) { showErr(errEl, 'Selecione um arquivo válido.'); return; }

  const btn = document.getElementById('btn-processar-import-relacao');
  btn.disabled = true; btn.textContent = 'Processando…';

  let criadas = 0, duplicadas = 0, erros = 0;
  const falhas = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const clienteNome = String(row[0] ?? '').trim();
    const codigoEmb = String(row[1] ?? '').trim().toUpperCase();
    const itemCodigo = String(row[2] ?? '').trim().toUpperCase();
    const multiplicador = Number(row[3]);
    const linhaPlanilha = i + 2; // +1 cabeçalho, +1 base 1

    const cliente = window._plClientes.find(c => c.nome?.trim().toLowerCase() === clienteNome.toLowerCase());
    if (!cliente) { erros++; falhas.push(`Linha ${linhaPlanilha}: cliente "${clienteNome}" não encontrado.`); continue; }
    const emb = window._plEmbCat.find(e => e.codigo === codigoEmb && e.clienteId === cliente.id);
    if (!emb) { erros++; falhas.push(`Linha ${linhaPlanilha}: embalagem "${codigoEmb}" não encontrada no catálogo para "${clienteNome}".`); continue; }
    if (!itemCodigo) { erros++; falhas.push(`Linha ${linhaPlanilha}: código do item vazio.`); continue; }
    if (!qtdInteiraValida(multiplicador) || multiplicador <= 0) { erros++; falhas.push(`Linha ${linhaPlanilha}: multiplicador inválido.`); continue; }

    const duplicada = window._plRelacoes.some(r => r.clienteId === cliente.id && r.itemCodigo === itemCodigo && r.embCatId === emb.id);
    if (duplicada) { duplicadas++; continue; }

    const ehPrimeiraDoItem = relacoesDoItemCliente(cliente.id, itemCodigo).length === 0;
    const data = {
      clienteId: cliente.id, clienteNome: cliente.nome,
      embCatId: emb.id, codigoEmbalagem: emb.codigo,
      itemCodigo, multiplicador,
      padrao: ehPrimeiraDoItem,
      criadoEm: serverTimestamp(),
      uid: window._plCurrentUser.uid
    };
    try {
      const ref = await addDoc(collection(db,'planejamento_relacoes'), data);
      window._plRelacoes.push({ id: ref.id, ...data });
      criadas++;
    } catch(e) { erros++; falhas.push(`Linha ${linhaPlanilha}: erro ao salvar (${e.message}).`); }
  }

  okEl.style.display = 'block';
  okEl.textContent = `✓ ${criadas} relação(ões) criada(s), ${duplicadas} já existente(s) (ignorada(s)), ${erros} com erro.`;
  if (falhas.length) {
    console.warn('[importação relação] falhas:', falhas);
    showErr(errEl, `Algumas linhas não foram importadas (veja o console/F12 para a lista completa): ` + falhas.slice(0,4).join(' | ') + (falhas.length > 4 ? ` (+${falhas.length-4} outra(s))` : ''));
  }
  renderRelacoes();
  btn.disabled = false; btn.textContent = 'Processar Importação';
};

// ══════════════════════════════════════════════════════════════════════
// PARTE 2 — PROJEÇÃO DE DEMANDA
// ══════════════════════════════════════════════════════════════════════

function renderResumoDemanda(){
  const el = document.getElementById('pl-demanda-resumo');
  if (!el) return;
  if (!window._plDemanda.length) {
    el.innerHTML = `<div class="empty-state"><p>Nenhuma projeção de demanda importada ainda.</p></div>`;
    return;
  }
  const itens = new Set(window._plDemanda.map(d => d.itemCodigo));
  const clientes = new Set(window._plDemanda.map(d => d.clienteId));
  const datas = window._plDemanda.map(d => d.dataEntrega).filter(Boolean).sort();
  const importadoEm = window._plDemanda[0]?.importadoEmLocal || '–';
  el.innerHTML = `
    <div class="detail-grid">
      <div class="detail-item"><label>REGISTROS IMPORTADOS</label><span>${window._plDemanda.length}</span></div>
      <div class="detail-item"><label>ITENS DISTINTOS</label><span>${itens.size}</span></div>
      <div class="detail-item"><label>CLIENTES DISTINTOS</label><span>${clientes.size}</span></div>
      <div class="detail-item"><label>PERÍODO COBERTO</label><span>${datas.length ? formatDataBR(datas[0]) + ' até ' + formatDataBR(datas[datas.length-1]) : '–'}</span></div>
      <div class="detail-item"><label>ÚLTIMA IMPORTAÇÃO</label><span>${esc(importadoEm)}</span></div>
    </div>`;
}
function formatDataBR(iso){ if(!iso) return '–'; const [y,m,d]=iso.split('-'); return `${d}/${m}/${y}`; }

window.baixarModeloDemanda = () => {
  const ws = XLSX.utils.aoa_to_sheet([['Código do Item','Cliente','Data de Entrega','Quantidade'],['ITEM001','Cliente Exemplo LTDA','15/03/2026',150]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Modelo');
  XLSX.writeFile(wb, 'modelo_projecao_demanda.xlsx');
};

window.handleImportDemandaFile = (input) => {
  const errEl = document.getElementById('modal-import-demanda-error'); errEl.style.display = 'none';
  document.getElementById('modal-import-demanda-success').style.display = 'none';
  document.getElementById('import-demanda-summary').style.display = 'none';
  const file = input.files?.[0];
  window._plImportDemandaRows = null;
  document.getElementById('btn-processar-import-demanda').disabled = true;
  const label = document.getElementById('import-demanda-file-label');
  if (!file) { label.textContent = '📄 Selecionar arquivo…'; return; }
  label.textContent = '📄 ' + file.name;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rows.length < 2) { showErr(errEl, 'Planilha vazia ou sem linhas de dados.'); return; }
      // colunas identificadas pelo NOME do cabeçalho — a ordem pode variar (ver instrução do usuário)
      const header = rows[0];
      const idxItem = acharColuna(header, ['item']);
      const idxCliente = acharColuna(header, ['cliente']);
      const idxData = acharColuna(header, ['data']);
      const idxQtd = acharColuna(header, ['quantidade','qtd']);
      if ([idxItem, idxCliente, idxData, idxQtd].some(i => i === -1)) {
        showErr(errEl, 'Não encontrei todas as colunas esperadas (Código do Item, Cliente, Data de Entrega, Quantidade) no cabeçalho da planilha. Verifique os títulos das colunas.');
        return;
      }
      const dataRows = rows.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
      if (!dataRows.length) { showErr(errEl, 'Nenhuma linha de dados encontrada na planilha.'); return; }
      window._plImportDemandaRows = { dataRows, idxItem, idxCliente, idxData, idxQtd };
      document.getElementById('btn-processar-import-demanda').disabled = false;
      document.getElementById('import-demanda-summary').style.display = 'block';
      document.getElementById('import-demanda-summary').textContent = `${dataRows.length} linha(s) encontrada(s), prontas para processar.`;
    } catch(e) { showErr(errEl, 'Erro ao ler o arquivo: ' + e.message); }
  };
  reader.onerror = () => showErr(errEl, 'Erro ao ler o arquivo.');
  reader.readAsArrayBuffer(file);
};

window.processarImportDemanda = async () => {
  const errEl = document.getElementById('modal-import-demanda-error'); errEl.style.display = 'none';
  const okEl = document.getElementById('modal-import-demanda-success'); okEl.style.display = 'none';
  const info = window._plImportDemandaRows;
  if (!info?.dataRows?.length) { showErr(errEl, 'Selecione um arquivo válido.'); return; }
  if (!confirm('Importar esta planilha vai SUBSTITUIR TODA a projeção de demanda atual (registros antigos serão apagados). Deseja continuar?')) return;

  const btn = document.getElementById('btn-processar-import-demanda');
  btn.disabled = true; btn.textContent = 'Processando…';

  const { dataRows, idxItem, idxCliente, idxData, idxQtd } = info;
  let erros = 0;
  const falhas = [];
  const validas = [];
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const linhaPlanilha = i + 2;
    const itemCodigo = String(row[idxItem] ?? '').trim().toUpperCase();
    const clienteNome = String(row[idxCliente] ?? '').trim();
    const dataEntrega = parseDataFlexivel(row[idxData]);
    const quantidade = Number(row[idxQtd]);

    if (!itemCodigo) { erros++; falhas.push(`Linha ${linhaPlanilha}: código do item vazio.`); continue; }
    const cliente = window._plClientes.find(c => c.nome?.trim().toLowerCase() === clienteNome.toLowerCase());
    if (!cliente) { erros++; falhas.push(`Linha ${linhaPlanilha}: cliente "${clienteNome}" não encontrado.`); continue; }
    if (!dataEntrega) { erros++; falhas.push(`Linha ${linhaPlanilha}: data de entrega inválida.`); continue; }
    if (!quantidade || quantidade <= 0) { erros++; falhas.push(`Linha ${linhaPlanilha}: quantidade inválida.`); continue; }

    validas.push({ itemCodigo, clienteId: cliente.id, clienteNome: cliente.nome, dataEntrega, quantidade });
  }

  if (!validas.length) {
    showErr(errEl, 'Nenhuma linha válida para importar. Verifique os erros no console (F12).');
    console.warn('[importação demanda] falhas:', falhas);
    btn.disabled = false; btn.textContent = 'Substituir Demanda pela Planilha';
    return;
  }

  try {
    // Ordem importante para segurança dos dados: primeiro GRAVA tudo de novo, e só DEPOIS apaga o
    // que existia antes — nessa ordem, se a importação falhar no meio do caminho (rede, etc.), a
    // demanda antiga continua intacta e o usuário pode simplesmente tentar de novo. Fazer o
    // contrário (apagar primeiro) arriscaria deixar o sistema sem nenhuma demanda em caso de falha.
    const antigos = await getDocs(collection(db,'planejamento_demanda'));
    const idsAntigos = antigos.docs.map(d => d.ref);

    // 1) grava a nova demanda em lotes (limite de 500 operações por batch do Firestore)
    const importadoEmLocal = formatDt(new Date());
    let batch = writeBatch(db);
    let opsNoBatch = 0;
    for (const v of validas) {
      const ref = doc(collection(db,'planejamento_demanda'));
      batch.set(ref, { ...v, importadoEm: serverTimestamp(), importadoEmLocal, uid: window._plCurrentUser.uid });
      opsNoBatch++;
      if (opsNoBatch === 450) { await batch.commit(); batch = writeBatch(db); opsNoBatch = 0; }
    }
    if (opsNoBatch > 0) { await batch.commit(); }

    // 2) só agora, com a nova demanda já gravada com sucesso, apaga a que existia antes
    batch = writeBatch(db);
    opsNoBatch = 0;
    for (const ref of idsAntigos) {
      batch.delete(ref);
      opsNoBatch++;
      if (opsNoBatch === 450) { await batch.commit(); batch = writeBatch(db); opsNoBatch = 0; }
    }
    if (opsNoBatch > 0) { await batch.commit(); }

    await loadDemanda();
    renderResumoDemanda();
    okEl.style.display = 'block';
    okEl.textContent = `✓ Demanda substituída: ${validas.length} registro(s) importado(s), ${erros} linha(s) com erro (ignorada(s)).`;
    if (falhas.length) {
      console.warn('[importação demanda] falhas:', falhas);
      showErr(errEl, `Algumas linhas não foram importadas: ` + falhas.slice(0,4).join(' | ') + (falhas.length > 4 ? ` (+${falhas.length-4} outra(s))` : ''));
    }
  } catch(e) {
    showErr(errEl, 'Erro ao gravar a demanda: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = 'Substituir Demanda pela Planilha';
  }
};

// ══════════════════════════════════════════════════════════════════════
// PARTE 3 — CÁLCULO DE NECESSIDADE POR PERÍODO
// ══════════════════════════════════════════════════════════════════════

// Fluxo: soma a demanda de cada (cliente, item) dentro do período → acha a embalagem PADRÃO
// daquele cliente+item → divide pelo multiplicador, sempre arredondando para cima.
// Embalagens alternativas nunca entram nesse cálculo (regra explícita do usuário).
function calcularNecessidadePorPeriodo(dataInicial, dataFinal){
  const porItemCliente = {}; // chave: clienteId + '|' + itemCodigo
  for (const d of window._plDemanda) {
    if (d.dataEntrega < dataInicial || d.dataEntrega > dataFinal) continue;
    const chave = d.clienteId + '|' + d.itemCodigo;
    if (!porItemCliente[chave]) porItemCliente[chave] = { clienteId: d.clienteId, clienteNome: d.clienteNome, itemCodigo: d.itemCodigo, demandaTotal: 0 };
    porItemCliente[chave].demandaTotal += Number(d.quantidade) || 0;
  }

  const resultado = [];
  for (const chave in porItemCliente) {
    const grupo = porItemCliente[chave];
    const padrao = window._plRelacoes.find(r => r.clienteId === grupo.clienteId && r.itemCodigo === grupo.itemCodigo && r.padrao);
    if (!padrao) {
      resultado.push({ ...grupo, embCatId: null, codigoEmbalagem: null, multiplicador: null, necessario: null, semPadrao: true });
      continue;
    }
    const necessario = Math.ceil(grupo.demandaTotal / padrao.multiplicador);
    resultado.push({ ...grupo, embCatId: padrao.embCatId, codigoEmbalagem: padrao.codigoEmbalagem, multiplicador: padrao.multiplicador, necessario, semPadrao: false });
  }
  resultado.sort((a,b) => (a.clienteNome||'').localeCompare(b.clienteNome||'') || a.itemCodigo.localeCompare(b.itemCodigo));
  return resultado;
}

window.calcularPeriodo = () => {
  const errEl = document.getElementById('pl-periodo-error'); errEl.style.display = 'none';
  const dataInicial = document.getElementById('pl-periodo-inicio').value;
  const dataFinal = document.getElementById('pl-periodo-fim').value;
  if (!dataInicial || !dataFinal) { showErr(errEl, 'Informe a data inicial e a data final.'); return; }
  if (dataInicial > dataFinal) { showErr(errEl, 'A data inicial não pode ser depois da data final.'); return; }

  const resultado = calcularNecessidadePorPeriodo(dataInicial, dataFinal);
  window._plUltimoCalculoPeriodo = { dataInicial, dataFinal, resultado };

  const tbody = document.getElementById('pl-periodo-grid');
  const empty = document.getElementById('pl-periodo-empty');
  if (!resultado.length) { tbody.innerHTML = ''; empty.style.display = 'block'; return; }
  empty.style.display = 'none';

  tbody.innerHTML = resultado.map(r => `
    <tr>
      <td data-label="Cliente">${esc(r.clienteNome)}</td>
      <td data-label="Item" style="font-family:var(--font-mono)">${esc(r.itemCodigo)}</td>
      <td data-label="Demanda no Período" style="font-family:var(--font-mono)">${r.demandaTotal}</td>
      <td data-label="Embalagem Padrão" style="font-family:var(--font-mono)">${r.semPadrao ? '<span style="color:var(--warn)">— sem padrão cadastrada —</span>' : esc(r.codigoEmbalagem)}</td>
      <td data-label="Multiplicador" style="font-family:var(--font-mono)">${r.semPadrao ? '–' : r.multiplicador}</td>
      <td data-label="Qtd. Necessária" style="font-family:var(--font-mono);font-weight:700;">${r.semPadrao ? '–' : r.necessario}</td>
    </tr>`).join('');

  document.getElementById('pl-btn-ir-comparacao').style.display = 'inline-flex';
};

window.exportarPeriodoXLSX = () => {
  const calc = window._plUltimoCalculoPeriodo;
  if (!calc?.resultado?.length) { window.showToast('Calcule um período primeiro.', true); return; }
  const linhas = [['Cliente','Item','Demanda no Período','Embalagem Padrão','Multiplicador','Qtd. Necessária']];
  calc.resultado.forEach(r => linhas.push([r.clienteNome, r.itemCodigo, r.demandaTotal, r.semPadrao ? 'sem padrão' : r.codigoEmbalagem, r.semPadrao ? '' : r.multiplicador, r.semPadrao ? '' : r.necessario]));
  const ws = XLSX.utils.aoa_to_sheet(linhas);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Necessidade');
  XLSX.writeFile(wb, `necessidade_embalagens_${ds()}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════════
// PARTE 4 — COMPARAÇÃO COM ESTOQUE (por cliente + embalagem)
// ══════════════════════════════════════════════════════════════════════

window.irParaComparacaoComPeriodo = () => {
  const calc = window._plUltimoCalculoPeriodo;
  if (calc) {
    document.getElementById('comp-periodo-inicio').value = calc.dataInicial;
    document.getElementById('comp-periodo-fim').value = calc.dataFinal;
  }
  switchPlTab('comparacao');
};

window.calcularComparacao = async () => {
  const errEl = document.getElementById('pl-comparacao-error'); errEl.style.display = 'none';
  const dataInicial = document.getElementById('comp-periodo-inicio').value;
  const dataFinal = document.getElementById('comp-periodo-fim').value;
  const plantaId = document.getElementById('comp-planta').value || 'matriz';
  if (!dataInicial || !dataFinal) { showErr(errEl, 'Informe a data inicial e a data final.'); return; }
  if (dataInicial > dataFinal) { showErr(errEl, 'A data inicial não pode ser depois da data final.'); return; }

  const btn = document.getElementById('btn-calcular-comparacao');
  btn.disabled = true; btn.textContent = 'Calculando…';
  try {
    window._plPlantaEscolhida = plantaId;
    window._plSaldos = await loadSaldosDaPlanta(plantaId);

    const todosOsItens = calcularNecessidadePorPeriodo(dataInicial, dataFinal);
    const semPadrao = todosOsItens.filter(r => r.semPadrao);
    const porItemCliente = todosOsItens.filter(r => !r.semPadrao);

    const avisoEl = document.getElementById('pl-comparacao-aviso-sem-padrao');
    if (semPadrao.length) {
      avisoEl.style.display = 'block';
      avisoEl.textContent = `⚠ ${semPadrao.length} item(ns) com demanda no período não entraram nesta comparação por não terem uma embalagem padrão cadastrada: ${semPadrao.slice(0,6).map(r=>r.itemCodigo).join(', ')}${semPadrao.length>6 ? ` (+${semPadrao.length-6} outro(s))` : ''}.`;
    } else {
      avisoEl.style.display = 'none';
    }

    // re-agrupa por (cliente, embalagem padrão) — várias itens podem compartilhar a mesma
    // embalagem padrão para o mesmo cliente, e a necessidade deles se SOMA nessa embalagem.
    const porClienteEmb = {};
    for (const r of porItemCliente) {
      const chave = r.clienteId + '|' + r.embCatId;
      if (!porClienteEmb[chave]) porClienteEmb[chave] = { clienteId: r.clienteId, clienteNome: r.clienteNome, embCatId: r.embCatId, codigoEmbalagem: r.codigoEmbalagem, necessario: 0, itens: [] };
      porClienteEmb[chave].necessario += r.necessario;
      porClienteEmb[chave].itens.push(r.itemCodigo);
    }

    const resultado = Object.values(porClienteEmb).map(g => {
      const saldo = window._plSaldos[g.embCatId] || { vazias: 0, cheias: 0 };
      const disponivel = saldo.vazias + saldo.cheias;
      const faltante = Math.max(0, g.necessario - disponivel);
      const excedente = Math.max(0, disponivel - g.necessario);
      return { ...g, vazias: saldo.vazias, cheias: saldo.cheias, disponivel, faltante, excedente };
    }).sort((a,b) => (a.clienteNome||'').localeCompare(b.clienteNome||'') || a.codigoEmbalagem.localeCompare(b.codigoEmbalagem));

    window._plUltimaComparacao = resultado;
    const tbody = document.getElementById('pl-comparacao-grid');
    const empty = document.getElementById('pl-comparacao-empty');
    if (!resultado.length) { tbody.innerHTML=''; empty.style.display='block'; return; }
    empty.style.display = 'none';

    tbody.innerHTML = resultado.map(r => `
      <tr>
        <td data-label="Cliente">${esc(r.clienteNome)}</td>
        <td data-label="Embalagem" style="font-family:var(--font-mono)" title="Itens: ${esc(r.itens.join(', '))}">${esc(r.codigoEmbalagem)}</td>
        <td data-label="Vazias" style="font-family:var(--font-mono)">${r.vazias}</td>
        <td data-label="Cheias" style="font-family:var(--font-mono)">${r.cheias}</td>
        <td data-label="Disponível" style="font-family:var(--font-mono);font-weight:700;">${r.disponivel}</td>
        <td data-label="Necessário" style="font-family:var(--font-mono)">${r.necessario}</td>
        <td data-label="Faltante" style="font-family:var(--font-mono);font-weight:700;color:${r.faltante>0?'var(--danger)':'var(--text2)'}">${r.faltante}</td>
        <td data-label="Excedente" style="font-family:var(--font-mono);color:${r.excedente>0?'var(--ok)':'var(--text2)'}">${r.excedente}</td>
      </tr>`).join('');
  } catch(e) {
    showErr(errEl, 'Erro ao calcular: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '✓ Calcular Comparação';
  }
};

window.exportarComparacaoXLSX = () => {
  const dados = window._plUltimaComparacao;
  if (!dados?.length) { window.showToast('Calcule a comparação primeiro.', true); return; }
  const linhas = [['Cliente','Embalagem','Vazias','Cheias','Disponível','Necessário','Faltante','Excedente']];
  dados.forEach(r => linhas.push([r.clienteNome, r.codigoEmbalagem, r.vazias, r.cheias, r.disponivel, r.necessario, r.faltante, r.excedente]));
  const ws = XLSX.utils.aoa_to_sheet(linhas);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Comparação');
  XLSX.writeFile(wb, `comparacao_estoque_demanda_${ds()}.xlsx`);
};
