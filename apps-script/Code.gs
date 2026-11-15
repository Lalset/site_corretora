/**
 * Painel de vendas — Raquel Andrade
 * Backend em Google Apps Script: guarda as vendas nesta planilha.
 *
 * Abas criadas automaticamente:
 *   Vendas            → uma linha por venda (colunas da direita ficam ocultas: são do sistema)
 *   Correspondentes   → nome e WhatsApp
 *   Config            → nome que assina as mensagens, intervalo de cobrança
 *
 * Pode editar à vontade na planilha: Cliente, Telefone, Imóvel, Valor e Anotações.
 * Etapa, documentos e histórico: mude pelo painel.
 */

const ABA_VENDAS = 'Vendas';
const ABA_CORR = 'Correspondentes';
const ABA_CFG = 'Config';

const CAB_VENDAS = [
  'Cliente', 'Telefone', 'Imóvel', 'Valor', 'Perfil', 'Correspondente', 'Etapa', 'Desde',
  'Dias na etapa', 'Situação', 'Última cobrança', 'Documentos', 'Docs faltando', 'Anotações',
  'Cadastrada em', 'Atualizada em',
  // colunas do sistema (ocultas)
  'id', 'status', 'etapaId', 'correspondenteId', 'perfis', 'docs', 'historico'
];
const N_VISIVEIS = 16;
const C = {};
CAB_VENDAS.forEach(function (h, i) { C[h] = i; });

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Minhas vendas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------- chamadas feitas pelo painel ---------- */

function carregar() {
  prepararAbas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const sh = ss.getSheetByName(ABA_VENDAS);
  const n = sh.getLastRow() - 1;
  const linhas = n > 0 ? sh.getRange(2, 1, n, CAB_VENDAS.length).getValues() : [];
  const vendas = linhas.filter(function (r) { return r[C.id]; }).map(function (r) {
    return {
      id: String(r[C.id]),
      cliente: texto_(r[C['Cliente']]),
      telefone: texto_(r[C['Telefone']]),
      imovel: texto_(r[C['Imóvel']]),
      valor: texto_(r[C['Valor']]),
      perfis: texto_(r[C.perfis]).split(',').filter(String),
      correspondenteId: texto_(r[C.correspondenteId]),
      etapa: texto_(r[C.etapaId]) || 'atendimento',
      etapaDesde: dia_(r[C['Desde']], tz) || hoje_(tz),
      status: texto_(r[C.status]) || 'ativa',
      ultimaCobranca: dia_(r[C['Última cobrança']], tz) || null,
      notas: texto_(r[C['Anotações']]),
      criadaEm: dia_(r[C['Cadastrada em']], tz) || hoje_(tz),
      docs: json_(r[C.docs], []),
      historico: json_(r[C.historico], [])
    };
  });

  const shC = ss.getSheetByName(ABA_CORR);
  const nc = shC.getLastRow() - 1;
  const correspondentes = (nc > 0 ? shC.getRange(2, 1, nc, 3).getValues() : [])
    .filter(function (r) { return r[0] || r[2]; })
    .map(function (r) {
      return { nome: texto_(r[0]), telefone: texto_(r[1]), id: texto_(r[2]) || Utilities.getUuid().slice(0, 8) };
    });

  const cfg = lerConfig_();
  return {
    settings: {
      nome: cfg.nome || 'Raquel',
      intervalo: Number(cfg.intervalo) || 2,
      correspondentes: correspondentes
    },
    vendas: vendas,
    url: ss.getUrl()
  };
}

/**
 * Recebe só o que mudou:
 * { vendas: [venda...], excluir: [id...], settings: {nome, intervalo} | null, correspondentes: [...] | null }
 */
function sincronizar(lote) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    prepararAbas_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(ABA_VENDAS);

    const vendas = (lote && lote.vendas) || [];
    if (vendas.length) {
      const ids = lerIds_(sh);
      vendas.forEach(function (v) {
        let i = ids.indexOf(String(v.id));
        if (i === -1) { i = ids.length; ids.push(String(v.id)); }
        const r = i + 2;
        sh.getRange(r, 1, 1, CAB_VENDAS.length).setValues([paraLinha_(v)]);
        const celDias = sh.getRange(r, C['Dias na etapa'] + 1);
        if (v.contarDias) celDias.setFormula('=IF(H' + r + '="","",TODAY()-H' + r + ')');
        else celDias.setValue('');
      });
    }

    const excluir = ((lote && lote.excluir) || []).map(String);
    if (excluir.length) {
      const ids = lerIds_(sh);
      for (let i = ids.length - 1; i >= 0; i--) {
        if (excluir.indexOf(ids[i]) !== -1) sh.deleteRow(i + 2);
      }
    }

    if (lote && lote.correspondentes) {
      const shC = ss.getSheetByName(ABA_CORR);
      if (shC.getLastRow() > 1) shC.getRange(2, 1, shC.getLastRow() - 1, 3).clearContent();
      const rows = lote.correspondentes.map(function (c) { return [seguro_(c.nome), seguro_(c.telefone), c.id]; });
      if (rows.length) shC.getRange(2, 1, rows.length, 3).setValues(rows);
    }

    if (lote && lote.settings) {
      const shG = ss.getSheetByName(ABA_CFG);
      shG.getRange(2, 1, 2, 2).setValues([
        ['nome', seguro_(lote.settings.nome || '')],
        ['intervalo', Number(lote.settings.intervalo) || 2]
      ]);
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- auxiliares ---------- */

function prepararAbas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(ABA_VENDAS)) {
    const primeira = ss.getSheets()[0];
    const sh = (ss.getSheets().length === 1 && primeira.getLastRow() === 0)
      ? primeira.setName(ABA_VENDAS)
      : ss.insertSheet(ABA_VENDAS, 0);
    sh.getRange(1, 1, 1, CAB_VENDAS.length).setValues([CAB_VENDAS])
      .setFontWeight('bold').setBackground('#0d2440').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('dd/mm/yyyy');
    sh.getRange('I:I').setNumberFormat('0');
    sh.getRange('K:K').setNumberFormat('dd/mm/yyyy');
    sh.getRange('O:O').setNumberFormat('dd/mm/yyyy');
    sh.getRange('P:P').setNumberFormat('dd/mm/yyyy hh:mm');
    sh.setColumnWidth(1, 180); sh.setColumnWidth(3, 220); sh.setColumnWidth(7, 220);
    sh.setColumnWidth(13, 260); sh.setColumnWidth(14, 260);
    sh.hideColumns(N_VISIVEIS + 1, CAB_VENDAS.length - N_VISIVEIS);
  }

  if (!ss.getSheetByName(ABA_CORR)) {
    const sh = ss.insertSheet(ABA_CORR);
    sh.getRange(1, 1, 1, 3).setValues([['Nome', 'WhatsApp', 'id']])
      .setFontWeight('bold').setBackground('#0d2440').setFontColor('#ffffff');
    sh.getRange('B:B').setNumberFormat('@');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 160);
    sh.hideColumns(3);
  }

  if (!ss.getSheetByName(ABA_CFG)) {
    const sh = ss.insertSheet(ABA_CFG);
    sh.getRange(1, 1, 3, 2).setValues([['Chave', 'Valor'], ['nome', 'Raquel'], ['intervalo', 2]]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
}

function lerIds_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, C.id + 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];
}

function lerConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CFG);
  const out = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function paraLinha_(v) {
  let hist = JSON.stringify(v.historico || []);
  while (hist.length > 45000 && v.historico.length > 1) { v.historico.pop(); hist = JSON.stringify(v.historico); }
  return [
    seguro_(v.cliente), seguro_(v.telefone), seguro_(v.imovel), seguro_(v.valor),
    v.perfilTxt || '', seguro_(v.corrNome || ''), v.etapaNome || '', data_(v.etapaDesde),
    '', v.situacao || '', data_(v.ultimaCobranca), v.docsTxt || '', seguro_(v.faltandoTxt || ''),
    seguro_(v.notas), data_(v.criadaEm), new Date(),
    String(v.id), v.status || 'ativa', v.etapa || '', v.correspondenteId || '',
    (v.perfis || []).join(','), JSON.stringify(v.docs || []), hist
  ];
}

// Evita que um texto começando com "=" vire fórmula na planilha
function seguro_(s) {
  s = s == null ? '' : String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function texto_(x) { return x == null ? '' : String(x); }
function json_(s, padrao) { try { return s ? JSON.parse(s) : padrao; } catch (e) { return padrao; } }
function data_(s) {
  if (!s) return '';
  const p = String(s).slice(0, 10).split('-').map(Number);
  return p.length === 3 && p[0] ? new Date(p[0], p[1] - 1, p[2]) : '';
}
function dia_(x, tz) {
  if (x instanceof Date && !isNaN(x)) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
  const s = String(x || '');
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}
function hoje_(tz) { return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); }
