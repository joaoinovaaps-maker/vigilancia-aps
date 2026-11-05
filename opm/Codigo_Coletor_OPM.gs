/**
 * COLETOR DO BANCO ANONIMIZADO DE OPM
 * UBS Ildefonso Pedroso - Dourados/MS
 *
 * Recebe os registros enviados pelo formulario (aba 7 - Banco de dados),
 * grava na planilha e envia os dados por e-mail ao endereco institucional.
 * Os registros NAO contem dados pessoais do paciente.
 *
 * Instalacao: Planilha > Extensoes > Apps Script > colar este codigo > salvar.
 * Depois: Implantar > Nova implantacao > App da Web
 *         Executar como: Eu | Quem pode acessar: Qualquer pessoa
 */

// ===================== CONFIGURACAO =====================
const CONFIG = {
  ABA: 'Registros OPM',
  // E-mail(s) institucional(is) de destino. Varios: separe por virgula.
  EMAIL_DESTINO: 'seu.email@institucional.gov.br',
  // true = envia um e-mail a cada solicitacao registrada/atualizada
  ENVIAR_A_CADA_REGISTRO: true,
  // true = envia novo e-mail tambem quando um registro existente e atualizado
  ENVIAR_ATUALIZACOES: false,
  // Anexa o registro em CSV no e-mail individual
  ANEXAR_CSV_NO_REGISTRO: true,
  // Pasta do Google Drive onde ficam as copias de seguranca (CSV) dos dados nao pessoais
  PASTA_DRIVE: 'Banco OPM - dados anonimizados',
  ASSUNTO_PREFIXO: '[OPM] ',
  NOME_REMETENTE: 'Banco OPM - UBS Ildefonso Pedroso'
};
// =========================================================

/** Recebe os dados do formulario. */
function doPost(e) {
  if (!e || !e.postData) {
    console.log('doPost só funciona quando o formulário envia dados. ' +
      'Para testar pelo editor, selecione a função "testarEmail" ou "testarRegistro" e clique em Executar.');
    return texto_('sem dados');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sh = obterAba_(data.heads);
    if (data.teste) return texto_('ok');

    const linha = data.cols.map(function (c) { return limpar_(data.row[c]); });
    const n = sh.getLastRow();
    const ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    const i = ids.indexOf(data.row.id);
    const novo = i < 0;

    if (novo) sh.appendRow(linha);
    else sh.getRange(i + 2, 1, 1, linha.length).setValues([linha]);

    if (CONFIG.ENVIAR_A_CADA_REGISTRO && (novo || CONFIG.ENVIAR_ATUALIZACOES)) {
      try {
        enviarEmailRegistro_(data.heads, linha, novo ? 'Nova solicitação' : 'Solicitação atualizada');
      } catch (err) {
        console.error('Falha no e-mail: ' + err);
      }
    }
    return texto_('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return texto_('Coletor OPM ativo'); }

// ===================== E-MAIL POR REGISTRO =====================
function enviarEmailRegistro_(heads, linha, status) {
  const get = function (nome) { const k = heads.indexOf(nome); return k >= 0 ? linha[k] : ''; };
  const id = get('ID do registro');
  const opm = get('Procedimento principal - nome') || 'OPM não informada';

  let tabela = '';
  heads.forEach(function (h, k) {
    const v = linha[k];
    if (v === '' || v === null || v === undefined) return;
    tabela += '<tr><td style="padding:4px 10px 4px 0;color:#555;border-bottom:1px solid #eee">' + esc_(h) +
      '</td><td style="padding:4px 0;border-bottom:1px solid #eee"><b>' + esc_(v) + '</b></td></tr>';
  });

  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1c2631">' +
    '<p><b>' + esc_(status) + '</b> registrada no banco anonimizado de OPM.</p>' +
    '<table style="border-collapse:collapse;font-size:13px">' + tabela + '</table>' +
    '<p style="font-size:12px;color:#777;margin-top:14px">Registro sem dados pessoais do paciente. ' +
    'Planilha consolidada: <a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '">abrir</a>.</p></div>';

  const opcoes = { to: CONFIG.EMAIL_DESTINO, subject: CONFIG.ASSUNTO_PREFIXO + status + ' - ' + opm + ' - ' + id,
    htmlBody: html, name: CONFIG.NOME_REMETENTE };
  if (CONFIG.ANEXAR_CSV_NO_REGISTRO) {
    opcoes.attachments = [Utilities.newBlob(
      '﻿' + [heads.map(csv_).join(';'), linha.map(csv_).join(';')].join('\r\n'),
      'text/csv', 'registro_' + id + '.csv')];
  }
  MailApp.sendEmail(opcoes);
}

// ===================== RELATORIO CONSOLIDADO =====================
/**
 * Envia por e-mail a planilha completa (XLSX) com um resumo por OPM, ESF e patologia.
 * Pode ser executado manualmente ou de forma automatica (ver criarEnvioSemanal).
 */
function enviarRelatorioConsolidado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.ABA);
  if (!sh || sh.getLastRow() < 2) { console.log('Sem registros para enviar.'); return; }

  const dados = sh.getDataRange().getValues();
  const heads = dados[0], linhas = dados.slice(1);
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  const blocos = ['Procedimento principal - nome', 'Equipe ESF', 'Patologia', 'Faixa etária', 'Habilidade do usuário']
    .map(function (col) { return resumoHtml_(col, contar_(heads, linhas, col)); }).join('');

  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1c2631">' +
    '<p>Relatório consolidado do banco anonimizado de OPM em ' + hoje + '.</p>' +
    '<p>Total de solicitações registradas: <b>' + linhas.length + '</b></p>' + blocos +
    '<p style="font-size:12px;color:#777">Planilha completa em anexo (XLSX). Sem dados pessoais do paciente.</p></div>';

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx&gid=' + sh.getSheetId();
  const xlsx = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName('Banco_OPM_' + hoje.replace(/\//g, '-') + '.xlsx');

  MailApp.sendEmail({ to: CONFIG.EMAIL_DESTINO, subject: CONFIG.ASSUNTO_PREFIXO + 'Relatório consolidado - ' + hoje,
    htmlBody: html, attachments: [xlsx], name: CONFIG.NOME_REMETENTE });
}

/** Agenda o relatorio consolidado para toda segunda-feira, por volta das 7h. Execute uma vez. */
function criarEnvioSemanal() {
  removerEnvioAutomatico();
  ScriptApp.newTrigger('enviarRelatorioConsolidado').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

/** Agenda o relatorio consolidado para o dia 1 de cada mes, por volta das 7h. Execute uma vez. */
function criarEnvioMensal() {
  removerEnvioAutomatico();
  ScriptApp.newTrigger('enviarRelatorioConsolidado').timeBased().onMonthDay(1).atHour(7).create();
}

// ===================== ARMAZENAMENTO NO DRIVE =====================
/**
 * Salva uma copia CSV da aba "Registros OPM" na pasta do Drive definida em CONFIG.PASTA_DRIVE.
 * Mantem um arquivo "Banco_OPM_atual.csv" (sempre substituido) e um arquivo datado por dia.
 */
function salvarCopiaNoDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.ABA);
  if (!sh || sh.getLastRow() < 2) { console.log('Sem registros para salvar.'); return; }
  const pasta = obterPasta_();
  const conteudo = '﻿' + sh.getDataRange().getValues().map(function (r) {
    return r.map(function (v) {
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      return csv_(v);
    }).join(';');
  }).join('\r\n');
  const dia = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  [['Banco_OPM_atual.csv'], ['Banco_OPM_' + dia + '.csv']].forEach(function (n) {
    const it = pasta.getFilesByName(n[0]);
    if (it.hasNext()) it.next().setContent(conteudo);
    else pasta.createFile(n[0], conteudo, MimeType.CSV);
  });
  console.log('Cópia salva no Drive: ' + pasta.getUrl());
}

/** Agenda a copia no Drive todos os dias por volta das 22h. Execute uma vez. */
function criarCopiaDiariaNoDrive() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'salvarCopiaNoDrive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('salvarCopiaNoDrive').timeBased().everyDays(1).atHour(22).create();
  salvarCopiaNoDrive();
}

/** Move a propria planilha para a pasta do Drive, deixando tudo no mesmo lugar. Execute uma vez. */
function organizarPlanilhaNaPasta() {
  const pasta = obterPasta_();
  DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).moveTo(pasta);
  console.log('Planilha movida para: ' + pasta.getUrl());
}

function obterPasta_() {
  const it = DriveApp.getFoldersByName(CONFIG.PASTA_DRIVE);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.PASTA_DRIVE);
}

/** Cancela o envio automatico do relatorio. */
function removerEnvioAutomatico() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarRelatorioConsolidado') ScriptApp.deleteTrigger(t);
  });
}

/** Teste: envia um e-mail simples ao destino configurado. Execute uma vez para autorizar o script. */
function testarEmail() {
  MailApp.sendEmail(CONFIG.EMAIL_DESTINO, CONFIG.ASSUNTO_PREFIXO + 'Teste do coletor',
    'O coletor do banco de OPM está configurado corretamente.\nPlanilha: ' +
    SpreadsheetApp.getActiveSpreadsheet().getUrl());
  console.log('Cota restante de e-mails hoje: ' + MailApp.getRemainingDailyQuota());
}

/** Teste completo pelo editor: simula um envio do formulario (grava linha de teste e envia e-mail). */
function testarRegistro() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ABA);
  if (!sh || sh.getLastRow() === 0) {
    console.log('A aba "' + CONFIG.ABA + '" ainda não tem cabeçalho. Primeiro implante o App da Web e clique em ' +
      '"Testar" na aba 7 do formulário; depois execute testarRegistro novamente.');
    return;
  }
  const heads = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const valores = { 'ID do registro': 'TESTE-' + Date.now(), 'Registrado em': new Date().toLocaleString('pt-BR'),
    'Data da solicitação': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    'Procedimento principal - nome': 'CADEIRA DE RODAS (TESTE)', 'Largura do assento (cm)': 44,
    'Comprimento do assento (cm)': 42, 'Altura do encosto (cm)': 40, 'Equipe ESF': 'ESF 60 - Ipê Amarelo' };
  const row = {};
  heads.forEach(function (h) { row[h] = valores[h] !== undefined ? valores[h] : ''; });
  row.id = valores['ID do registro'];
  const r = doPost({ postData: { contents: JSON.stringify({ cols: heads, heads: heads, row: row }) } });
  console.log('Resultado: ' + r.getContent() + '. Confira a planilha e o e-mail. Apague a linha de teste depois.');
}

// ===================== AUXILIARES =====================
function obterAba_(heads) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.ABA) || ss.insertSheet(CONFIG.ABA);
  if (sh.getLastRow() === 0 && heads && heads.length) {
    sh.appendRow(heads);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, heads.length).setFontWeight('bold').setBackground('#e5f2ec');
  }
  return sh;
}
function limpar_(v) {
  if (v === undefined || v === null) return '';
  return (typeof v === 'string' && /^[=+@-]/.test(v)) ? "'" + v : v;
}
function contar_(heads, linhas, col) {
  const k = heads.indexOf(col), m = {};
  if (k < 0) return [];
  linhas.forEach(function (r) {
    String(r[k] || 'Não informado').split(/, | \| /).forEach(function (v) { v = v || 'Não informado'; m[v] = (m[v] || 0) + 1; });
  });
  return Object.keys(m).map(function (x) { return [x, m[x]]; }).sort(function (a, b) { return b[1] - a[1]; });
}
function resumoHtml_(titulo, pares) {
  if (!pares.length) return '';
  return '<p style="margin:14px 0 4px"><b>' + esc_(titulo.replace(' - nome', '')) + '</b></p>' +
    '<table style="border-collapse:collapse;font-size:13px">' + pares.map(function (p) {
      return '<tr><td style="padding:3px 16px 3px 0;border-bottom:1px solid #eee">' + esc_(p[0]) +
        '</td><td style="text-align:right;border-bottom:1px solid #eee">' + p[1] + '</td></tr>';
    }).join('') + '</table>';
}
function csv_(v) {
  const s = String(v === null || v === undefined ? '' : v);
  return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}
function texto_(t) { return ContentService.createTextOutput(t); }
