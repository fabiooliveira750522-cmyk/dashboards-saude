/**
 * Monitoramento da Rede Materno-Infantil de Alagoas
 * ---------------------------------------------------
 * - Menu no topo: Salvar dia / Limpar entrada / Instalar gatilho automático
 * - Varredura horária (6h-22h): arquiva a linha de entrada de cada unidade no histórico.
 * - Regra: 1 registro por DATA. Se a data já existe no histórico, SOBRESCREVE (última versão vence).
 * - Salvar bloqueia data FUTURA (maior que hoje). Datas passadas são permitidas.
 * - Limpar só funciona se a data da entrada já estiver arquivada no histórico.
 *
 * Coordenadas da planilha (fixas):
 *   Linha de cabeçalho de sub-colunas: 3
 *   LINHA DE ENTRADA (editável):       4
 *   Rótulo do histórico:               5
 *   HISTÓRICO começa na linha:         6
 *   Total de colunas:                  22  (A..V)  Data = coluna 1 (A)
 */

var ROW_ENTRY   = 4;
var ROW_HIST    = 6;
var N_COLS      = 22;
var COL_DATA    = 1;   // coluna A
var NON_UNIT    = ['PAINEL', 'CADASTRO', 'HISTÓRICO', 'INSTRUÇÕES'];

var ABA_HIST    = 'HISTÓRICO';   // aba consolidada
var HIST_HDR    = 3;             // linha do cabeçalho na aba HISTÓRICO
var HIST_INI    = 4;             // primeira linha de dados na aba HISTÓRICO
var HIST_COLS   = N_COLS + 1;    // Unidade + 22 colunas

/* ----------------------- MENU ----------------------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 Monitoramento')
    .addItem('💾 Salvar dia (aba atual)', 'salvarDiaAtivo')
    .addItem('🧹 Limpar entrada (aba atual)', 'limparEntradaAtivo')
    .addSeparator()
    .addItem('🔄 Salvar TODAS as unidades agora', 'varreduraManual')
    .addSeparator()
    .addItem('⚙️ Instalar gatilho automático (1x)', 'instalarGatilhoHorario')
    .addItem('🔒 Proteger histórico (aviso)', 'protegerHistoricoManual')
    .addItem('🛑 Remover gatilho automático', 'removerGatilhoHorario')
    .addToUi();
}

/* ------------------- UTILIDADES --------------------- */
function ehAbaUnidade_(sh) {
  return NON_UNIT.indexOf(sh.getName()) === -1;
}

function soData_(d) {                       // zera hora -> comparação por dia
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function lerEntrada_(sh) {
  return sh.getRange(ROW_ENTRY, 1, 1, N_COLS).getValues()[0];
}

// Retorna a linha (número) no histórico cuja Data == data alvo, ou -1 se não achar.
function acharLinhaHistorico_(sh, dataAlvo) {
  var ultima = sh.getLastRow();
  if (ultima < ROW_HIST) return -1;
  var n = ultima - ROW_HIST + 1;
  var datas = sh.getRange(ROW_HIST, COL_DATA, n, 1).getValues();
  var alvo = soData_(dataAlvo).getTime();
  for (var i = 0; i < datas.length; i++) {
    var v = datas[i][0];
    if (v instanceof Date && soData_(v).getTime() === alvo) {
      return ROW_HIST + i;
    }
  }
  return -1;
}

// Primeira linha vazia do histórico (coluna Data vazia).
function primeiraLinhaVazia_(sh) {
  var ultima = sh.getLastRow();
  if (ultima < ROW_HIST) return ROW_HIST;
  var n = ultima - ROW_HIST + 1;
  var datas = sh.getRange(ROW_HIST, COL_DATA, n, 1).getValues();
  for (var i = 0; i < datas.length; i++) {
    if (datas[i][0] === '' || datas[i][0] === null) return ROW_HIST + i;
  }
  return ultima + 1;
}

/* --------------- ARQUIVAR (núcleo) ------------------ */
// Arquiva a entrada de UMA aba. Retorna objeto {ok, msg}.
// modo: 'manual' mostra alertas; 'auto' fica silencioso.
function arquivarAba_(sh, modo) {
  var entrada = lerEntrada_(sh);
  var data = entrada[COL_DATA - 1];

  if (!(data instanceof Date)) {
    return { ok: false, msg: 'A célula Data (linha ' + ROW_ENTRY + ') está vazia ou não é uma data válida.' };
  }
  // Bloqueia data futura
  if (soData_(data).getTime() > soData_(new Date()).getTime()) {
    return { ok: false, msg: 'Não é possível salvar uma data futura. Corrija a Data antes de salvar.' };
  }

  var linha = acharLinhaHistorico_(sh, data);
  if (linha === -1) linha = primeiraLinhaVazia_(sh);   // nova data -> nova linha
  // grava valores (não fórmulas) para congelar o registro
  sh.getRange(linha, 1, 1, N_COLS).setValues([entrada]);
  sh.getRange(linha, COL_DATA).setNumberFormat('dd/mm/yyyy');
  // consolida também na aba HISTÓRICO (unidade + data como chave)
  consolidarHistorico_(sh.getName(), entrada, data);
  return { ok: true, msg: 'Dia ' + Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
                          ' arquivado no histórico (linha ' + linha + ').' };
}

// Insere/atualiza a linha (unidade + data) na aba HISTÓRICO consolidada.
function consolidarHistorico_(nomeUnidade, entrada, data) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HIST);
  if (!sh) return;                        // aba ainda não existe: ignora sem erro
  var linha = acharLinhaHistConsolidado_(sh, nomeUnidade, data);
  if (linha === -1) linha = primeiraLinhaVaziaHist_(sh);
  var valores = [nomeUnidade].concat(entrada);   // Unidade + 22 colunas
  sh.getRange(linha, 1, 1, HIST_COLS).setValues([valores]);
  sh.getRange(linha, 2).setNumberFormat('dd/mm/yyyy');   // coluna B = Data
}

function acharLinhaHistConsolidado_(sh, nomeUnidade, data) {
  var ultima = sh.getLastRow();
  if (ultima < HIST_INI) return -1;
  var n = ultima - HIST_INI + 1;
  var dados = sh.getRange(HIST_INI, 1, n, 2).getValues();   // col A=Unidade, B=Data
  var alvo = soData_(data).getTime();
  for (var i = 0; i < dados.length; i++) {
    var u = dados[i][0], d = dados[i][1];
    if (u === nomeUnidade && d instanceof Date && soData_(d).getTime() === alvo) {
      return HIST_INI + i;
    }
  }
  return -1;
}

function primeiraLinhaVaziaHist_(sh) {
  var ultima = sh.getLastRow();
  if (ultima < HIST_INI) return HIST_INI;
  var n = ultima - HIST_INI + 1;
  var col = sh.getRange(HIST_INI, 1, n, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (col[i][0] === '' || col[i][0] === null) return HIST_INI + i;
  }
  return ultima + 1;
}

/* --------------- AÇÕES DO MENU ---------------------- */
function salvarDiaAtivo() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (!ehAbaUnidade_(sh)) { ui.alert('Abra a aba de uma unidade para salvar o dia.'); return; }
  var r = arquivarAba_(sh, 'manual');
  ui.alert(r.ok ? '✅ ' + r.msg : '⚠️ ' + r.msg);
}

function limparEntradaAtivo() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSheet();
  if (!ehAbaUnidade_(sh)) { ui.alert('Abra a aba de uma unidade para limpar.'); return; }

  var entrada = lerEntrada_(sh);
  var data = entrada[COL_DATA - 1];
  if (!(data instanceof Date)) {
    // sem data: nada arquivado, mas nada a perder -> limpa direto
    sh.getRange(ROW_ENTRY, 1, 1, N_COLS).clearContent();
    return;
  }
  var linha = acharLinhaHistorico_(sh, data);
  if (linha === -1) {
    ui.alert('⚠️ O dia atual (' + Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
             ') ainda NÃO foi salvo no histórico. Clique em "Salvar dia" antes de limpar.');
    return;
  }
  var resp = ui.alert('Limpar a linha de entrada?', 'O dia ' +
             Utilities.formatDate(data, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
             ' já está no histórico. Deseja limpar a entrada para um novo lançamento?',
             ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    sh.getRange(ROW_ENTRY, 1, 1, N_COLS).clearContent();
  }
}

function varreduraManual() {
  var res = varreduraTodas_('manual');
  SpreadsheetApp.getUi().alert('🔄 ' + res.salvas + ' unidade(s) arquivada(s).\n' + res.detalhe);
}

/* --------------- VARREDURA (todas) ------------------ */
function varreduraTodas_(modo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abas = ss.getSheets();
  var salvas = 0, detalhe = '';
  for (var i = 0; i < abas.length; i++) {
    var sh = abas[i];
    if (!ehAbaUnidade_(sh)) continue;
    var entrada = lerEntrada_(sh);
    var data = entrada[COL_DATA - 1];
    if (!(data instanceof Date)) continue;                 // sem data -> ignora
    if (soData_(data).getTime() > soData_(new Date()).getTime()) continue; // futura -> ignora
    var r = arquivarAba_(sh, 'auto');
    if (r.ok) { salvas++; detalhe += '• ' + sh.getName() + ': ' + r.msg + '\n'; }
  }
  return { salvas: salvas, detalhe: detalhe };
}

// Chamada pelo gatilho de tempo.
function varreduraAutomatica() {
  var h = new Date().getHours();
  if (h < 6 || h > 22) return;      // roda só entre 6h e 22h
  varreduraTodas_('auto');
}

/* --------------- GATILHO DE TEMPO ------------------- */
function instalarGatilhoHorario() {
  removerGatilhoHorario();          // evita duplicar
  ScriptApp.newTrigger('varreduraAutomatica')
    .timeBased()
    .everyHours(1)
    .create();
  protegerHistoricoTodas_();        // aplica proteção-com-aviso no histórico
  SpreadsheetApp.getUi().alert('✅ Gatilho automático instalado (arquiva de hora em hora, 6h–22h) '
    + 'e histórico protegido com aviso em todas as unidades.');
}

/* --------- PROTEÇÃO COM AVISO NO HISTÓRICO ---------- */
// Protege a área de histórico (linha 6 pra baixo) com AVISO — quem editar manualmente
// recebe um alerta "tem certeza?", mas o script (que roda como o próprio usuário) continua gravando.
function protegerHistoricoTodas_() {
  var abas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < abas.length; i++) {
    var sh = abas[i];
    if (!ehAbaUnidade_(sh)) continue;
    // remove proteções antigas dessa área para não empilhar
    var prot = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var p = 0; p < prot.length; p++) {
      if (prot[p].getDescription() === 'Histórico (não editar)') prot[p].remove();
    }
    var rng = sh.getRange(ROW_HIST, 1, sh.getMaxRows() - ROW_HIST + 1, N_COLS);
    var protecao = rng.protect()
      .setDescription('Histórico (não editar)')
      .setWarningOnly(true);   // AVISO, não bloqueio -> script continua funcionando
  }
}

// Item de menu para (re)aplicar só a proteção, se precisar.
function protegerHistoricoManual() {
  protegerHistoricoTodas_();
  SpreadsheetApp.getUi().alert('🔒 Histórico protegido com aviso em todas as unidades.');
}

function removerGatilhoHorario() {
  var gs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < gs.length; i++) {
    if (gs[i].getHandlerFunction() === 'varreduraAutomatica') ScriptApp.deleteTrigger(gs[i]);
  }
}
