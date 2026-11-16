/**
 * Comparativo F&I — site com histórico de fechamentos (Google Apps Script).
 * Os fechamentos ficam na aba FECHAMENTOS desta planilha; quem pode salvar/excluir
 * está na aba EDITORES (um e-mail por linha). O dono da planilha sempre pode.
 */
const ABA_FECH = 'FECHAMENTOS';
const ABA_EDIT = 'EDITORES';
const PEDACO = 45000; // uma célula guarda até 50.000 caracteres: o JSON é dividido em partes
const CAB = ['id', 'nome', 'criadoEm', 'arquivo', 'autor', 'partes'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Comparativo F&I GB Norte')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function planilha_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function aba_(nome, cabecalho) {
  const ss = planilha_();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Rode uma vez pelo editor do Apps Script (botão Executar) para criar as abas e autorizar. */
function configurar() {
  aba_(ABA_FECH, CAB);
  const ed = aba_(ABA_EDIT, ['E-mail de quem pode salvar e excluir fechamentos']);
  const dono = Session.getEffectiveUser().getEmail();
  if (dono && ed.getLastRow() < 2) ed.getRange(2, 1).setValue(dono);
}

function emailAtual_() { return (Session.getActiveUser().getEmail() || '').toLowerCase(); }

function ehEditor_() {
  const email = emailAtual_();
  if (!email) return false;
  if (email === (Session.getEffectiveUser().getEmail() || '').toLowerCase()) return true;
  const sh = aba_(ABA_EDIT, ['E-mail de quem pode salvar e excluir fechamentos']);
  if (sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .some(r => String(r[0]).trim().toLowerCase() === email);
}

function quemSou() { return {email: emailAtual_(), editor: ehEditor_()}; }

/** Lista os fechamentos (sem os dados), do mais recente para o mais antigo. */
function listarFechamentos() {
  const sh = aba_(ABA_FECH, CAB);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => r[0])
    .map(r => ({id: String(r[0]), nome: String(r[1]), criadoEm: String(r[2]), arquivo: String(r[3] || ''), autor: String(r[4] || '')}))
    .sort((a, b) => b.criadoEm.localeCompare(a.criadoEm));
}

function linhaDe_(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

/** Devolve os dados (texto JSON) de um fechamento. */
function obterFechamento(id) {
  const sh = aba_(ABA_FECH, CAB);
  const lin = linhaDe_(sh, String(id));
  if (lin < 0) throw new Error('Fechamento não encontrado.');
  const partes = Number(sh.getRange(lin, 6).getValue()) || 0;
  if (!partes) return null;
  return sh.getRange(lin, 7, 1, partes).getDisplayValues()[0].join('');
}

/** Guarda um fechamento novo. Nunca altera os anteriores. */
function salvarFechamento(nome, arquivo, dadosJson) {
  if (!ehEditor_()) throw new Error('SEM_PERMISSAO');
  nome = String(nome || '').trim().slice(0, 80);
  if (!nome) throw new Error('Dê um nome ao fechamento.');
  const dados = JSON.parse(dadosJson); // valida
  if (!dados || !dados.porMes) throw new Error('Dados inválidos.');
  const texto = JSON.stringify(dados);
  const partes = [];
  for (let i = 0; i < texto.length; i += PEDACO) partes.push(texto.slice(i, i + PEDACO));
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    const sh = aba_(ABA_FECH, CAB);
    const id = Utilities.getUuid();
    const criadoEm = new Date().toISOString();
    const linha = [id, nome, criadoEm, String(arquivo || ''), emailAtual_(), partes.length].concat(partes);
    const lin = sh.getLastRow() + 1;
    sh.getRange(lin, 1, 1, linha.length).setNumberFormat('@').setValues([linha]);
    return {id: id, nome: nome, criadoEm: criadoEm, arquivo: String(arquivo || ''), autor: emailAtual_()};
  } finally { trava.releaseLock(); }
}

function excluirFechamento(id) {
  if (!ehEditor_()) throw new Error('SEM_PERMISSAO');
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    const sh = aba_(ABA_FECH, CAB);
    const lin = linhaDe_(sh, String(id));
    if (lin > 0) sh.deleteRow(lin);
    return true;
  } finally { trava.releaseLock(); }
}
