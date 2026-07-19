/**
 * CONTROLE DE EMPRESAS — automações (Google Apps Script)
 * Para a planilha CONTROLE_EMPRESAS_2026.
 *
 * Layout:
 *   CADASTRO: A=Empresa B=CNPJ C=Regime D=Faz Folha? E=Faz SPED? F=Ativa?
 *             G=Perfil H=IE I=IM J=Senha        (dados a partir da linha 3)
 *   Abas de etapa: A=MÊS B=EMPRESA C=CNPJ(fórmula) D=REGIME(fórmula)
 *             demais colunas = status            (dados a partir da linha 3)
 *   PAINEL: D4 = mês selecionado
 *
 * Tudo num arquivo só: as janelas (HTML) são geradas aqui dentro.
 */

// ============================================================
//  CONFIGURAÇÃO
// ============================================================

const LINHA_INICIAL = 3;
const LIMITE_LINHAS = 1500;   // teto das faixas (listas e cores vão até aqui na planilha)

// colFeito = coluna que sinaliza "concluído"  | colsMoeda = colunas em R$
// colTotal = coluna com fórmula de soma       | colData   = coluna de data
// flag     = quem entra nesta aba
const ABAS_ETAPA = {
  '1. FOLHA':       { flag: 'folha', ultimaCol: 9,  colFeito: 9 },
  '2. SPED':        { flag: 'sped',  ultimaCol: 11, colFeito: 11, colData: 9 },
  '3. FATURAMENTO': { flag: null,    ultimaCol: 10, colFeito: 9,
                      colsMoeda: [5, 6, 7, 8], colTotal: 8 },
  '4. CONSULTAS':   { flag: null,    ultimaCol: 6,  colFeito: 6 },
};

const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO',
               'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];

const REGIMES = ['Simples Nacional', 'Simples Híbrido', 'Lucro Presumido', 'MEI'];

const COR_AZUL = '#2E5496';
const COR_NAVY = '#1F3864';

// ============================================================
//  MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧮 Modo Contador')
    .addItem('Abrir novo mês', 'dialogoAbrirMes')
    .addItem('Cadastrar empresa nova', 'dialogoCadastro')
    .addSeparator()
    .addItem('Renomear empresa', 'dialogoRenomear')
    .addItem('Ativar / desativar empresa', 'dialogoAtivar')
    .addItem('Excluir empresa de vez', 'dialogoExcluir')
    .addItem('Excluir um mês inteiro', 'dialogoExcluirMes')
    .addSeparator()
    .addItem('Conferir faturamentos suspeitos', 'dialogoAnomalias')
    .addItem('Mostrar / atualizar calendário', 'mostrarCalendario')
    .addItem('Bloco de notas', 'abrirBlocoDeNotas')
    .addItem('Virar o ano', 'virarOAno')
    .addSeparator()
    .addItem('Como usar', 'dialogoSobre')
    .addToUi();

  // abre já no mês atual (você continua podendo trocar pelo dropdown)
  try { irParaMesAtual(); } catch (e) { /* onOpen em modo limitado às vezes não escreve */ }
}

// ============================================================
//  AUXILIARES
// ============================================================

function planilha() { return SpreadsheetApp.getActive(); }

/** Devolve quais das abas pedidas não existem.
 *  O nome da aba é contrato: se alguém renomeia, getSheetByName devolve null
 *  e a função pularia o trabalho sem avisar. Melhor recusar na entrada do que
 *  dizer que fez. */
function abasFaltando(nomes) {
  const ss = planilha();
  return nomes.filter(n => !ss.getSheetByName(n));
}

/** Abas de etapa + as pedidas a mais. */
function abasNecessarias(extras) {
  return Object.keys(ABAS_ETAPA).concat(extras || []);
}

function msgAbasFaltando(faltando) {
  return 'Não achei ' + (faltando.length === 1 ? 'esta aba:' : 'estas abas:') +
         '\n• ' + faltando.join('\n• ') +
         '\n\nAlguém renomeou ou apagou? O nome precisa bater exatamente,\n' +
         'incluindo o número e o ponto (ex.: "3. FATURAMENTO").\n' +
         'Corrija o nome da aba e tente de novo.';
}

/** Acha a última linha preenchida numa coluna.
 *  Usa getLastRow() como teto (em vez de escanear LIMITE_LINHAS inteiro):
 *  o Google já sabe onde a aba termina, então isso não lê célula nenhuma
 *  além do que existe de fato — bem mais rápido conforme a planilha cresce. */
function ultimaLinhaCol(aba, letra) {
  const teto = Math.min(aba.getLastRow(), LIMITE_LINHAS);
  if (teto < LINHA_INICIAL) return LINHA_INICIAL - 1;
  const valores = aba.getRange(letra + LINHA_INICIAL + ':' + letra + teto).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (valores[i][0] !== '' && valores[i][0] !== null) return LINHA_INICIAL + i;
  }
  return LINHA_INICIAL - 1;
}

/** Abas de etapa: a coluna EMPRESA (B) manda. */
function ultimaLinha(aba) { return ultimaLinhaCol(aba, 'B'); }

/** Lê o CADASTRO -> [{empresa, folha, sped, ativa, linha}, ...] */
function lerCadastro() {
  const aba = planilha().getSheetByName('CADASTRO');
  const teto = Math.min(aba.getLastRow(), LIMITE_LINHAS);
  if (teto < LINHA_INICIAL) return [];
  const dados = aba.getRange(LINHA_INICIAL, 1, teto - LINHA_INICIAL + 1, 6).getValues();
  const lista = [];
  dados.forEach((l, i) => {
    if (l[0] === '' || l[0] === null) return;
    const ativaTxt = String(l[5]).trim().toUpperCase();
    lista.push({
      empresa: l[0],
      folha: String(l[3]).trim().toUpperCase() === 'SIM',
      sped:  String(l[4]).trim().toUpperCase() === 'SIM',
      ativa: ativaTxt !== 'NÃO' && ativaTxt !== 'NAO',   // em branco = ativa
      linha: LINHA_INICIAL + i,
    });
  });
  return lista;
}

function acharEmpresa(nome) {
  const alvo = String(nome).trim().toUpperCase();
  const achadas = lerCadastro().filter(e =>
    String(e.empresa).trim().toUpperCase() === alvo);
  return achadas.length > 0 ? achadas[0] : null;
}

/** Empresas (+ se estão ativas) para os dropdowns dos comandos.
 *  Em ordem alfabética: o CADASTRO é por ordem de entrada, e achar numa
 *  lista de ~55 nomes fica muito mais rápido assim. */
function listarEmpresas() {
  return lerCadastro()
    .map(e => ({ nome: String(e.empresa), ativa: e.ativa }))
    .sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function entraNaEtapa(empresa, flag) {
  if (flag === null) return true;
  if (flag === 'folha') return empresa.folha;
  if (flag === 'sped') return empresa.sped;
  return true;
}

function mesDoPainel() {
  return String(planilha().getSheetByName('PAINEL').getRange('D4').getValue()).trim();
}

/** Meses que já existem numa aba, em ordem de calendário. */
function mesesDaAba(aba) {
  const fim = ultimaLinha(aba);
  if (fim < LINHA_INICIAL) return [];
  const achados = {};
  aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, 1).getValues()
     .forEach(l => {
       const m = String(l[0]).trim().toUpperCase();
       if (MESES.indexOf(m) !== -1) achados[m] = true;
     });
  return MESES.filter(m => achados[m]);
}

/** Meses já abertos na planilha (união das abas), em ordem de calendário.
 *  Usado pela janela de cadastro. */
function mesesAbertos() {
  const achados = {};
  for (const nomeAba in ABAS_ETAPA) {
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;
    mesesDaAba(aba).forEach(m => { achados[m] = true; });
  }
  return MESES.filter(m => achados[m]);
}

/** Formata linhas novas de uma aba de etapa.
 *  Explícito de propósito: funciona com a planilha vazia e não herda
 *  o estilo da linha vizinha. */
function formatarEtapa(aba, nomeAba, linhaInicio, qtd) {
  const cfg = ABAS_ETAPA[nomeAba];
  aba.getRange(linhaInicio, 1, qtd, cfg.ultimaCol)
     .setFontFamily('Arial').setFontSize(10).setFontColor('#000000')
     .setBackground('#ffffff').setFontWeight('normal')
     .setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBorder(true, true, true, true, true, true,
                '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

  aba.getRange(linhaInicio, 1, qtd, 1).setFontWeight('bold');                  // MÊS
  aba.getRange(linhaInicio, 2, qtd, 1).setFontWeight('bold')
     .setHorizontalAlignment('left');                                          // EMPRESA

  (cfg.colsMoeda || []).forEach(c =>
    aba.getRange(linhaInicio, c, qtd, 1).setNumberFormat('R$ #,##0.00'));
  if (cfg.colTotal) aba.getRange(linhaInicio, cfg.colTotal, qtd, 1).setFontWeight('bold');
  if (cfg.colData)  aba.getRange(linhaInicio, cfg.colData, qtd, 1).setNumberFormat('dd/mm/yyyy');
}

function formatarComparativo(aba, linha) {
  aba.getRange(linha, 2, 1, 14)
     .setFontFamily('Arial').setFontSize(10).setFontColor('#000000')
     .setBackground('#ffffff').setFontWeight('normal')
     .setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBorder(true, true, true, true, true, true,
                '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  aba.getRange(linha, 2).setFontWeight('bold').setHorizontalAlignment('left');
  aba.getRange(linha, 3, 1, 13).setNumberFormat('R$ #,##0.00');
  aba.getRange(linha, 15).setFontWeight('bold');   // TOTAL ANO
}

/** Refaz o filtro cobrindo todas as linhas de dados.
 *  Sem isso, meses novos não aparecem na lista do funil. */
function ajustarFiltro(aba, ultimaCol) {
  const fim = ultimaLinha(aba);
  const filtro = aba.getFilter();
  if (filtro) filtro.remove();
  if (fim < LINHA_INICIAL) return;
  aba.getRange(2, 1, fim - 1, ultimaCol).createFilter();
}

/** Insere a empresa na tabela do COMPARATIVO.
 *  O TOTAL GERAL puxa direto do FATURAMENTO, então não precisa ser refeito. */
function adicionarNoComparativo(nome) {
  const aba = planilha().getSheetByName('COMPARATIVO');
  if (!aba) return false;

  const colB = aba.getRange('B1:B' + aba.getLastRow()).getValues()
                  .map(l => String(l[0]).trim());
  const linhaCabecalho = colB.indexOf('EMPRESA') + 1;
  const linhaTotal     = colB.indexOf('TOTAL GERAL') + 1;
  if (linhaCabecalho === 0 || linhaTotal === 0) return false;

  for (let r = linhaCabecalho + 1; r < linhaTotal; r++) {
    if (colB[r - 1].toUpperCase() === nome.toUpperCase()) return true;   // já existe
  }

  aba.insertRowBefore(linhaTotal);
  const nova = linhaTotal;

  formatarComparativo(aba, nova);
  aba.getRange(nova, 2).setValue(nome);

  const ini = LINHA_INICIAL, fim = LIMITE_LINHAS;
  const formulas = [];
  for (let m = 0; m < MESES.length; m++) {
    formulas.push(
      "=IFERROR(SUMIFS('3. FATURAMENTO'!$H$" + ini + ":$H$" + fim +
      ";'3. FATURAMENTO'!$B$" + ini + ":$B$" + fim + ";$B" + nova +
      ";'3. FATURAMENTO'!$A$" + ini + ":$A$" + fim + ";\"" + MESES[m] + "\");0)");
  }
  formulas.push('=SUM(C' + nova + ':N' + nova + ')');
  aba.getRange(nova, 3, 1, 13).setFormulas([formulas]);
  return true;
}

// ============================================================
//  ESTILO DAS JANELAS
// ============================================================

function estiloDialogo() {
  return '<style>' +
  '* { box-sizing: border-box; }' +
  'body { font-family: Arial, sans-serif; margin: 0; padding: 18px 20px;' +
  '  color: #262626; font-size: 13px; background: #fff; }' +
  'h3 { margin: 0 0 4px 0; color: ' + COR_NAVY + '; font-size: 15px; }' +
  '.sub { color: #808080; font-size: 11.5px; margin-bottom: 14px; }' +
  'select:disabled { background: #F2F2F2; color: #808080; }' +
  'label { display: block; font-weight: bold; color: ' + COR_NAVY + ';' +
  '  margin: 10px 0 4px 0; font-size: 12px; }' +
  'input[type=text], select { width: 100%; padding: 8px 9px; border: 1px solid #C9D2E3;' +
  '  border-radius: 4px; font-size: 13px; font-family: Arial, sans-serif; background: #fff; }' +
  'input[type=text]:focus, select:focus { outline: none; border-color: ' + COR_AZUL + ';' +
  '  box-shadow: 0 0 0 2px rgba(46,84,150,.15); }' +
  '.linha { display: flex; gap: 12px; }' +
  '.linha > div { flex: 1; }' +
  '.toggle { display: flex; gap: 6px; }' +
  '.toggle button { flex: 1; padding: 7px 0; border: 1px solid #C9D2E3; background: #fff;' +
  '  border-radius: 4px; cursor: pointer; font-size: 12px; font-family: Arial, sans-serif; }' +
  '.toggle button.on { background: #C6EFCE; border-color: #7FBF8F; color: #006100; font-weight: bold; }' +
  '.toggle button.off { background: #F2F2F2; border-color: #D0D0D0; color: #808080; font-weight: bold; }' +
  '.aviso { background: #FFF8E7; border-left: 3px solid #BF8F00; color: #5C4500;' +
  '  padding: 8px 10px; font-size: 11.5px; margin: 14px 0 4px 0; border-radius: 3px; }' +
  '.botoes { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }' +
  'button.pri { background: ' + COR_AZUL + '; color: #fff; border: none; padding: 9px 18px;' +
  '  border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold;' +
  '  font-family: Arial, sans-serif; }' +
  'button.pri:hover { background: ' + COR_NAVY + '; }' +
  'button.pri:disabled { background: #A9B7CE; cursor: default; }' +
  'button.sec { background: #fff; color: #595959; border: 1px solid #C9D2E3; padding: 9px 16px;' +
  '  border-radius: 4px; cursor: pointer; font-size: 13px; font-family: Arial, sans-serif; }' +
  'button.perigo { background: #9C0006; color: #fff; border: none; padding: 9px 18px;' +
  '  border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold;' +
  '  font-family: Arial, sans-serif; }' +
  'button.perigo:disabled { background: #E0B4B7; cursor: default; }' +
  '.alerta { background: #FFC7CE; border-left: 3px solid #9C0006; color: #9C0006;' +
  '  padding: 9px 11px; font-size: 12px; margin: 12px 0; border-radius: 3px; white-space: pre-line; }' +
  '.caixa { background: #F7F9FC; border: 1px solid #D9E1F2; border-radius: 4px;' +
  '  padding: 10px 12px; font-size: 12px; margin: 12px 0; white-space: pre-line; }' +
  '#status { margin-top: 12px; font-size: 12px; padding: 9px 11px; border-radius: 4px;' +
  '  display: none; white-space: pre-line; }' +
  '#status.ok { display: block; background: #C6EFCE; color: #006100; }' +
  '#status.erro { display: block; background: #FFC7CE; color: #9C0006; }' +
  '#status.load { display: block; background: #EEF1F7; color: #2E5496; }' +
  '</style>';
}

// ============================================================
//  1. ABRIR NOVO MÊS  (janela com lista de meses)
// ============================================================

function dialogoAbrirMes() {
  const html = HtmlService.createHtmlOutput(htmlAbrirMes())
                          .setWidth(430).setHeight(310);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrir novo mês');
}

function htmlAbrirMes() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">As empresas ativas entram automaticamente nas abas certas.</div>' +
  '<label>Mês</label>' +
  '<select id="mes"></select>' +
  '<div class="aviso">Os meses anteriores continuam guardados. Se o mês já existir, ele é pulado.</div>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Cancelar</button>' +
  '  <button class="pri" id="ok" onclick="abrir()">Abrir mês</button>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var MESES = ' + JSON.stringify(MESES) + ';' +
  'var sel = document.getElementById("mes");' +
  'for (var i = 0; i < MESES.length; i++) {' +
  '  var o = document.createElement("option");' +
  '  o.value = MESES[i]; o.text = MESES[i]; sel.add(o);' +
  '}' +
  'function msg(txt, tipo) {' +
  '  var s = document.getElementById("status");' +
  '  s.className = tipo; s.textContent = txt;' +
  '}' +
  'function abrir() {' +
  '  document.getElementById("ok").disabled = true;' +
  '  msg("Abrindo o mês, aguarde...", "load");' +
  '  google.script.run' +
  '    .withSuccessHandler(function (r) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg(r.msg, r.ok ? "ok" : "erro");' +
  '    })' +
  '    .withFailureHandler(function (e) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg("Erro: " + e.message, "erro");' +
  '    })' +
  '    .executarAbrirMes(sel.value);' +
  '}' +
  '</script></body></html>';
}

/** Chamada pela janela. Devolve {ok, msg} em vez de abrir alerta. */
function executarAbrirMes(mes) {
  mes = String(mes).trim().toUpperCase();
  if (MESES.indexOf(mes) === -1) return { ok: false, msg: 'Mês inválido.' };

  const faltando = abasFaltando(abasNecessarias(['CADASTRO', 'PAINEL']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const empresas = lerCadastro().filter(e => e.ativa);
  if (empresas.length === 0) {
    return { ok: false, msg: 'Nenhuma empresa ativa no CADASTRO.\nCadastre primeiro.' };
  }

  const resumo = [];

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);

    const fim = ultimaLinha(aba);

    const jaTem = fim >= LINHA_INICIAL
      ? aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, 1)
           .getValues().map(l => String(l[0]).trim()).indexOf(mes) !== -1
      : false;
    if (jaTem) { resumo.push('• ' + nomeAba + ': já existia, pulei'); continue; }

    const doMes = empresas.filter(e => entraNaEtapa(e, cfg.flag));
    if (doMes.length === 0) { resumo.push('• ' + nomeAba + ': nenhuma empresa'); continue; }

    const linha = fim + 1;
    if (linha + doMes.length > LIMITE_LINHAS) {
      return { ok: false, msg: 'A aba ' + nomeAba + ' passou do limite de ' +
                               LIMITE_LINHAS + ' linhas.' };
    }

    formatarEtapa(aba, nomeAba, linha, doMes.length);
    aba.getRange(linha, 1, doMes.length, 2).setValues(doMes.map(e => [mes, e.empresa]));

    const fCnpj = [], fRegime = [];
    for (let i = 0; i < doMes.length; i++) {
      const r = linha + i;
      fCnpj.push(['=IFERROR(INDEX(CADASTRO!$B:$B;MATCH($B' + r + ';CADASTRO!$A:$A;0));"")']);
      fRegime.push(['=IFERROR(INDEX(CADASTRO!$C:$C;MATCH($B' + r + ';CADASTRO!$A:$A;0));"")']);
    }
    aba.getRange(linha, 3, doMes.length, 1).setFormulas(fCnpj);
    aba.getRange(linha, 4, doMes.length, 1).setFormulas(fRegime);

    if (cfg.colTotal) {
      const fTotal = [];
      for (let i = 0; i < doMes.length; i++) {
        const r = linha + i;
        fTotal.push(['=IF(COUNT(E' + r + ':G' + r + ')=0;"";SUM(E' + r + ':G' + r + '))']);
      }
      aba.getRange(linha, cfg.colTotal, doMes.length, 1).setFormulas(fTotal);
    }

    ajustarFiltro(aba, cfg.ultimaCol);
    resumo.push('• ' + nomeAba + ': ' + doMes.length + ' empresas');
  }

  planilha().getSheetByName('PAINEL').getRange('D4').setValue(mes);
  try { atualizarCalendario(); } catch (e) { /* calendário nunca trava o mês */ }
  return { ok: true, msg: mes + ' aberto!\n' + resumo.join('\n') };
}

// ============================================================
//  2. CADASTRAR EMPRESA  (formulário único)
// ============================================================

function dialogoCadastro() {
  const html = HtmlService.createHtmlOutput(htmlCadastro())
                          .setWidth(470).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cadastrar empresa');
}

function htmlCadastro() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Dados da empresa e em quais etapas ela entra.</div>' +
  '<label>Nome da empresa</label>' +
  '<input type="text" id="nome" autofocus>' +
  '<div class="linha">' +
  '  <div><label>CNPJ</label><input type="text" id="cnpj"></div>' +
  '  <div><label>Regime</label><select id="regime"></select></div>' +
  '</div>' +
  '<div id="blocoMes" style="display:none">' +
  '  <label>Entra a partir de</label>' +
  '  <select id="mesIni"></select>' +
  '</div>' +
  '<div class="linha">' +
  '  <div><label>Faz a FOLHA?</label><div class="toggle">' +
  '    <button id="folhaSim" onclick="setFolha(true)">Sim</button>' +
  '    <button id="folhaNao" onclick="setFolha(false)">Não</button></div></div>' +
  '  <div><label>Faz o SPED?</label><div class="toggle">' +
  '    <button id="spedSim" onclick="setSped(true)">Sim</button>' +
  '    <button id="spedNao" onclick="setSped(false)">Não</button></div></div>' +
  '</div>' +
  '<div class="aviso" id="dica">Ela entra nesse mês e em todos os meses já abertos depois dele. ' +
  'Senha, perfil e inscrições você preenche direto no CADASTRO.</div>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Fechar</button>' +
  '  <button class="pri" id="ok" onclick="salvar()">Cadastrar</button>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var REGIMES = ' + JSON.stringify(REGIMES) + ';' +
  'var selReg = document.getElementById("regime");' +
  'for (var i = 0; i < REGIMES.length; i++) {' +
  '  var o = document.createElement("option");' +
  '  o.value = REGIMES[i]; o.text = REGIMES[i]; selReg.add(o);' +
  '}' +
  'var fazFolha = false, fazSped = true;' +
  'function pinta() {' +
  '  document.getElementById("folhaSim").className = fazFolha ? "on" : "";' +
  '  document.getElementById("folhaNao").className = fazFolha ? "" : "off";' +
  '  document.getElementById("spedSim").className  = fazSped  ? "on" : "";' +
  '  document.getElementById("spedNao").className  = fazSped  ? "" : "off";' +
  '}' +
  'function setFolha(v) { fazFolha = v; pinta(); }' +
  'function setSped(v)  { fazSped  = v; pinta(); }' +
  'pinta();' +
  'var selMes = document.getElementById("mesIni");' +
  'google.script.run.withSuccessHandler(function (meses) {' +
  '  if (meses.length === 0) {' +
  '    document.getElementById("blocoMes").style.display = "none";' +
  '    document.getElementById("dica").textContent =' +
  '      "Nenhum mês foi aberto ainda: a empresa fica só no cadastro e entra quando' +
  ' você usar Abrir novo mês. É assim mesmo quando se começa do zero.";' +
  '    return;' +
  '  }' +
  '  selMes.innerHTML = "";' +
  '  var o0 = document.createElement("option");' +
  '  o0.value = ""; o0.text = "— Não adicionar a nenhum mês —";' +
  '  selMes.add(o0);' +
  '  for (var i = 0; i < meses.length; i++) {' +
  '    var o = document.createElement("option");' +
  '    o.value = meses[i]; o.text = meses[i]; selMes.add(o);' +
  '  }' +
  '  selMes.selectedIndex = meses.length;' +   // padrão: o mês mais recente aberto
  '  document.getElementById("blocoMes").style.display = "block";' +
  '}).mesesAbertos();' +
  'function msg(txt, tipo) {' +
  '  var s = document.getElementById("status");' +
  '  s.className = tipo; s.textContent = txt;' +
  '}' +
  'function salvar() {' +
  '  var nome = document.getElementById("nome").value.trim();' +
  '  if (!nome) { msg("Digite o nome da empresa.", "erro"); return; }' +
  '  document.getElementById("ok").disabled = true;' +
  '  msg("Cadastrando...", "load");' +
  '  google.script.run' +
  '    .withSuccessHandler(function (r) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg(r.msg, r.ok ? "ok" : "erro");' +
  '      if (r.ok) {' +
  '        document.getElementById("nome").value = "";' +
  '        document.getElementById("cnpj").value = "";' +
  '        document.getElementById("nome").focus();' +
  '      }' +
  '    })' +
  '    .withFailureHandler(function (e) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg("Erro: " + e.message, "erro");' +
  '    })' +
  '    .executarCadastro({' +
  '      nome: nome,' +
  '      cnpj: document.getElementById("cnpj").value.trim(),' +
  '      regime: selReg.value,' +
  '      folha: fazFolha ? "Sim" : "Não",' +
  '      sped: fazSped ? "Sim" : "Não",' +
  '      mesInicial: selMes.value' +
  '    });' +
  '}' +
  'document.getElementById("nome").addEventListener("keydown", function (e) {' +
  '  if (e.key === "Enter") salvar();' +
  '});' +
  '</script></body></html>';
}

/** Chamada pela janela. Devolve {ok, msg}. */
function executarCadastro(d) {
  const nome = String(d.nome).trim();
  if (!nome) return { ok: false, msg: 'Digite o nome da empresa.' };

  const faltando = abasFaltando(abasNecessarias(['CADASTRO', 'COMPARATIVO']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  if (acharEmpresa(nome)) return { ok: false, msg: '"' + nome + '" já está no cadastro.' };
  if (REGIMES.indexOf(d.regime) === -1) return { ok: false, msg: 'Regime inválido.' };

  // CADASTRO
  const cad = planilha().getSheetByName('CADASTRO');
  const linhaCad = ultimaLinhaCol(cad, 'A') + 1;
  cad.getRange(linhaCad, 1, 1, 6)
     .setValues([[nome, d.cnpj, d.regime, d.folha, d.sped, 'Sim']]);
  cad.getRange(linhaCad, 1, 1, 10)
     .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
     .setHorizontalAlignment('center')
     .setBorder(true, true, true, true, true, true,
                '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  cad.getRange(linhaCad, 1).setFontWeight('bold').setHorizontalAlignment('left');

  // Em quais meses ela entra: do mês escolhido em diante, só os já abertos
  const iIni = d.mesInicial ? MESES.indexOf(String(d.mesInicial).trim().toUpperCase()) : -1;
  const empresa = { empresa: nome, folha: d.folha === 'Sim', sped: d.sped === 'Sim' };
  const inseridas = [];
  const mesesUsados = {};

  if (iIni >= 0) {
    for (const nomeAba in ABAS_ETAPA) {
      const cfg = ABAS_ETAPA[nomeAba];
      if (!entraNaEtapa(empresa, cfg.flag)) continue;

      const aba = planilha().getSheetByName(nomeAba);

      const alvos = mesesDaAba(aba).filter(m => MESES.indexOf(m) >= iIni);
      if (alvos.length === 0) continue;

      const linha = ultimaLinha(aba) + 1;
      if (linha + alvos.length > LIMITE_LINHAS) {
        return { ok: false, msg: 'A aba ' + nomeAba + ' passou do limite de ' +
                                 LIMITE_LINHAS + ' linhas.' };
      }

      formatarEtapa(aba, nomeAba, linha, alvos.length);
      aba.getRange(linha, 1, alvos.length, 2).setValues(alvos.map(m => [m, nome]));

      const fCnpj = [], fRegime = [], fTotal = [];
      for (let i = 0; i < alvos.length; i++) {
        const r = linha + i;
        fCnpj.push(['=IFERROR(INDEX(CADASTRO!$B:$B;MATCH($B' + r + ';CADASTRO!$A:$A;0));"")']);
        fRegime.push(['=IFERROR(INDEX(CADASTRO!$C:$C;MATCH($B' + r + ';CADASTRO!$A:$A;0));"")']);
        fTotal.push(['=IF(COUNT(E' + r + ':G' + r + ')=0;"";SUM(E' + r + ':G' + r + '))']);
      }
      aba.getRange(linha, 3, alvos.length, 1).setFormulas(fCnpj);
      aba.getRange(linha, 4, alvos.length, 1).setFormulas(fRegime);
      if (cfg.colTotal) aba.getRange(linha, cfg.colTotal, alvos.length, 1).setFormulas(fTotal);

      ajustarFiltro(aba, cfg.ultimaCol);
      alvos.forEach(m => { mesesUsados[m] = true; });
      inseridas.push(nomeAba.replace(/^\d\.\s*/, '') + ' (' + alvos.length + ')');
    }
  }

  // A aba existe (checado na entrada), então um "false" aqui só pode ser
  // layout fora do esperado — vale avisar em vez de omitir da lista.
  const noComparativo = adicionarNoComparativo(nome);
  if (noComparativo) inseridas.push('COMPARATIVO');

  const meses = MESES.filter(m => mesesUsados[m]);
  const onde = meses.length === 0
    ? 'Só no cadastro (nenhum mês)'
    : (meses.length === 1 ? 'Mês: ' + meses[0]
                          : 'Meses: ' + meses[0] + ' a ' + meses[meses.length - 1] +
                            ' (' + meses.length + ')');

  const total = lerCadastro().length;
  const aviso = noComparativo ? '' :
    '\n\n⚠ NÃO entrou no COMPARATIVO: não achei as linhas "EMPRESA" e\n' +
    '"TOTAL GERAL" na coluna B. O faturamento dela não vai somar lá.';

  return { ok: true, msg: '✔ ' + nome + ' cadastrada (' + d.regime + ')' +
                          '\n' + onde +
                          '\nAbas: ' + inseridas.join(', ') +
                          '\nTotal no cadastro: ' + total + aviso };
}

// ============================================================
//  3. RENOMEAR EMPRESA
// ============================================================

function dialogoRenomear() {
  const html = HtmlService.createHtmlOutput(htmlRenomear())
                          .setWidth(430).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'Renomear empresa');
}

function htmlRenomear() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Troca o nome em tudo de uma vez: cadastro, as 4 abas e o comparativo.</div>' +
  '<label>Empresa</label>' +
  '<select id="emp"><option value="">carregando...</option></select>' +
  '<label>Novo nome</label>' +
  '<input type="text" id="novo" autocomplete="off">' +
  '<div class="aviso">O nome é a chave que liga tudo. A troca vale para todos os meses já lançados.</div>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Cancelar</button>' +
  '  <button class="pri" id="ok" onclick="salvar()">Renomear</button>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var sel = document.getElementById("emp");' +
  'function preenche(lista, escolher) {' +
  '  sel.innerHTML = "";' +
  '  if (!lista.length) {' +
  '    var v = document.createElement("option"); v.value = ""; v.text = "— Nenhuma empresa —"; sel.add(v);' +
  '    document.getElementById("ok").disabled = true; return;' +
  '  }' +
  '  for (var i = 0; i < lista.length; i++) {' +
  '    var o = document.createElement("option"); o.value = lista[i].nome; o.text = lista[i].nome; sel.add(o);' +
  '  }' +
  '  if (escolher) sel.value = escolher;' +
  '}' +
  'google.script.run.withSuccessHandler(function (lista) { preenche(lista); }).listarEmpresas();' +
  'function msg(txt, tipo) { var s = document.getElementById("status"); s.className = tipo; s.textContent = txt; }' +
  'function salvar() {' +
  '  var velho = sel.value;' +
  '  var novo = document.getElementById("novo").value.trim();' +
  '  if (!velho) { msg("Escolha a empresa.", "erro"); return; }' +
  '  if (!novo)  { msg("Digite o novo nome.", "erro"); return; }' +
  '  document.getElementById("ok").disabled = true;' +
  '  msg("Renomeando...", "load");' +
  '  google.script.run' +
  '    .withSuccessHandler(function (r) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg(r.msg, r.ok ? "ok" : "erro");' +
  '      if (r.ok) {' +
  '        document.getElementById("novo").value = "";' +
  '        google.script.run.withSuccessHandler(function (lista) { preenche(lista, novo); }).listarEmpresas();' +
  '      }' +
  '    })' +
  '    .withFailureHandler(function (e) {' +
  '      document.getElementById("ok").disabled = false;' +
  '      msg("Erro: " + e.message, "erro");' +
  '    })' +
  '    .executarRenomear(velho, novo);' +
  '}' +
  'document.getElementById("novo").addEventListener("keydown", function (e) {' +
  '  if (e.key === "Enter") salvar();' +
  '});' +
  '</script></body></html>';
}

/** Chamada pela janela. O nome é a chave que liga tudo: troca no CADASTRO,
 *  nas 4 abas (todos os meses) e no COMPARATIVO de uma vez. Devolve {ok, msg}.
 *  Renomear em 3 das 4 abas é pior que não renomear: as linhas da aba que
 *  faltou viram órfãs e o faturamento delas some do COMPARATIVO. */
function executarRenomear(velho, novo) {
  velho = String(velho).trim();
  novo  = String(novo).trim();
  if (!velho) return { ok: false, msg: 'Escolha a empresa.' };
  if (!novo)  return { ok: false, msg: 'Digite o novo nome.' };

  const faltando = abasFaltando(abasNecessarias(['CADASTRO', 'COMPARATIVO']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const emp = acharEmpresa(velho);
  if (!emp) return { ok: false, msg: 'Não achei "' + velho + '" no CADASTRO.' };

  // Só barra se o novo nome for de OUTRA empresa. Assim continua dando para
  // corrigir só a grafia (maiúsc/minúsc/acento) da própria empresa.
  const outra = acharEmpresa(novo);
  if (outra && outra.linha !== emp.linha) {
    return { ok: false, msg: 'Já existe uma empresa chamada "' + novo + '".' };
  }

  const alvo = velho.toUpperCase();
  const trocas = [];

  planilha().getSheetByName('CADASTRO').getRange(emp.linha, 1).setValue(novo);
  trocas.push('CADASTRO: 1');

  for (const nomeAba in ABAS_ETAPA) {
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;
    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;

    const faixa = aba.getRange(LINHA_INICIAL, 2, fim - LINHA_INICIAL + 1, 1);
    const valores = faixa.getValues();
    let qtd = 0;
    for (let i = 0; i < valores.length; i++) {
      if (String(valores[i][0]).trim().toUpperCase() === alvo) { valores[i][0] = novo; qtd++; }
    }
    if (qtd > 0) { faixa.setValues(valores); trocas.push(nomeAba + ': ' + qtd); }
  }

  const comp = planilha().getSheetByName('COMPARATIVO');
  if (comp) {
    const colB = comp.getRange('B1:B' + comp.getLastRow()).getValues();
    for (let i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim().toUpperCase() === alvo) {
        comp.getRange(i + 1, 2).setValue(novo);
        trocas.push('COMPARATIVO: 1');
        break;
      }
    }
  }

  return { ok: true, msg: '✔ Renomeada!\n"' + velho + '"  ->  "' + novo +
                          '"\n\nLinhas trocadas:\n' + trocas.join('\n') };
}

// ============================================================
//  4. ATIVAR / DESATIVAR
// ============================================================

function dialogoAtivar() {
  const html = HtmlService.createHtmlOutput(htmlAtivar())
                          .setWidth(430).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'Ativar / desativar empresa');
}

function htmlAtivar() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Tira (ou devolve) a empresa dos meses novos, sem apagar o histórico.</div>' +
  '<label>Empresa</label>' +
  '<select id="emp"><option value="">carregando...</option></select>' +
  '<div id="cartao" class="caixa" style="display:none"></div>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Cancelar</button>' +
  '  <button class="pri" id="ok" onclick="alternar()" disabled>Selecione</button>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var porNome = {};' +
  'var sel = document.getElementById("emp");' +
  'var btn = document.getElementById("ok");' +
  'function rotulo(e) { return e.nome + (e.ativa ? "" : "  (inativa)"); }' +
  'google.script.run.withSuccessHandler(function (dados) {' +
  '  sel.innerHTML = "";' +
  '  if (!dados.length) {' +
  '    var v = document.createElement("option"); v.value = ""; v.text = "— Nenhuma empresa —"; sel.add(v); return;' +
  '  }' +
  '  var o0 = document.createElement("option"); o0.value = ""; o0.text = "— escolha —"; sel.add(o0);' +
  '  for (var i = 0; i < dados.length; i++) {' +
  '    porNome[dados[i].nome] = dados[i];' +
  '    var o = document.createElement("option"); o.value = dados[i].nome; o.text = rotulo(dados[i]); sel.add(o);' +
  '  }' +
  '}).listarEmpresas();' +
  'sel.addEventListener("change", pinta);' +
  'function pinta() {' +
  '  var e = porNome[sel.value];' +
  '  var cartao = document.getElementById("cartao");' +
  '  var s = document.getElementById("status"); s.className = ""; s.textContent = "";' +
  '  if (!e) { cartao.style.display = "none"; btn.disabled = true; btn.textContent = "Selecione"; btn.className = "pri"; return; }' +
  '  cartao.style.display = "block";' +
  '  if (e.ativa) {' +
  '    cartao.textContent = e.nome + " está ATIVA.\\nDesativar: ela deixa de entrar nos meses novos. O histórico continua guardado.";' +
  '    btn.textContent = "Desativar"; btn.className = "perigo";' +
  '  } else {' +
  '    cartao.textContent = e.nome + " está INATIVA.\\nReativar: ela volta a entrar nos próximos meses.";' +
  '    btn.textContent = "Reativar"; btn.className = "pri";' +
  '  }' +
  '  btn.disabled = false;' +
  '}' +
  'function msg(txt, tipo) { var s = document.getElementById("status"); s.className = tipo; s.textContent = txt; }' +
  'function alternar() {' +
  '  var nome = sel.value;' +
  '  if (!nome) { msg("Escolha a empresa.", "erro"); return; }' +
  '  btn.disabled = true; msg("Aplicando...", "load");' +
  '  google.script.run' +
  '    .withSuccessHandler(function (r) {' +
  '      if (r.ok) {' +
  '        porNome[nome].ativa = r.ativa;' +
  '        for (var i = 0; i < sel.options.length; i++) {' +
  '          if (sel.options[i].value === nome) sel.options[i].text = rotulo(porNome[nome]);' +
  '        }' +
  '        pinta(); msg(r.msg, "ok");' +
  '      } else { btn.disabled = false; msg(r.msg, "erro"); }' +
  '    })' +
  '    .withFailureHandler(function (e) { btn.disabled = false; msg("Erro: " + e.message, "erro"); })' +
  '    .executarAtivar(nome);' +
  '}' +
  '</script></body></html>';
}

/** Chamada pela janela. Alterna ativa/inativa no CADASTRO. Devolve {ok, msg, ativa}. */
function executarAtivar(nome) {
  nome = String(nome).trim();
  if (!nome) return { ok: false, msg: 'Escolha a empresa.' };

  const faltando = abasFaltando(['CADASTRO']);
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const emp = acharEmpresa(nome);
  if (!emp) return { ok: false, msg: 'Não achei "' + nome + '" no CADASTRO.' };

  const novoAtiva = !emp.ativa;
  planilha().getSheetByName('CADASTRO')
            .getRange(emp.linha, 6).setValue(novoAtiva ? 'Sim' : 'Não');

  return { ok: true, ativa: novoAtiva,
           msg: '✔ ' + emp.empresa + ' agora está ' + (novoAtiva ? 'ATIVA' : 'INATIVA') + '.' };
}

// ============================================================
//  5. EXCLUIR DE VEZ
// ============================================================

function dialogoExcluir() {
  const html = HtmlService.createHtmlOutput(htmlExcluir())
                          .setWidth(460).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, 'Excluir empresa de vez');
}

function htmlExcluir() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Apaga o cadastro e TODO o histórico da empresa. Use só para cadastro errado.</div>' +
  '<label>Empresa</label>' +
  '<select id="emp"><option value="">carregando...</option></select>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Cancelar</button>' +
  '  <button class="pri" id="btnVer" onclick="verificar()" disabled>Ver o que será apagado</button>' +
  '</div>' +
  '<div id="info"></div>' +
  '<div id="bloco" style="display:none">' +
  '  <label id="lblConf" style="display:none">Digite o nome para confirmar</label>' +
  '  <input type="text" id="conf" style="display:none" autocomplete="off">' +
  '  <div class="botoes">' +
  '    <button class="perigo" id="btnDel" onclick="apagar()">Apagar de vez</button>' +
  '  </div>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var sel = document.getElementById("emp");' +
  'var empVerificada = null, exigeTexto = false;' +
  'function carrega(escolher) {' +
  '  google.script.run.withSuccessHandler(function (lista) {' +
  '    sel.innerHTML = "";' +
  '    if (!lista.length) {' +
  '      var v = document.createElement("option"); v.value = ""; v.text = "— Nenhuma empresa —"; sel.add(v);' +
  '      document.getElementById("btnVer").disabled = true; return;' +
  '    }' +
  '    var o0 = document.createElement("option"); o0.value = ""; o0.text = "— escolha —"; sel.add(o0);' +
  '    for (var i = 0; i < lista.length; i++) { var o = document.createElement("option"); o.value = lista[i].nome; o.text = lista[i].nome; sel.add(o); }' +
  '    if (escolher) sel.value = escolher;' +
  '    document.getElementById("btnVer").disabled = !sel.value;' +
  '  }).listarEmpresas();' +
  '}' +
  'carrega();' +
  'sel.addEventListener("change", function () {' +
  '  document.getElementById("btnVer").disabled = !sel.value;' +
  '  document.getElementById("bloco").style.display = "none";' +
  '  document.getElementById("info").innerHTML = "";' +
  '  var s = document.getElementById("status"); s.className = ""; s.textContent = "";' +
  '});' +
  'function verificar() {' +
  '  document.getElementById("btnVer").disabled = true;' +
  '  google.script.run.withSuccessHandler(function (r) {' +
  '    document.getElementById("btnVer").disabled = false;' +
  '    if (!r.existe) { document.getElementById("info").innerHTML = \'<div class="alerta">\' + r.resumo + "</div>"; return; }' +
  '    empVerificada = r.nome; exigeTexto = r.temHistorico;' +
  '    var cls = r.temHistorico ? "alerta" : "caixa";' +
  '    document.getElementById("info").innerHTML = \'<div class="\' + cls + \'">\' + r.resumo + "</div>";' +
  '    document.getElementById("bloco").style.display = "block";' +
  '    var lbl = document.getElementById("lblConf"), campo = document.getElementById("conf");' +
  '    if (r.temHistorico) {' +
  '      lbl.style.display = "block"; campo.style.display = "block";' +
  '      lbl.textContent = "Digite " + r.nome + " para confirmar";' +
  '      campo.value = ""; campo.focus();' +
  '    } else { lbl.style.display = "none"; campo.style.display = "none"; }' +
  '  }).analisarEmpresa(sel.value);' +
  '}' +
  'function apagar() {' +
  '  var txt = exigeTexto ? document.getElementById("conf").value.trim() : empVerificada;' +
  '  document.getElementById("btnDel").disabled = true;' +
  '  var s = document.getElementById("status"); s.className = "load"; s.textContent = "Apagando...";' +
  '  google.script.run.withSuccessHandler(function (r) {' +
  '    document.getElementById("btnDel").disabled = false;' +
  '    s.className = r.ok ? "ok" : "erro"; s.textContent = r.msg;' +
  '    if (r.ok) {' +
  '      document.getElementById("bloco").style.display = "none";' +
  '      document.getElementById("info").innerHTML = "";' +
  '      carrega();' +
  '    }' +
  '  }).executarExcluirEmpresa(empVerificada, txt);' +
  '}' +
  '</script></body></html>';
}

/** Mostra quanto histórico a empresa tem antes de apagar. Devolve {existe, ...}. */
function analisarEmpresa(nome) {
  nome = String(nome).trim();
  const emp = acharEmpresa(nome);
  if (!emp) return { existe: false, resumo: 'Não achei "' + nome + '" no CADASTRO.' };

  const alvo = nome.toUpperCase();
  let total = 0; const detalhe = [];
  for (const nomeAba in ABAS_ETAPA) {
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;
    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;
    const valores = aba.getRange(LINHA_INICIAL, 2, fim - LINHA_INICIAL + 1, 1).getValues();
    const qtd = valores.filter(l => String(l[0]).trim().toUpperCase() === alvo).length;
    if (qtd > 0) { total += qtd; detalhe.push('• ' + nomeAba + ': ' + qtd + ' linha(s)'); }
  }

  const temHistorico = total > 0;
  const estado = emp.ativa ? 'ATIVA' : 'inativa';
  const resumo = emp.empresa + ' (' + estado + ')\n\n' +
    (temHistorico
      ? 'Isso apaga o cadastro e ' + total + ' linha(s) de histórico:\n' + detalhe.join('\n') +
        '\n\nISSO NÃO TEM VOLTA. Se a empresa só saiu do escritório, cancele e use\n' +
        '"Ativar / desativar" — assim o histórico é preservado.'
      : 'Ela não tem histórico em nenhum mês. Só sairá do cadastro (e do comparativo).');

  return { existe: true, nome: emp.empresa, temHistorico: temHistorico, total: total, resumo: resumo };
}

/** Apaga a empresa e TODO o histórico dela. Se houver histórico, exige o nome
 *  digitado. Só para cadastro errado. Devolve {ok, msg}.
 *  Apagar de algumas abas e não de outras deixa histórico solto — por isso
 *  recusa na entrada se faltar alguma aba de contrato. */
function executarExcluirEmpresa(nome, confirmacao) {
  nome = String(nome).trim();

  const faltando = abasFaltando(abasNecessarias(['CADASTRO', 'COMPARATIVO']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const emp = acharEmpresa(nome);
  if (!emp) return { ok: false, msg: 'Não achei "' + nome + '" no CADASTRO.' };

  const analise = analisarEmpresa(nome);
  if (analise.temHistorico &&
      String(confirmacao).trim().toUpperCase() !== emp.empresa.toUpperCase()) {
    return { ok: false, msg: 'Confirmação errada. Digite exatamente: ' + emp.empresa };
  }

  const alvo = nome.toUpperCase();

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;
    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;

    const valores = aba.getRange(LINHA_INICIAL, 2, fim - LINHA_INICIAL + 1, 1).getValues();
    for (let i = valores.length - 1; i >= 0; i--) {          // de baixo para cima
      if (String(valores[i][0]).trim().toUpperCase() === alvo) aba.deleteRow(LINHA_INICIAL + i);
    }
    ajustarFiltro(aba, cfg.ultimaCol);
  }

  const comp = planilha().getSheetByName('COMPARATIVO');
  if (comp) {
    const colB = comp.getRange('B1:B' + comp.getLastRow()).getValues();
    for (let i = colB.length - 1; i >= 0; i--) {
      if (String(colB[i][0]).trim().toUpperCase() === alvo) { comp.deleteRow(i + 1); break; }
    }
  }

  const cad = planilha().getSheetByName('CADASTRO');
  const colA = cad.getRange('A1:A' + cad.getLastRow()).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim().toUpperCase() === alvo) { cad.deleteRow(i + 1); break; }
  }

  return { ok: true, msg: '✔ ' + emp.empresa + ' foi apagada de vez.' };
}

// ============================================================
//  6. VIRAR O ANO
// ============================================================

/** Cria uma cópia do arquivo para o ano novo: mantém o CADASTRO e as
 *  empresas do COMPARATIVO, e zera as abas de etapa. O arquivo atual
 *  não é alterado — ele fica como histórico do ano que passou. */
function virarOAno() {
  const ui = SpreadsheetApp.getUi();

  const r = ui.prompt('Virar o ano', 'Ano do arquivo NOVO (ex.: 2027):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const ano = r.getResponseText().trim();
  if (!/^\d{4}$/.test(ano)) { ui.alert('Digite um ano com 4 dígitos.'); return; }

  const conf = ui.alert('Virar o ano',
    'Vou criar uma CÓPIA chamada "CONTROLE EMPRESAS ' + ano + '".\n\n' +
    'Na cópia: o CADASTRO e as empresas do COMPARATIVO são mantidos,\n' +
    'e as abas de etapa ficam zeradas para começar o ano.\n\n' +
    'Este arquivo aqui NÃO é alterado — fica como histórico.\n\nPode fazer?',
    ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  const copia = planilha().copy('CONTROLE EMPRESAS ' + ano);

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = copia.getSheetByName(nomeAba);
    if (!aba) continue;

    const filtro = aba.getFilter();
    if (filtro) filtro.remove();

    const qtd = LIMITE_LINHAS - LINHA_INICIAL + 1;
    if (cfg.colTotal) {
      // preserva a fórmula do TOTAL, limpando dos dois lados dela
      aba.getRange(LINHA_INICIAL, 1, qtd, cfg.colTotal - 1).clearContent();
      aba.getRange(LINHA_INICIAL, cfg.colTotal + 1, qtd, cfg.ultimaCol - cfg.colTotal).clearContent();
    } else {
      aba.getRange(LINHA_INICIAL, 1, qtd, cfg.ultimaCol).clearContent();
    }
  }

  copia.getSheetByName('PAINEL').getRange('D4').setValue('JANEIRO');

  ui.alert('Ano virado!\n\nArquivo novo: CONTROLE EMPRESAS ' + ano +
           '\n\nAbra ele e use ⚙️ Controle → Abrir novo mês → JANEIRO.\n\nLink:\n' + copia.getUrl());
}

// ============================================================
//  7. ALERTA DE PENDÊNCIAS POR E-MAIL
// ============================================================
//  Não tem botão no menu. Para receber automático:
//  Apps Script -> ícone de relógio (Acionadores) -> Adicionar acionador
//  -> função: enviarPendencias | origem: Baseado no tempo | ex.: semanal
// ============================================================

function levantarPendencias(mes) {
  const blocos = [];
  let total = 0;

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;

    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;

    const dados = aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, cfg.ultimaCol).getValues();
    const pendentes = dados
      .filter(l => String(l[0]).trim() === mes)
      .filter(l => {
        const st = String(l[cfg.colFeito - 1]).trim();
        return st === 'Pendente' || st === 'Erro' || st === '';
      })
      .map(l => l[1]);

    if (pendentes.length > 0) {
      total += pendentes.length;
      blocos.push('▸ ' + nomeAba + ' — ' + pendentes.length + ' pendente(s):\n   ' +
                  pendentes.join('\n   '));
    }
  }

  const texto = total === 0
    ? 'Tudo em dia em ' + mes + '! Nenhuma pendência.'
    : 'Pendências de ' + mes + ' (' + total + ' no total):\n\n' + blocos.join('\n\n');

  return { texto: texto, total: total };
}

function enviarPendencias() {
  const mes = mesDoPainel();
  const r = levantarPendencias(mes);
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'Controle de empresas — pendências de ' + mes +
             (r.total > 0 ? ' (' + r.total + ')' : ' — tudo em dia'),
    body: r.texto + '\n\n' + planilha().getUrl() +
          '\n\n—\nEnviado automaticamente pela planilha de controle.',
  });
}

// ============================================================
//  8. EXCLUIR UM MÊS INTEIRO
// ============================================================
//  Serve para quando você abre o mês errado. É destrutivo:
//  antes de apagar, o script mostra o que existe naquele mês e,
//  se houver trabalho feito, exige confirmação digitada.
// ============================================================

function dialogoExcluirMes() {
  const html = HtmlService.createHtmlOutput(htmlExcluirMes())
                          .setWidth(460).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, 'Excluir um mês inteiro');
}

function htmlExcluirMes() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Apaga todas as linhas do mês nas 4 abas. Use quando abrir o mês errado.</div>' +
  '<label>Mês</label>' +
  '<select id="mes" disabled><option value="">carregando...</option></select>' +
  '<div class="botoes">' +
  '  <button class="sec" onclick="google.script.host.close()">Cancelar</button>' +
  '  <button class="pri" id="btnVer" onclick="verificar()">Verificar</button>' +
  '</div>' +
  '<div id="info"></div>' +
  '<div id="bloco" style="display:none">' +
  '  <label id="lblConf">Digite o nome do mês para confirmar</label>' +
  '  <input type="text" id="conf">' +
  '  <div class="botoes">' +
  '    <button class="perigo" id="btnDel" onclick="apagar()">Apagar definitivamente</button>' +
  '  </div>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'var sel = document.getElementById("mes");' +
  'var mesVerificado = null, exigeTexto = false;' +
  'google.script.run.withSuccessHandler(function (meses) {' +
  '  sel.innerHTML = "";' +
  '  if (meses.length === 0) {' +
  '    var v = document.createElement("option");' +
  '    v.value = ""; v.text = "— Nenhum mês aberto —"; sel.add(v);' +
  '    document.getElementById("btnVer").disabled = true;' +
  '    return;' +
  '  }' +
  '  sel.disabled = false;' +
  '  for (var i = 0; i < meses.length; i++) {' +
  '    var o = document.createElement("option");' +
  '    o.value = meses[i]; o.text = meses[i]; sel.add(o);' +
  '  }' +
  '  sel.selectedIndex = meses.length - 1;' +
  '}).mesesAbertos();' +
  'sel.addEventListener("change", function () {' +
  '  document.getElementById("bloco").style.display = "none";' +
  '  document.getElementById("info").innerHTML = "";' +
  '  document.getElementById("status").className = "";' +
  '  document.getElementById("status").textContent = "";' +
  '});' +
  'function verificar() {' +
  '  document.getElementById("btnVer").disabled = true;' +
  '  google.script.run.withSuccessHandler(function (r) {' +
  '    document.getElementById("btnVer").disabled = false;' +
  '    mesVerificado = r.mes; exigeTexto = r.temTrabalho;' +
  '    var cls = r.temTrabalho ? "alerta" : "caixa";' +
  '    document.getElementById("info").innerHTML =' +
  '      \'<div class="\' + cls + \'">\' + r.resumo + "</div>";' +
  '    document.getElementById("bloco").style.display = "block";' +
  '    var lbl = document.getElementById("lblConf");' +
  '    var campo = document.getElementById("conf");' +
  '    if (r.temTrabalho) {' +
  '      lbl.style.display = "block"; campo.style.display = "block";' +
  '      lbl.textContent = "Digite " + r.mes + " para confirmar";' +
  '      campo.value = ""; campo.focus();' +
  '    } else {' +
  '      lbl.style.display = "none"; campo.style.display = "none";' +
  '    }' +
  '  }).analisarMes(sel.value);' +
  '}' +
  'function apagar() {' +
  '  var txt = exigeTexto ? document.getElementById("conf").value.trim() : mesVerificado;' +
  '  document.getElementById("btnDel").disabled = true;' +
  '  var s = document.getElementById("status");' +
  '  s.className = "load"; s.textContent = "Apagando...";' +
  '  google.script.run.withSuccessHandler(function (r) {' +
  '    document.getElementById("btnDel").disabled = false;' +
  '    s.className = r.ok ? "ok" : "erro"; s.textContent = r.msg;' +
  '    if (r.ok) {' +
  '      document.getElementById("bloco").style.display = "none";' +
  '      document.getElementById("info").innerHTML = "";' +
  '    }' +
  '  }).executarExcluirMes(mesVerificado, txt);' +
  '}' +
  '</script></body></html>';
}

/** Mostra o que existe no mês antes de apagar. */
function analisarMes(mes) {
  mes = String(mes).trim().toUpperCase();
  const detalhe = [];
  let totalLinhas = 0, preenchidos = 0, faturamento = 0;

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);
    if (!aba) continue;
    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;

    const dados = aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, cfg.ultimaCol).getValues();
    const doMes = dados.filter(l => String(l[0]).trim().toUpperCase() === mes);
    if (doMes.length === 0) continue;

    totalLinhas += doMes.length;
    detalhe.push('• ' + nomeAba + ': ' + doMes.length + ' linha(s)');

    // trabalho feito = status marcado (fora "Pendente" e vazio)
    doMes.forEach(l => {
      for (let c = 4; c < cfg.ultimaCol; c++) {
        if (cfg.colTotal && c === cfg.colTotal - 1) continue;
        const v = String(l[c]).trim();
        if (v !== '' && v !== 'Pendente') preenchidos++;
      }
      if (cfg.colTotal) {
        const t = l[cfg.colTotal - 1];
        if (typeof t === 'number') faturamento += t;
      }
    });
  }

  if (totalLinhas === 0) {
    return { mes: mes, temTrabalho: false, vazio: true,
             resumo: 'Não existe nenhuma linha de ' + mes + '. Nada a apagar.' };
  }

  const temTrabalho = preenchidos > 0 || faturamento > 0;
  const moeda = 'R$ ' + faturamento.toFixed(2).replace('.', ',');

  const resumo = (temTrabalho ? '⚠ ATENÇÃO: esse mês TEM trabalho feito.\n\n' : 'Mês sem trabalho lançado.\n\n') +
    'Vou apagar ' + totalLinhas + ' linha(s):\n' + detalhe.join('\n') +
    '\n\nStatus marcados: ' + preenchidos +
    '\nFaturamento lançado: ' + moeda +
    (temTrabalho ? '\n\nISSO NÃO TEM VOLTA. Se foi engano, dá para voltar\npor Arquivo → Histórico de versões.' : '');

  return { mes: mes, temTrabalho: temTrabalho, vazio: false, resumo: resumo };
}

/** Apaga o mês. Se houver trabalho feito, exige o nome do mês digitado. */
function executarExcluirMes(mes, confirmacao) {
  mes = String(mes).trim().toUpperCase();
  if (MESES.indexOf(mes) === -1) return { ok: false, msg: 'Mês inválido.' };

  const faltando = abasFaltando(abasNecessarias(['PAINEL']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const analise = analisarMes(mes);
  if (analise.vazio) return { ok: false, msg: 'Não existe nenhuma linha de ' + mes + '.' };

  if (analise.temTrabalho &&
      String(confirmacao).trim().toUpperCase() !== mes) {
    return { ok: false, msg: 'Confirmação errada. Digite exatamente: ' + mes };
  }

  let apagadas = 0;

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);
    const fim = ultimaLinha(aba);
    if (fim < LINHA_INICIAL) continue;

    const colA = aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, 1).getValues();
    for (let i = colA.length - 1; i >= 0; i--) {          // de baixo para cima
      if (String(colA[i][0]).trim().toUpperCase() === mes) {
        aba.deleteRow(LINHA_INICIAL + i);
        apagadas++;
      }
    }
    ajustarFiltro(aba, cfg.ultimaCol);
  }

  // Se o PAINEL apontava para o mês apagado, joga para o último que sobrou
  const painel = planilha().getSheetByName('PAINEL');
  if (String(painel.getRange('D4').getValue()).trim().toUpperCase() === mes) {
    const restantes = mesesAbertos();
    painel.getRange('D4').setValue(restantes.length ? restantes[restantes.length - 1] : 'JANEIRO');
  }
  try { atualizarCalendario(); } catch (e) { /* calendário nunca trava a exclusão */ }

  return { ok: true, msg: '✔ ' + mes + ' apagado. ' + apagadas + ' linha(s) removidas.\n' +
                          'Agora dá para abrir o mês de novo, se quiser.' };
}

// ============================================================
//  9. DETECÇÃO DE FATURAMENTOS SUSPEITOS (ANOMALIAS)
// ============================================================
//  Compara o faturamento de cada empresa/mês com a MÉDIA dos
//  OUTROS meses dela. Se destoar muito, marca como suspeito —
//  provável erro de digitação (zero a mais/a menos) ou cliente
//  com variação forte que vale conferir.
//
//  Não altera a planilha: só lê a 3. FATURAMENTO e mostra o
//  resultado numa janela. É um "vale conferir", não um erro.
// ============================================================

const ANOMALIA_LIMITE = 0.50;   // ±50% da média dispara o alerta
const ANOMALIA_MIN_MESES = 3;   // só analisa empresas com 3+ meses lançados

/** Varre o FATURAMENTO e devolve a lista de valores suspeitos.
 *  Considera só o TOTAL (coluna H) de cada empresa/mês. */
function detectarAnomalias() {
  const fat = planilha().getSheetByName('3. FATURAMENTO');
  if (!fat) return { erro: 'Não achei a aba "3. FATURAMENTO".' };

  const fim = ultimaLinha(fat);
  if (fim < LINHA_INICIAL) return { erro: 'A aba 3. FATURAMENTO está vazia.' };

  // A=MÊS B=EMPRESA ... H=TOTAL  -> lê colunas 1..8
  const dados = fat.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, 8).getValues();

  // agrupa por empresa: { empresa: [{mes, valor, linha}, ...] }
  const porEmpresa = {};
  dados.forEach((l, i) => {
    const mes = String(l[0]).trim().toUpperCase();
    const empresa = String(l[1]).trim();
    const total = Number(l[7]);
    if (!empresa || MESES.indexOf(mes) === -1) return;
    if (!(total > 0)) return;                       // ignora vazio/zero
    if (!porEmpresa[empresa]) porEmpresa[empresa] = [];
    porEmpresa[empresa].push({ mes: mes, valor: total, linha: LINHA_INICIAL + i });
  });

  const suspeitos = [];
  let analisadas = 0, semHistorico = 0;

  for (const empresa in porEmpresa) {
    const meses = porEmpresa[empresa];
    if (meses.length < ANOMALIA_MIN_MESES) { semHistorico++; continue; }
    analisadas++;

    // para cada mês, compara com a média dos OUTROS meses da empresa
    meses.forEach(m => {
      const outros = meses.filter(x => x !== m);
      const media = outros.reduce((s, x) => s + x.valor, 0) / outros.length;
      if (media <= 0) return;
      const desvio = (m.valor - media) / media;     // + acima, - abaixo
      if (Math.abs(desvio) >= ANOMALIA_LIMITE) {
        suspeitos.push({
          empresa: empresa, mes: m.mes, valor: m.valor,
          media: media, desvio: desvio, linha: m.linha,
        });
      }
    });
  }

  // ordena pelo desvio mais gritante primeiro
  suspeitos.sort((a, b) => Math.abs(b.desvio) - Math.abs(a.desvio));

  return { suspeitos: suspeitos, analisadas: analisadas, semHistorico: semHistorico };
}

function dialogoAnomalias() {
  const html = HtmlService.createHtmlOutput(htmlAnomalias())
                          .setWidth(560).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Conferir faturamentos suspeitos');
}

function htmlAnomalias() {
  const r = detectarAnomalias();

  let corpo;
  if (r.erro) {
    corpo = '<div class="alerta">' + r.erro + '</div>';
  } else if (r.suspeitos.length === 0) {
    corpo = '<div class="caixa">✅ Nenhum faturamento fora do padrão.\n\n' +
            'Analisei ' + r.analisadas + ' empresa(s) com ' + ANOMALIA_MIN_MESES +
            '+ meses de histórico.' +
            (r.semHistorico ? '\n' + r.semHistorico + ' empresa(s) ainda têm pouco histórico para comparar.' : '') +
            '</div>';
  } else {
    const linhas = r.suspeitos.map(s => {
      const acima = s.desvio > 0;
      const pct = Math.round(Math.abs(s.desvio) * 100);
      const seta = acima ? '▲' : '▼';
      const cls = acima ? 'up' : 'down';
      const rotulo = acima ? 'acima' : 'abaixo';
      return '<tr>' +
        '<td class="emp">' + s.empresa + '</td>' +
        '<td>' + s.mes + '</td>' +
        '<td class="num">' + fmtR(s.valor) + '</td>' +
        '<td class="num media">' + fmtR(s.media) + '</td>' +
        '<td class="num ' + cls + '">' + seta + ' ' + pct + '% ' + rotulo + '</td>' +
        '</tr>';
    }).join('');

    corpo =
      '<div class="resumo"><b>' + r.suspeitos.length + '</b> valor(es) para conferir ' +
      '— fora de ±' + Math.round(ANOMALIA_LIMITE * 100) + '% da média da própria empresa.</div>' +
      '<div class="tabela-wrap"><table>' +
      '<thead><tr><th>Empresa</th><th>Mês</th><th>Lançado</th><th>Média dela</th><th>Diferença</th></tr></thead>' +
      '<tbody>' + linhas + '</tbody></table></div>' +
      '<div class="nota">Isto é um "vale conferir", não um erro. Pode ser digitação ' +
      '(um zero a mais/a menos) ou variação real do cliente. Confira na aba 3. FATURAMENTO.</div>';
  }

  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() +
  '<style>' +
  '.resumo{font-size:13px;margin-bottom:12px;color:#262626}' +
  '.tabela-wrap{max-height:320px;overflow:auto;border:1px solid #D9E1F2;border-radius:6px}' +
  'table{width:100%;border-collapse:collapse;font-size:12px}' +
  'th{background:#1F3864;color:#fff;padding:7px 9px;text-align:left;position:sticky;top:0;font-size:11px}' +
  'td{padding:6px 9px;border-bottom:1px solid #EEF1F7}' +
  'td.emp{font-weight:bold;color:#1F3864}' +
  'td.num{text-align:right;font-variant-numeric:tabular-nums}' +
  'td.media{color:#808080}' +
  'td.up{color:#9C0006;font-weight:bold}' +      // acima da média = pode ser zero a mais
  'td.down{color:#BF8F00;font-weight:bold}' +    // abaixo = pode ser zero a menos / queda
  'tr:hover td{background:#FBFCFE}' +
  '.nota{margin-top:12px;font-size:11.5px;color:#808080;line-height:1.5}' +
  '</style></head><body>' +
  '<div class="sub">Compara cada faturamento com a média dos outros meses da mesma empresa.</div>' +
  corpo +
  '<div class="botoes"><button class="pri" onclick="google.script.host.close()">Fechar</button></div>' +
  '</body></html>';
}

/** Formata número como R$ pt-BR (para as janelas). */
function fmtR(v) {
  return 'R$ ' + Number(v).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ============================================================
//  10. COMO USAR  (guia rápido + easter egg)
// ============================================================
//  Tela "Como usar": guia compacto das funções do menu.
//  No fim, um campo secreto: digitar a palavra mágica revela
//  quem fez a planilha, com confetes. 🥚
// ============================================================

const AUTOR = 'Jhonatan';
const VERSAO = '2026.1';

function dialogoSobre() {
  const html = HtmlService.createHtmlOutput(htmlSobre())
                          .setWidth(440).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Como usar');
}

/** Itens do guia rápido: [emoji, título, descrição curta]. */
function itensGuia() {
  return [
    ['🗓️', 'Abrir novo mês', 'Monta o mês novo com as empresas ativas nas abas certas.'],
    ['🏢', 'Cadastrar empresa', 'Adiciona uma empresa e escolhe a partir de qual mês ela entra.'],
    ['✏️', 'Renomear empresa', 'Troca o nome em tudo de uma vez (cadastro, abas e comparativo).'],
    ['🔘', 'Ativar / desativar', 'Tira a empresa dos meses novos sem perder o histórico.'],
    ['🔎', 'Conferir suspeitos', 'Aponta faturamentos fora do padrão da empresa (possível erro).'],
    ['🔄', 'Virar o ano', 'Cria a cópia do próximo ano com o cadastro e zera as etapas.'],
  ];
}

/** As linhas do easter egg (aparecem uma a uma, com efeito de digitação). */
function linhasCreditos() {
  return [
    'Você descobriu quem fez a planilha!',
    '',
    '<b>' + AUTOR.toUpperCase() + '</b>',
    '',
    'Obrigado por chegar até aqui. 🎉',
  ];
}

function htmlSobre() {
  const guia = itensGuia().map(g =>
    '<div class="item"><div class="ico">' + g[0] + '</div>' +
    '<div class="txt"><b>' + g[1] + '</b><span>' + g[2] + '</span></div></div>'
  ).join('');

  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() +
  '<style>' +
  '.topo{text-align:center;padding:2px 0 10px}' +
  '.escudo{width:56px;height:56px;margin:0 auto 10px;border-radius:14px;' +
  '  background:linear-gradient(135deg,#2E5496,#1F3864);display:flex;align-items:center;' +
  '  justify-content:center;font-size:28px;box-shadow:0 6px 18px rgba(31,56,100,.35)}' +
  '.topo h2{margin:0;color:#1F3864;font-size:17px}' +
  '.topo .v{color:#808080;font-size:11.5px;margin-top:2px}' +
  '.item{display:flex;gap:11px;align-items:flex-start;padding:9px 11px;margin-bottom:7px;' +
  '  background:#F7F9FC;border:1px solid #E3E8F2;border-radius:8px}' +
  '.item .ico{width:26px;height:26px;flex-shrink:0;border-radius:7px;background:#EEF2FA;' +
  '  display:flex;align-items:center;justify-content:center;font-size:14px}' +
  '.item .txt{display:flex;flex-direction:column;font-size:12px;line-height:1.4}' +
  '.item .txt b{color:#1F3864}' +
  '.item .txt span{color:#595959;font-size:11.5px}' +
  '.linha{height:1px;background:#E3E8F2;margin:14px 0 10px}' +
  '.magia{text-align:center}' +
  '.magia input{width:170px;text-align:center;letter-spacing:1px}' +
  '.magia .dica{font-size:10.5px;color:#B8B8B8;margin-top:5px;font-style:italic}' +
  '#confete{position:fixed;inset:0;pointer-events:none;overflow:hidden}' +
  '.c{position:absolute;top:-12px;width:9px;height:9px;border-radius:2px;animation:cair linear forwards}' +
  '@keyframes cair{to{transform:translateY(600px) rotate(540deg);opacity:0}}' +
  '.festa{animation:pulo .5s ease}' +
  '@keyframes pulo{0%,100%{transform:scale(1)}30%{transform:scale(1.18) rotate(-8deg)}60%{transform:scale(1.1) rotate(6deg)}}' +
  '#creditos{display:none;margin-top:12px;min-height:96px;text-align:center}' +
  '#creditos .ln{font-size:13px;color:#2E5496;line-height:1.9;opacity:0;transition:opacity .3s}' +
  '#creditos .ln.on{opacity:1}' +
  '#creditos .ln b{color:#1F3864;font-size:20px;letter-spacing:1px}' +
  '.cursor{display:inline-block;width:7px;background:#2E5496;margin-left:1px;animation:pisca .6s step-end infinite}' +
  '@keyframes pisca{50%{opacity:0}}' +
  '</style></head><body>' +
  '<div id="confete"></div>' +
  '<div class="topo">' +
  '  <div class="escudo" id="escudo">🧮</div>' +
  '  <h2>Como usar</h2>' +
  '  <div class="v">Modo Contador · v' + VERSAO + '</div>' +
  '</div>' +
  guia +
  '<div class="linha"></div>' +
  '<div class="magia">' +
  '  <input type="text" id="palavra" placeholder="?" autocomplete="off">' +
  '  <div class="dica">descubra quem fez a planilha 🤫</div>' +
  '</div>' +
  '<div id="creditos"></div>' +
  '<div class="botoes"><button class="pri" onclick="google.script.host.close()">Fechar</button></div>' +
  '<script>' +
  'var CORES = ["#2E5496","#1F3864","#C6EFCE","#F2C94C","#9C0006","#8FAADC"];' +
  'function confete(){' +
  '  var box = document.getElementById("confete");' +
  '  for (var i=0;i<80;i++){' +
  '    var c = document.createElement("div"); c.className="c";' +
  '    c.style.left = Math.random()*100+"%";' +
  '    c.style.background = CORES[Math.floor(Math.random()*CORES.length)];' +
  '    c.style.animationDuration = (1.6+Math.random()*1.6)+"s";' +
  '    c.style.animationDelay = (Math.random()*0.6)+"s";' +
  '    box.appendChild(c);' +
  '    (function(el){ setTimeout(function(){ el.remove(); }, 3600); })(c);' +
  '  }' +
  '}' +
  'var CREDITOS = ' + JSON.stringify(linhasCreditos()) + ';' +
  'var jaRolou = false;' +
  // cada linha aparece com fade suave, uma após a outra (robusto e simples)
  'function rolarCreditos(){' +
  '  var box = document.getElementById("creditos");' +
  '  box.style.display = "block"; box.innerHTML = "";' +
  '  CREDITOS.forEach(function(txt, i){' +
  '    var d = document.createElement("div"); d.className = "ln";' +
  '    d.innerHTML = (txt === "") ? "&nbsp;" : txt;' +
  '    box.appendChild(d);' +
  '    setTimeout(function(){ d.classList.add("on"); }, 250 + i * 550);' +
  '  });' +
  '}' +
  'function festa(){' +
  '  confete();' +
  '  var e = document.getElementById("escudo");' +
  '  e.textContent = "🎉"; e.classList.add("festa");' +
  '  setTimeout(function(){ e.classList.remove("festa"); e.textContent = "🧮"; }, 1400);' +
  '  if (!jaRolou){ jaRolou = true; rolarCreditos(); }' +
  '  else { confete(); }' +
  '}' +
  'var inp = document.getElementById("palavra");' +
  'inp.addEventListener("input", function(){' +
  '  if (inp.value.trim().toLowerCase() === "jhonatan"){' +
  '    festa(); inp.value=""; inp.blur();' +
  '  }' +
  '});' +
  '</script></body></html>';
}

// ============================================================
//  11. CALENDÁRIO DE VENCIMENTOS  (no PAINEL, colunas H:N)
// ============================================================
//  Lê a aba VENCIMENTOS (MÊS | DIA | OBRIGAÇÃO | CATEGORIA) e pinta o
//  calendário do mês selecionado (PAINEL!D4) a partir da coluna H.
//  É pintado com VALORES fixos — sem fórmula viva —, então não entra no
//  recálculo da planilha. Repinta: no botão do menu, ao trocar o mês em
//  D4 (onEdit) e ao abrir/excluir um mês.
// ============================================================

const CAL_COL         = 8;    // coluna H
const CAL_LIN_TITULO  = 7;    // H7 (mês / ano) — alinhado com a tabela de etapas
const CAL_LIN_CAB     = 8;    // H8 (SEG..DOM)
const CAL_LIN_GRADE   = 9;    // H9 (primeira semana; 6 semanas até H14)
const CAL_LIN_LEGENDA = 16;   // H16
const CAL_MAX_VENC    = 500;  // teto de linhas lidas na aba VENCIMENTOS

// As categorias moram na própria aba VENCIMENTOS (colunas F:G), do lado das
// obrigações — você adiciona, renomeia e troca a cor ali mesmo, sem tocar no
// código e sem precisar abrir outra aba. A cor vem da bolinha (emoji).
const CAT_COL_NOME  = 6;    // VENCIMENTOS coluna F (nome da categoria)
const CAT_COL_EMOJI = 7;    // VENCIMENTOS coluna G (bolinha)
const CAT_MAX       = 30;   // teto de categorias

// paleta de bolinhas pra escolher (as únicas cores de círculo que há em emoji)
const BOLINHAS_DISPONIVEIS = ['🔴', '🟠', '🟡', '🟢', '🔵', '🟣', '🟤', '⚫', '⚪'];

// usadas só na primeira vez, pra semear a lista (depois é tudo manual)
const CATEGORIAS_PADRAO = [
  ['DAS',               '🟡'],
  ['DCTFWeb',           '🔵'],
  ['SPED',              '🔴'],
  ['EFD-Reinf',         '🟢'],
  ['Impostos Federais', '🟠'],
  ['Outros',            '🟣'],
];

/** Lê as categorias (nome + bolinha) da aba VENCIMENTOS, colunas F:G. */
function lerCategorias() {
  const venc = planilha().getSheetByName('VENCIMENTOS');
  if (!venc) return [];
  const dados = venc.getRange(LINHA_INICIAL, CAT_COL_NOME, CAT_MAX, 2).getValues();
  const cats = [];
  dados.forEach(l => {
    const nome = String(l[0]).trim();
    if (nome) cats.push({ nome: nome, emoji: String(l[1]).trim() });
  });
  return cats;
}

/** Garante a mini-tabela de categorias na aba VENCIMENTOS (colunas F:G):
 *  título, cabeçalho, semente na primeira vez e o dropdown de bolinha.
 *  Idempotente: não sobrescreve o que você editou. */
function garantirCategorias() {
  const venc = planilha().getSheetByName('VENCIMENTOS');
  if (!venc) return;

  venc.getRange(1, CAT_COL_NOME).setValue('CATEGORIAS (edite aqui)')
      .setFontWeight('bold').setFontColor(COR_NAVY);
  venc.getRange(2, CAT_COL_NOME, 1, 2).setValues([['CATEGORIA', 'BOLINHA']])
      .setFontWeight('bold').setBackground(COR_NAVY).setFontColor('#ffffff')
      .setHorizontalAlignment('center');

  if (lerCategorias().length === 0) {
    venc.getRange(LINHA_INICIAL, CAT_COL_NOME, CATEGORIAS_PADRAO.length, 2)
        .setValues(CATEGORIAS_PADRAO);
  }

  const dvBolinha = SpreadsheetApp.newDataValidation()
                      .requireValueInList(BOLINHAS_DISPONIVEIS, true).build();
  venc.getRange(LINHA_INICIAL, CAT_COL_EMOJI, CAT_MAX, 1).setDataValidation(dvBolinha);

  venc.setColumnWidth(5, 30);                 // E: respiro entre as duas tabelas
  venc.setColumnWidth(CAT_COL_NOME, 160);     // F
  venc.setColumnWidth(CAT_COL_EMOJI, 75);     // G

  // remove a tabela antiga de categorias que ficava na LISTAS (agora vive aqui)
  const listas = planilha().getSheetByName('LISTAS');
  if (listas) listas.getRange(1, 7, CAT_MAX + 1, 2).clearContent().clearDataValidations();
}

/** Aponta o dropdown de CATEGORIA (coluna D) para a lista de categorias
 *  ao lado (coluna F), pra ele crescer sozinho quando você adiciona uma nova. */
function aplicarDropdownCategoria() {
  const venc = planilha().getSheetByName('VENCIMENTOS');
  if (!venc) return;
  const dvCat = SpreadsheetApp.newDataValidation()
      .requireValueInRange(venc.getRange(LINHA_INICIAL, CAT_COL_NOME, CAT_MAX, 1), true).build();
  venc.getRange('D3:D' + CAL_MAX_VENC).setDataValidation(dvCat);
}

/** Ano do arquivo — tirado do nome (ex.: "...2026..."); cai no ano atual
 *  se não achar. Assim, ao "Virar o ano", o calendário acompanha sozinho. */
function anoDoArquivo() {
  const m = String(planilha().getName()).match(/(\d{4})/);
  return m ? Number(m[1]) : new Date().getFullYear();
}

/** Chave de dia (ano-mês-dia) para comparar datas sem hora. */
function ymd(d) {
  return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
}

/** Item de menu: garante a aba de config, desenha a moldura e pinta. */
function mostrarCalendario() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) { SpreadsheetApp.getUi().alert('Não achei a aba PAINEL.'); return; }

  garantirAbaVencimentos();
  garantirCategorias();
  aplicarDropdownCategoria();
  montarCalendario();
  atualizarCalendario();
  planilha().setActiveSheet(planilha().getSheetByName('VENCIMENTOS') || painel);

  SpreadsheetApp.getUi().alert('Calendário pronto!\n\n' +
    'Está tudo na aba VENCIMENTOS: os prazos (colunas A–D) e, ao lado, a lista\n' +
    'de categorias e cores (colunas F–G) — pode renomear, adicionar e trocar a\n' +
    'bolinha à vontade. O calendário, a legenda e o dropdown se atualizam sozinhos.');
}

/** Cria a aba VENCIMENTOS (se não existir) com cabeçalho, dropdowns e um
 *  exemplo. É de onde o calendário lê os prazos. */
function garantirAbaVencimentos() {
  const ss = planilha();
  let aba = ss.getSheetByName('VENCIMENTOS');
  if (aba) return aba;

  aba = ss.insertSheet('VENCIMENTOS');

  aba.getRange('A1').setValue(
      'VENCIMENTOS — lance aqui os prazos de cada mês (o calendário do PAINEL lê daqui)')
     .setFontWeight('bold').setFontColor(COR_NAVY).setFontSize(12);

  aba.getRange('A2:D2').setValues([['MÊS', 'DIA', 'OBRIGAÇÃO', 'CATEGORIA']])
     .setFontWeight('bold').setBackground(COR_NAVY).setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  const dvMes = SpreadsheetApp.newDataValidation().requireValueInList(MESES, true).build();
  const dvDia = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 31).build();
  aba.getRange('A3:A' + CAL_MAX_VENC).setDataValidation(dvMes);
  aba.getRange('B3:B' + CAL_MAX_VENC).setDataValidation(dvDia);
  // o dropdown de CATEGORIA (coluna D) é aplicado por aplicarDropdownCategoria()

  // Exemplo (o do ABRIL que você mandou — pode editar/apagar à vontade).
  aba.getRange('A3:D8').setValues([
    ['ABRIL', 7,  'Folha',                  'Outros'],
    ['ABRIL', 14, 'SPED EFD Contribuições', 'EFD-Reinf'],
    ['ABRIL', 15, 'SPED ICMS',              'SPED'],
    ['ABRIL', 20, 'Simples (DAS)',          'DAS'],
    ['ABRIL', 25, 'DCTFWeb',                'DCTFWeb'],
    ['ABRIL', 25, 'PIS e Cofins',           'Outros'],
  ]);

  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 55);
  aba.setColumnWidth(3, 230);
  aba.setColumnWidth(4, 120);
  aba.setFrozenRows(2);

  return aba;
}

/** Desenha a moldura fixa do calendário (título, cabeçalho, legenda,
 *  larguras e bordas). Idempotente: pode rodar quantas vezes quiser. */
function montarCalendario() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) return;

  // limpa a faixa do calendário (H:N) antes de desenhar — assim, mudar a
  // posição não deixa sobra do lugar antigo (merges, valores, bordas)
  painel.getRange(1, CAL_COL, 20, 7).breakApart().clearContent().clearFormat();

  // larguras só das colunas do calendário (H:N estão livres no PAINEL)
  for (let c = CAL_COL; c < CAL_COL + 7; c++) painel.setColumnWidth(c, 64);

  // título (mescla H2:N2) — o texto do mês entra no atualizarCalendario
  const titulo = painel.getRange(CAL_LIN_TITULO, CAL_COL, 1, 7);
  titulo.breakApart().merge();
  titulo.setFontWeight('bold').setFontColor('#ffffff').setFontSize(13)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBackground(COR_NAVY);

  // cabeçalho dos dias da semana
  painel.getRange(CAL_LIN_CAB, CAL_COL, 1, 7)
        .setValues([['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SÁB', 'DOM']])
        .setFontWeight('bold').setFontColor('#808080').setFontSize(10)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBackground('#EEF2FA');

  // moldura da grade (cabeçalho + 6 semanas)
  painel.getRange(CAL_LIN_CAB, CAL_COL, 7, 7)
        .setBorder(true, true, true, true, true, true,
                   '#D9E1F2', SpreadsheetApp.BorderStyle.SOLID);

  // legenda (mescla; montada a partir das categorias da aba VENCIMENTOS)
  const legenda = painel.getRange(CAL_LIN_LEGENDA, CAL_COL, 1, 7);
  legenda.breakApart().merge();
  const textoLegenda = lerCategorias().map(c => c.emoji + ' ' + c.nome).join('   ');
  painel.getRange(CAL_LIN_LEGENDA, CAL_COL).setValue(textoLegenda);
  legenda.setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/** Repinta o calendário do mês selecionado. Só valores/cores — sem fórmula. */
function atualizarCalendario() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) return;

  const mesTexto = String(mesDoPainel()).trim().toUpperCase();
  const mesNum = MESES.indexOf(mesTexto) + 1;
  if (mesNum === 0) return;                       // D4 sem mês válido: não mexe

  const ano = anoDoArquivo();

  // mapa categoria -> bolinha (lido uma vez)
  const mapaCat = {};
  lerCategorias().forEach(c => { mapaCat[c.nome.toUpperCase()] = c.emoji; });

  // dia -> bolinhas, lendo a aba VENCIMENTOS só do mês exibido
  const bolinhas = {};
  const venc = planilha().getSheetByName('VENCIMENTOS');
  if (venc) {
    const fim = Math.min(venc.getLastRow(), CAL_MAX_VENC);
    if (fim >= LINHA_INICIAL) {
      const dados = venc.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, 4).getValues();
      dados.forEach(l => {
        if (String(l[0]).trim().toUpperCase() !== mesTexto) return;
        const dia = Number(l[1]);
        if (!(dia >= 1 && dia <= 31)) return;
        const emoji = mapaCat[String(l[3]).trim().toUpperCase()] || '';
        if (emoji) bolinhas[dia] = (bolinhas[dia] || '') + emoji;
      });
    }
  }

  // título
  painel.getRange(CAL_LIN_TITULO, CAL_COL).setValue(mesTexto + ' / ' + ano);

  // a grade começa na segunda-feira da semana do dia 1
  const primeiro = new Date(ano, mesNum - 1, 1);
  const desloca = (primeiro.getDay() + 6) % 7;    // 0 = seg ... 6 = dom
  const inicio = new Date(ano, mesNum - 1, 1 - desloca);

  const hojeStr = ymd(new Date());
  const valores = [], fundos = [], fontes = [];

  for (let sem = 0; sem < 6; sem++) {
    const vRow = [], bRow = [], fRow = [];
    for (let dow = 0; dow < 7; dow++) {
      const d = new Date(inicio.getFullYear(), inicio.getMonth(),
                         inicio.getDate() + sem * 7 + dow);
      const noMes  = (d.getMonth() + 1 === mesNum);
      const ehHoje = noMes && ymd(d) === hojeStr;
      let txt = String(d.getDate());
      if (noMes && bolinhas[d.getDate()]) txt += ' ' + bolinhas[d.getDate()];
      vRow.push(txt);
      bRow.push(ehHoje ? COR_NAVY : '#ffffff');
      fRow.push(ehHoje ? '#ffffff' : (noMes ? '#000000' : '#B7B7B7'));
    }
    valores.push(vRow); fundos.push(bRow); fontes.push(fRow);
  }

  painel.getRange(CAL_LIN_GRADE, CAL_COL, 6, 7)
        .setValues(valores).setBackgrounds(fundos).setFontColors(fontes)
        .setFontFamily('Arial').setFontSize(11)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/** Repinta o calendário quando você troca o mês em PAINEL!D4, ou quando
 *  edita a aba VENCIMENTOS. Simples trigger: nunca atrapalha a edição. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const aba = e.range.getSheet();
    const nome = aba.getName();
    if (nome === 'VENCIMENTOS') { atualizarCalendario(); return; }
    if (nome === 'PAINEL' && e.range.getA1Notation() === 'D4') atualizarCalendario();
  } catch (err) { /* nunca quebra a digitação do usuário */ }
}

/** Coloca o seletor de mês (PAINEL!D4) no mês atual. Chamado ao abrir a
 *  planilha. Não trava nada: você segue trocando o mês pelo dropdown quando
 *  quiser — só no próximo abrir ele volta pro mês de hoje. */
function irParaMesAtual() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) return;
  painel.getRange('D4').setValue(MESES[new Date().getMonth()]);
  try { atualizarCalendario(); } catch (e) { /* calendário repinta depois se falhar aqui */ }
}

// ============================================================
//  12. BLOCO DE NOTAS  (barra lateral)
// ============================================================
//  Anotações gerais guardadas numa aba NOTAS (oculta):
//    A=ID B=DATA C=PRIORIDADE D=EMPRESA E=DESCRIÇÃO F=STATUS G=DATA CONCLUSÃO
//  A lateral é HTML montada só quando aberta — não pesa na planilha.
//  Abre pelo menu ou por um desenho com o script abrirBlocoDeNotas atribuído.
// ============================================================

const NOTAS_ABA = 'NOTAS';
const NOTAS_PRIORIDADES = ['Alta', 'Média', 'Baixa'];

function abrirBlocoDeNotas() {
  const ativo = planilha().getActiveSheet();
  garantirAbaNotas();
  planilha().setActiveSheet(ativo);           // não pula pra aba oculta
  const html = HtmlService.createHtmlOutput(htmlBlocoNotas()).setWidth(440).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bloco de Notas');
}

/** Cria a aba NOTAS (oculta) com cabeçalho, se ainda não existir. */
function garantirAbaNotas() {
  const ss = planilha();
  let aba = ss.getSheetByName(NOTAS_ABA);
  if (aba) return aba;
  aba = ss.insertSheet(NOTAS_ABA);
  aba.getRange(1, 1, 1, 7).setValues([[
    'ID', 'DATA', 'PRIORIDADE', 'EMPRESA', 'DESCRIÇÃO', 'STATUS', 'DATA CONCLUSÃO']])
     .setFontWeight('bold').setBackground(COR_NAVY).setFontColor('#ffffff');
  aba.getRange('B2:B1000').setNumberFormat('dd/mm/yyyy');
  aba.getRange('G2:G1000').setNumberFormat('dd/mm/yyyy');
  aba.setColumnWidth(5, 320);
  aba.setFrozenRows(1);
  aba.hideSheet();
  return aba;
}

function fmtData(v) {
  if (!(v instanceof Date)) return v === null || v === undefined ? '' : String(v);
  return ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2) + '/' + v.getFullYear();
}
function fmtDataISO(v) {
  if (!(v instanceof Date)) return '';
  return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
}
function parseDataISO(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date();
}

/** Devolve {pendentes, concluidas}, cada nota com data já formatada. */
function listarNotas() {
  const aba = garantirAbaNotas();
  const fim = aba.getLastRow();
  const pendentes = [], concluidas = [], avisos = [];
  if (fim >= 2) {
    aba.getRange(2, 1, fim - 1, 7).getValues().forEach(l => {
      if (l[0] === '' || l[0] === null) return;
      const st = String(l[5]).trim().toUpperCase();
      const nota = {
        id: Number(l[0]),
        data: fmtData(l[1]),
        dataISO: fmtDataISO(l[1]),
        ordem: (l[1] instanceof Date) ? l[1].getTime() : 0,
        prioridade: String(l[2]).trim() || 'Média',
        empresa: String(l[3]).trim(),
        descricao: String(l[4]),
        dataConclusao: fmtData(l[6]),
      };
      if (st === 'AVISO') avisos.push(nota);
      else if (st === 'CONCLUÍDA' || st === 'CONCLUIDA') concluidas.push(nota);
      else pendentes.push(nota);
    });
  }
  pendentes.sort((a, b) => b.ordem - a.ordem);      // mais recente primeiro
  avisos.sort((a, b) => b.ordem - a.ordem);
  concluidas.sort((a, b) => b.ordem - a.ordem);
  return { pendentes: pendentes, concluidas: concluidas, avisos: avisos };
}

function acharLinhaNota(aba, id) {
  const fim = aba.getLastRow();
  if (fim < 2) return 0;
  const ids = aba.getRange(2, 1, fim - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (Number(ids[i][0]) === Number(id)) return 2 + i;
  return 0;
}
function proximoIdNota(aba) {
  const fim = aba.getLastRow();
  if (fim < 2) return 1;
  let max = 0;
  aba.getRange(2, 1, fim - 1, 1).getValues().forEach(r => { const n = Number(r[0]); if (n > max) max = n; });
  return max + 1;
}

/** Cria (id vazio) ou edita (id existente) uma nota. Devolve {ok, msg}. */
function salvarNota(obj) {
  const aba = garantirAbaNotas();
  const desc = String(obj.descricao || '').trim();
  if (!desc) return { ok: false, msg: 'Escreva a descrição da nota.' };
  const prioridade = NOTAS_PRIORIDADES.indexOf(obj.prioridade) !== -1 ? obj.prioridade : 'Média';
  const empresa = String(obj.empresa || '').trim();
  const data = parseDataISO(obj.dataISO);
  const id = Number(obj.id);
  const ehAviso = obj.tipo === 'aviso';

  if (id > 0) {
    const linha = acharLinhaNota(aba, id);
    if (!linha) return { ok: false, msg: 'Nota não encontrada.' };
    aba.getRange(linha, 2, 1, 4).setValues([[data, prioridade, empresa, desc]]);
    const stAtual = String(aba.getRange(linha, 6).getValue()).trim().toUpperCase();
    const novoSt = ehAviso ? 'Aviso'
                 : ((stAtual === 'CONCLUÍDA' || stAtual === 'CONCLUIDA') ? 'Concluída' : 'Pendente');
    aba.getRange(linha, 6).setValue(novoSt);
    if (novoSt !== 'Concluída') aba.getRange(linha, 7).setValue('');
  } else {
    const linha = Math.max(aba.getLastRow(), 1) + 1;
    aba.getRange(linha, 1, 1, 7)
       .setValues([[proximoIdNota(aba), data, prioridade, empresa, desc, ehAviso ? 'Aviso' : 'Pendente', '']]);
  }
  return { ok: true };
}

function apagarNota(id) {
  const aba = garantirAbaNotas();
  const linha = acharLinhaNota(aba, id);
  if (linha) aba.deleteRow(linha);
  return { ok: true };
}

/** Alterna Pendente <-> Concluída e grava/limpa a data de conclusão. */
function alternarConcluida(id) {
  const aba = garantirAbaNotas();
  const linha = acharLinhaNota(aba, id);
  if (!linha) return { ok: false };
  const st = String(aba.getRange(linha, 6).getValue()).trim().toUpperCase();
  if (st === 'AVISO') return { ok: true };          // aviso não conclui
  const concluir = !(st === 'CONCLUÍDA' || st === 'CONCLUIDA');
  aba.getRange(linha, 6).setValue(concluir ? 'Concluída' : 'Pendente');
  aba.getRange(linha, 7).setValue(concluir ? new Date() : '');
  return { ok: true };
}

function htmlBlocoNotas() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">
<style>
* { box-sizing: border-box; }
html, body { height: 100%; }
body { margin: 0; background: #f6f8fb; color: #3c4757; font-family: 'Google Sans', Roboto, Arial, sans-serif; font-size: 13px; display: flex; flex-direction: column; }
::-webkit-scrollbar { width: 8px; } ::-webkit-scrollbar-thumb { background: #c6d0dc; border-radius: 8px; }
button { font-family: inherit; cursor: pointer; }
.novaWrap { padding: 16px 14px 12px; background: #fff; border-bottom: 1px solid #eef1f5; flex-shrink: 0; }
.btnNova { width: 100%; background: #1e8e3e; color: #fff; border: none; border-radius: 8px; padding: 10px 14px; font-size: 13.5px; font-weight: 600; display: flex; align-items: center; justify-content: center; gap: 7px; box-shadow: 0 1px 2px rgba(20,40,70,.12); }
.btnNova:hover { background: #187a35; }
.filtros { padding: 11px 14px 4px; display: flex; gap: 6px; flex-wrap: wrap; flex-shrink: 0; }
.chip { padding: 5px 12px; border-radius: 20px; font-size: 11.5px; font-weight: 600; cursor: pointer; border: 1px solid #dbe2ea; color: #5c6b7d; background: #fff; }
.chip.on { background: #1a3c6e; color: #fff; border-color: #1a3c6e; }
.lista { padding: 8px 14px 16px; display: flex; flex-direction: column; gap: 10px; flex: 1; overflow-y: auto; min-height: 0; }
.nota { background: #fff; border-radius: 10px; border: 1px solid #e4e9f0; border-left: 4px solid #ccc; padding: 12px 13px; }
.nota .top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
.nota .prio { display: flex; align-items: center; gap: 7px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .3px; }
.nota .prio .dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
.nota .date { font-size: 11px; color: #8a97a8; }
.nota .emp { font-size: 12.5px; font-weight: 600; color: #1a3c6e; margin-bottom: 3px; }
.nota .desc { font-size: 13px; line-height: 1.42; white-space: pre-wrap; word-break: break-word; }
.nota.done .desc { text-decoration: line-through; color: #9aa7b8; }
.nota .acoes { display: flex; align-items: center; margin-top: 10px; padding-top: 9px; border-top: 1px solid #eef1f5; }
.nota .chk { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #5c6b7d; cursor: pointer; }
.nota .chk input { width: 15px; height: 15px; accent-color: #1e8e3e; cursor: pointer; }
.nota .icos { display: flex; gap: 13px; margin-left: auto; }
.nota .icos span { cursor: pointer; display: inline-flex; }
.vazio { text-align: center; padding: 42px 12px; color: #9aa7b8; font-size: 13px; line-height: 1.6; }
#overlay { position: fixed; inset: 0; background: #fff; padding: 16px; overflow-y: auto; display: none; }
#overlay h3 { margin: 2px 0 10px; color: #1a3c6e; font-size: 15.5px; }
label { display: block; font-weight: 600; color: #1a3c6e; margin: 12px 0 4px; font-size: 12px; }
select, input[type=date], textarea { width: 100%; padding: 8px; border: 1px solid #dbe2ea; border-radius: 6px; font-size: 13px; font-family: inherit; }
textarea { resize: vertical; }
.tipoT, .prioT { display: flex; gap: 6px; }
.tipoT button, .prioT button { flex: 1; padding: 8px 0; border: 1px solid #D0D0D0; background: #F2F2F2; color: #808080; border-radius: 6px; font-size: 12px; font-weight: 600; }
.tipoT button.onT { background: #E4EEFB; border-color: #1a3c6e; color: #1a3c6e; }
.tipoT button.onV { background: #EEE7FB; border-color: #8250df; color: #6f42c1; }
.prioT button.onA { background: #FCE4E2; border-color: #e0362c; color: #c0281f; }
.prioT button.onM { background: #FBEFDD; border-color: #f0a400; color: #9a6b00; }
.prioT button.onB { background: #E4EEFB; border-color: #4a89dc; color: #2f6cb5; }
.formBtns { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }
.cinza { background: #fff; color: #5c6b7d; border: 1px solid #dbe2ea; border-radius: 7px; padding: 9px 16px; font-size: 13px; }
.salvar { background: #1e8e3e; color: #fff; border: none; border-radius: 7px; padding: 9px 20px; font-size: 13px; font-weight: 600; }
#msg { margin-top: 12px; font-size: 12px; padding: 8px 10px; border-radius: 6px; display: none; }
#msg.erro { display: block; background: #FFDDE0; color: #9C0006; }
</style></head><body>
<div class="novaWrap">
  <button class="btnNova" onclick="novaNota()"><svg width="15" height="15" viewBox="0 0 24 24" fill="none"><path d="M12 5v14M5 12h14" stroke="#fff" stroke-width="2.4" stroke-linecap="round"></path></svg>Nova Nota</button>
</div>
<div class="filtros" id="filtros"></div>
<div class="lista" id="lista"></div>
<div id="overlay">
  <h3 id="formTit">Nova Nota</h3>
  <label>Tipo</label>
  <div class="tipoT">
    <button type="button" id="tTarefa" onclick="setTipo('tarefa')">Tarefa</button>
    <button type="button" id="tAviso" onclick="setTipo('aviso')">Aviso</button>
  </div>
  <label>Empresa (opcional)</label>
  <select id="fEmp"></select>
  <div id="prioBloco">
    <label>Prioridade</label>
    <div class="prioT">
      <button type="button" id="bA" onclick="setPrio('Alta')">Alta</button>
      <button type="button" id="bM" onclick="setPrio('Média')">Média</button>
      <button type="button" id="bB" onclick="setPrio('Baixa')">Baixa</button>
    </div>
  </div>
  <label>Data</label>
  <input type="date" id="fData">
  <label>Descrição</label>
  <textarea id="fDesc" rows="4"></textarea>
  <div class="formBtns">
    <button class="cinza" onclick="cancelar()">Cancelar</button>
    <button class="salvar" onclick="salvar()">Salvar</button>
  </div>
  <div id="msg"></div>
</div>
<script>
var editId = 0, prioSel = 'Média', tipoSel = 'tarefa', filtro = 'todas', cache = { pendentes: [], concluidas: [], avisos: [] };
var PRIO = { 'Alta':{c:'#e0362c',t:'#c0281f'}, 'Média':{c:'#f0a400',t:'#9a6b00'}, 'Baixa':{c:'#4a89dc',t:'#2f6cb5'} };
var DONE = { c:'#34a853', t:'#1e7e34' };
var AVISO = { c:'#8250df', t:'#6f42c1' };
var FILTROS = [ {k:'todas',r:'Todas'}, {k:'pendentes',r:'Pendentes'}, {k:'avisos',r:'Avisos'}, {k:'concluidas',r:'Concluídas'} ];
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function hojeISO(){ var d=new Date(); return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2); }
function msg(t){ var m=document.getElementById('msg'); m.className='erro'; m.textContent=t; }
google.script.run.withSuccessHandler(function(lista){
  var s=document.getElementById('fEmp');
  var o=document.createElement('option'); o.value=''; o.text='— nenhuma —'; s.add(o);
  lista.forEach(function(e){ var op=document.createElement('option'); op.value=e.nome; op.text=e.nome; s.add(op); });
}).listarEmpresas();
var svgEdit = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none"><path d="M4 20h4l11-11-4-4L4 16v4z" stroke="#7a8aa0" stroke-width="1.7" stroke-linejoin="round"></path></svg>';
var svgDel  = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none"><path d="M5 7h14M9 7V5h6v2M7 7l1 13h8l1-13" stroke="#c2506a" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"></path></svg>';
function card(n){
  var cor, rot;
  if(n.tipo==='aviso'){ cor=AVISO; rot='Aviso'; }
  else if(n.tipo==='concluida'){ cor=DONE; rot='Concluída'; }
  else { cor=PRIO[n.prioridade]||PRIO['Média']; rot=n.prioridade; }
  var emp = n.empresa ? '<div class="emp">'+esc(n.empresa)+'</div>' : '';
  var chk = (n.tipo==='aviso') ? '' : '<label class="chk"><input type="checkbox" '+(n.tipo==='concluida'?'checked':'')+' onchange="toggle('+n.id+')"> concluída</label>';
  return '<div class="nota '+(n.tipo==='concluida'?'done':'')+'" style="border-left-color:'+cor.c+'">'+
    '<div class="top"><span class="prio" style="color:'+cor.t+'"><span class="dot" style="background:'+cor.c+'"></span>'+esc(rot)+'</span>'+
    '<span class="date">'+esc(n.data)+'</span></div>'+ emp +
    '<div class="desc">'+esc(n.descricao)+'</div>'+
    '<div class="acoes">'+chk+
    '<div class="icos"><span title="Editar" onclick="editar('+n.id+')">'+svgEdit+'</span>'+
    '<span title="Apagar" onclick="apagar('+n.id+')">'+svgDel+'</span></div></div></div>';
}
function render(dados){
  cache = dados;
  var pend = dados.pendentes.map(function(n){ n.tipo='pendente'; return n; });
  var avis = dados.avisos.map(function(n){ n.tipo='aviso'; return n; });
  var conc = dados.concluidas.map(function(n){ n.tipo='concluida'; return n; });
  document.getElementById('filtros').innerHTML = FILTROS.map(function(f){
    return '<div class="chip '+(filtro===f.k?'on':'')+'" onclick="setFiltro(\\''+f.k+'\\')">'+f.r+'</div>';
  }).join('');
  var vis;
  if(filtro==='pendentes') vis=pend;
  else if(filtro==='avisos') vis=avis;
  else if(filtro==='concluidas') vis=conc;
  else vis=pend.concat(avis).concat(conc);
  var lista=document.getElementById('lista');
  if(vis.length){ lista.innerHTML = vis.map(card).join(''); }
  else { lista.innerHTML = '<div class="vazio">Nada aqui ainda.<br>Clique em “Nova Nota”.</div>'; }
}
function carregar(){ google.script.run.withSuccessHandler(render).listarNotas(); }
function setFiltro(k){ filtro=k; render(cache); }
function setPrio(p){
  prioSel=p;
  document.getElementById('bA').className=(p==='Alta')?'onA':'';
  document.getElementById('bM').className=(p==='Média')?'onM':'';
  document.getElementById('bB').className=(p==='Baixa')?'onB':'';
}
function setTipo(t){
  tipoSel=t;
  document.getElementById('tTarefa').className=(t==='tarefa')?'onT':'';
  document.getElementById('tAviso').className=(t==='aviso')?'onV':'';
  document.getElementById('prioBloco').style.display=(t==='aviso')?'none':'block';
}
function novaNota(){
  editId=0;
  document.getElementById('formTit').textContent='Nova Nota';
  document.getElementById('fEmp').value='';
  document.getElementById('fData').value=hojeISO();
  document.getElementById('fDesc').value='';
  document.getElementById('msg').className='';
  setTipo('tarefa'); setPrio('Média');
  document.getElementById('overlay').style.display='block';
}
function editar(id){
  var n=cache.pendentes.concat(cache.concluidas,cache.avisos).filter(function(x){return x.id===id;})[0];
  if(!n) return;
  editId=id;
  document.getElementById('formTit').textContent='Editar Nota';
  document.getElementById('fEmp').value=n.empresa||'';
  document.getElementById('fData').value=n.dataISO||hojeISO();
  document.getElementById('fDesc').value=n.descricao||'';
  document.getElementById('msg').className='';
  setTipo(n.tipo==='aviso'?'aviso':'tarefa'); setPrio(n.prioridade||'Média');
  document.getElementById('overlay').style.display='block';
}
function cancelar(){ document.getElementById('overlay').style.display='none'; }
function salvar(){
  var desc=document.getElementById('fDesc').value.trim();
  if(!desc){ msg('Escreva a descrição.'); return; }
  google.script.run.withSuccessHandler(function(r){
    if(r && r.ok===false){ msg(r.msg||'Erro ao salvar.'); return; }
    cancelar(); carregar();
  }).withFailureHandler(function(e){ msg('Erro: '+e.message); })
    .salvarNota({ id:editId, tipo:tipoSel, dataISO:document.getElementById('fData').value, prioridade:prioSel, empresa:document.getElementById('fEmp').value, descricao:desc });
}
function toggle(id){ google.script.run.withSuccessHandler(carregar).alternarConcluida(id); }
function apagar(id){ if(confirm('Apagar esta anotação?')) google.script.run.withSuccessHandler(carregar).apagarNota(id); }
carregar();
</script>
</body></html>`;
}
