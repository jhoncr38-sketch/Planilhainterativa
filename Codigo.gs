/**
 * CONTROLE DE EMPRESAS — automações (Google Apps Script)
 * Para a planilha CONTROLE_EMPRESAS_2026.
 *
 * Layout:
 *   CADASTRO: A=Empresa B=CNPJ C=Regime (fixas); as demais (Faz Folha?,
 *             Faz SPED?, Faz EFD Contrib.?, Ativa?, Perfil, IE, IM, Senha)
 *             são achadas pelo cabeçalho  (dados a partir da linha 3)
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
// colIcms / colContrib = colunas da EFD ICMS/IPI e da EFD Contribuições
const ABAS_ETAPA = {
  '1. FOLHA':       { flag: 'folha', ultimaCol: 9,  colFeito: 9 },
  '2. SPED':        { flag: 'sped',  ultimaCol: 11, colFeito: 11, colData: 9,
                      colIcms: 7, colContrib: 8 },
  '3. FATURAMENTO': { flag: null,    ultimaCol: 10, colFeito: 9,
                      colsMoeda: [5, 6, 7, 8], colTotal: 8 },
  '4. CONSULTAS':   { flag: null,    ultimaCol: 6,  colFeito: 6 },
};

const MESES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO',
               'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];

const REGIMES = ['Simples Nacional', 'Simples Híbrido', 'Lucro Presumido', 'MEI'];

// CADASTRO: EMPRESA, CNPJ e REGIME ficam fixos em A, B e C (as fórmulas das
// abas de etapa buscam neles). As outras colunas são achadas pelo CABEÇALHO
// da linha 2 — assim dá para inserir uma coluna nova sem quebrar o script.
const CAD_CABECALHOS = {
  folha: 'FAZ FOLHA?', sped: 'FAZ SPED?', contrib: 'FAZ EFD CONTRIB.?', ativa: 'ATIVA?',
  perfil: 'PERFIL', ie: 'INSCRIÇÃO ESTADUAL', im: 'INSCRIÇÃO MUNICIPAL', senha: 'SENHA',
};
const CAD_OBRIGATORIAS = ['folha', 'sped', 'ativa'];

const NAO_SE_APLICA = 'Não se aplica';

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
    .addItem('Importar empresas (em massa)', 'dialogoImportar')
    .addSeparator()
    .addItem('Renomear empresa', 'dialogoRenomear')
    .addItem('Ativar / desativar empresa', 'dialogoAtivar')
    .addItem('Excluir empresa de vez', 'dialogoExcluir')
    .addItem('Excluir um mês inteiro', 'dialogoExcluirMes')
    .addSeparator()
    .addItem('Ficha da empresa', 'dialogoFicha')
    .addItem('Conferir faturamentos suspeitos', 'dialogoAnomalias')
    .addItem('Bloco de notas', 'abrirBlocoDeNotas')
    .addItem('Virar o ano', 'virarOAno')
    .addSeparator()
    .addItem('Como usar', 'dialogoSobre')
    .addToUi();

  // abre já no mês atual (você continua podendo trocar pelo dropdown)
  try { irParaMesAtual(); } catch (e) { /* onOpen em modo limitado às vezes não escreve */ }
}

/** Coloca o seletor de mês (PAINEL!D4) no mês atual. Chamado ao abrir a
 *  planilha. Não trava nada: você segue trocando o mês pelo dropdown quando
 *  quiser — só no próximo abrir ele volta pro mês de hoje. */
function irParaMesAtual() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) return;
  painel.getRange('D4').setValue(MESES[new Date().getMonth()]);
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

/** Posição (1 = A) de cada coluna do CADASTRO, achada pelo cabeçalho da
 *  linha 2. 0 = a coluna não existe (ex.: FAZ EFD CONTRIB.? antes de criada).
 *  Recusa se faltar uma obrigatória: ler a coluna errada (ex.: tratar PERFIL
 *  como ATIVA?) é pior que parar. */
function colunasCadastro() {
  const aba = planilha().getSheetByName('CADASTRO');
  const cab = aba.getRange(2, 1, 1, Math.max(aba.getLastColumn(), 1)).getValues()[0]
                 .map(v => String(v).replace(/\s+/g, ' ').trim().toUpperCase());
  const col = { empresa: 1, cnpj: 2, regime: 3, total: cab.length };
  for (const k in CAD_CABECALHOS) col[k] = cab.indexOf(CAD_CABECALHOS[k]) + 1;
  const faltam = CAD_OBRIGATORIAS.filter(k => !col[k]);
  if (faltam.length > 0) {
    throw new Error('No CADASTRO (linha 2) não achei a coluna ' +
      faltam.map(k => '"' + CAD_CABECALHOS[k] + '"').join(', ') +
      '. Alguém renomeou o cabeçalho? O texto precisa bater exatamente.');
  }
  return col;
}

/** Lê o CADASTRO -> [{empresa, folha, sped, contrib, ativa, linha}, ...]
 *  contrib (EFD Contribuições): true = faz, false = não faz, null = em branco. */
function lerCadastro() {
  const aba = planilha().getSheetByName('CADASTRO');
  const teto = Math.min(aba.getLastRow(), LIMITE_LINHAS);
  if (teto < LINHA_INICIAL) return [];
  const col = colunasCadastro();
  const dados = aba.getRange(LINHA_INICIAL, 1, teto - LINHA_INICIAL + 1, col.total).getValues();
  const campo = (l, c) => c ? String(l[c - 1]).trim().toUpperCase() : '';
  const ehNao = t => t === 'NÃO' || t === 'NAO';
  const lista = [];
  dados.forEach((l, i) => {
    if (l[0] === '' || l[0] === null) return;
    const contribTxt = campo(l, col.contrib);
    lista.push({
      empresa: l[0],
      folha: campo(l, col.folha) === 'SIM',
      sped:  campo(l, col.sped) === 'SIM',
      contrib: contribTxt === 'SIM' ? true : ehNao(contribTxt) ? false : null,
      ativa: !ehNao(campo(l, col.ativa)),                 // em branco = ativa
      linha: LINHA_INICIAL + i,
    });
  });
  return lista;
}

/** Valor padrão de FAZ EFD CONTRIB.? pelo regime: o Simples (e o MEI) é
 *  dispensado da EFD Contribuições; o Lucro Presumido entrega. */
function contribPadrao(regime) {
  const r = String(regime).trim();
  if (r === 'Lucro Presumido') return 'Sim';
  if (r === 'Simples Nacional' || r === 'Simples Híbrido' || r === 'MEI') return 'Não';
  return '';
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
  if (flag === 'sped') return empresa.sped || empresa.contrib === true;   // SPED fiscal ou só a EFD Contribuições
  return true;
}

/** Na 2. SPED, linha nova já nasce com "Não se aplica" na EFD que a empresa
 *  não faz: EFD ICMS/IPI quando ela entrou só pela EFD Contribuições
 *  (FAZ SPED? = Não), e EFD CONTRIBUIÇÕES quando FAZ EFD CONTRIB.? = Não.
 *  FAZ EFD CONTRIB.? em branco = como sempre foi (nada marcado). */
function marcarNaoSeAplica(aba, nomeAba, linha, empresas) {
  const cfg = ABAS_ETAPA[nomeAba];
  if (!cfg.colIcms) return;
  const icms    = empresas.map(e => [e.sped ? '' : NAO_SE_APLICA]);
  const contrib = empresas.map(e => [e.contrib === false ? NAO_SE_APLICA : '']);
  if (icms.some(v => v[0]))    aba.getRange(linha, cfg.colIcms, empresas.length, 1).setValues(icms);
  if (contrib.some(v => v[0])) aba.getRange(linha, cfg.colContrib, empresas.length, 1).setValues(contrib);
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

/** Refaz o filtro cobrindo todas as linhas de dados.
 *  Sem isso, meses novos não aparecem na lista do funil. */
function ajustarFiltro(aba, ultimaCol) {
  const fim = ultimaLinha(aba);
  const filtro = aba.getFilter();
  if (filtro) filtro.remove();
  if (fim < LINHA_INICIAL) return;
  aba.getRange(2, 1, fim - 1, ultimaCol).createFilter();
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
    marcarNaoSeAplica(aba, nomeAba, linha, doMes);

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
  let temContrib = false;     // o botão só aparece depois que a coluna existir
  try { temContrib = colunasCadastro().contrib > 0; } catch (e) { /* cabeçalho estranho: a gravação avisa */ }
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
  (temContrib ?
  '  <div><label>Faz EFD Contrib.?</label><div class="toggle">' +
  '    <button id="contribSim" onclick="setContrib(true)">Sim</button>' +
  '    <button id="contribNao" onclick="setContrib(false)">Não</button></div></div>' : '') +
  '</div>' +
  '<div class="aviso" id="dica">Ela entra nesse mês e em todos os meses já abertos depois dele. ' +
  (temContrib ? 'SPED = EFD ICMS/IPI. Quem faz só a EFD Contribuições: SPED = Não e EFD Contrib. = Sim. ' : '') +
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
  'var TEM_CONTRIB = ' + temContrib + ';' +
  'var fazFolha = false, fazSped = true, fazContrib = false;' +
  'function pinta() {' +
  '  document.getElementById("folhaSim").className = fazFolha ? "on" : "";' +
  '  document.getElementById("folhaNao").className = fazFolha ? "" : "off";' +
  '  document.getElementById("spedSim").className  = fazSped  ? "on" : "";' +
  '  document.getElementById("spedNao").className  = fazSped  ? "" : "off";' +
  '  if (!TEM_CONTRIB) return;' +
  '  document.getElementById("contribSim").className = fazContrib ? "on" : "";' +
  '  document.getElementById("contribNao").className = fazContrib ? "" : "off";' +
  '}' +
  'function setFolha(v)   { fazFolha   = v; pinta(); }' +
  'function setSped(v)    { fazSped    = v; pinta(); }' +
  'function setContrib(v) { fazContrib = v; pinta(); }' +
  // EFD Contribuições segue o regime (Lucro Presumido = Sim), mas dá para trocar
  'selReg.addEventListener("change", function () { setContrib(selReg.value === "Lucro Presumido"); });' +
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
  '      contrib: TEM_CONTRIB ? (fazContrib ? "Sim" : "Não") : "",' +
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

  const faltando = abasFaltando(abasNecessarias(['CADASTRO']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  if (acharEmpresa(nome)) return { ok: false, msg: '"' + nome + '" já está no cadastro.' };
  if (REGIMES.indexOf(d.regime) === -1) return { ok: false, msg: 'Regime inválido.' };

  // CADASTRO (colunas pelo cabeçalho)
  const cad = planilha().getSheetByName('CADASTRO');
  const col = colunasCadastro();
  const contrib = col.contrib ? (d.contrib === 'Sim' ? 'Sim' : d.contrib === 'Não' ? 'Não' : '') : '';
  const linhaCad = ultimaLinhaCol(cad, 'A') + 1;
  const nova = new Array(col.total).fill('');
  nova[0] = nome; nova[1] = d.cnpj; nova[2] = d.regime;
  nova[col.folha - 1] = d.folha; nova[col.sped - 1] = d.sped; nova[col.ativa - 1] = 'Sim';
  if (col.contrib) nova[col.contrib - 1] = contrib;
  cad.getRange(linhaCad, 1, 1, col.total).setValues([nova]);
  cad.getRange(linhaCad, 1, 1, col.total)
     .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
     .setHorizontalAlignment('center')
     .setBorder(true, true, true, true, true, true,
                '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  cad.getRange(linhaCad, 1).setFontWeight('bold').setHorizontalAlignment('left');

  // Em quais meses ela entra: do mês escolhido em diante, só os já abertos
  const iIni = d.mesInicial ? MESES.indexOf(String(d.mesInicial).trim().toUpperCase()) : -1;
  const empresa = { empresa: nome, folha: d.folha === 'Sim', sped: d.sped === 'Sim',
                    contrib: contrib === 'Sim' ? true : contrib === 'Não' ? false : null };
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
      marcarNaoSeAplica(aba, nomeAba, linha, alvos.map(() => empresa));

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

  const meses = MESES.filter(m => mesesUsados[m]);
  const onde = meses.length === 0
    ? 'Só no cadastro (nenhum mês)'
    : (meses.length === 1 ? 'Mês: ' + meses[0]
                          : 'Meses: ' + meses[0] + ' a ' + meses[meses.length - 1] +
                            ' (' + meses.length + ')');

  const total = lerCadastro().length;

  return { ok: true, msg: '✔ ' + nome + ' cadastrada (' + d.regime + ')' +
                          '\n' + onde +
                          (inseridas.length ? '\nAbas: ' + inseridas.join(', ') : '') +
                          '\nTotal no cadastro: ' + total };
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
  '<div class="sub">Troca o nome em tudo de uma vez: o cadastro e as 4 abas.</div>' +
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

/** Chamada pela janela. O nome é a chave que liga tudo: troca no CADASTRO
 *  e nas 4 abas (todos os meses) de uma vez. Devolve {ok, msg}.
 *  Renomear em 3 das 4 abas é pior que não renomear: as linhas da aba que
 *  faltou viram órfãs (perdem CNPJ/regime e somem da ficha da empresa). */
function executarRenomear(velho, novo) {
  velho = String(velho).trim();
  novo  = String(novo).trim();
  if (!velho) return { ok: false, msg: 'Escolha a empresa.' };
  if (!novo)  return { ok: false, msg: 'Digite o novo nome.' };

  const faltando = abasFaltando(abasNecessarias(['CADASTRO']));
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
            .getRange(emp.linha, colunasCadastro().ativa).setValue(novoAtiva ? 'Sim' : 'Não');

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
      : 'Ela não tem histórico em nenhum mês. Só sairá do cadastro.');

  return { existe: true, nome: emp.empresa, temHistorico: temHistorico, total: total, resumo: resumo };
}

/** Apaga a empresa e TODO o histórico dela. Se houver histórico, exige o nome
 *  digitado. Só para cadastro errado. Devolve {ok, msg}.
 *  Apagar de algumas abas e não de outras deixa histórico solto — por isso
 *  recusa na entrada se faltar alguma aba de contrato. */
function executarExcluirEmpresa(nome, confirmacao) {
  nome = String(nome).trim();

  const faltando = abasFaltando(abasNecessarias(['CADASTRO']));
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

/** Cria uma cópia do arquivo para o ano novo: mantém o CADASTRO e zera
 *  as abas de etapa. O arquivo atual não é alterado — ele fica como
 *  histórico do ano que passou. */
function virarOAno() {
  const ui = SpreadsheetApp.getUi();

  const r = ui.prompt('Virar o ano', 'Ano do arquivo NOVO (ex.: 2027):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const ano = r.getResponseText().trim();
  if (!/^\d{4}$/.test(ano)) { ui.alert('Digite um ano com 4 dígitos.'); return; }

  const conf = ui.alert('Virar o ano',
    'Vou criar uma CÓPIA chamada "CONTROLE EMPRESAS ' + ano + '".\n\n' +
    'Na cópia: o CADASTRO é mantido e as abas de etapa ficam\n' +
    'zeradas para começar o ano.\n\n' +
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
           '\n\nAbra ele e use 🧮 Modo Contador → Abrir novo mês → JANEIRO.\n\nLink:\n' + copia.getUrl());
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

    // trabalho feito = status marcado (fora "Pendente", "Não se aplica" e vazio)
    doMes.forEach(l => {
      for (let c = 4; c < cfg.ultimaCol; c++) {
        if (cfg.colTotal && c === cfg.colTotal - 1) continue;
        const v = String(l[c]).trim();
        if (v !== '' && v !== 'Pendente' && v !== NAO_SE_APLICA) preenchidos++;   // N/A já nasce marcado
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
    ['✏️', 'Renomear empresa', 'Troca o nome em tudo de uma vez (cadastro e as 4 abas).'],
    ['🔘', 'Ativar / desativar', 'Tira a empresa dos meses novos sem perder o histórico.'],
    ['📇', 'Ficha da empresa', 'O ano inteiro de uma empresa numa tela: status, faturamento e notas.'],
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
//  11. BLOCO DE NOTAS  (barra lateral)
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

// ============================================================
//  12. IMPORTAR EMPRESAS EM MASSA  (aba IMPORTAR + barra lateral)
// ============================================================
//  Você cola a lista na aba IMPORTAR (EMPRESA | CNPJ | REGIME | PERFIL |
//  IE | IM) e clica em Importar. O script cadastra todas de uma vez no
//  CADASTRO, pulando as que já existem.
//  Padrões: Ativa=Sim; Folha/SPED=Sim (MEI=Não); EFD Contrib.=Sim só no
//  Lucro Presumido; senha em branco.
// ============================================================

const IMPORTAR_ABA = 'IMPORTAR';
const IMPORTAR_CAB = ['EMPRESA', 'CNPJ', 'REGIME', 'PERFIL',
                      'INSCRIÇÃO ESTADUAL', 'INSCRIÇÃO MUNICIPAL'];

/** "simples nacional" -> "Simples Nacional", etc. Desconhecido passa como veio. */
function normalizarRegime(r) {
  const s = String(r).trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('hibrid') >= 0 || s.indexOf('híbrid') >= 0) return 'Simples Híbrido';
  if (s.indexOf('simples') >= 0) return 'Simples Nacional';
  if (s.indexOf('presumido') >= 0) return 'Lucro Presumido';
  if (s === 'mei') return 'MEI';
  return String(r).trim();
}

/** Cria a aba IMPORTAR (se não existir) com cabeçalho e colunas de texto. */
function garantirAbaImportar() {
  const ss = planilha();
  let aba = ss.getSheetByName(IMPORTAR_ABA);
  if (aba) return aba;
  aba = ss.insertSheet(IMPORTAR_ABA);
  aba.getRange('A1:F1000').setNumberFormat('@');   // tudo texto: preserva CNPJ/IE
  aba.getRange(1, 1, 1, IMPORTAR_CAB.length).setValues([IMPORTAR_CAB])
     .setFontWeight('bold').setBackground(COR_NAVY).setFontColor('#ffffff');
  aba.setColumnWidth(1, 240); aba.setColumnWidth(2, 150);
  aba.setColumnWidth(5, 150); aba.setColumnWidth(6, 150);
  aba.setFrozenRows(1);
  return aba;
}

/** Lê a aba IMPORTAR -> [{nome,cnpj,regime,perfil,ie,im}]. */
function lerImportar() {
  const aba = garantirAbaImportar();
  const fim = aba.getLastRow();
  if (fim < 2) return [];
  const linhas = [];
  aba.getRange(2, 1, fim - 1, 6).getValues().forEach(l => {
    const nome = String(l[0]).trim();
    if (!nome) return;
    linhas.push({
      nome: nome, cnpj: String(l[1]).trim(), regime: normalizarRegime(l[2]),
      perfil: String(l[3]).trim(), ie: String(l[4]).trim(), im: String(l[5]).trim(),
    });
  });
  return linhas;
}

/** Confere a aba IMPORTAR antes de cadastrar. Devolve contagens + prévia. */
function analisarImportacao() {
  const faltando = abasFaltando(['CADASTRO']);
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const linhas = lerImportar();
  if (linhas.length === 0) {
    return { ok: true, vazio: true,
      msg: 'A aba IMPORTAR está vazia.\n\nCole sua lista lá (a partir da linha 2) e clique em Conferir de novo.' };
  }

  const existentes = {};
  lerCadastro().forEach(e => { existentes[String(e.empresa).trim().toUpperCase()] = true; });

  let novas = 0, duplicadas = 0;
  const vistos = {}, previa = [];
  linhas.forEach(l => {
    const chave = l.nome.toUpperCase();
    if (existentes[chave] || vistos[chave]) { duplicadas++; return; }
    vistos[chave] = true; novas++;
    if (previa.length < 5) previa.push(l.nome + ' — ' + (l.regime || '(sem regime)'));
  });

  return { ok: true, vazio: false, total: linhas.length, novas: novas, duplicadas: duplicadas, previa: previa };
}

/** Cadastro em massa: cada nova empresa entra no CADASTRO.
 *  Pula duplicadas. Devolve {ok, msg}. */
function executarImportacao() {
  const faltando = abasFaltando(['CADASTRO']);
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const linhas = lerImportar();
  if (linhas.length === 0) return { ok: false, msg: 'A aba IMPORTAR está vazia.' };

  const cad = planilha().getSheetByName('CADASTRO');
  const col = colunasCadastro();
  const existentes = {};
  lerCadastro().forEach(e => { existentes[String(e.empresa).trim().toUpperCase()] = true; });

  let cadastradas = 0, duplicadas = 0;
  const vistos = {};

  linhas.forEach(l => {
    const chave = l.nome.toUpperCase();
    if (existentes[chave] || vistos[chave]) { duplicadas++; return; }
    vistos[chave] = true;

    const ehMei = (l.regime === 'MEI') || (l.perfil.toUpperCase() === 'MEI');
    const nova = new Array(col.total).fill('');
    nova[0] = l.nome; nova[1] = l.cnpj; nova[2] = l.regime;
    nova[col.folha - 1] = ehMei ? 'Não' : 'Sim';
    nova[col.sped - 1]  = ehMei ? 'Não' : 'Sim';
    nova[col.ativa - 1] = 'Sim';
    if (col.contrib) nova[col.contrib - 1] = ehMei ? 'Não' : contribPadrao(l.regime);
    if (col.perfil)  nova[col.perfil - 1] = l.perfil;
    if (col.ie)      nova[col.ie - 1] = l.ie;
    if (col.im)      nova[col.im - 1] = l.im;

    const linhaCad = ultimaLinhaCol(cad, 'A') + 1;
    // texto ANTES de gravar: depois o Google já teria comido o zero à esquerda
    [col.cnpj, col.ie, col.senha].filter(c => c).forEach(c => cad.getRange(linhaCad, c).setNumberFormat('@'));
    cad.getRange(linhaCad, 1, 1, col.total).setValues([nova]);
    cad.getRange(linhaCad, 1, 1, col.total)
       .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
       .setHorizontalAlignment('center')
       .setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
    cad.getRange(linhaCad, 1).setFontWeight('bold').setHorizontalAlignment('left');

    cadastradas++;
  });

  // limpa os dados da aba IMPORTAR (mantém o cabeçalho)
  const imp = garantirAbaImportar();
  const fim = imp.getLastRow();
  if (fim >= 2) imp.getRange(2, 1, fim - 1, 6).clearContent();

  let msg = '✔ Importação concluída!\n\n' +
    '• Cadastradas: ' + cadastradas + '\n' +
    '• Já existiam (puladas): ' + duplicadas;
  msg +='\n\nPara colocá-las num mês, use "Abrir novo mês" — ele puxa todas as ativas.';
  return { ok: true, msg: msg };
}

/** Abre a barra lateral de importação e leva você pra aba IMPORTAR. */
function dialogoImportar() {
  planilha().setActiveSheet(garantirAbaImportar());
  const html = HtmlService.createHtmlOutput(htmlImportar()).setTitle('Importar empresas');
  SpreadsheetApp.getUi().showSidebar(html);
}

function htmlImportar() {
  return '<!DOCTYPE html><html><head><base target="_top">' + estiloDialogo() + '</head><body>' +
  '<div class="sub">Cole sua lista na aba <b>IMPORTAR</b> (já aberta ao lado), a partir da linha 2 — colunas: EMPRESA · CNPJ · REGIME · PERFIL · IE · IM. Depois clique em Conferir.</div>' +
  '<div class="aviso">Padrões: Ativa = Sim · Folha/SPED = Sim (MEI = Não) · EFD Contrib. = Sim só no Lucro Presumido · senha em branco · regime normalizado. Empresas que já existem são puladas.</div>' +
  '<div class="botoes">' +
  '  <button class="pri" id="btnVer" onclick="conferir()">Conferir</button>' +
  '</div>' +
  '<div id="info"></div>' +
  '<div id="bloco" style="display:none">' +
  '  <div class="botoes"><button class="pri" id="btnImp" onclick="importar()">Importar</button></div>' +
  '</div>' +
  '<div id="status"></div>' +
  '<script>' +
  'function msg(t,c){var s=document.getElementById("status");s.className=c;s.textContent=t;}' +
  'function conferir(){' +
  '  document.getElementById("btnVer").disabled=true;' +
  '  document.getElementById("bloco").style.display="none";' +
  '  document.getElementById("status").className="";document.getElementById("status").textContent="";' +
  '  google.script.run.withSuccessHandler(function(r){' +
  '    document.getElementById("btnVer").disabled=false;' +
  '    if(!r.ok){ document.getElementById("info").innerHTML=\'<div class="alerta">\'+r.msg+"</div>"; return; }' +
  '    if(r.vazio){ document.getElementById("info").innerHTML=\'<div class="caixa">\'+r.msg+"</div>"; return; }' +
  '    var prev = r.previa.length ? "\\n\\nPrimeiras:\\n• "+r.previa.join("\\n• ") : "";' +
  '    document.getElementById("info").innerHTML=\'<div class="caixa">\'+("Encontrei "+r.total+" linha(s):\\n• Novas: "+r.novas+"\\n• Já existem (puladas): "+r.duplicadas+prev)+"</div>";' +
  '    if(r.novas>0){ document.getElementById("bloco").style.display="block"; document.getElementById("btnImp").textContent="Importar "+r.novas+" empresa(s)"; }' +
  '  }).withFailureHandler(function(e){ document.getElementById("btnVer").disabled=false; msg("Erro: "+e.message,"erro"); }).analisarImportacao();' +
  '}' +
  'function importar(){' +
  '  document.getElementById("btnImp").disabled=true;' +
  '  msg("Importando, aguarde...","load");' +
  '  google.script.run.withSuccessHandler(function(r){' +
  '    document.getElementById("btnImp").disabled=false;' +
  '    msg(r.msg, r.ok?"ok":"erro");' +
  '    if(r.ok){ document.getElementById("bloco").style.display="none"; document.getElementById("info").innerHTML=""; }' +
  '  }).withFailureHandler(function(e){ document.getElementById("btnImp").disabled=false; msg("Erro: "+e.message,"erro"); }).executarImportacao();' +
  '}' +
  '</script></body></html>';
}

// ============================================================
//  13. FICHA DA EMPRESA  (o ano inteiro de uma empresa numa janela)
// ============================================================
//  Junta numa tela só o que fica espalhado em 4 abas + CADASTRO + NOTAS:
//  dados do cadastro, o status de cada etapa mês a mês, o faturamento e
//  as anotações da empresa. Só lê — não altera nada na planilha.
//  Se o cursor estiver numa linha de empresa ao abrir, ela já vem escolhida.
// ============================================================

// status -> [fundo, texto, rótulo curto]. Mesmas cores da formatação
// condicional das abas de etapa, para a ficha "bater" com a planilha.
const FICHA_STATUS = {
  'Concluído':     ['#C6EFCE', '#006100', '✔'],
  'Importada':     ['#C6EFCE', '#006100', 'Imp.'],
  'Retificada':    ['#BDD7EE', '#1F3864', 'Retif.'],
  'Baixada':       ['#FFF2CC', '#7F6000', 'Baix.'],
  'Agendada':      ['#FFF2CC', '#7F6000', 'Agend.'],
  'Pendente':      ['#FFEB9C', '#9C6500', 'Pend.'],
  'Sem movimento': ['#E7E6E6', '#595959', 'S/M'],
  'Não se aplica': ['#F2F2F2', '#808080', 'N/A'],
  'Erro':          ['#FFC7CE', '#9C0006', 'Erro'],
};

function dialogoFicha() {
  const html = HtmlService.createHtmlOutput(htmlFicha(empresaDaSelecao()))
                          .setWidth(940).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Ficha da empresa');
}

/** Empresa da linha onde está o cursor (abas de etapa ou CADASTRO), para
 *  a ficha já abrir nela. '' se não der para saber. */
function empresaDaSelecao() {
  try {
    const aba = planilha().getActiveSheet();
    const nomeAba = aba.getName();
    const col = nomeAba === 'CADASTRO' ? 1 : ABAS_ETAPA[nomeAba] ? 2 : 0;
    const linha = aba.getActiveCell().getRow();
    if (!col || linha < LINHA_INICIAL) return '';
    const emp = acharEmpresa(aba.getRange(linha, col).getValue());
    return emp ? String(emp.empresa) : '';
  } catch (e) { return ''; }
}

/** Célula -> valor que a janela mostra. google.script.run não transporta
 *  Date, então data vira texto aqui. */
function celulaFicha(v, tipo) {
  if (tipo === 'moeda') return typeof v === 'number' ? v : '';
  if (v instanceof Date) return ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2);
  return v === null || v === undefined ? '' : String(v).trim();
}

/** Tudo o que a ficha mostra de uma empresa. Lê cada aba uma vez só.
 *  Nas grades: null = a empresa não tem linha naquele mês (mês não aberto
 *  ou ela não entra na etapa); '' = a linha existe mas está em branco. */
function fichaEmpresa(nome) {
  const faltando = abasFaltando(abasNecessarias(['CADASTRO']));
  if (faltando.length > 0) return { ok: false, msg: msgAbasFaltando(faltando) };

  const emp = acharEmpresa(nome);
  if (!emp) return { ok: false, msg: 'Não achei "' + nome + '" no CADASTRO.' };
  const alvo = String(emp.empresa).trim().toUpperCase();
  const txt = v => (v === null || v === undefined) ? '' : String(v).trim();

  // linha do CADASTRO (colunas pelo cabeçalho) — a SENHA fica de fora de propósito
  const col = colunasCadastro();
  const cad = planilha().getSheetByName('CADASTRO').getRange(emp.linha, 1, 1, col.total).getValues()[0];
  const doCad = c => c ? txt(cad[c - 1]) : '';

  const abertos = MESES.map(() => false);
  const etapas = [];
  let faturamento = MESES.map(() => null);
  let emAberto = 0;

  for (const nomeAba in ABAS_ETAPA) {
    const cfg = ABAS_ETAPA[nomeAba];
    const aba = planilha().getSheetByName(nomeAba);
    const cab = aba.getRange(2, 1, 1, cfg.ultimaCol).getValues()[0];
    const fim = ultimaLinha(aba);
    const dados = fim < LINHA_INICIAL ? []
      : aba.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, cfg.ultimaCol).getValues();

    const porMes = MESES.map(() => null);          // linha da empresa em cada mês
    dados.forEach(l => {
      const m = MESES.indexOf(String(l[0]).trim().toUpperCase());
      if (m === -1) return;
      abertos[m] = true;
      if (!porMes[m] && String(l[1]).trim().toUpperCase() === alvo) porMes[m] = l;
    });

    const colunas = [];
    for (let c = 5; c <= cfg.ultimaCol; c++) {       // da coluna E em diante
      const tipo = (cfg.colsMoeda || []).indexOf(c) !== -1 ? 'moeda'
                 : c === cfg.colData ? 'data' : 'status';
      colunas.push({
        nome: txt(cab[c - 1]),
        tipo: tipo,
        principal: c === cfg.colFeito,
        valores: porMes.map(l => l ? celulaFicha(l[c - 1], tipo) : null),
      });
    }

    // mesma regra do e-mail de pendências: Pendente, Erro ou em branco
    porMes.forEach(l => {
      if (!l) return;
      const st = String(l[cfg.colFeito - 1]).trim();
      if (st === 'Pendente' || st === 'Erro' || st === '') emAberto++;
    });
    if (cfg.colTotal) faturamento = porMes.map(l => l ? celulaFicha(l[cfg.colTotal - 1], 'moeda') : null);

    etapas.push({ nome: nomeAba, entra: entraNaEtapa(emp, cfg.flag), colunas: colunas });
  }

  const lancados = faturamento.filter(v => typeof v === 'number');
  const totalAno = lancados.reduce((s, v) => s + v, 0);

  // notas: só lê se a aba existir (listarNotas criaria a aba NOTAS)
  let notas = { pendentes: [], avisos: [], concluidas: [] };
  if (planilha().getSheetByName(NOTAS_ABA)) {
    const todas = listarNotas();
    const daEmpresa = n => n.empresa.toUpperCase() === alvo;
    notas = { pendentes:  todas.pendentes.filter(daEmpresa),
              avisos:     todas.avisos.filter(daEmpresa),
              concluidas: todas.concluidas.filter(daEmpresa) };
  }

  return {
    ok: true,
    empresa: String(emp.empresa), ativa: emp.ativa, folha: emp.folha, sped: emp.sped,
    contrib: emp.contrib, temContrib: col.contrib > 0,
    cnpj: doCad(col.cnpj), regime: doCad(col.regime), perfil: doCad(col.perfil),
    ie: doCad(col.ie), im: doCad(col.im),
    meses: MESES, abertos: abertos, etapas: etapas,
    faturamento: faturamento, totalAno: totalAno, mesesLancados: lancados.length,
    media: lancados.length ? totalAno / lancados.length : 0,
    emAberto: emAberto, notas: notas,
  };
}

function htmlFicha(escolhida) {
  const json = v => JSON.stringify(v).replace(/</g, '\\u003c');
  return `<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">${estiloDialogo()}
<style>
body { padding: 14px 18px; }
.topo { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
.topo select { width: 360px; }
.topo .dica { color: #808080; font-size: 11px; }
.cab { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.cab .nome { font-size: 17px; font-weight: bold; color: ${COR_NAVY}; }
.tag { font-size: 10.5px; font-weight: bold; padding: 3px 8px; border-radius: 10px; background: #EEF2FA; color: ${COR_AZUL}; }
.tag.ok { background: #C6EFCE; color: #006100; }
.tag.off { background: #E7E6E6; color: #595959; }
.dados { display: flex; gap: 18px; flex-wrap: wrap; font-size: 12px; color: #595959; margin: 6px 0 12px; }
.dados b { color: ${COR_NAVY}; margin-right: 3px; }
.kpis { display: flex; gap: 10px; margin-bottom: 14px; }
.kpi { flex: 1; background: #F7F9FC; border: 1px solid #E3E8F2; border-radius: 8px; padding: 9px 12px; }
.kpi .r { font-size: 11px; color: #595959; }
.kpi .v { font-size: 17px; font-weight: bold; color: ${COR_NAVY}; margin: 2px 0; font-variant-numeric: tabular-nums; }
.kpi .s { font-size: 10.5px; color: #808080; }
.kpi.ruim { background: #FFF1F2; border-color: #F4B6BD; } .kpi.ruim .v { color: #9C0006; }
.kpi.bom .v { color: #006100; }
.kpi.atencao .v { color: #9C6500; }
h4 { margin: 14px 0 4px; color: ${COR_NAVY}; font-size: 13px; }
.dicaGrade { font-size: 11px; color: #808080; margin-bottom: 6px; }
table.grade { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11.5px; }
.grade th { background: ${COR_NAVY}; color: #fff; padding: 6px 2px; font-size: 10.5px; }
.grade th.rot { text-align: left; padding-left: 8px; width: 196px; }
.grade th.fechado { background: #8A97A8; }
.grade td { text-align: center; padding: 4px 2px; border-bottom: 1px solid #EEF1F7; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.grade td.rot { text-align: left; padding-left: 8px; }
.grade tr.grupo td { border-top: 1px solid #D9E1F2; }
.grade tr.grupo td.rot { font-weight: bold; color: ${COR_NAVY}; cursor: pointer; }
.grade tr.grupo:hover td { background: #FBFCFE; }
.grade .seta { display: inline-block; width: 12px; color: ${COR_AZUL}; }
.grade .col { font-weight: normal; color: #808080; font-size: 10.5px; margin-left: 5px; }
.grade tr.sub td { font-size: 11px; background: #FAFBFD; }
.grade tr.sub td.rot { padding-left: 26px; color: #595959; }
.grade tr.fat td { border-top: 2px solid #C9D2E3; font-weight: bold; color: ${COR_NAVY}; }
.grade td.naoentra { color: #A0A0A0; font-style: italic; text-align: left; padding-left: 8px; }
.grade tr.off td.rot { color: #A0A0A0; }
.st { display: inline-block; min-width: 38px; padding: 2px 4px; border-radius: 4px; font-size: 10.5px; font-weight: bold; background: #EEF1F7; color: #3c4757; }
.nada { color: #D0D0D0; }
.branco { color: #A0A0A0; }
.num { font-variant-numeric: tabular-nums; }
.legenda { display: flex; flex-wrap: wrap; gap: 6px 12px; font-size: 10.5px; color: #595959; margin-top: 8px; }
.legenda .st { min-width: 0; margin-right: 3px; }
.nota { background: #fff; border: 1px solid #E4E9F0; border-left: 4px solid #ccc; border-radius: 8px; padding: 8px 11px; margin-bottom: 7px; }
.nota .top { display: flex; justify-content: space-between; font-size: 10.5px; margin-bottom: 3px; }
.nota .prio { font-weight: bold; text-transform: uppercase; letter-spacing: .3px; }
.nota .date { color: #8A97A8; }
.nota .desc { font-size: 12.5px; white-space: pre-wrap; word-break: break-word; }
.nota.done .desc { text-decoration: line-through; color: #9AA7B8; }
.vermais { font-size: 11.5px; color: ${COR_AZUL}; cursor: pointer; margin: 4px 0 8px; }
.vazio { color: #9AA7B8; font-size: 12px; padding: 6px 0; }
.inicio { text-align: center; color: #9AA7B8; font-size: 13px; padding: 90px 0; }
</style></head><body>
<div class="topo">
  <select id="emp"></select>
  <span class="dica">Dica: com o cursor na linha de uma empresa, a ficha já abre nela.</span>
</div>
<div id="conteudo"></div>
<div class="botoes"><button class="pri" onclick="google.script.host.close()">Fechar</button></div>
<script>
var EMPRESAS = ${json(listarEmpresas())};
var INICIAL = ${json(escolhida)};
var STATUS = ${json(FICHA_STATUS)};
var MES_CURTO = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
var PRIO = { 'Alta':{c:'#e0362c',t:'#c0281f'}, 'Média':{c:'#f0a400',t:'#9a6b00'}, 'Baixa':{c:'#4a89dc',t:'#2f6cb5'} };
var AVISO = { c:'#8250df', t:'#6f42c1' }, FEITA = { c:'#34a853', t:'#1e7e34' };
var pedido = 0;   // descarta resposta atrasada se você trocar de empresa rápido

function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function reais(v){ return 'R$ ' + Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function curto(v){
  var a = Math.abs(v);
  if (a >= 1e6) return (v/1e6).toLocaleString('pt-BR',{maximumFractionDigits:1}) + ' mi';
  if (a >= 1e3) return (v/1e3).toLocaleString('pt-BR',{maximumFractionDigits:1}) + ' mil';
  return v.toLocaleString('pt-BR',{maximumFractionDigits:0});
}

function celula(v, tipo){
  if (v === null) return '<td class="nada" title="sem linha neste mês">—</td>';
  if (v === '') return '<td class="branco" title="em branco">·</td>';
  if (tipo === 'moeda') return '<td class="num" title="' + reais(v) + '">' + curto(v) + '</td>';
  if (tipo === 'data') return '<td>' + esc(v) + '</td>';
  var s = STATUS[v];
  if (!s) return '<td title="' + esc(v) + '"><span class="st">' + esc(v) + '</span></td>';
  return '<td title="' + esc(v) + '"><span class="st" style="background:' + s[0] + ';color:' + s[1] + '">' + s[2] + '</span></td>';
}

function cabecalho(f){
  var tags = '<span class="tag">' + esc(f.regime || 'sem regime') + '</span>' +
    (f.perfil ? '<span class="tag">Perfil ' + esc(f.perfil) + '</span>' : '') +
    (f.ativa ? '<span class="tag ok">ATIVA</span>' : '<span class="tag off">INATIVA</span>');
  var dados = [['CNPJ', f.cnpj], ['IE', f.ie], ['IM', f.im], ['Folha', f.folha ? 'Sim' : 'Não'], ['SPED', f.sped ? 'Sim' : 'Não']];
  if (f.temContrib) dados.push(['EFD Contrib.', f.contrib === true ? 'Sim' : f.contrib === false ? 'Não' : '']);
  dados = dados.map(function(d){ return '<span><b>' + d[0] + '</b>' + esc(d[1] || '—') + '</span>'; }).join('');
  return '<div class="cab"><span class="nome">' + esc(f.empresa) + '</span>' + tags + '</div><div class="dados">' + dados + '</div>';
}

function kpi(rot, val, cls, sub){
  return '<div class="kpi ' + cls + '"><div class="r">' + rot + '</div><div class="v">' + val + '</div><div class="s">' + sub + '</div></div>';
}
function kpis(f){
  var np = f.notas.pendentes.length;
  return '<div class="kpis">' +
    kpi('Em aberto no ano', f.emAberto, f.emAberto ? 'ruim' : 'bom', 'Pendente, Erro ou em branco') +
    kpi('Faturamento no ano', reais(f.totalAno), '', f.mesesLancados + ' mês(es) lançado(s)') +
    kpi('Média por mês', reais(f.media), '', 'dos meses lançados') +
    kpi('Notas pendentes', np, np ? 'atencao' : '', f.notas.avisos.length + ' aviso(s)') +
    '</div>';
}

function grade(f){
  var h = '<table class="grade"><thead><tr><th class="rot">Etapa</th>';
  f.meses.forEach(function(m, i){
    h += '<th class="' + (f.abertos[i] ? '' : 'fechado') + '" title="' + m + (f.abertos[i] ? '' : ' — ainda não aberto') + '">' + MES_CURTO[i] + '</th>';
  });
  h += '</tr></thead><tbody>';
  f.etapas.forEach(function(e, k){
    var principal = e.colunas.filter(function(c){ return c.principal; })[0];
    var outras = e.colunas.filter(function(c){ return !c.principal; });
    var temLinha = principal.valores.some(function(v){ return v !== null; });
    if (!e.entra && !temLinha) {
      h += '<tr class="grupo off"><td class="rot">' + esc(e.nome) + '</td><td colspan="12" class="naoentra">não entra nesta etapa (pelo cadastro)</td></tr>';
      return;
    }
    h += '<tr class="grupo" onclick="abre(' + k + ')" title="Clique para ver as colunas desta etapa"><td class="rot"><span class="seta" id="seta' + k + '">▸</span>' +
         esc(e.nome) + '<span class="col">' + esc(principal.nome.toLowerCase()) + '</span></td>';
    principal.valores.forEach(function(v){ h += celula(v, principal.tipo); });
    h += '</tr>';
    outras.forEach(function(c){
      h += '<tr class="sub s' + k + '" style="display:none"><td class="rot" title="' + esc(c.nome) + '">' + esc(c.nome) + '</td>';
      c.valores.forEach(function(v){ h += celula(v, c.tipo); });
      h += '</tr>';
    });
  });
  h += '<tr class="fat"><td class="rot">Faturamento (total)</td>';
  f.faturamento.forEach(function(v){ h += celula(v, 'moeda'); });
  return h + '</tr></tbody></table>';
}

function legenda(){
  var h = '<div class="legenda">';
  for (var nome in STATUS) {
    var s = STATUS[nome];
    h += '<span><span class="st" style="background:' + s[0] + ';color:' + s[1] + '">' + s[2] + '</span>' + esc(nome) + '</span>';
  }
  return h + '<span><b class="nada">—</b> sem linha no mês</span><span><b class="branco">·</b> em branco</span></div>';
}

function cartao(n, cor, rot, feita){
  return '<div class="nota' + (feita ? ' done' : '') + '" style="border-left-color:' + cor.c + '">' +
    '<div class="top"><span class="prio" style="color:' + cor.t + '">' + esc(rot) + '</span><span class="date">' + esc(n.data) + '</span></div>' +
    '<div class="desc">' + esc(n.descricao) + '</div></div>';
}
function notas(f){
  var n = f.notas;
  var h = '<h4>Notas da empresa</h4>';
  var itens = n.pendentes.map(function(x){ return cartao(x, PRIO[x.prioridade] || PRIO['Média'], x.prioridade); })
    .concat(n.avisos.map(function(x){ return cartao(x, AVISO, 'Aviso'); }));
  if (!itens.length && !n.concluidas.length)
    return h + '<div class="vazio">Nenhuma nota ligada a esta empresa. No Bloco de notas, escolha a empresa ao criar a nota.</div>';
  h += itens.length ? itens.join('') : '<div class="vazio">Nenhuma nota pendente.</div>';
  if (n.concluidas.length) {
    h += '<div class="vermais" onclick="var c=document.getElementById(\\'conc\\');c.style.display=c.style.display===\\'none\\'?\\'block\\':\\'none\\'">ver / esconder ' +
         n.concluidas.length + ' concluída(s)</div><div id="conc" style="display:none">' +
         n.concluidas.map(function(x){ return cartao(x, FEITA, 'Concluída', true); }).join('') + '</div>';
  }
  return h;
}

function abre(k){
  var linhas = document.querySelectorAll('.s' + k);
  if (!linhas.length) return;
  var mostrar = linhas[0].style.display === 'none';
  for (var i = 0; i < linhas.length; i++) linhas[i].style.display = mostrar ? '' : 'none';
  document.getElementById('seta' + k).textContent = mostrar ? '▾' : '▸';
}

function carregar(nome){
  var meu = ++pedido;
  var alvo = document.getElementById('conteudo');
  if (!nome) { alvo.innerHTML = '<div class="inicio">Escolha uma empresa na lista acima.</div>'; return; }
  alvo.innerHTML = '<div id="status" class="load">Montando a ficha de ' + esc(nome) + '...</div>';
  google.script.run.withSuccessHandler(function(f){
    if (meu !== pedido) return;
    if (!f.ok) { alvo.innerHTML = '<div class="alerta">' + esc(f.msg) + '</div>'; return; }
    alvo.innerHTML = cabecalho(f) + kpis(f) +
      '<h4>Status por mês</h4><div class="dicaGrade">Cada linha mostra a coluna final da etapa. Clique na etapa para ver as outras colunas.</div>' +
      grade(f) + legenda() + notas(f);
  }).withFailureHandler(function(e){
    if (meu === pedido) alvo.innerHTML = '<div class="alerta">Erro: ' + esc(e.message) + '</div>';
  }).fichaEmpresa(nome);
}

var sel = document.getElementById('emp');
var o0 = document.createElement('option'); o0.value = ''; o0.text = '— escolha a empresa —'; sel.add(o0);
EMPRESAS.forEach(function(e){
  var o = document.createElement('option'); o.value = e.nome; o.text = e.nome + (e.ativa ? '' : '  (inativa)'); sel.add(o);
});
sel.addEventListener('change', function(){ carregar(sel.value); });
sel.value = INICIAL;
carregar(INICIAL);
</script>
</body></html>`;
}

// ============================================================
//  14. MANUTENÇÃO — limpezas de uma vez só (rodar pelo editor)
// ============================================================
//  Funções que ajustam a planilha quando o script muda (tirar o que saiu,
//  criar coluna nova). Rode cada uma UMA VEZ pelo editor do Apps Script
//  (escolha o nome dela ao lado do botão "Executar"). Depois de rodadas,
//  podem ficar aqui sem problema.
//
//  removerFaturamentoDoComparativo: a Ficha da empresa substituiu a parte
//  de faturamento do COMPARATIVO. Ela apaga o bloco inteiro de faturamento —
//  título, linhas de empresa e TOTAL GERAL — e deixa só a tabela de
//  PENDÊNCIAS. Nada digitado se perde: o bloco era só fórmula somando a aba
//  3. FATURAMENTO. Funciona com a tabela original e também se as linhas de
//  empresa já tiverem sido tiradas antes. Rodar de novo não faz nada.
// ============================================================

function removerFaturamentoDoComparativo() {
  const aba = planilha().getSheetByName('COMPARATIVO');
  if (!aba) throw new Error('Não achei a aba COMPARATIVO. Nada foi apagado.');

  const ultCol = aba.getLastColumn();
  const colB = aba.getRange(1, 2, aba.getLastRow(), 1).getValues()
                  .map(l => String(l[0]).trim().toUpperCase());
  const linhaTotal = colB.indexOf('TOTAL GERAL') + 1;
  const linhaCab   = colB.indexOf('EMPRESA') + 1;
  if (!linhaTotal) {
    if (linhaCab) throw new Error('Achei o cabeçalho "EMPRESA" mas não a linha "TOTAL GERAL" — layout inesperado. Nada foi apagado.');
    console.log('Nada a fazer: o faturamento já foi removido do COMPARATIVO.');
    return;
  }

  // do título do bloco ("FATURAMENTO POR EMPRESA..." ou "FATURAMENTO TOTAL...")
  // até a linha antes das PENDÊNCIAS — leva junto o espaço em branco do meio
  const linhaTitulo = colB.findIndex((t, i) => i + 1 < linhaTotal && t.indexOf('FATURAMENTO') === 0) + 1;
  const linhaPend   = colB.findIndex((t, i) => i + 1 > linhaTotal && t.indexOf('PENDÊNCIAS') === 0) + 1;
  const ini = linhaTitulo || linhaCab || linhaTotal;
  const fim = linhaPend ? linhaPend - 1 : linhaTotal;

  // trava: nada digitado à mão pode ir junto. Liberados só o título, o
  // cabeçalho dos meses e a coluna B (nomes) até o TOTAL GERAL; o resto
  // tem que ser fórmula ou vazio.
  const faixa = aba.getRange(ini, 1, fim - ini + 1, ultCol);
  const formulas = faixa.getFormulas(), valores = faixa.getValues();
  for (let i = 0; i < valores.length; i++) {
    const r = ini + i;
    if (r === linhaTitulo) continue;
    if (MESES.every((m, j) => String(valores[i][2 + j]).trim().toUpperCase() === m)) continue;   // cabeçalho
    for (let c = 0; c < ultCol; c++) {
      if (c === 1 && r <= linhaTotal) continue;
      if (!formulas[i][c] && valores[i][c] !== '' && valores[i][c] !== null) {
        throw new Error('A célula ' + letraColuna(c + 1) + r + ' do COMPARATIVO tem um valor digitado à mão (' +
                        valores[i][c] + '). Por segurança nada foi apagado — confira essa célula.');
      }
    }
  }

  aba.setFrozenRows(0);                  // o congelamento era do cabeçalho do faturamento
  aba.deleteRows(ini, fim - ini + 1);

  console.log('Pronto: ' + (fim - ini + 1) + ' linha(s) removidas — o faturamento saiu do COMPARATIVO. ' +
              'A tabela de PENDÊNCIAS ficou.');
}

/** Tira o calendário de vencimentos do PAINEL (o recurso saiu do script).
 *  Rode UMA VEZ pelo editor, como a função acima. Limpa só as células do
 *  calendário (H7:N14) e da legenda (H16:N16) — o resto do PAINEL, como o
 *  ícone e o texto do Bloco de notas logo abaixo, fica intacto. A aba oculta
 *  VENCIMENTOS não é apagada (tem os prazos que você digitou).
 *  Rodar de novo não faz nada. */
function removerCalendario() {
  const painel = planilha().getSheetByName('PAINEL');
  if (!painel) throw new Error('Não achei a aba PAINEL. Nada foi apagado.');

  // assinatura do calendário: SEG..DOM em H8:N8. Sem ela, não mexe em nada.
  const cab = painel.getRange('H8:N8').getValues()[0].map(v => String(v).trim().toUpperCase());
  if (cab.join(',') !== 'SEG,TER,QUA,QUI,SEX,SÁB,DOM') {
    console.log('Nada a fazer: não achei o calendário no PAINEL.');
    return;
  }

  painel.getRange('H7:N14').breakApart().clearContent().clearFormat();   // título, dias e 6 semanas
  painel.getRange('H16:N16').breakApart().clearContent().clearFormat();  // legenda
  PropertiesService.getDocumentProperties().deleteProperty('CALENDARIO');

  console.log('Pronto: o calendário saiu do PAINEL.');
}

/** Cria a coluna "FAZ EFD CONTRIB.?" no CADASTRO, logo depois de "FAZ SPED?",
 *  com a mesma lista Sim/Não e as mesmas cores, e já preenche pelo regime:
 *  Lucro Presumido = Sim; Simples, Simples Híbrido e MEI = Não (o Simples é
 *  dispensado da EFD Contribuições). Confira e ajuste o que precisar.
 *  Rode UMA VEZ pelo editor. Rodar de novo não faz nada. */
function adicionarColunaEfdContrib() {
  const cad = planilha().getSheetByName('CADASTRO');
  if (!cad) throw new Error('Não achei a aba CADASTRO. Nada foi alterado.');
  const antes = colunasCadastro();
  if (antes.contrib) {
    console.log('Nada a fazer: a coluna "' + CAD_CABECALHOS.contrib + '" já existe (coluna ' +
                letraColuna(antes.contrib) + ').');
    return;
  }

  const ref = antes.sped, nova = antes.sped + 1;
  cad.insertColumnAfter(ref);
  const linhas = cad.getMaxRows() - 1;                     // da linha 2 (cabeçalho) até o fim
  const origem = cad.getRange(2, ref, linhas, 1), destino = cad.getRange(2, nova, linhas, 1);
  origem.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  origem.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  cad.setColumnWidth(nova, cad.getColumnWidth(ref));
  cad.getRange(2, nova).setValue(CAD_CABECALHOS.contrib);

  // cores Sim/Não: se a cópia de formato não trouxe, copia as regras da FAZ SPED?
  const regras = cad.getConditionalFormatRules();
  const cobre = c => regras.some(r => r.getRanges().some(g => g.getColumn() <= c && g.getLastColumn() >= c));
  if (!cobre(nova)) {
    const faixa = cad.getRange(LINHA_INICIAL, nova, cad.getMaxRows() - LINHA_INICIAL + 1, 1);
    const copias = regras
      .filter(r => r.getRanges().some(g => g.getColumn() === ref && g.getLastColumn() === ref))
      .map(r => r.copy().setRanges([faixa]).build());
    cad.setConditionalFormatRules(regras.concat(copias));
  }

  // preenche pelo regime e junta a lista para conferir
  const fim = ultimaLinhaCol(cad, 'A');
  let sim = 0, nao = 0;
  const conferir = [];
  if (fim >= LINHA_INICIAL) {
    const col = colunasCadastro();
    const dados = cad.getRange(LINHA_INICIAL, 1, fim - LINHA_INICIAL + 1, col.total).getValues();
    const vals = dados.map(l => {
      const v = l[0] === '' ? '' : contribPadrao(l[2]);
      if (v === 'Sim') sim++; else if (v === 'Não') nao++;
      // Lucro Presumido sem IE costuma fazer SÓ a EFD Contribuições
      const ie = col.ie ? String(l[col.ie - 1]).trim().toUpperCase() : '';
      if (v === 'Sim' && (ie === '' || ie === 'S/IE') &&
          String(l[col.sped - 1]).trim().toUpperCase() === 'SIM') conferir.push(l[0]);
      return [v];
    });
    cad.getRange(LINHA_INICIAL, nova, vals.length, 1).setValues(vals);
  }

  console.log('Pronto: coluna "' + CAD_CABECALHOS.contrib + '" criada na coluna ' + letraColuna(nova) +
              ' do CADASTRO. Preenchida pelo regime: ' + sim + ' Sim, ' + nao + ' Não.');
  if (conferir.length) {
    console.log('Confira: estas são Lucro Presumido sem IE — se fazem SÓ a EFD Contribuições, ' +
                'mude FAZ SPED? para Não (aí a EFD ICMS/IPI já vem "Não se aplica"): ' + conferir.join(', '));
  }
}

/** 1 -> A, 27 -> AA (só para as mensagens). */
function letraColuna(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}
