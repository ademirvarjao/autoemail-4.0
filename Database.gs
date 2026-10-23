/**
 * AUTO EMAIL 4.0 — armazenamento estruturado, migração e relatórios.
 * Google Sheets é usado como persistência para manter o projeto autocontido no Apps Script.
 */
const DB_SCHEMA_ = Object.freeze({
  CAMPANHAS: Object.freeze({
    name: 'AE • Campanhas',
    headers: Object.freeze([
      'ID da Campanha', 'Iniciada em', 'Concluída em', 'Assunto', 'Planilha de Origem',
      'Linhas Selecionadas', 'Mensagens Aceitas', 'Falhas', 'Destinatários Estimados', 'Cota Gmail Antes do Envio',
      'Aberturas Únicas', 'Eventos de Pixel', 'Taxa de Abertura (%)', 'Taxa de Falha (%)',
      'Duração (segundos)', 'Rastreamento Ativo', 'Status', 'Fonte', 'Notas'
    ])
  }),
  DESTINATARIOS: Object.freeze({
    name: 'AE • Destinatários',
    headers: Object.freeze([
      'ID Interno', 'ID da Campanha', 'ID de Rastreamento', 'Data/Hora Envio',
      'Destinatário principal', 'CC', 'Empresa', 'Assunto', 'Status',
      'Data 1ª Abertura', 'Data Última Abertura', 'Eventos de Pixel', 'Erro',
      'Linha de Origem', 'Fonte'
    ])
  }),
  EVENTOS: Object.freeze({
    name: 'AE • Eventos',
    headers: Object.freeze([
      'ID do Evento', 'ID da Campanha', 'ID de Rastreamento', 'Tipo de Evento',
      'Data/Hora', 'Observação'
    ])
  })
});

const RELATORIO_HEADERS_ = Object.freeze([
  'Data/Hora', 'Assunto do Email', 'Total', 'Sucesso', 'Falhas', 'Detalhes dos Erros',
  'ID da Campanha', 'Taxa de Sucesso (%)', 'Aberturas Únicas', 'Taxa de Abertura (%)',
  'Leitura consultiva', 'Status de Rastreamento'
]);

const RASTREAMENTO_HEADERS_ = Object.freeze([
  'Data Envio', 'Destinatário', 'Empresa', 'Assunto', 'Status Atual',
  'Data da 1ª Leitura', 'Data da Última Leitura', 'ID', 'ID da Campanha',
  'CC', 'Aberturas (eventos)', 'Erro envio', 'Linha Origem'
]);

function ensureDatabaseAndMigrate_() {
  ensureDatabase_();
  migrarDadosLegados_();
}

function ensureDatabase_() {
  const planilha = obterPlanilha_();
  Object.keys(DB_SCHEMA_).forEach(function (chave) {
    const definicao = DB_SCHEMA_[chave];
    const aba = obterOuCriarAba_(planilha, definicao.name);
    const esquemaAlterado = garantirCabecalhos_(aba, definicao.headers, true);
    if (esquemaAlterado) {
      formatarAbaDeDados_(aba, definicao.headers.length, true);
      try { aba.setTabColor(CONFIG.CORES.DESTAQUE); } catch (erro) { /* cor é apenas visual */ }
    }
  });
  return planilha;
}

function obterOuCriarAba_(planilha, nome) {
  return planilha.getSheetByName(nome) || planilha.insertSheet(nome);
}

function garantirCabecalhos_(aba, cabecalhos, oculta) {
  const ultimaColuna = aba.getLastColumn();
  let esquemaAlterado = false;
  if (aba.getLastRow() === 0 || ultimaColuna === 0) {
    aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos.slice()]);
    esquemaAlterado = true;
  } else {
    const existentes = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0]
      .map(function (valor) { return String(valor).trim(); });
    const conjunto = new Set(existentes);
    const faltantes = cabecalhos.filter(function (cabecalho) { return !conjunto.has(cabecalho); });
    if (faltantes.length) {
      aba.getRange(1, ultimaColuna + 1, 1, faltantes.length).setValues([faltantes]);
      esquemaAlterado = true;
    }
  }

  if (aba.getFrozenRows() !== 1) {
    aba.setFrozenRows(1);
    esquemaAlterado = true;
  }
  if (esquemaAlterado) {
    const colunas = Math.max(aba.getLastColumn(), cabecalhos.length);
    aba.getRange(1, 1, 1, colunas)
      .setBackground(CONFIG.CORES.FUNDO_CABECALHO)
      .setFontColor(CONFIG.CORES.TEXTO_CABECALHO)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');
    aba.setRowHeight(1, 34);
  }
  if (oculta && !aba.isSheetHidden()) {
    try { aba.hideSheet(); } catch (erro) { console.warn('A aba não pôde ser ocultada: ' + erro); }
  }
  return esquemaAlterado;
}

function formatarAbaDeDados_(aba, colunasPrincipais, oculta) {
  const ultimaLinha = Math.max(aba.getLastRow(), 1);
  const ultimaColuna = Math.max(aba.getLastColumn(), colunasPrincipais);
  if (ultimaLinha > 1) {
    aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).setVerticalAlignment('middle');
  }
  if (!oculta && !aba.getFilter() && ultimaColuna > 0) {
    try { aba.getRange(1, 1, ultimaLinha, ultimaColuna).createFilter(); } catch (erro) { /* filtro opcional */ }
  }
}

function garantirAbaOperacional_(nome, cabecalhos, cor) {
  const planilha = obterPlanilha_();
  const aba = obterOuCriarAba_(planilha, nome);
  const esquemaAlterado = garantirCabecalhos_(aba, cabecalhos, false);
  if (esquemaAlterado) {
    if (cor) {
      try { aba.setTabColor(cor); } catch (erro) { /* cor é apenas visual */ }
    }
    const totalColunasAlteradas = aba.getLastColumn();
    const totalLinhasAlteradas = aba.getLastRow();
    if (totalColunasAlteradas > 0 && totalLinhasAlteradas > 1) {
      aba.getRange(2, 1, totalLinhasAlteradas - 1, totalColunasAlteradas).setVerticalAlignment('middle');
    }
  }

  const totalColunas = aba.getLastColumn();
  const totalLinhas = Math.max(aba.getLastRow(), 1);
  const filtro = aba.getFilter();
  if (!filtro || filtro.getRange().getNumRows() < totalLinhas || filtro.getRange().getNumColumns() < totalColunas) {
    if (filtro) filtro.remove();
    const linhasDoFiltro = Math.max(totalLinhas, Math.min(aba.getMaxRows(), 1000));
    try { aba.getRange(1, 1, linhasDoFiltro, totalColunas).createFilter(); } catch (erro) { /* filtro opcional */ }
  }
  return aba;
}

function obterCabecalhosAba_(aba) {
  const totalColunas = aba.getLastColumn();
  return totalColunas ? aba.getRange(1, 1, 1, totalColunas).getDisplayValues()[0].map(function (v) { return String(v).trim(); }) : [];
}

function obterIndiceCabecalho_(cabecalhos, nome) {
  const desejado = normalizarCabecalho_(nome);
  return cabecalhos.findIndex(function (cabecalho) { return normalizarCabecalho_(cabecalho) === desejado; });
}

function criarObjetoLinha_(cabecalhos, valores) {
  const objeto = {};
  cabecalhos.forEach(function (cabecalho, indice) {
    if (cabecalho) objeto[cabecalho] = valores[indice];
  });
  return objeto;
}

function listarObjetosAba_(aba) {
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 2 || ultimaColuna < 1) return [];
  const cabecalhos = obterCabecalhosAba_(aba);
  const valores = aba.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
  return valores.map(function (linha, indice) {
    const objeto = criarObjetoLinha_(cabecalhos, linha);
    objeto.__rowNumber = indice + 2;
    return objeto;
  });
}

function inserirObjetoAba_(aba, registro) {
  const cabecalhos = obterCabecalhosAba_(aba);
  const linha = cabecalhos.map(function (cabecalho) {
    return Object.prototype.hasOwnProperty.call(registro, cabecalho) ? registro[cabecalho] : '';
  });
  aba.appendRow(linha);
  return aba.getLastRow();
}

function localizarLinhaPorValor_(aba, nomeCabecalho, valor) {
  if (valor == null || valor === '' || aba.getLastRow() < 2) return -1;
  const cabecalhos = obterCabecalhosAba_(aba);
  const indice = obterIndiceCabecalho_(cabecalhos, nomeCabecalho);
  if (indice < 0) return -1;
  const encontrado = aba.getRange(2, indice + 1, aba.getLastRow() - 1, 1)
    .createTextFinder(String(valor))
    .matchEntireCell(true)
    .useRegularExpression(false)
    .findNext();
  return encontrado ? encontrado.getRow() : -1;
}

function obterObjetoPorValor_(aba, nomeCabecalho, valor) {
  const numeroLinha = localizarLinhaPorValor_(aba, nomeCabecalho, valor);
  if (numeroLinha < 0) return null;
  const cabecalhos = obterCabecalhosAba_(aba);
  const valores = aba.getRange(numeroLinha, 1, 1, cabecalhos.length).getValues()[0];
  const objeto = criarObjetoLinha_(cabecalhos, valores);
  objeto.__rowNumber = numeroLinha;
  return objeto;
}

function atualizarObjetoPorValor_(aba, nomeCabecalho, valor, alteracoes) {
  const numeroLinha = localizarLinhaPorValor_(aba, nomeCabecalho, valor);
  if (numeroLinha < 0) return -1;
  const cabecalhos = obterCabecalhosAba_(aba);
  Object.keys(alteracoes).forEach(function (chave) {
    const indice = obterIndiceCabecalho_(cabecalhos, chave);
    if (indice >= 0) aba.getRange(numeroLinha, indice + 1).setValue(alteracoes[chave]);
  });
  return numeroLinha;
}

function registrarInicioCampanha_(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const planilha = ensureDatabase_();
    const aba = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
    const agora = dados.iniciadaEm || new Date();
    inserirObjetoAba_(aba, {
      'ID da Campanha': dados.id,
      'Iniciada em': agora,
      'Concluída em': '',
      'Assunto': dados.assunto || '',
      'Planilha de Origem': dados.planilhaOrigem || '',
      'Linhas Selecionadas': numeroSeguro_(dados.totalLinhas),
      'Mensagens Aceitas': 0,
      'Falhas': 0,
      'Destinatários Estimados': numeroSeguro_(dados.destinatariosEstimados),
      'Cota Gmail Antes do Envio': dados.cotaAntesEnvio == null ? '' : numeroSeguro_(dados.cotaAntesEnvio),
      'Aberturas Únicas': 0,
      'Eventos de Pixel': 0,
      'Taxa de Abertura (%)': '',
      'Taxa de Falha (%)': '',
      'Duração (segundos)': '',
      'Rastreamento Ativo': Boolean(dados.rastreamentoAtivo),
      'Status': 'EM_ANDAMENTO',
      'Fonte': 'LIVE',
      'Notas': dados.rastreamentoAtivo ? 'Pixel de abertura habilitado.' : 'Pixel não habilitado; envios seguem registrados.'
    });
  } finally {
    lock.releaseLock();
  }
}

function finalizarRegistroCampanha_(id, resultado) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const planilha = ensureDatabase_();
    const aba = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
    const sucesso = numeroSeguro_(resultado.sucesso);
    const falhas = numeroSeguro_(resultado.falhas);
    const unicas = numeroSeguro_(resultado.aberturasUnicas);
    const eventos = numeroSeguro_(resultado.eventosPixel);
    const rastreio = Boolean(resultado.rastreamentoAtivo);
    const baseAbertura = rastreio ? sucesso : 0;
    atualizarObjetoPorValor_(aba, 'ID da Campanha', id, {
      'Concluída em': new Date(),
      'Mensagens Aceitas': sucesso,
      'Falhas': falhas,
      'Aberturas Únicas': unicas,
      'Eventos de Pixel': eventos,
      'Taxa de Abertura (%)': baseAbertura ? unicas / baseAbertura : '',
      'Taxa de Falha (%)': (sucesso + falhas) ? falhas / (sucesso + falhas) : 0,
      'Duração (segundos)': numeroSeguro_(resultado.duracaoSegundos),
      'Status': resultado.status || 'CONCLUÍDA',
      'Notas': resultado.notas || ''
    });
  } finally {
    lock.releaseLock();
  }
}

function registrarDestinatario_(registro) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const planilha = ensureDatabase_();
    const abaDb = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const linhaDb = inserirObjetoAba_(abaDb, registro);
    const abaRelatorio = garantirAbaOperacional_(CONFIG.ABA_RASTREAMENTO, RASTREAMENTO_HEADERS_, '#087e8b');
    const linhaVisivel = inserirObjetoAba_(abaRelatorio, {
      'Data Envio': registro['Data/Hora Envio'],
      'Destinatário': registro['Destinatário principal'] || '',
      'Empresa': registro['Empresa'] || '',
      'Assunto': registro['Assunto'] || '',
      'Status Atual': registro.Status || 'PENDENTE',
      'Data da 1ª Leitura': registro['Data 1ª Abertura'] || '',
      'Data da Última Leitura': registro['Data Última Abertura'] || '',
      'ID': registro['ID de Rastreamento'] || '',
      'ID da Campanha': registro['ID da Campanha'] || '',
      'CC': registro.CC || '',
      'Aberturas (eventos)': numeroSeguro_(registro['Eventos de Pixel']),
      'Erro envio': registro.Erro || '',
      'Linha Origem': registro['Linha de Origem'] || ''
    });
    return { databaseRow: linhaDb, reportRow: linhaVisivel };
  } finally {
    lock.releaseLock();
  }
}

function atualizarStatusDestinatario_(idInterno, status, erro, linhaVisivel) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const planilha = ensureDatabase_();
    const abaDb = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const atual = obterObjetoPorValor_(abaDb, 'ID Interno', idInterno);
    const statusFinal = status === 'ENVIADO' && atual && atual['Data 1ª Abertura'] ? 'ABERTO' : status;
    atualizarObjetoPorValor_(abaDb, 'ID Interno', idInterno, { Status: statusFinal, Erro: erro || '' });
    const abaRelatorio = planilha.getSheetByName(CONFIG.ABA_RASTREAMENTO);
    if (abaRelatorio && linhaVisivel && linhaVisivel > 1) {
      const cabecalhos = obterCabecalhosAba_(abaRelatorio);
      const indiceStatus = obterIndiceCabecalho_(cabecalhos, 'Status Atual');
      const indiceErro = obterIndiceCabecalho_(cabecalhos, 'Erro envio');
      if (indiceStatus >= 0) abaRelatorio.getRange(linhaVisivel, indiceStatus + 1).setValue(statusFinal);
      if (indiceErro >= 0) abaRelatorio.getRange(linhaVisivel, indiceErro + 1).setValue(erro || '');
      const linha = abaRelatorio.getRange(linhaVisivel, 1, 1, abaRelatorio.getLastColumn());
      linha.setBackground(statusFinal === 'FALHA' ? CONFIG.CORES.ERRO : statusFinal === 'ABERTO' ? CONFIG.CORES.ABERTO : '#ffffff');
    }
  } finally {
    lock.releaseLock();
  }
}

function registrarRelatorioCampanha_(dados) {
  const aba = garantirAbaOperacional_(CONFIG.ABA_RELATORIO, RELATORIO_HEADERS_, '#1f6b78');
  const taxaSucesso = (dados.total || 0) ? dados.sucesso / dados.total : 0;
  const taxaAbertura = dados.rastreamentoAtivo && dados.sucesso ? dados.aberturasUnicas / dados.sucesso : '';
  const recomendacao = gerarLeituraConsultiva_(dados);
  const linha = inserirObjetoAba_(aba, {
    'Data/Hora': new Date(),
    'Assunto do Email': dados.assunto || '',
    'Total': numeroSeguro_(dados.total),
    'Sucesso': numeroSeguro_(dados.sucesso),
    'Falhas': numeroSeguro_(dados.falhas),
    'Detalhes dos Erros': dados.detalhes || 'Sem erros registrados.',
    'ID da Campanha': dados.campaignId || '',
    'Taxa de Sucesso (%)': taxaSucesso,
    'Aberturas Únicas': numeroSeguro_(dados.aberturasUnicas),
    'Taxa de Abertura (%)': taxaAbertura,
    'Leitura consultiva': recomendacao,
    'Status de Rastreamento': dados.rastreamentoAtivo ? 'Ativo; sinal indicativo' : 'Não ativado'
  });
  const totalColunas = aba.getLastColumn();
  aba.getRange(linha, 1, 1, totalColunas)
    .setBackground(dados.falhas ? CONFIG.CORES.ERRO : CONFIG.CORES.SUCESSO)
    .setVerticalAlignment('middle');
  aba.getRange(linha, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  const indiceTaxaSucesso = obterIndiceCabecalho_(obterCabecalhosAba_(aba), 'Taxa de Sucesso (%)');
  const indiceTaxaAbertura = obterIndiceCabecalho_(obterCabecalhosAba_(aba), 'Taxa de Abertura (%)');
  if (indiceTaxaSucesso >= 0) aba.getRange(linha, indiceTaxaSucesso + 1).setNumberFormat('0.0%');
  if (indiceTaxaAbertura >= 0) aba.getRange(linha, indiceTaxaAbertura + 1).setNumberFormat('0.0%');
  ajustarColunasRelatorio_(aba);
  return recomendacao;
}

function gerarLeituraConsultiva_(dados) {
  const totalTentativas = numeroSeguro_(dados.sucesso) + numeroSeguro_(dados.falhas);
  const taxaFalha = totalTentativas ? dados.falhas / totalTentativas : 0;
  const taxaAbertura = dados.rastreamentoAtivo && dados.sucesso
    ? numeroSeguro_(dados.aberturasUnicas) / dados.sucesso
    : null;
  const orientacoes = [];

  if (!dados.rastreamentoAtivo) {
    orientacoes.push('Ative o Web App para medir sinais de abertura em próximos disparos.');
  } else if (taxaAbertura === 0) {
    orientacoes.push('Nenhuma abertura foi registrada; confirme o deployment e lembre que bloqueios de imagens podem ocultar leituras.');
  } else if (taxaAbertura < 0.1) {
    orientacoes.push('O sinal de abertura foi baixo; revise segmentação, clareza do assunto e horário antes de reenviar.');
  } else if (taxaAbertura < 0.25) {
    orientacoes.push('Compare o assunto e o público com campanhas anteriores antes de ajustar a mensagem.');
  } else {
    orientacoes.push('O sinal de abertura foi consistente; acompanhe respostas e próximas ações, não apenas o pixel.');
  }

  if (taxaFalha > 0.05) {
    orientacoes.push('Priorize a higienização dos e-mails e a revisão das linhas com falha.');
  } else if (dados.falhas > 0) {
    orientacoes.push('Revise as poucas falhas listadas no relatório para evitar recorrência.');
  } else if (totalTentativas) {
    orientacoes.push('Envio processado sem falhas registradas pelo Gmail.');
  }

  orientacoes.push('Aberturas são estimativas: proxies de imagem podem inflar ou bloquear a medição; não comprovam leitura humana.');
  return orientacoes.join(' ');
}

function ajustarColunasRelatorio_(aba) {
  const larguras = {
    'Data/Hora': 155, 'Assunto do Email': 260, 'Total': 80, 'Sucesso': 85,
    'Falhas': 80, 'Detalhes dos Erros': 330, 'ID da Campanha': 230,
    'Taxa de Sucesso (%)': 125, 'Aberturas Únicas': 125, 'Taxa de Abertura (%)': 130,
    'Leitura consultiva': 480, 'Status de Rastreamento': 170
  };
  const cabecalhos = obterCabecalhosAba_(aba);
  cabecalhos.forEach(function (nome, indice) {
    if (larguras[nome]) aba.setColumnWidth(indice + 1, larguras[nome]);
  });
  aba.setFrozenRows(1);
}

function migrarDadosLegados_() {
  const propriedades = PropertiesService.getScriptProperties();
  if (propriedades.getProperty(CONFIG.CHAVE_MIGRACAO_V4) === 'true') return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    if (propriedades.getProperty(CONFIG.CHAVE_MIGRACAO_V4) === 'true') return;
    const planilha = obterPlanilha_();
    const abaCampanhas = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
    const existentes = new Set(listarObjetosAba_(abaCampanhas).map(function (item) { return String(item['ID da Campanha'] || ''); }));

    const abaResumo = planilha.getSheetByName(CONFIG.ABA_RELATORIO);
    if (abaResumo && abaResumo.getLastRow() > 1) {
      const cabecalhos = obterCabecalhosAba_(abaResumo);
      const linhas = abaResumo.getRange(2, 1, abaResumo.getLastRow() - 1, abaResumo.getLastColumn()).getValues();
      const idxData = obterIndiceCabecalho_(cabecalhos, 'Data/Hora');
      const idxAssunto = obterIndiceCabecalho_(cabecalhos, 'Assunto do Email');
      const idxTotal = obterIndiceCabecalho_(cabecalhos, 'Total');
      const idxSucesso = obterIndiceCabecalho_(cabecalhos, 'Sucesso');
      const idxFalhas = obterIndiceCabecalho_(cabecalhos, 'Falhas');
      const idxDetalhes = obterIndiceCabecalho_(cabecalhos, 'Detalhes dos Erros');
      const idxId = obterIndiceCabecalho_(cabecalhos, 'ID da Campanha');

      linhas.forEach(function (linha, indice) {
        const linhaPlanilha = indice + 2;
        const idExistente = idxId >= 0 ? String(linha[idxId] || '').trim() : '';
        const id = idExistente || 'LEGACY-REPORT-' + linhaPlanilha;
        if (existentes.has(id)) return;
        const data = idxData >= 0 ? parseDataLegada_(linha[idxData]) : null;
        if (idxTotal < 0 && idxSucesso < 0) return;
        inserirObjetoAba_(abaCampanhas, {
          'ID da Campanha': id,
          'Iniciada em': data || '',
          'Concluída em': data || '',
          'Assunto': idxAssunto >= 0 ? String(linha[idxAssunto] || '') : '',
          'Planilha de Origem': '',
          'Linhas Selecionadas': idxTotal >= 0 ? numeroSeguro_(linha[idxTotal]) : 0,
          'Mensagens Aceitas': idxSucesso >= 0 ? numeroSeguro_(linha[idxSucesso]) : 0,
          'Falhas': idxFalhas >= 0 ? numeroSeguro_(linha[idxFalhas]) : 0,
          'Destinatários Estimados': '',
          'Aberturas Únicas': 0,
          'Eventos de Pixel': 0,
          'Taxa de Abertura (%)': '',
          'Taxa de Falha (%)': '',
          'Duração (segundos)': '',
          'Rastreamento Ativo': false,
          'Status': 'HISTÓRICO IMPORTADO',
          'Fonte': 'LEGACY_REPORT',
          'Notas': 'Resumo importado da aba v3.2; não foi associado automaticamente a destinatários.' +
            (idxDetalhes >= 0 && linha[idxDetalhes] ? ' Detalhes: ' + String(linha[idxDetalhes]) : '')
        });
        existentes.add(id);
      });
    }

    const abaDetalhes = planilha.getSheetByName(CONFIG.ABA_RASTREAMENTO);
    const abaDbDestinatarios = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const idsDestinatariosExistentes = new Set(listarObjetosAba_(abaDbDestinatarios).map(function (item) { return String(item['ID Interno'] || ''); }));
    if (abaDetalhes && abaDetalhes.getLastRow() > 1) {
      const cabecalhos = obterCabecalhosAba_(abaDetalhes);
      const linhas = abaDetalhes.getRange(2, 1, abaDetalhes.getLastRow() - 1, abaDetalhes.getLastColumn()).getValues();
      const idxData = obterIndiceCabecalho_(cabecalhos, 'Data Envio');
      const idxDest = obterIndiceCabecalho_(cabecalhos, 'Destinatário');
      const idxEmpresa = obterIndiceCabecalho_(cabecalhos, 'Empresa');
      const idxAssunto = obterIndiceCabecalho_(cabecalhos, 'Assunto');
      const idxStatus = obterIndiceCabecalho_(cabecalhos, 'Status Atual');
      const idxPrimeira = obterIndiceCabecalho_(cabecalhos, 'Data da 1ª Leitura');
      const idxUltima = obterIndiceCabecalho_(cabecalhos, 'Data da Última Leitura');
      const idxId = obterIndiceCabecalho_(cabecalhos, 'ID');
      const idxCampaign = obterIndiceCabecalho_(cabecalhos, 'ID da Campanha');
      const idxCc = obterIndiceCabecalho_(cabecalhos, 'CC');
      let importados = 0;
      let aberturas = 0;
      let primeiroEnvio = null;
      let ultimoEnvio = null;

      linhas.forEach(function (linha, indice) {
        const linhaPlanilha = indice + 2;
        const trackingId = idxId >= 0 ? String(linha[idxId] || '').trim() : '';
        const campaignIdLegado = idxCampaign >= 0 ? String(linha[idxCampaign] || '').trim() : '';
        const idInterno = trackingId || 'LEGACY-TRACK-' + linhaPlanilha;
        if (idsDestinatariosExistentes.has(idInterno)) return;
        const dataEnvio = idxData >= 0 ? parseDataLegada_(linha[idxData]) : null;
        const primeiraAbertura = idxPrimeira >= 0 ? parseDataLegada_(linha[idxPrimeira]) : null;
        const ultimaAbertura = idxUltima >= 0 ? parseDataLegada_(linha[idxUltima]) : null;
        const statusOriginal = idxStatus >= 0 ? String(linha[idxStatus] || '') : 'ENVIADO';
        const foiAberto = Boolean(primeiraAbertura || /lido|aberto/i.test(statusOriginal));
        const campaignId = campaignIdLegado || 'LEGACY-TRACKING';
        const agora = dataEnvio || new Date();

        inserirObjetoAba_(abaDbDestinatarios, {
          'ID Interno': idInterno,
          'ID da Campanha': campaignId,
          'ID de Rastreamento': trackingId,
          'Data/Hora Envio': dataEnvio || '',
          'Destinatário principal': idxDest >= 0 ? String(linha[idxDest] || '') : '',
          'CC': idxCc >= 0 ? String(linha[idxCc] || '') : '',
          'Empresa': idxEmpresa >= 0 ? String(linha[idxEmpresa] || '') : '',
          'Assunto': idxAssunto >= 0 ? String(linha[idxAssunto] || '') : '',
          'Status': foiAberto ? 'ABERTO' : statusOriginal || 'ENVIADO',
          'Data 1ª Abertura': primeiraAbertura || '',
          'Data Última Abertura': ultimaAbertura || primeiraAbertura || '',
          'Eventos de Pixel': foiAberto ? 1 : 0,
          'Erro': '',
          'Linha de Origem': linhaPlanilha,
          'Fonte': 'LEGACY_V3'
        });
        idsDestinatariosExistentes.add(idInterno);
        importados++;
        if (foiAberto) {
          aberturas++;
          inserirObjetoAba_(planilha.getSheetByName(DB_SCHEMA_.EVENTOS.name), {
            'ID do Evento': 'LEGACY-OPEN-' + linhaPlanilha,
            'ID da Campanha': campaignId,
            'ID de Rastreamento': trackingId,
            'Tipo de Evento': 'OPEN',
            'Data/Hora': primeiraAbertura || agora,
            'Observação': 'Evento reconstruído da primeira abertura disponível no relatório v3.2.'
          });
        }
        if (dataEnvio && (!primeiroEnvio || dataEnvio < primeiroEnvio)) primeiroEnvio = dataEnvio;
        if (dataEnvio && (!ultimoEnvio || dataEnvio > ultimoEnvio)) ultimoEnvio = dataEnvio;
      });

      if (importados && !existentes.has('LEGACY-TRACKING')) {
        inserirObjetoAba_(abaCampanhas, {
          'ID da Campanha': 'LEGACY-TRACKING',
          'Iniciada em': primeiroEnvio || '',
          'Concluída em': ultimoEnvio || '',
          'Assunto': 'Histórico de rastreamento importado',
          'Planilha de Origem': '',
          'Linhas Selecionadas': importados,
          'Mensagens Aceitas': importados,
          'Falhas': 0,
          'Destinatários Estimados': importados,
          'Aberturas Únicas': aberturas,
          'Eventos de Pixel': aberturas,
          'Taxa de Abertura (%)': importados ? aberturas / importados : 0,
          'Taxa de Falha (%)': 0,
          'Duração (segundos)': '',
          'Rastreamento Ativo': true,
          'Status': 'HISTÓRICO IMPORTADO',
          'Fonte': 'LEGACY_TRACKING',
          'Notas': 'O v3.2 não guardava associação confiável entre o resumo do disparo e cada destinatário. Esta série histórica fica separada para evitar dupla contagem.'
        });
        existentes.add('LEGACY-TRACKING');
      }
    }

    propriedades.setProperty(CONFIG.CHAVE_MIGRACAO_V4, 'true');
  } finally {
    lock.releaseLock();
  }
}

function getDashboardData(periodo) {
  try {
    ensureDatabaseAndMigrate_();
    const planilha = obterPlanilha_();
    const abaCampanhas = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
    const abaDestinatarios = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const campanhasTodas = listarObjetosAba_(abaCampanhas).map(mapearCampanha_);
    const destinatarios = listarObjetosAba_(abaDestinatarios);
    const dias = periodo === 'all' ? 3650 : Math.max(7, Math.min(365, Number(periodo) || 30));
    const agora = new Date();
    const inicio = periodo === 'all' ? new Date(0) : new Date(agora.getTime() - dias * 24 * 60 * 60 * 1000);
    const inicioAnterior = periodo === 'all' ? new Date(0) : new Date(agora.getTime() - 2 * dias * 24 * 60 * 60 * 1000);

    const campanhasLiveTodas = campanhasTodas.filter(function (campanha) { return campanha.fonte === 'LIVE'; });
    const campanhasPeriodo = campanhasLiveTodas.filter(function (campanha) {
      return campanha.dataInicio && campanha.dataInicio >= inicio;
    });
    const campanhasPeriodoAnterior = campanhasLiveTodas.filter(function (campanha) {
      return campanha.dataInicio && campanha.dataInicio >= inicioAnterior && campanha.dataInicio < inicio;
    });
    const idsLivePeriodo = new Set(campanhasPeriodo.map(function (campanha) { return campanha.id; }));
    const idsLiveAnterior = new Set(campanhasPeriodoAnterior.map(function (campanha) { return campanha.id; }));
    const destinatariosPeriodo = destinatarios.filter(function (item) { return idsLivePeriodo.has(String(item['ID da Campanha'] || '')); });
    const destinatariosAnterior = destinatarios.filter(function (item) { return idsLiveAnterior.has(String(item['ID da Campanha'] || '')); });
    const resumo = resumirCampanhas_(campanhasPeriodo, destinatariosPeriodo);
    const anterior = resumirCampanhas_(campanhasPeriodoAnterior, destinatariosAnterior);
    const campanhasRecentes = campanhasTodas.slice().sort(function (a, b) {
      return (b.dataInicio ? b.dataInicio.getTime() : 0) - (a.dataInicio ? a.dataInicio.getTime() : 0);
    }).slice(0, 10).map(function (campanha) {
      const registros = destinatarios.filter(function (item) { return String(item['ID da Campanha'] || '') === campanha.id; });
      const abertos = registros.filter(function (item) { return Boolean(item['Data 1ª Abertura']); }).length;
      const eventos = registros.reduce(function (soma, item) { return soma + numeroSeguro_(item['Eventos de Pixel']); }, 0);
      const rastreio = campanha.rastreamentoAtivo;
      return {
        id: campanha.id,
        assunto: campanha.assunto || '(Sem assunto)',
        data: dataParaIso_(campanha.dataInicio),
        status: campanha.status,
        fonte: campanha.fonte,
        enviado: campanha.mensagensAceitas,
        falhas: campanha.falhas,
        aberturasUnicas: abertos,
        eventosPixel: eventos,
        taxaAbertura: rastreio && campanha.mensagensAceitas ? abertos / campanha.mensagensAceitas * 100 : null,
        recomendacao: recomendacaoDashboard_(campanha, abertos, eventos)
      };
    });

    const historicoLegado = campanhasTodas.filter(function (campanha) { return campanha.fonte !== 'LIVE'; });
    const historico = {
      campanhasResumo: historicoLegado.filter(function (campanha) { return campanha.fonte === 'LEGACY_REPORT'; }).length,
      enviosReportados: historicoLegado.filter(function (campanha) { return campanha.fonte === 'LEGACY_REPORT'; })
        .reduce(function (soma, campanha) { return soma + campanha.mensagensAceitas; }, 0),
      falhasReportadas: historicoLegado.filter(function (campanha) { return campanha.fonte === 'LEGACY_REPORT'; })
        .reduce(function (soma, campanha) { return soma + campanha.falhas; }, 0),
      aberturasImportadas: historicoLegado.filter(function (campanha) { return campanha.fonte === 'LEGACY_TRACKING'; })
        .reduce(function (soma, campanha) { return soma + campanha.aberturasUnicas; }, 0),
      destinatariosImportados: destinatarios.filter(function (item) { return String(item.Fonte || '').indexOf('LEGACY') === 0; }).length
    };

    return {
      success: true,
      periodo: periodo || '30',
      atualizadoEm: agora.toISOString(),
      kpis: resumo,
      comparacao: periodo === 'all' ? null : {
        mensagensAceitas: variacaoPercentual_(resumo.mensagensAceitas, anterior.mensagensAceitas),
        aberturasUnicas: variacaoPercentual_(resumo.aberturasUnicas, anterior.aberturasUnicas),
        falhas: variacaoPercentual_(resumo.falhas, anterior.falhas)
      },
      trend: montarTendencia_(campanhasPeriodo, destinatariosPeriodo, 7),
      recomendacoes: gerarRecomendacoesGerais_(resumo, campanhasPeriodo, Boolean(getTrackingUrl())),
      campanhas: campanhasRecentes,
      historico: historico,
      status: {
        trackingEnabled: Boolean(getTrackingUrl()),
        storage: 'Google Sheets',
        dataSheet: getSheetDados().getName()
      }
    };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro) };
  }
}

function mapearCampanha_(item) {
  return {
    id: String(item['ID da Campanha'] || ''),
    dataInicio: parseDataLegada_(item['Iniciada em']),
    dataFim: parseDataLegada_(item['Concluída em']),
    assunto: String(item.Assunto || ''),
    mensagensAceitas: numeroSeguro_(item['Mensagens Aceitas']),
    falhas: numeroSeguro_(item.Falhas),
    linhas: numeroSeguro_(item['Linhas Selecionadas']),
    destinatariosEstimados: numeroSeguro_(item['Destinatários Estimados']),
    aberturasUnicas: numeroSeguro_(item['Aberturas Únicas']),
    eventosPixel: numeroSeguro_(item['Eventos de Pixel']),
    rastreamentoAtivo: item['Rastreamento Ativo'] === true || String(item['Rastreamento Ativo']).toLowerCase() === 'true',
    status: String(item.Status || ''),
    fonte: String(item.Fonte || 'LIVE'),
    notas: String(item.Notas || '')
  };
}

function resumirCampanhas_(campanhas, destinatarios) {
  const totalMensagens = campanhas.reduce(function (soma, campanha) { return soma + campanha.mensagensAceitas; }, 0);
  const totalFalhas = campanhas.reduce(function (soma, campanha) { return soma + campanha.falhas; }, 0);
  const totalLinhas = campanhas.reduce(function (soma, campanha) { return soma + campanha.linhas; }, 0);
  const campanhasRastreadas = campanhas.filter(function (campanha) { return campanha.rastreamentoAtivo; });
  const idsRastreados = new Set(campanhasRastreadas.map(function (campanha) { return campanha.id; }));
  const linhasRastreadas = destinatarios.filter(function (item) { return idsRastreados.has(String(item['ID da Campanha'] || '')); });
  const abertos = linhasRastreadas.filter(function (item) { return Boolean(item['Data 1ª Abertura']); }).length;
  const eventos = linhasRastreadas.reduce(function (soma, item) { return soma + numeroSeguro_(item['Eventos de Pixel']); }, 0);
  const baseRastreavel = campanhasRastreadas.reduce(function (soma, campanha) { return soma + campanha.mensagensAceitas; }, 0);
  const totalTentativas = totalMensagens + totalFalhas;
  return {
    campanhas: campanhas.length,
    linhasSelecionadas: totalLinhas,
    mensagensAceitas: totalMensagens,
    falhas: totalFalhas,
    destinatariosEstimados: campanhas.reduce(function (soma, campanha) { return soma + campanha.destinatariosEstimados; }, 0),
    aberturasUnicas: abertos,
    eventosPixel: eventos,
    taxaSucesso: totalTentativas ? totalMensagens / totalTentativas * 100 : 0,
    taxaFalha: totalTentativas ? totalFalhas / totalTentativas * 100 : 0,
    taxaAbertura: baseRastreavel ? abertos / baseRastreavel * 100 : null,
    mensagensRastreaveis: baseRastreavel
  };
}

function variacaoPercentual_(atual, anterior) {
  if (!anterior && !atual) return 0;
  if (!anterior) return null;
  return (atual - anterior) / anterior * 100;
}

function montarTendencia_(campanhas, destinatarios, totalDias) {
  const resultado = [];
  const agora = new Date();
  const diaAtual = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  for (let deslocamento = totalDias - 1; deslocamento >= 0; deslocamento--) {
    const inicioDia = new Date(diaAtual.getTime() - deslocamento * 24 * 60 * 60 * 1000);
    const fimDia = new Date(inicioDia.getTime() + 24 * 60 * 60 * 1000);
    const campanhasDia = campanhas.filter(function (campanha) {
      return campanha.dataInicio && campanha.dataInicio >= inicioDia && campanha.dataInicio < fimDia;
    });
    const aberturasDia = destinatarios.filter(function (item) {
      const dataAbertura = parseDataLegada_(item['Data 1ª Abertura']);
      return dataAbertura && dataAbertura >= inicioDia && dataAbertura < fimDia;
    });
    resultado.push({
      data: inicioDia.toISOString(),
      label: Utilities.formatDate(inicioDia, Session.getScriptTimeZone(), 'dd/MM'),
      enviados: campanhasDia.reduce(function (soma, campanha) { return soma + campanha.mensagensAceitas; }, 0),
      aberturas: aberturasDia.length
    });
  }
  return resultado;
}

function recomendacaoDashboard_(campanha, aberturas, eventos) {
  if (campanha.fonte !== 'LIVE') return 'Histórico importado do v3.2; mantido separado para evitar associações e taxas não confiáveis.';
  if (campanha.status === 'EM_ANDAMENTO') return 'Disparo em andamento. Aguarde a conclusão para avaliar os indicadores.';
  const tentativas = campanha.mensagensAceitas + campanha.falhas;
  if (tentativas && campanha.falhas / tentativas > 0.05) return 'Ação prioritária: revisar os endereços e as linhas com falha antes do próximo envio.';
  if (!campanha.rastreamentoAtivo) return 'O envio foi registrado, mas aberturas não são mensuráveis sem o Web App ativo.';
  if (!aberturas) return 'Sem sinal de abertura. Valide o deployment e considere bloqueios de imagem antes de concluir que a mensagem não foi lida.';
  if (campanha.mensagensAceitas && aberturas / campanha.mensagensAceitas < 0.1) return 'Teste assunto, segmentação e horário em uma próxima campanha; interprete a abertura como estimativa.';
  if (eventos > aberturas * 2) return 'Há muitos carregamentos repetidos do pixel; priorize aberturas únicas e respostas reais na análise.';
  return 'Sinal positivo de abertura. Acompanhe respostas e ações geradas para medir o resultado de negócio.';
}

function gerarRecomendacoesGerais_(kpis, campanhas, trackingAtivo) {
  const recomendacoes = [];
  if (!trackingAtivo) {
    recomendacoes.push({
      tipo: 'configuracao',
      titulo: 'Ative o rastreamento de abertura',
      texto: 'Publique o projeto como App da Web e valide a URL no dashboard. Os envios já ficam armazenados sem essa etapa, mas não haverá sinal de abertura.'
    });
  }
  if (!campanhas.length) {
    recomendacoes.push({
      tipo: 'inicio',
      titulo: 'Comece com um envio de teste',
      texto: 'Confira assunto, personalização, anexos e destinatários com uma amostra pequena antes do primeiro disparo.'
    });
  } else if (kpis.taxaFalha > 5) {
    recomendacoes.push({
      tipo: 'alerta',
      titulo: 'Reduza as falhas antes do próximo disparo',
      texto: 'A taxa de falha do período passou de 5%. Revise a coluna de e-mail, duplicidades e as mensagens de erro no relatório detalhado.'
    });
  } else if (kpis.taxaFalha > 0) {
    recomendacoes.push({
      tipo: 'atencao',
      titulo: 'Revise os registros com falha',
      texto: 'As falhas são pontuais, mas vale corrigir a origem dos endereços antes de reutilizar a lista.'
    });
  }
  if (trackingAtivo && kpis.mensagensRastreaveis > 0 && kpis.aberturasUnicas === 0) {
    recomendacoes.push({
      tipo: 'atencao',
      titulo: 'Abertura zero não significa leitura zero',
      texto: 'Clientes de e-mail podem bloquear imagens ou carregar pixels por proxy. Verifique a configuração e use respostas/cliques para validar interesse.'
    });
  } else if (trackingAtivo && kpis.taxaAbertura !== null && kpis.taxaAbertura >= 25) {
    recomendacoes.push({
      tipo: 'positivo',
      titulo: 'Bom sinal de atenção',
      texto: 'A taxa estimada de abertura está acima de 25%. Compare assunto e segmento com o histórico e avalie a próxima ação do destinatário.'
    });
  } else if (trackingAtivo && kpis.taxaAbertura !== null && kpis.taxaAbertura < 10 && kpis.mensagensRastreaveis > 0) {
    recomendacoes.push({
      tipo: 'atencao',
      titulo: 'Teste uma hipótese por vez',
      texto: 'Considere testar assunto, horário ou segmentação, mantendo o restante constante para entender o que influencia o sinal de abertura.'
    });
  }
  recomendacoes.push({
    tipo: 'nota',
    titulo: 'Leitura consultiva dos indicadores',
    texto: 'A taxa de abertura é estimativa baseada no carregamento de imagens: pode ser inflada por proxies ou subestimada por bloqueios. Não comprova leitura humana nem entrega final.'
  });
  return recomendacoes;
}

function getTrackingReport(filtros) {
  try {
    ensureDatabaseAndMigrate_();
    filtros = filtros || {};
    const planilha = obterPlanilha_();
    const aba = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const busca = normalizarCabecalho_(filtros.search || '');
    const statusFiltro = String(filtros.status || 'TODOS').toUpperCase();
    const periodo = filtros.period || '30';
    const inicio = periodo === 'all' ? null : new Date(Date.now() - Math.max(1, Number(periodo) || 30) * 24 * 60 * 60 * 1000);
    let linhas = listarObjetosAba_(aba).map(function (item) {
      const enviadoEm = parseDataLegada_(item['Data/Hora Envio']);
      return {
        idCampanha: String(item['ID da Campanha'] || ''),
        trackingId: String(item['ID de Rastreamento'] || ''),
        destinatario: String(item['Destinatário principal'] || ''),
        cc: String(item.CC || ''),
        empresa: String(item.Empresa || ''),
        assunto: String(item.Assunto || ''),
        status: String(item.Status || ''),
        enviadoEm: dataParaIso_(enviadoEm),
        primeiraAbertura: dataParaIso_(item['Data 1ª Abertura']),
        ultimaAbertura: dataParaIso_(item['Data Última Abertura']),
        eventosPixel: numeroSeguro_(item['Eventos de Pixel']),
        erro: String(item.Erro || ''),
        origem: String(item.Fonte || 'LIVE')
      };
    });
    if (inicio) linhas = linhas.filter(function (item) { return !item.enviadoEm || new Date(item.enviadoEm) >= inicio; });
    if (statusFiltro !== 'TODOS') linhas = linhas.filter(function (item) { return item.status.toUpperCase() === statusFiltro; });
    if (busca) {
      linhas = linhas.filter(function (item) {
        return normalizarCabecalho_([item.destinatario, item.cc, item.empresa, item.assunto, item.idCampanha, item.erro].join(' ')).indexOf(busca) >= 0;
      });
    }
    linhas.sort(function (a, b) { return (b.enviadoEm ? new Date(b.enviadoEm).getTime() : 0) - (a.enviadoEm ? new Date(a.enviadoEm).getTime() : 0); });
    const total = linhas.length;
    const abertas = linhas.filter(function (item) { return Boolean(item.primeiraAbertura); }).length;
    const falhas = linhas.filter(function (item) { return item.status.toUpperCase() === 'FALHA'; }).length;
    const eventos = linhas.reduce(function (soma, item) { return soma + item.eventosPixel; }, 0);
    const rastreaveis = linhas.filter(function (item) { return Boolean(item.trackingId); });
    const aberturasRastreaveis = rastreaveis.filter(function (item) { return Boolean(item.primeiraAbertura); }).length;
    return {
      success: true,
      total: total,
      aberturasUnicas: abertas,
      falhas: falhas,
      eventosPixel: eventos,
      taxaAbertura: rastreaveis.length ? aberturasRastreaveis / rastreaveis.length * 100 : 0,
      rows: filtros.exportAll ? linhas : linhas.slice(0, 500)
    };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro), rows: [] };
  }
}

function exportarRelatorioCSV(filtros) {
  filtros = Object.assign({}, filtros || {}, { exportAll: true });
  if (!filtros.period) filtros.period = 'all';
  const resultado = getTrackingReport(filtros);
  if (!resultado.success) return resultado;
  const cabecalhos = [
    'Data de envio', 'Destinatário principal', 'CC', 'Empresa', 'Assunto',
    'Status', 'Primeira abertura', 'Última abertura', 'Eventos de pixel',
    'ID da campanha', 'Erro', 'Origem'
  ];
  const linhas = [cabecalhos].concat(resultado.rows.map(function (item) {
    return [item.enviadoEm, item.destinatario, item.cc, item.empresa, item.assunto, item.status,
      item.primeiraAbertura, item.ultimaAbertura, item.eventosPixel, item.idCampanha, item.erro, item.origem];
  }));
  return { success: true, filename: 'autoemail-rastreamento.csv', csv: linhas.map(function (linha) { return linha.map(csvCell_).join(';'); }).join('\r\n') };
}

function registrarLogBatch(assunto, total, sucessos, erros, detalhesArray) {
  const campaignId = 'LEGACY-API-' + Utilities.getUuid();
  const lista = Array.isArray(detalhesArray) ? detalhesArray : [];
  ensureDatabaseAndMigrate_();
  registrarInicioCampanha_({
    id: campaignId, assunto: assunto, totalLinhas: total,
    destinatariosEstimados: total, rastreamentoAtivo: false,
    planilhaOrigem: getSheetDados().getName()
  });
  const dados = {
    campaignId: campaignId,
    assunto: assunto,
    total: total,
    sucesso: sucessos,
    falhas: erros,
    detalhes: lista.length ? lista.join(' | ') : 'Sem erros registrados.',
    aberturasUnicas: 0,
    rastreamentoAtivo: false
  };
  finalizarRegistroCampanha_(campaignId, {
    sucesso: sucessos, falhas: erros, aberturasUnicas: 0, eventosPixel: 0,
    rastreamentoAtivo: false, duracaoSegundos: 0, status: 'CONCLUÍDA', notas: dados.detalhes
  });
  registrarRelatorioCampanha_(dados);
}

function registrarEnvioIndividual(id, destinatarios, empresa, assunto, status) {
  ensureDatabaseAndMigrate_();
  const campaignId = 'LEGACY-API';
  const planilha = obterPlanilha_();
  const abaCampanhas = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
  if (!obterObjetoPorValor_(abaCampanhas, 'ID da Campanha', campaignId)) {
    registrarInicioCampanha_({ id: campaignId, assunto: assunto, totalLinhas: 0, destinatariosEstimados: 0, rastreamentoAtivo: Boolean(getTrackingUrl()) });
  }
  const registro = {
    'ID Interno': String(id || Utilities.getUuid()),
    'ID da Campanha': campaignId,
    'ID de Rastreamento': String(id || ''),
    'Data/Hora Envio': new Date(),
    'Destinatário principal': String(destinatarios || ''),
    'CC': '', 'Empresa': String(empresa || ''), 'Assunto': String(assunto || ''),
    'Status': String(status || 'ENVIADO'), 'Data 1ª Abertura': '', 'Data Última Abertura': '',
    'Eventos de Pixel': 0, 'Erro': '', 'Linha de Origem': '', 'Fonte': 'LEGACY_API'
  };
  return registrarDestinatario_(registro);
}

function ajustarColunasRastreamento_() {
  const planilha = obterPlanilha_();
  const aba = planilha.getSheetByName(CONFIG.ABA_RASTREAMENTO);
  if (!aba) return;
  const larguras = [155, 280, 180, 260, 130, 155, 155, 245, 235, 220, 130, 280, 100];
  larguras.forEach(function (largura, indice) { if (indice < aba.getLastColumn()) aba.setColumnWidth(indice + 1, largura); });
}
