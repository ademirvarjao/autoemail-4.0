/**
 * AUTO EMAIL 4.0 — endpoint público de pixel e persistência de eventos.
 * O endpoint armazena somente o token aleatório da mensagem e horários; não coleta IP ou user-agent.
 */
function doGet(e) {
  const id = e && e.parameter ? String(e.parameter.id || '').trim() : '';
  if (id) {
    try { registrarAbertura(id); }
    catch (erro) { console.error('Falha ao registrar evento de abertura: ' + erro); }
  }

  const gif = 'R0lGODlhAQABAPAAAAAAAAAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw==';
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;padding:0"><img src="data:image/gif;base64,' + gif + '" width="1" height="1" alt="" style="display:block;width:1px;height:1px;border:0"></body></html>'
  );
}

/** Nome mantido para compatibilidade com a implantação v3.2. */
function registrarAbertura(id) {
  const token = String(id == null ? '' : id).trim();
  if (!token || token.length > 120) return false;
  ensureDatabaseAndMigrate_();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return false;
  try {
    const planilha = obterPlanilha_();
    const abaDestinatarios = planilha.getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
    const registro = obterObjetoPorValor_(abaDestinatarios, 'ID de Rastreamento', token);
    if (!registro) return false;

    const agora = new Date();
    const primeiraAbertura = parseDataLegada_(registro['Data 1ª Abertura']);
    const totalEventos = numeroSeguro_(registro['Eventos de Pixel']) + 1;
    const atualizacoes = {
      'Data 1ª Abertura': primeiraAbertura || agora,
      'Data Última Abertura': agora,
      'Eventos de Pixel': totalEventos,
      'Status': 'ABERTO'
    };
    atualizarObjetoPorValor_(abaDestinatarios, 'ID de Rastreamento', token, atualizacoes);

    const campaignId = String(registro['ID da Campanha'] || '');
    const abaEventos = planilha.getSheetByName(DB_SCHEMA_.EVENTOS.name);
    inserirObjetoAba_(abaEventos, {
      'ID do Evento': Utilities.getUuid(),
      'ID da Campanha': campaignId,
      'ID de Rastreamento': token,
      'Tipo de Evento': 'OPEN',
      'Data/Hora': agora,
      'Observação': 'Carregamento do pixel; pode ser automático ou realizado por proxy do cliente de e-mail.'
    });

    const abaDetalhes = planilha.getSheetByName(CONFIG.ABA_RASTREAMENTO);
    if (abaDetalhes) {
      const linha = localizarLinhaPorValor_(abaDetalhes, 'ID', token);
      if (linha > 0) {
        const cabecalhos = obterCabecalhosAba_(abaDetalhes);
        const indiceStatus = obterIndiceCabecalho_(cabecalhos, 'Status Atual');
        const indicePrimeira = obterIndiceCabecalho_(cabecalhos, 'Data da 1ª Leitura');
        const indiceUltima = obterIndiceCabecalho_(cabecalhos, 'Data da Última Leitura');
        const indiceEventos = obterIndiceCabecalho_(cabecalhos, 'Aberturas (eventos)');
        if (indiceStatus >= 0) abaDetalhes.getRange(linha, indiceStatus + 1).setValue('ABERTO');
        if (indicePrimeira >= 0) abaDetalhes.getRange(linha, indicePrimeira + 1).setValue(primeiraAbertura || agora);
        if (indiceUltima >= 0) abaDetalhes.getRange(linha, indiceUltima + 1).setValue(agora);
        if (indiceEventos >= 0) abaDetalhes.getRange(linha, indiceEventos + 1).setValue(totalEventos);
        abaDetalhes.getRange(linha, 1, 1, abaDetalhes.getLastColumn()).setBackground(CONFIG.CORES.ABERTO);
      }
    }

    const abaCampanhas = planilha.getSheetByName(DB_SCHEMA_.CAMPANHAS.name);
    const campanha = obterObjetoPorValor_(abaCampanhas, 'ID da Campanha', campaignId);
    if (campanha) {
      const eventosCampanha = numeroSeguro_(campanha['Eventos de Pixel']) + 1;
      const aberturasCampanha = numeroSeguro_(campanha['Aberturas Únicas']) + (primeiraAbertura ? 0 : 1);
      const mensagens = numeroSeguro_(campanha['Mensagens Aceitas']);
      atualizarObjetoPorValor_(abaCampanhas, 'ID da Campanha', campaignId, {
        'Aberturas Únicas': aberturasCampanha,
        'Eventos de Pixel': eventosCampanha,
        'Taxa de Abertura (%)': campanha.Status === 'EM_ANDAMENTO' || !mensagens ? '' : aberturasCampanha / mensagens
      });
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

function getTrackingUrl() {
  const propriedades = PropertiesService.getScriptProperties();
  if (propriedades.getProperty('AUTOEMAIL_TRACKING_DISABLED') === 'true') return '';
  const configurada = propriedades.getProperty('URL_WEB_APP');
  if (configurada) return configurada;
  try {
    return ScriptApp.getService().getUrl() || CONFIG.URL_WEB_APP || '';
  } catch (erro) {
    return CONFIG.URL_WEB_APP || '';
  }
}

function setTrackingUrl(url) {
  const valor = String(url == null ? '' : url).trim();
  const propriedades = PropertiesService.getScriptProperties();
  if (!valor) {
    propriedades.deleteProperty('URL_WEB_APP');
    propriedades.setProperty('AUTOEMAIL_TRACKING_DISABLED', 'true');
    return { success: true, enabled: false, message: 'Rastreamento desativado para novos envios.' };
  }
  const padraoUrlAppsScript = /^https:\/\/script\.google\.com\/(?:macros\/s\/[^/]+\/exec|a\/macros\/[^/]+\/s\/[^/]+\/exec)(?:\?.*)?$/i;
  if (!padraoUrlAppsScript.test(valor)) {
    return { success: false, message: 'Informe a URL publicada do Apps Script, terminada em /exec.' };
  }
  propriedades.setProperty('URL_WEB_APP', valor);
  propriedades.deleteProperty('AUTOEMAIL_TRACKING_DISABLED');
  return { success: true, enabled: true, message: 'URL de rastreamento salva.' };
}
