/**
 * AUTO EMAIL 4.0 — configuração, menu e abertura da interface.
 * Os nomes das abas legadas foram preservados para manter compatibilidade com v3.2.
 */
const CONFIG = Object.freeze({
  VERSAO: '4.0.0',
  ABA_TEMPLATES: '💾 Meus Templates',
  ABA_RELATORIO: '📊 Relatório de Envios',
  ABA_RASTREAMENTO: '📊 Rastreamento Detalhado',
  ABA_DADOS: 'Contatos',
  NOME_REMETENTE: 'JR Contabilidade',
  DELAY_ENVIO_MS: 1500,
  MISSING_VAR_FALLBACK: '',
  LIMITE_TEMPLATE_PLANILHA: 45000,
  LIMITE_ANEXOS_BYTES: 20 * 1024 * 1024,
  LIMITE_ANEXO_BYTES: 10 * 1024 * 1024,
  MAX_DESTINATARIOS_POR_MENSAGEM: 100,
  URL_WEB_APP: '',
  CHAVE_PLANILHA_ATIVA: 'AE_SPREADSHEET_ID',
  CHAVE_MIGRACAO_V4: 'AE_V4_LEGACY_MIGRATION_DONE',
  ABAS_BANCO: Object.freeze({
    CAMPANHAS: 'AE • Campanhas',
    DESTINATARIOS: 'AE • Destinatários',
    EVENTOS: 'AE • Eventos'
  }),
  CORES: Object.freeze({
    PRIMARIA: '#15324b',
    DESTAQUE: '#087e8b',
    SUCESSO: '#d9ead3',
    ERRO: '#f4cccc',
    ABERTO: '#d8eff0',
    FUNDO_CABECALHO: '#15324b',
    TEXTO_CABECALHO: '#ffffff'
  })
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Auto Email 4.0')
    .addItem('📊 Abrir dashboard e painel', 'abrirPainel')
    .addSeparator()
    .addItem('🗄️ Inicializar banco de dados', 'inicializarBancoDeDados')
    .addItem('🛠️ Verificar rastreamento de abertura', 'mostrarUrlRastreamento')
    .addToUi();
  ocultarAbaTemplates();
}

function abrirPainel() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  if (planilha) {
    PropertiesService.getScriptProperties().setProperty(CONFIG.CHAVE_PLANILHA_ATIVA, planilha.getId());
  }

  const template = HtmlService.createTemplateFromFile('Painel');
  template.userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  const html = template.evaluate()
    .setWidth(1120)
    .setHeight(720)
    .setTitle('Auto Email 4.0 — Central de comunicação');

  SpreadsheetApp.getUi().showModalDialog(html, 'Auto Email 4.0');
}

function obterFotoPerfilGoogle_(emailUsuario) {
  try {
    const resposta = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v2/userinfo?alt=json', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resposta.getResponseCode() !== 200) return '';

    const perfil = JSON.parse(resposta.getContentText() || '{}');
    const emailPerfil = String(perfil.email || '').trim();
    if (emailUsuario && emailPerfil && emailUsuario.toLowerCase() !== emailPerfil.toLowerCase()) return '';

    const foto = String(perfil.picture || '').trim();
    return /^https:\/\/([a-z0-9-]+\.)*googleusercontent\.com\//i.test(foto) ? foto : '';
  } catch (erro) {
    console.warn('Não foi possível carregar a foto do perfil Google: ' + erro);
    return '';
  }
}

function getUserProfilePhoto() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  return obterFotoPerfilGoogle_(email);
}

function inicializarBancoDeDados() {
  try {
    ensureDatabaseAndMigrate_();
    ocultarAbasBanco_();
    const planilha = obterPlanilha_();
    planilha.toast('Banco de dados pronto. Os dados históricos disponíveis foram preservados.', 'Auto Email 4.0', 6);
    SpreadsheetApp.getUi().alert(
      'Banco de dados inicializado',
      'A base estruturada foi criada nesta planilha. As abas técnicas ficam ocultas; os relatórios legados continuam disponíveis.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (erro) {
    SpreadsheetApp.getUi().alert('Não foi possível inicializar o banco de dados', erro.message || String(erro), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function mostrarUrlRastreamento() {
  let url = '';
  try {
    url = getTrackingUrl();
  } catch (erro) {
    console.error('Falha ao consultar URL de rastreamento: ' + erro);
  }

  const mensagem = url
    ? 'Rastreamento disponível. URL do Web App:\n\n' + url + '\n\nUse o deployment ativo do Apps Script. A medição depende de o cliente de e-mail carregar imagens externas.'
    : 'O rastreamento ainda não está ativo. Implante o projeto como App da Web e configure a URL no dashboard.\n\nA planilha registra os envios mesmo sem o pixel; somente as aberturas não serão mensuráveis.';
  SpreadsheetApp.getUi().alert('Rastreamento de abertura', mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
}

function ocultarAbaTemplates() {
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    if (!planilha) return;
    const aba = planilha.getSheetByName(CONFIG.ABA_TEMPLATES);
    if (aba && !aba.isSheetHidden()) aba.hideSheet();
  } catch (erro) {
    // A criação/ocultação é repetida quando o usuário inicializa os recursos pelo menu.
    console.warn('Não foi possível ocultar a aba de templates: ' + erro);
  }
}

function ocultarAbasBanco_() {
  const planilha = obterPlanilha_();
  Object.keys(CONFIG.ABAS_BANCO).forEach(function (chave) {
    const aba = planilha.getSheetByName(CONFIG.ABAS_BANCO[chave]);
    if (aba && !aba.isSheetHidden()) {
      try { aba.hideSheet(); } catch (erro) { console.warn(erro); }
    }
  });
}

function obterPlanilha_() {
  const propriedades = PropertiesService.getScriptProperties();
  const id = propriedades.getProperty(CONFIG.CHAVE_PLANILHA_ATIVA);
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) {
    if (id !== ativa.getId()) propriedades.setProperty(CONFIG.CHAVE_PLANILHA_ATIVA, ativa.getId());
    return ativa;
  }

  if (id) return SpreadsheetApp.openById(id);
  throw new Error('Não foi possível localizar a planilha vinculada. Abra o painel e inicialize o banco de dados antes de publicar o Web App.');
}

function getAppStatus() {
  try {
    const planilha = obterPlanilha_();
    ensureDatabaseAndMigrate_();
    let cotaRestante = null;
    try { cotaRestante = MailApp.getRemainingDailyQuota(); } catch (erro) { /* será solicitada autorização ao enviar */ }
    const url = getTrackingUrl();
    const abaDados = getSheetDados();
    return {
      success: true,
      version: CONFIG.VERSAO,
      spreadsheetName: planilha.getName(),
      dataSheetName: abaDados ? abaDados.getName() : '',
      trackingEnabled: Boolean(url),
      trackingUrl: url || '',
      remainingQuota: cotaRestante,
      storage: 'Google Sheets'
    };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro) };
  }
}
