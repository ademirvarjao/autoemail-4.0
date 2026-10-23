const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const sourceFiles = ['Utils.gs', 'Config.gs', 'Database.gs', 'Templates.gs', 'Email.gs', 'WebApp.gs'];
const combinedAppsScript = sourceFiles.map((name) => fs.readFileSync(path.join(root, name), 'utf8')).join('\n\n');
const sandbox = {
  CONFIG: { MISSING_VAR_FALLBACK: '[não encontrado]' },
  Utilities: {
    formatDate: (date) => date.toISOString().slice(0, 10),
  },
  Session: { getScriptTimeZone: () => 'America/Sao_Paulo' },
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(root, 'Utils.gs'), 'utf8'), sandbox);

test('cabeçalhos são normalizados sem acentos, HTML ou espaços extras', () => {
  assert.equal(sandbox.normalizarCabecalho_('  <b>Razão&nbsp; Social</b>  '), 'razao social');
  assert.equal(sandbox.normalizarCabecalho_('E-mail'), 'e-mail');
});

test('identifica os formatos de coluna de e-mail usados em planilhas', () => {
  assert.equal(sandbox.identificarColunaEmail(['Empresa', 'E-mail']), 1);
  assert.equal(sandbox.identificarColunaEmail(['Razão Social', 'E mail principal']), 1);
  assert.equal(sandbox.identificarColunaEmail(['Contato', 'Email address']), 1);
  assert.equal(sandbox.identificarColunaEmail(['Nome', 'Endereço']), -1);
});

test('valida, divide e deduplica destinatários separados por vírgula ou ponto e vírgula', () => {
  assert.equal(sandbox.isValidEmail('ana.silva@exemplo.com'), true);
  assert.equal(sandbox.isValidEmail('ana@@exemplo.com'), false);
  assert.deepEqual(
    Array.from(sandbox.dividirEmails_('ana@exemplo.com; Bia@exemplo.com, ana@exemplo.com; inválido')),
    ['ana@exemplo.com', 'Bia@exemplo.com'],
  );
  assert.deepEqual(Array.from(sandbox.obterEmailsInvalidos_('ok@exemplo.com; errado;')), ['errado']);
});

test('substitui variáveis ignorando acentos e maiúsculas/minúsculas', () => {
  const resultado = sandbox.personalizarTexto('Olá {{ NOME }} — {{ Razão Social }} — {{ ausente }}', ['Nome', 'Razão Social'], ['Ana', 'ACME']);
  assert.equal(resultado, 'Olá Ana — ACME — [não encontrado]');
});

test('escapa valores da planilha quando a variável é inserida no corpo HTML', () => {
  const resultado = sandbox.personalizarHtml('<p>{{ Nome }}</p>', ['Nome'], ['<img src=x onerror="alert(1)"> &']);
  assert.equal(resultado, '<p>&lt;img src=x onerror=&quot;alert(1)&quot;&gt; &amp;</p>');
});

test('lê datas legadas brasileiras e protege valores CSV contra fórmulas', () => {
  assert.equal(sandbox.parseDataLegada_('23/10/2026 14:05:09').toISOString(), '2026-10-23T14:05:09.000Z');
  assert.equal(sandbox.parseDataLegada_('data inválida'), null);
  assert.equal(sandbox.csvCell_('=HYPERLINK("x")'), '"\'=HYPERLINK(""x"")"');
});

test('dashboard calcula abertura somente sobre mensagens rastreáveis', () => {
  const dbSandbox = {
    CONFIG: {
      MISSING_VAR_FALLBACK: '',
      ABAS_BANCO: { CAMPANHAS: 'AE • Campanhas', DESTINATARIOS: 'AE • Destinatários', EVENTOS: 'AE • Eventos' },
      CORES: { PRIMARIA: '#15324b', DESTAQUE: '#087e8b', SUCESSO: '#d9ead3', ERRO: '#f4cccc', ABERTO: '#d8eff0', FUNDO_CABECALHO: '#15324b', TEXTO_CABECALHO: '#fff' },
      CHAVE_MIGRACAO_V4: 'migration', CHAVE_PLANILHA_ATIVA: 'sheet',
      ABA_RELATORIO: '📊 Relatório de Envios', ABA_RASTREAMENTO: '📊 Rastreamento Detalhado', ABA_TEMPLATES: '💾 Meus Templates'
    },
    Utilities: { formatDate: (date) => date.toISOString().slice(0, 10) },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo' }
  };
  vm.createContext(dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Utils.gs'), 'utf8'), dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Database.gs'), 'utf8'), dbSandbox);
  const campaigns = [
    { id: 'campanha-a', mensagensAceitas: 8, falhas: 2, linhas: 10, destinatariosEstimados: 12, rastreamentoAtivo: true },
    { id: 'campanha-b', mensagensAceitas: 3, falhas: 1, linhas: 4, destinatariosEstimados: 4, rastreamentoAtivo: false }
  ];
  const recipients = [
    { 'ID da Campanha': 'campanha-a', 'Data 1ª Abertura': '2026-10-21T10:00:00.000Z', 'Eventos de Pixel': 3 },
    { 'ID da Campanha': 'campanha-a', 'Data 1ª Abertura': '', 'Eventos de Pixel': 0 },
    { 'ID da Campanha': 'campanha-b', 'Data 1ª Abertura': '2026-10-21T11:00:00.000Z', 'Eventos de Pixel': 1 }
  ];
  const resumo = dbSandbox.resumirCampanhas_(campaigns, recipients);
  assert.equal(resumo.mensagensAceitas, 11);
  assert.equal(resumo.falhas, 3);
  assert.equal(resumo.aberturasUnicas, 1);
  assert.equal(resumo.eventosPixel, 3);
  assert.equal(resumo.mensagensRastreaveis, 8);
  assert.equal(resumo.taxaAbertura, 12.5);
});

test('relatório consultivo explicita quando a abertura não pode ser medida', () => {
  const dbSandbox = {
    CONFIG: { ABAS_BANCO: { CAMPANHAS: 'C', DESTINATARIOS: 'D', EVENTOS: 'E' }, CORES: {} },
    Utilities: { formatDate: (date) => date.toISOString().slice(0, 10) },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo' }
  };
  vm.createContext(dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Utils.gs'), 'utf8'), dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Database.gs'), 'utf8'), dbSandbox);
  const texto = dbSandbox.gerarLeituraConsultiva_({ sucesso: 8, falhas: 2, rastreamentoAtivo: false, aberturasUnicas: 0 });
  assert.match(texto, /Ative o Web App/i);
  assert.match(texto, /não comprovam leitura humana/i);
  assert.match(texto, /higienização dos e-mails/i);
});

test('todos os arquivos Apps Script e o JavaScript da interface têm sintaxe válida', () => {
  assert.doesNotThrow(() => new vm.Script(combinedAppsScript, { filename: 'AppsScript.bundle.gs' }));
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/gi)];
  assert.equal(scripts.length, 1, 'a interface deve conter um único bloco JavaScript inline');
  assert.doesNotThrow(() => new vm.Script(scripts[0][1], { filename: 'Painel.inline.js' }));
});

test('painel remove os avisos decorativos, amplia a tipografia e mantém os relatórios', () => {
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const css = html.split('<style>', 2)[1].split('</style>', 1)[0];
  assert.doesNotMatch(html, /class="sidebar-tip"|id="recommendations"|renderizarRecomendacoes/);
  assert.match(css, /body \{[^}]*font-size: 15px;/);
  assert.match(html, /id="campaignRows"/);
  assert.match(html, /id="trackingRows"/);
  assert.match(html, /<th>Próximo passo<\/th>/);
  assert.match(html, /campanha\.recomendacao/);
});

test('estado vazio do gráfico usa valores numéricos e informa o intervalo correspondente', () => {
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const obterFuncao = (nome) => {
    const funcao = html.match(new RegExp(`    function ${nome}\\([^\\n]*\\) \\{[\\s\\S]*?\\n    \\}`));
    assert.ok(funcao, `função ${nome} encontrada no painel`);
    return funcao[0];
  };
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext([
    obterFuncao('valorTendencia'),
    obterFuncao('haAtividadeNaTendencia'),
    obterFuncao('mensagemSemAtividadeNaTendencia')
  ].join('\n\n'), sandbox);

  assert.equal(sandbox.haAtividadeNaTendencia([
    { enviados: '0', aberturas: '0' },
    { enviados: 0, aberturas: null }
  ]), false);
  assert.equal(sandbox.haAtividadeNaTendencia([{ enviados: '2', aberturas: '0' }]), true);
  assert.equal(sandbox.mensagemSemAtividadeNaTendencia([
    { label: '17/10' }, { label: '23/10' }
  ]), 'Nenhuma atividade registrada entre 17/10 e 23/10.');
  assert.match(html, /\.chart-wrap \{[^}]*position: relative;/);
  assert.doesNotMatch(html, /Ainda não há atividade registrada/);
});

test('foto de perfil é associada ao e-mail ativo e usa somente imagem hospedada pelo Google', () => {
  const config = fs.readFileSync(path.join(root, 'Config.gs'), 'utf8');
  const funcao = config.match(/function obterFotoPerfilGoogle_\(emailUsuario\) \{[\s\S]*?\n\}/);
  assert.ok(funcao, 'função de consulta da foto deve existir');

  let perfil = { email: 'ana@example.com', picture: 'https://lh3.googleusercontent.com/a/profile-photo' };
  const sandbox = {
    ScriptApp: { getOAuthToken: () => 'oauth-token' },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify(perfil) }) },
    console: { warn: () => {} }
  };
  vm.createContext(sandbox);
  vm.runInContext(funcao[0], sandbox);
  assert.equal(sandbox.obterFotoPerfilGoogle_('ana@example.com'), perfil.picture);
  assert.equal(sandbox.obterFotoPerfilGoogle_('outra@example.com'), '');
  perfil = { email: 'ana@example.com', picture: 'https://example.com/photo.jpg' };
  assert.equal(sandbox.obterFotoPerfilGoogle_('ana@example.com'), '');
});

test('rodapé exibe o e-mail completo e remove a legenda de armazenamento', () => {
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const css = html.split('<style>', 2)[1].split('</style>', 1)[0];
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
  const config = fs.readFileSync(path.join(root, 'Config.gs'), 'utf8');
  assert.match(html, /data-user-email="<\?= userEmail \?>"/);
  assert.match(html, /id="footerAvatarImage"/);
  assert.match(html, /getUserProfilePhoto\(\)/);
  assert.match(config, /function getUserProfilePhoto\(\)/);
  assert.match(css, /\.footer-copy strong \{[^}]*overflow-wrap: anywhere;[^}]*white-space: normal;/);
  assert.doesNotMatch(html, /Dados persistidos na planilha/);
  assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/userinfo.profile'));
  assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/script.external_request'));
});

test('a pré-visualização mostra progresso até carregar o iframe e trata falhas', () => {
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const css = html.split('<style>', 2)[1].split('</style>', 1)[0];
  assert.match(html, /id="previewLoading" class="preview-loading" hidden role="status"/);
  assert.match(html, /id="previewPaper" class="preview-paper" aria-busy="false"/);
  assert.match(html, /definirCarregamentoPreview\(true, 'Gerando pré-visualização…'\)/);
  assert.match(html, /state\.previewFrameRequestId === state\.previewRequestId\) definirCarregamentoPreview\(false\)/);
  assert.match(html, /withFailureHandler\(function \(erro\) \{[\s\S]*?definirCarregamentoPreview\(false\);/);
  assert.match(html, /if \(requestId !== state\.previewRequestId\) return;/);
  assert.match(css, /\.preview-paper \{[^}]*position: relative;/);
  assert.match(css, /\.preview-loading\[hidden\] \{ display: none; \}/);
});

test('o corpo do editor evita rolagem horizontal desnecessária e preserva o overflow real', () => {
  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  const css = html.split('<style>', 2)[1].split('</style>', 1)[0];
  assert.match(css, /\.editor-content \{[^}]*overflow-x: auto; overflow-y: auto;[^}]*overflow-wrap: anywhere;/);
  assert.match(css, /\.editor-content img, \.editor-content table[^\{]*\{ max-width: 100%; \}/);
  assert.match(css, /\.editor-content pre \{[^}]*white-space: pre-wrap;/);
});

test('exportações do rastreamento preservam filtros e incluem todos os registros em CSV, XLSX e PDF', () => {
  const exportRows = Array.from({ length: 501 }, (_, index) => ({
    enviadoEm: '2026-10-23T12:00:00.000Z', destinatario: index === 0 ? '=formula@exemplo.com' : 'contato' + index + '@exemplo.com',
    cc: '', empresa: 'Empresa ' + index, assunto: 'Assunto', status: 'FALHA',
    primeiraAbertura: '', ultimaAbertura: '', eventosPixel: 0, idCampanha: 'campanha-' + index,
    erro: '', origem: 'LIVE'
  }));
  let filtrosRecebidos;
  let matrizPlanilha;
  let arquivoRemovido = false;
  let urlPdf = '';
  let urlXlsx = '';
  let opcoesFetchPdf;
  let opcoesFetchXlsx;
  const range = {
    setValues: (values) => { matrizPlanilha = values; return range; }
  };
  const rangeProxy = new Proxy(range, {
    get(target, propriedade) {
      if (propriedade in target) return target[propriedade];
      return () => rangeProxy;
    }
  });
  const aba = {
    getRange: () => rangeProxy,
    getSheetId: () => 42,
    setName: () => {}, setFrozenRows: () => {}, setRowHeight: () => {}, setColumnWidth: () => {}
  };
  const blob = {
    setContentType: (mimeType) => { blob.mimeType = mimeType; return blob; },
    setName: (name) => { blob.name = name; return blob; },
    getBytes: () => blob.bytes
  };
  const driveFile = {
    getAs: () => { throw new Error('XLSX must be exported through the Sheets export endpoint'); },
    setTrashed: (value) => { arquivoRemovido = value; }
  };
  const dbSandbox = {
    CONFIG: { MISSING_VAR_FALLBACK: '' },
    Utilities: {
      formatDate: (date) => date.toISOString().slice(0, 10),
      base64Encode: (bytes) => Buffer.from(bytes).toString('base64')
    },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo' },
    console: { error: () => {}, warn: () => {} },
    SpreadsheetApp: {
      create: () => ({ getId: () => 'temp-sheet-id', getSheets: () => [aba] }),
      flush: () => {}
    },
    DriveApp: { getFileById: () => driveFile },
    ScriptApp: { getOAuthToken: () => 'oauth-token' },
    UrlFetchApp: {
      fetch: (url, options) => {
        if (/format=xlsx/.test(url)) { urlXlsx = url; opcoesFetchXlsx = options; blob.bytes = [80, 75, 3, 4]; }
        else { urlPdf = url; blob.bytes = [37, 80, 68, 70]; opcoesFetchPdf = options; }
        return { getResponseCode: () => 200, getBlob: () => blob };
      }
    }
  };
  vm.createContext(dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Utils.gs'), 'utf8'), dbSandbox);
  vm.runInContext(fs.readFileSync(path.join(root, 'Database.gs'), 'utf8'), dbSandbox);
  dbSandbox.getTrackingReport = (filtros) => {
    filtrosRecebido = JSON.parse(JSON.stringify(filtros));
    return { success: true, total: exportRows.length, rows: exportRows };
  };

  const filtros = { search: 'contato', status: 'FALHA', period: '90' };
  for (const formato of ['xlsx', 'pdf']) {
    arquivoRemovido = false;
    const resultado = formato === 'xlsx'
      ? dbSandbox.exportarRelatorioXLSX(filtros)
      : dbSandbox.exportarRelatorioPDF(filtros);
    assert.equal(resultado.success, true);
    assert.equal(resultado.total, 501);
    assert.equal(resultado.mimeType, formato === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/pdf');
    assert.match(resultado.filename, new RegExp('\\.' + formato + '$'));
    assert.deepEqual(filtrosRecebido, { search: 'contato', status: 'FALHA', period: '90', exportAll: true });
    assert.equal(matrizPlanilha.length, 502, 'a planilha deve conter cabeçalho e os 501 registros');
    assert.equal(matrizPlanilha[1].length, 12);
    assert.equal(matrizPlanilha[1][1], "'=formula@exemplo.com", 'valores com aparência de fórmula devem ser tratados como texto');
    assert.equal(arquivoRemovido, true, 'a planilha temporária deve ser removida após a conversão');
  }
  assert.match(urlXlsx, /format=xlsx/);
  assert.match(urlPdf, /format=pdf/);
  assert.match(urlPdf, /gid=42/);
  assert.equal(opcoesFetchXlsx.headers.Authorization, 'Bearer oauth-token');
  assert.equal(opcoesFetchPdf.headers.Authorization, 'Bearer oauth-token');

  const csv = dbSandbox.exportarRelatorioCSV(filtros);
  assert.equal(csv.success, true);
  assert.equal(csv.total, 501);
  assert.deepEqual(filtrosRecebido, { search: 'contato', status: 'FALHA', period: '90', exportAll: true });

  const html = fs.readFileSync(path.join(root, 'Painel.html'), 'utf8');
  assert.match(html, /exportarRastreamento\('csv'\)/);
  assert.match(html, /exportarRastreamento\('xlsx'\)/);
  assert.match(html, /exportarRastreamento\('pdf'\)/);
  assert.match(html, /baixarArquivoBase64\(res\.base64/);
});
