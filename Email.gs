/**
 * AUTO EMAIL 4.0 — leitura da planilha, preview e envio de e-mails.
 * Preserva as funções públicas da interface Auto Email v3.2.
 */
function getSheetDados() {
  const planilha = obterPlanilha_();
  if (CONFIG.ABA_DADOS) {
    const configurada = planilha.getSheetByName(CONFIG.ABA_DADOS);
    if (configurada) return configurada;
  }

  const ativa = planilha.getActiveSheet();
  if (ativa && !abaDeSistema_(ativa.getName()) && identificarColunaEmail(ativa.getRange(1, 1, 1, Math.max(ativa.getLastColumn(), 1)).getDisplayValues()[0]) >= 0) {
    return ativa;
  }

  const candidata = planilha.getSheets().find(function (aba) {
    if (abaDeSistema_(aba.getName()) || aba.getLastColumn() < 1) return false;
    return identificarColunaEmail(aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0]) >= 0;
  });
  if (candidata) return candidata;
  return ativa || planilha.getSheets()[0];
}

function abaDeSistema_(nome) {
  return nome === CONFIG.ABA_TEMPLATES || nome === CONFIG.ABA_RELATORIO || nome === CONFIG.ABA_RASTREAMENTO ||
    Object.keys(CONFIG.ABAS_BANCO).some(function (chave) { return CONFIG.ABAS_BANCO[chave] === nome; });
}

function getSheetDataDetalhada_() {
  const aba = getSheetDados();
  if (!aba || aba.getLastRow() < 1 || aba.getLastColumn() < 1) {
    return { sheet: aba, headers: [], rows: [] };
  }
  const totalLinhas = aba.getLastRow();
  const totalColunas = aba.getLastColumn();
  const valores = aba.getRange(1, 1, totalLinhas, totalColunas).getValues();
  const cabecalhos = valores[0];
  const rows = [];
  for (let indice = 1; indice < valores.length; indice++) {
    const numeroLinha = indice + 1;
    if (aba.getFilter() && aba.isRowHiddenByFilter(numeroLinha)) continue;
    rows.push({ values: valores[indice], sourceRow: numeroLinha });
  }
  return { sheet: aba, headers: cabecalhos, rows: rows };
}

function getSheetDataComFiltro() {
  const dados = getSheetDataDetalhada_();
  if (!dados.headers.length) return [];
  return [dados.headers].concat(dados.rows.map(function (linha) { return linha.values; }));
}

function getVariaveisDisponiveis() {
  const dados = getSheetDataDetalhada_();
  return dados.headers.map(function (header) { return String(header == null ? '' : header).replace(/\s+/g, ' ').trim(); }).filter(Boolean);
}

function contarDestinatarios() {
  const dados = getSheetDataDetalhada_();
  return dados.rows.length;
}

function gerarPreview(htmlEditor, assuntoEditor, index) {
  try {
    const dados = getSheetDataDetalhada_();
    if (!dados.headers.length || !dados.rows.length) return { error: 'Nenhuma linha visível para pré-visualização.' };
    const maxIndex = dados.rows.length - 1;
    const safeIndex = Math.max(0, Math.min(Number(index) || 0, maxIndex));
    const linha = dados.rows[safeIndex].values;
    const indiceEmail = identificarColunaEmail(dados.headers);
    const celulaEmail = indiceEmail >= 0 ? String(linha[indiceEmail] || '') : '';
    const emailsValidos = dividirEmails_(celulaEmail);
    const cc = emailsValidos.slice(1);
    return {
      index: safeIndex,
      total: dados.rows.length,
      para: emailsValidos[0] || (indiceEmail < 0 ? 'Coluna de e-mail não encontrada' : 'Endereço vazio ou inválido'),
      cc: cc.length ? cc.join(', ') : '-',
      htmlPreview: personalizarHtml(htmlEditor, dados.headers, linha),
      assuntoPreview: personalizarTexto(assuntoEditor, dados.headers, linha)
    };
  } catch (erro) {
    return { error: erro.message || String(erro) };
  }
}

function enviarEmailTeste(html, assunto, destinatario, anexosBase64, imagensInline) {
  try {
    if (!isValidEmail(destinatario)) throw new Error('Informe um endereço de e-mail válido para o teste.');
    const dados = getSheetDataDetalhada_();
    let corpoFinal = String(html || '');
    let assuntoFinal = String(assunto || '');
    if (dados.rows.length) {
      corpoFinal = personalizarHtml(corpoFinal, dados.headers, dados.rows[0].values);
      assuntoFinal = personalizarTexto(assuntoFinal, dados.headers, dados.rows[0].values);
    }
    if (!assuntoFinal.trim()) throw new Error('O assunto do e-mail não pode ficar vazio.');

    const anexos = processarBlobs(anexosBase64 || []);
    const inline = prepararBlobsInline(imagensInline || []);
    const corpoDeTeste = corpoFinal + '<br><br><hr><p style="color:#64748b;font-size:11px;font-family:Arial,sans-serif">[MODO TESTE] Exemplo personalizado com a primeira linha visível da planilha. Este teste não é contabilizado no relatório de campanha.</p>';
    const opcoes = { htmlBody: corpoDeTeste, name: CONFIG.NOME_REMETENTE };
    if (anexos.length) opcoes.attachments = anexos;
    if (Object.keys(inline).length) opcoes.inlineImages = inline;

    GmailApp.sendEmail(destinatario.trim(), '[TESTE] ' + assuntoFinal, 'Seu cliente de e-mail não suporta HTML.', opcoes);
    return { success: true, message: 'Teste enviado para ' + destinatario.trim() + '.' };
  } catch (erro) {
    console.error(erro);
    return { success: false, message: 'Falha no envio de teste: ' + (erro.message || erro) };
  }
}

function executarEnvioEmMassa(htmlFinal, assuntoFinal, ccGlobalString, anexosBase64, imagensInline) {
  let campaignId = '';
  let sucessoParcial = 0;
  let falhasParciais = 0;
  let rastreamentoCampanha = false;
  const inicioExecucao = new Date();
  try {
    ensureDatabaseAndMigrate_();
    const dados = getSheetDataDetalhada_();
    if (!dados.rows.length) return { success: false, message: 'Não há linhas visíveis para envio.' };
    const indiceEmail = identificarColunaEmail(dados.headers);
    if (indiceEmail < 0) return { success: false, message: 'Não encontrei a coluna "Email", "E-mail" ou "E mail" na primeira linha.' };
    if (!String(assuntoFinal || '').trim()) return { success: false, message: 'O assunto do e-mail não pode ficar vazio.' };

    const emailsCcInvalidos = obterEmailsInvalidos_(ccGlobalString || '');
    if (emailsCcInvalidos.length) return { success: false, message: 'Revise o CC global. Endereço(s) inválido(s): ' + emailsCcInvalidos.join(', ') };
    const ccGlobal = dividirEmails_(ccGlobalString || '');
    const anexos = processarBlobs(anexosBase64 || []);
    const imagens = prepararBlobsInline(imagensInline || []);
    const indiceEmpresa = encontrarColunaEmpresa_(dados.headers);

    const preparados = dados.rows.map(function (registro) {
      const valores = registro.values;
      const celula = String(valores[indiceEmail] == null ? '' : valores[indiceEmail]).trim();
      const invalidos = obterEmailsInvalidos_(celula);
      const emails = dividirEmails_(celula);
      const erroEndereco = invalidos.length
        ? 'Endereço(s) inválido(s): ' + invalidos.join(', ')
        : (!emails.length ? 'Nenhum endereço de e-mail válido encontrado.' : '');
      const destinatario = emails[0] || '';
      const cc = combinarEmails_(emails.slice(1).concat(ccGlobal));
      let erro = erroEndereco;
      if (!erro && cc.length + 1 > CONFIG.MAX_DESTINATARIOS_POR_MENSAGEM) {
        erro = 'A mensagem excede o limite de segurança de ' + CONFIG.MAX_DESTINATARIOS_POR_MENSAGEM + ' destinatários.';
      }
      return {
        sourceRow: registro.sourceRow,
        to: destinatario,
        cc: cc,
        erro: erro,
        assunto: personalizarTexto(assuntoFinal, dados.headers, valores),
        html: personalizarHtml(htmlFinal, dados.headers, valores),
        empresa: indiceEmpresa >= 0 && valores[indiceEmpresa] ? String(valores[indiceEmpresa]) : '',
        recipients: erro ? 0 : 1 + cc.length
      };
    });

    const destinatariosEstimados = preparados.reduce(function (soma, item) { return soma + item.recipients; }, 0);
    let cotaRestante = null;
    try { cotaRestante = MailApp.getRemainingDailyQuota(); } catch (erro) {
      return { success: false, message: 'Não foi possível consultar a cota do Gmail. Autorize as permissões de envio e tente novamente.' };
    }
    if (destinatariosEstimados > cotaRestante) {
      return {
        success: false,
        message: 'Envio cancelado por segurança. A cota restante é de ' + cotaRestante +
          ' destinatário(s), mas este disparo pode consumir ' + destinatariosEstimados +
          ' (incluindo CCs). Corrija a seleção ou tente novamente após a renovação da cota.'
      };
    }

    const trackingUrl = getTrackingUrl();
    const trackingAtivo = Boolean(trackingUrl);
    rastreamentoCampanha = trackingAtivo;
    campaignId = Utilities.getUuid();
    registrarInicioCampanha_({
      id: campaignId,
      iniciadaEm: inicioExecucao,
      assunto: String(assuntoFinal),
      planilhaOrigem: dados.sheet ? dados.sheet.getName() : '',
      totalLinhas: preparados.length,
      destinatariosEstimados: destinatariosEstimados,
      cotaAntesEnvio: cotaRestante,
      rastreamentoAtivo: trackingAtivo
    });

    let sucesso = 0;
    let falhas = 0;
    let destinatariosAceitos = 0;
    const detalhesErros = [];
    let interromperPorCota = false;

    preparados.forEach(function (item, indice) {
      let registroDb = null;
      let linhaVisivel = null;
      const trackingId = trackingAtivo && !item.erro ? Utilities.getUuid() : '';
      const idInterno = Utilities.getUuid();
      const erroPreEnvio = interromperPorCota
        ? 'Não processado porque o Gmail interrompeu o disparo ao atingir um limite.'
        : item.erro;
      const registro = {
        'ID Interno': idInterno,
        'ID da Campanha': campaignId,
        'ID de Rastreamento': trackingId,
        'Data/Hora Envio': new Date(),
        'Destinatário principal': item.to,
        'CC': item.cc.join('; '),
        'Empresa': item.empresa,
        'Assunto': item.assunto,
        'Status': erroPreEnvio ? 'FALHA' : 'PENDENTE',
        'Data 1ª Abertura': '',
        'Data Última Abertura': '',
        'Eventos de Pixel': 0,
        'Erro': erroPreEnvio,
        'Linha de Origem': item.sourceRow,
        'Fonte': 'LIVE'
      };

      try {
        const gravado = registrarDestinatario_(registro);
        registroDb = gravado.databaseRow;
        linhaVisivel = gravado.reportRow;
        if (erroPreEnvio) {
          falhas++;
          detalhesErros.push('Linha ' + item.sourceRow + ': ' + erroPreEnvio);
        } else {
          let corpo = item.html;
          if (trackingId) corpo += criarPixelRastreamento_(trackingUrl, trackingId);
          const opcoes = { htmlBody: corpo, name: CONFIG.NOME_REMETENTE };
          if (item.cc.length) opcoes.cc = item.cc.join(',');
          if (anexos.length) opcoes.attachments = anexos;
          if (Object.keys(imagens).length) opcoes.inlineImages = imagens;
          GmailApp.sendEmail(item.to, item.assunto, 'Seu cliente de e-mail não suporta HTML.', opcoes);
          sucesso++;
          destinatariosAceitos += item.recipients;
          try { atualizarStatusDestinatario_(idInterno, 'ENVIADO', '', linhaVisivel); }
          catch (falhaStatus) {
            detalhesErros.push('Aviso na linha ' + item.sourceRow + ': a mensagem foi aceita pelo Gmail, mas o status na planilha não pôde ser atualizado.');
            console.error('Falha ao atualizar status depois do envio: ' + falhaStatus);
          }
        }
      } catch (erro) {
        const mensagem = String(erro && erro.message ? erro.message : erro);
        falhas++;
        detalhesErros.push('Linha ' + item.sourceRow + ': ' + mensagem);
        if (registroDb) {
          try { atualizarStatusDestinatario_(idInterno, 'FALHA', mensagem, linhaVisivel); } catch (falhaRegistro) { console.error(falhaRegistro); }
        }
        if (/limit exceeded|too many times|service invoked too many times|quota/i.test(mensagem)) {
          interromperPorCota = true;
          detalhesErros.push('O processamento foi interrompido por limite/cota do Google; as próximas linhas foram registradas sem envio.');
        }
      }
      sucessoParcial = sucesso;
      falhasParciais = falhas;

      if (indice % 5 === 0) {
        try { dados.sheet.toast('Processando ' + (indice + 1) + ' de ' + preparados.length + ' linhas…', 'Auto Email 4.0', 3); } catch (ignorado) { /* painel pode estar em execução fora da planilha */ }
      }
      if (!erroPreEnvio && indice < preparados.length - 1 && !interromperPorCota && CONFIG.DELAY_ENVIO_MS > 0) {
        Utilities.sleep(CONFIG.DELAY_ENVIO_MS);
      }
    });

    const duracao = Math.max(0, Math.round((new Date().getTime() - inicioExecucao.getTime()) / 1000));
    const registrosCampanha = listarObjetosAba_(obterPlanilha_().getSheetByName(DB_SCHEMA_.DESTINATARIOS.name))
      .filter(function (item) { return String(item['ID da Campanha'] || '') === campaignId; });
    const aberturasUnicas = registrosCampanha.filter(function (item) { return Boolean(item['Data 1ª Abertura']); }).length;
    const eventosPixel = registrosCampanha.reduce(function (soma, item) { return soma + numeroSeguro_(item['Eventos de Pixel']); }, 0);
    const notas = trackingAtivo
      ? 'Pixel habilitado. Aberturas são sinais estimados e podem ser afetadas por proxies ou bloqueio de imagens.'
      : 'Pixel não configurado. Os envios e falhas foram armazenados; aberturas não são mensuráveis.';
    const resumo = {
      campaignId: campaignId,
      assunto: String(assuntoFinal),
      total: preparados.length,
      sucesso: sucesso,
      falhas: falhas,
      detalhes: limitarTexto_(detalhesErros.join(' | ') || 'Sem erros registrados.', 45000),
      aberturasUnicas: aberturasUnicas,
      rastreamentoAtivo: trackingAtivo
    };

    finalizarRegistroCampanha_(campaignId, {
      sucesso: sucesso,
      falhas: falhas,
      aberturasUnicas: aberturasUnicas,
      eventosPixel: eventosPixel,
      rastreamentoAtivo: trackingAtivo,
      duracaoSegundos: duracao,
      status: 'CONCLUÍDA',
      notas: notas
    });

    let recomendacao = '';
    let falhaRelatorio = '';
    try { recomendacao = registrarRelatorioCampanha_(resumo); }
    catch (erro) { falhaRelatorio = erro.message || String(erro); console.error('Falha ao atualizar o relatório visível: ' + erro); }

    try { dados.sheet.toast('Disparo concluído: ' + sucesso + ' mensagem(ns) aceita(s), ' + falhas + ' falha(s).', 'Auto Email 4.0', 6); } catch (ignorado) { /* notificação opcional */ }
    let mensagem = 'Processo concluído.\n\nMensagens aceitas pelo Gmail: ' + sucesso + '\nFalhas registradas: ' + falhas +
      '\nDestinatários estimados (incluindo CC): ' + destinatariosEstimados + '\nAberturas únicas registradas: ' + aberturasUnicas;
    mensagem += trackingAtivo
      ? '\n\nO rastreamento está ativo. A abertura é um indicador estimado, não uma confirmação de leitura humana.'
      : '\n\nO rastreamento não está configurado; o envio foi registrado, mas a abertura não pode ser medida.';
    mensagem += '\nID da campanha: ' + campaignId;
    if (recomendacao) mensagem += '\n\nLeitura consultiva: ' + recomendacao;
    if (falhaRelatorio) mensagem += '\n\nAviso: o banco foi atualizado, mas o relatório visível não pôde ser atualizado: ' + falhaRelatorio;
    return {
      success: true,
      message: mensagem,
      campaignId: campaignId,
      sent: sucesso,
      failed: falhas,
      estimatedRecipients: destinatariosEstimados,
      acceptedRecipients: destinatariosAceitos,
      openedUnique: aberturasUnicas,
      recommendation: recomendacao
    };
  } catch (erro) {
    console.error(erro);
    if (campaignId) {
      try {
        const abaDestinatarios = obterPlanilha_().getSheetByName(DB_SCHEMA_.DESTINATARIOS.name);
        const registros = listarObjetosAba_(abaDestinatarios).filter(function (item) { return String(item['ID da Campanha'] || '') === campaignId; });
        const aberturas = registros.filter(function (item) { return Boolean(item['Data 1ª Abertura']); }).length;
        const eventos = registros.reduce(function (soma, item) { return soma + numeroSeguro_(item['Eventos de Pixel']); }, 0);
        finalizarRegistroCampanha_(campaignId, {
          sucesso: sucessoParcial,
          falhas: falhasParciais,
          aberturasUnicas: aberturas,
          eventosPixel: eventos,
          rastreamentoAtivo: rastreamentoCampanha,
          duracaoSegundos: Math.max(0, Math.round((new Date().getTime() - inicioExecucao.getTime()) / 1000)),
          status: sucessoParcial ? 'PARCIAL' : 'ERRO',
          notas: 'Processo interrompido: ' + (erro.message || erro) + ' Mensagens aceitas antes da interrupção: ' + sucessoParcial + '.'
        });
        try {
          registrarRelatorioCampanha_({
            campaignId: campaignId, assunto: '', total: sucessoParcial + falhasParciais,
            sucesso: sucessoParcial, falhas: falhasParciais,
            detalhes: 'Processo interrompido: ' + (erro.message || erro),
            aberturasUnicas: aberturas, rastreamentoAtivo: rastreamentoCampanha
          });
        } catch (falhaRelatorio) { console.error(falhaRelatorio); }
      } catch (falhaFinalizacao) { console.error(falhaFinalizacao); }
    }
    const mensagem = 'O processo foi interrompido: ' + (erro.message || erro) +
      (sucessoParcial ? ' Já foram aceitas ' + sucessoParcial + ' mensagem(ns) pelo Gmail e registradas ' + falhasParciais + ' falha(s). Consulte o ID da campanha antes de tentar novamente.' : ' Nenhuma mensagem foi confirmada como enviada.');
    return { success: false, partial: sucessoParcial > 0, message: mensagem, campaignId: campaignId };
  }
}

function criarPixelRastreamento_(url, id) {
  const separador = String(url).indexOf('?') >= 0 ? '&' : '?';
  return '<img src="' + escaparHtml_(url + separador + 'id=' + encodeURIComponent(id)) +
    '" width="1" height="1" alt="" style="display:block;width:1px;height:1px;border:0;opacity:0" />';
}

function encontrarColunaEmpresa_(cabecalhos) {
  const aceitos = new Set(['empresa', 'razao social', 'cliente', 'company', 'organizacao']);
  return (cabecalhos || []).findIndex(function (cabecalho) { return aceitos.has(normalizarCabecalho_(cabecalho)); });
}

function combinarEmails_(lista) {
  const vistos = new Set();
  return (lista || []).filter(function (email) {
    const valor = String(email || '').trim();
    if (!isValidEmail(valor)) return false;
    const chave = valor.toLowerCase();
    if (vistos.has(chave)) return false;
    vistos.add(chave);
    return true;
  });
}

function limitarTexto_(texto, limite) {
  const valor = String(texto == null ? '' : texto);
  if (valor.length <= limite) return valor;
  return valor.substring(0, limite - 40) + ' … (detalhes truncados para caber no relatório)';
}

function processarBlobs(anexosBase64) {
  if (!anexosBase64 || !anexosBase64.length) return [];
  let totalBytes = 0;
  return anexosBase64.map(function (objeto) {
    if (!objeto || !objeto.data) throw new Error('Um dos anexos não contém dados válidos.');
    const bytes = Utilities.base64Decode(String(objeto.data));
    totalBytes += bytes.length;
    if (bytes.length > CONFIG.LIMITE_ANEXO_BYTES) throw new Error('O arquivo "' + String(objeto.name || 'anexo') + '" excede o limite de 10 MB.');
    if (totalBytes > CONFIG.LIMITE_ANEXOS_BYTES) throw new Error('O tamanho total dos anexos excede 20 MB.');
    const mimeType = String(objeto.type || 'application/octet-stream');
    const nome = String(objeto.name || 'anexo');
    return Utilities.newBlob(bytes, mimeType, nome);
  });
}

function prepararBlobsInline(listaImagens) {
  if (!listaImagens || !listaImagens.length) return {};
  const inlineImages = {};
  let totalBytes = 0;
  listaImagens.forEach(function (imagem) {
    if (!imagem || !imagem.data || !imagem.cid) throw new Error('Uma das imagens inseridas está inválida.');
    if (!/^image\//i.test(String(imagem.type || ''))) throw new Error('Imagens inline devem usar um tipo MIME de imagem.');
    const bytes = Utilities.base64Decode(String(imagem.data));
    totalBytes += bytes.length;
    if (bytes.length > 5 * 1024 * 1024) throw new Error('Cada imagem inline deve ter até 5 MB.');
    if (totalBytes > 10 * 1024 * 1024) throw new Error('O tamanho total das imagens inline excede 10 MB.');
    const nomeCid = String(imagem.cid).replace(/[^a-zA-Z0-9_-]/g, '_');
    inlineImages[nomeCid] = Utilities.newBlob(bytes, String(imagem.type), nomeCid);
  });
  return inlineImages;
}

function getAssinaturaGmail() {
  try {
    const resposta = Gmail.Users.Settings.SendAs.list('me');
    const contas = resposta && resposta.sendAs ? resposta.sendAs : [];
    const principal = contas.find(function (conta) { return conta.isPrimary; }) || contas[0];
    if (principal && principal.signature) return { success: true, signature: principal.signature };
    return { success: false, message: 'Nenhuma assinatura encontrada na conta Gmail principal.' };
  } catch (erro) {
    console.error('Falha ao consultar assinatura: ' + erro);
    return { success: false, message: 'Não foi possível consultar a assinatura. Ative o serviço avançado Gmail API e autorize gmail.settings.basic.' };
  }
}
