/**
 * AUTO EMAIL 4.0 — templates na planilha, com conteúdo extenso armazenado no Drive.
 * Mantém os nomes das funções públicas usados pela interface v3.2.
 */
function getPastaTemplates() {
  const nomePasta = 'Templates AutoEmail (Sistema)';
  const pastas = DriveApp.getFoldersByName(nomePasta);
  return pastas.hasNext() ? pastas.next() : DriveApp.createFolder(nomePasta);
}

function obterAbaTemplates_(criar) {
  const planilha = obterPlanilha_();
  let aba = planilha.getSheetByName(CONFIG.ABA_TEMPLATES);
  if (!aba && criar) {
    aba = planilha.insertSheet(CONFIG.ABA_TEMPLATES);
    aba.getRange(1, 1, 1, 4).setValues([['Nome', 'Assunto', 'Conteudo', 'Data Atualização']]);
    aba.getRange('A1:D1').setFontWeight('bold').setBackground(CONFIG.CORES.FUNDO_CABECALHO).setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 220);
    aba.setColumnWidth(2, 320);
    aba.setColumnWidth(3, 420);
    aba.setColumnWidth(4, 170);
  }
  if (aba && !aba.isSheetHidden()) {
    try { aba.hideSheet(); } catch (erro) { /* mantém a operação disponível se for a única aba visível */ }
  }
  return aba;
}

function salvarNovoTemplate(nomeOriginal, nomeNovo, assunto, html) {
  const nome = String(nomeNovo == null ? '' : nomeNovo).trim();
  const assuntoLimpo = String(assunto == null ? '' : assunto);
  const conteudoHtml = String(html == null ? '' : html);
  if (!nome) return { success: false, message: 'Informe um nome para o template.' };
  if (nome.length > 120) return { success: false, message: 'O nome do template deve ter até 120 caracteres.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let novoArquivo = null;
  try {
    const aba = obterAbaTemplates_(true);
    const ultimaLinha = aba.getLastRow();
    const linhas = ultimaLinha > 1 ? aba.getRange(2, 1, ultimaLinha - 1, 4).getValues() : [];
    const original = String(nomeOriginal == null ? '' : nomeOriginal).trim();
    let indiceOriginal = -1;
    let conteudoAntigo = '';

    linhas.forEach(function (linha, indice) {
      const nomeLinha = String(linha[0] || '').trim();
      if (nomeLinha === original) {
        indiceOriginal = indice + 2;
        conteudoAntigo = String(linha[2] || '');
      }
      if (nomeLinha.toLowerCase() === nome.toLowerCase() && nomeLinha !== original) {
        indiceOriginal = -2;
      }
    });

    if (indiceOriginal === -2) return { success: false, message: 'Já existe um template com esse nome.' };
    if (original && indiceOriginal < 0) return { success: false, message: 'O template que seria atualizado não foi encontrado.' };

    let conteudoFinal = conteudoHtml;
    if (conteudoHtml.length > CONFIG.LIMITE_TEMPLATE_PLANILHA) {
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const nomeArquivo = 'Template_' + sanitizeFileName(nome) + '_' + timestamp + '.html';
      novoArquivo = getPastaTemplates().createFile(nomeArquivo, conteudoHtml, MimeType.HTML);
      conteudoFinal = 'DRIVE:' + novoArquivo.getId();
    }

    const registro = [nome, assuntoLimpo, conteudoFinal, new Date()];
    if (indiceOriginal > 0) {
      aba.getRange(indiceOriginal, 1, 1, 4).setValues([registro]);
    } else {
      aba.appendRow(registro);
    }

    if (indiceOriginal > 0 && conteudoAntigo.indexOf('DRIVE:') === 0) {
      const idAntigo = conteudoAntigo.substring('DRIVE:'.length);
      if (idAntigo && (!novoArquivo || novoArquivo.getId() !== idAntigo)) {
        try { DriveApp.getFileById(idAntigo).setTrashed(true); } catch (erro) { console.warn('Não foi possível remover a versão antiga do Drive: ' + erro); }
      }
    }

    return {
      success: true,
      message: indiceOriginal > 0 ? 'Template atualizado com sucesso.' : 'Template salvo com sucesso.',
      name: nome
    };
  } catch (erro) {
    if (novoArquivo) {
      try { novoArquivo.setTrashed(true); } catch (ignorado) { /* evita arquivo órfão quando possível */ }
    }
    console.error(erro);
    return { success: false, message: 'Erro ao salvar o template: ' + (erro.message || erro) };
  } finally {
    lock.releaseLock();
  }
}

function renomearTemplateBackend(nomeAntigo, nomeNovo) {
  const antigo = String(nomeAntigo == null ? '' : nomeAntigo).trim();
  const novo = String(nomeNovo == null ? '' : nomeNovo).trim();
  if (!antigo || !novo) return { success: false, message: 'Informe o nome atual e o novo nome.' };
  if (novo.length > 120) return { success: false, message: 'O nome deve ter até 120 caracteres.' };
  if (antigo.toLowerCase() === novo.toLowerCase()) return { success: true, message: 'O nome permaneceu igual.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const aba = obterAbaTemplates_(false);
    if (!aba) return { success: false, message: 'Base de templates não encontrada.' };
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return { success: false, message: 'Template original não encontrado.' };
    const linhas = aba.getRange(2, 1, ultimaLinha - 1, 4).getValues();
    const existe = linhas.some(function (linha) { return String(linha[0] || '').trim().toLowerCase() === novo.toLowerCase(); });
    if (existe) return { success: false, message: 'Já existe um template com esse nome.' };
    const indice = linhas.findIndex(function (linha) { return String(linha[0] || '').trim() === antigo; });
    if (indice < 0) return { success: false, message: 'Template original não encontrado.' };
    aba.getRange(indice + 2, 1).setValue(novo);
    aba.getRange(indice + 2, 4).setValue(new Date());
    return { success: true, message: 'Template renomeado com sucesso.' };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro) };
  } finally {
    lock.releaseLock();
  }
}

function listarMeusTemplates() {
  const aba = obterAbaTemplates_(false);
  if (!aba || aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues()
    .map(function (linha) { return String(linha[0] || '').trim(); })
    .filter(Boolean);
}

function carregarTemplateEspecifico(nome) {
  try {
    const busca = String(nome == null ? '' : nome).trim();
    if (!busca) throw new Error('Nome de template inválido.');
    const aba = obterAbaTemplates_(false);
    if (!aba || aba.getLastRow() < 2) throw new Error('Base de templates não encontrada.');
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();
    const linha = dados.find(function (item) { return String(item[0] || '').trim() === busca; });
    if (!linha) return { success: false, message: 'Template não encontrado.' };
    return {
      success: true,
      message: 'Template carregado.',
      assunto: String(linha[1] || ''),
      corpo: obterConteudoTemplate_(String(linha[2] || ''))
    };
  } catch (erro) {
    console.error(erro);
    return { success: false, message: 'Erro ao carregar o template: ' + (erro.message || erro) };
  }
}

function obterConteudoTemplate_(conteudo) {
  if (conteudo.indexOf('DRIVE:') !== 0) return conteudo;
  const id = conteudo.substring('DRIVE:'.length);
  if (!id) throw new Error('Referência de template no Drive inválida.');
  return DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8');
}

function excluirTemplate(nome) {
  const busca = String(nome == null ? '' : nome).trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const aba = obterAbaTemplates_(false);
    if (!aba || aba.getLastRow() < 2) return { success: false, message: 'Aba de templates não encontrada.' };
    const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();
    const indice = linhas.findIndex(function (linha) { return String(linha[0] || '').trim() === busca; });
    if (indice < 0) return { success: false, message: 'Template não encontrado para exclusão.' };
    const conteudo = String(linhas[indice][2] || '');
    aba.deleteRow(indice + 2);
    if (conteudo.indexOf('DRIVE:') === 0) {
      try { DriveApp.getFileById(conteudo.substring('DRIVE:'.length)).setTrashed(true); } catch (erro) { console.warn('O arquivo do template não pôde ser removido do Drive: ' + erro); }
    }
    return { success: true, message: 'Template excluído.' };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro) };
  } finally {
    lock.releaseLock();
  }
}

function exportarTemplatesLista(nomesArray) {
  try {
    const selecionados = new Set((Array.isArray(nomesArray) ? nomesArray : []).map(function (nome) { return String(nome); }));
    const aba = obterAbaTemplates_(false);
    if (!aba || aba.getLastRow() < 2) return { success: true, data: [] };
    const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues();
    const templates = [];
    dados.forEach(function (linha) {
      const nome = String(linha[0] || '');
      if (!selecionados.has(nome)) return;
      templates.push({ nome: nome, assunto: String(linha[1] || ''), corpo: obterConteudoTemplate_(String(linha[2] || '')) });
    });
    return { success: true, data: templates };
  } catch (erro) {
    return { success: false, message: erro.message || String(erro) };
  }
}

function importarTemplatesGeral(listaTemplates) {
  if (!Array.isArray(listaTemplates)) return { success: false, message: 'O arquivo deve conter uma lista JSON de templates.' };
  if (listaTemplates.length > 500) return { success: false, message: 'Importação limitada a 500 templates por arquivo.' };
  let importados = 0;
  const falhas = [];
  const existentes = new Map(listarMeusTemplates().map(function (nomeExistente) { return [nomeExistente.toLowerCase(), nomeExistente]; }));
  listaTemplates.forEach(function (template, indice) {
    if (!template || !String(template.nome || '').trim() || template.corpo == null) {
      falhas.push('Item ' + (indice + 1) + ': nome ou conteúdo ausente.');
      return;
    }
    const nome = String(template.nome).trim();
    const nomeExistente = existentes.get(nome.toLowerCase()) || null;
    const resultado = salvarNovoTemplate(nomeExistente, nome, template.assunto || '', String(template.corpo));
    if (resultado.success) { importados++; existentes.set(nome.toLowerCase(), nome); }
    else falhas.push(nome + ': ' + resultado.message);
  });
  return {
    success: falhas.length === 0,
    imported: importados,
    message: importados + ' template(s) importado(s).' + (falhas.length ? ' Itens não importados: ' + falhas.join(' | ') : '')
  };
}

function sanitizeFileName(nome) {
  return String(nome == null ? '' : nome)
    .replace(/[\\/:*?"<>|#]+/g, '_')
    .replace(/\s+/g, ' ')
    .trim() || 'template';
}
