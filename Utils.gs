/**
 * AUTO EMAIL 4.0 — funções puras compartilhadas pelo projeto.
 * Estas funções não acessam serviços do Google e podem ser validadas localmente.
 */

function normalizarCabecalho_(valor) {
  return String(valor == null ? '' : valor)
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&#160;/gi, ' ')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isValidEmail(email) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(email == null ? '' : email).trim());
}

function dividirEmails_(valor) {
  const vistos = new Set();
  return String(valor == null ? '' : valor)
    .split(/[;,]/)
    .map(function (email) { return email.trim(); })
    .filter(function (email) {
      if (!email || !isValidEmail(email)) return false;
      const chave = email.toLowerCase();
      if (vistos.has(chave)) return false;
      vistos.add(chave);
      return true;
    });
}

function obterEmailsInvalidos_(valor) {
  return String(valor == null ? '' : valor)
    .split(/[;,]/)
    .map(function (email) { return email.trim(); })
    .filter(function (email) { return email && !isValidEmail(email); });
}

function identificarColunaEmail(cabecalhos) {
  const aceitos = /^(e[- ]?mail(?: principal)?|email address|correio eletronico)$/;
  return (cabecalhos || []).findIndex(function (cabecalho) {
    return aceitos.test(normalizarCabecalho_(cabecalho).replace(/\s+/g, ' '));
  });
}

function escaparHtml_(valor) {
  return String(valor == null ? '' : valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function valorParaTexto_(valor) {
  if (valor == null) return '';
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(valor);
}

function substituirVariaveis_(texto, cabecalhos, linha, escaparValores) {
  if (texto == null || texto === '') return '';
  const mapaColunas = Object.create(null);

  (cabecalhos || []).forEach(function (cabecalho, indice) {
    const chave = normalizarCabecalho_(cabecalho);
    if (chave && mapaColunas[chave] === undefined) mapaColunas[chave] = indice;
  });

  const fallback = typeof CONFIG !== 'undefined' && CONFIG.MISSING_VAR_FALLBACK != null
    ? CONFIG.MISSING_VAR_FALLBACK
    : '';

  return String(texto).replace(/\{\{([\s\S]+?)\}\}/g, function (match, conteudo) {
    const chave = normalizarCabecalho_(conteudo);
    const indice = mapaColunas[chave];
    if (indice === undefined || !linha || indice >= linha.length) return fallback;
    const valor = valorParaTexto_(linha[indice]);
    return escaparValores ? escaparHtml_(valor) : valor;
  });
}

/** Mantida com o nome original para compatibilidade com templates e integrações v3.2. */
function personalizarTexto(texto, cabecalhos, linha) {
  return substituirVariaveis_(texto, cabecalhos, linha, false);
}

/** Substitui variáveis em HTML escapando valores oriundos da planilha. */
function personalizarHtml(html, cabecalhos, linha) {
  return substituirVariaveis_(html, cabecalhos, linha, true);
}

function parseDataLegada_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return valor;
  if (valor == null || valor === '') return null;

  const texto = String(valor).trim();
  const local = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (local) {
    return new Date(
      Number(local[3]), Number(local[2]) - 1, Number(local[1]),
      Number(local[4] || 0), Number(local[5] || 0), Number(local[6] || 0)
    );
  }

  const iso = new Date(texto);
  return isNaN(iso.getTime()) ? null : iso;
}

function dataParaIso_(valor) {
  const data = valor instanceof Date ? valor : parseDataLegada_(valor);
  return data && !isNaN(data.getTime()) ? data.toISOString() : '';
}

function numeroSeguro_(valor) {
  const numero = Number(valor);
  return isFinite(numero) ? numero : 0;
}

function csvCell_(valor) {
  let texto = valor == null ? '' : String(valor);
  if (/^[\s]*[=+@-]/.test(texto)) texto = "'" + texto;
  return '"' + texto.replace(/"/g, '""') + '"';
}
