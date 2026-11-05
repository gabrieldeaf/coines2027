/**
 * COINES 2027 - API de inscrição, presença, lanche, kit e certificados
 * Google Apps Script vinculado a uma planilha (Extensões > Apps Script).
 *
 * Depois de colar este arquivo:
 *  1. Ajuste o CONFIG abaixo.
 *  2. Execute a função setup() uma vez (cria as abas e a senha do admin).
 *  3. Implante como App da Web (Executar como: Eu / Acesso: Qualquer pessoa).
 */

const CONFIG = {
  EVENTO: 'COINES 2027',
  INSTITUICAO: 'Instituto Nacional de Educação de Surdos (INES)',
  CIDADE: 'Rio de Janeiro',
  ASSINATURA: 'Coordenação do COINES 2027',
  DIAS_EVENTO: ['2027-09-22', '2027-09-23', '2027-09-24'], // formato AAAA-MM-DD
  PERIODO_TEXTO: '22 a 24 de setembro de 2027',
  FREQ_MINIMA: 75,              // % mínimo para liberar o certificado do evento
  CARGA_HORARIA_EVENTO: 24,     // horas
  VALIDAR_HORARIO_MINICURSO: true, // só aceita o QR do minicurso no dia e horário dele
  TOLERANCIA_MIN: 30,           // minutos antes do início e depois do fim
  HORA_TARDE: '12:00',          // a partir deste horário o admin sugere o lanche da tarde
  FUSO: 'America/Sao_Paulo',
  PREFIXO_QR: 'COINES27',

  // Sorteio, confirmação e lista de espera
  VAGAS: 500,                   // vagas presenciais
  PRAZO_CONFIRMACAO_DIAS: 5,    // prazo sugerido para os sorteados confirmarem
  PRAZO_CHAMADA_DIAS: 2,        // prazo para quem é chamado da lista de espera
  CHAMADA_AUTOMATICA: true,     // de hora em hora: encerra prazos vencidos e chama a lista de espera
  AVISAR_ESPERA: true,          // manda e-mail para quem ficou na lista de espera
  EXIGIR_CONFIRMACAO: true,     // depois do sorteio, presença, kit, lanche e minicurso só para confirmados
  SITE_URL: 'https://gabrieldeaf.github.io/coines2027/',
  REMETENTE: 'COINES 2027',     // nome que aparece como remetente dos e-mails
  RESPONDER_PARA: ''            // opcional: e-mail para respostas (ex.: coines@ines.gov.br)
};

const ST = {
  INSCRITO: 'Inscrito',
  SORTEADO: 'Sorteado',
  CONFIRMADO: 'Confirmado',
  RECUSOU: 'Recusou',
  EXPIRADO: 'Expirado',
  ESPERA: 'Espera'
};

const ABAS = {
  Inscritos: ['Codigo', 'CPF', 'Nome', 'Email', 'Telefone', 'Instituicao', 'Identidade', 'Atuacao', 'Cidade', 'UF', 'DataInscricao',
    'Status', 'Origem', 'PosicaoEspera', 'PrazoConfirmacao', 'Resposta', 'RespostaEm', 'Token'],
  Emails: ['Criado', 'Tipo', 'Codigo', 'Email', 'Nome', 'Status', 'Enviado', 'Erro'],
  Presenca: ['DataHora', 'Dia', 'Codigo', 'CPF', 'Nome'],
  Kit: ['DataHora', 'Codigo', 'CPF', 'Nome'],
  Lanche: ['DataHora', 'Dia', 'Turno', 'Codigo', 'CPF', 'Nome'],
  Minicursos: ['ID', 'Titulo', 'Data', 'HoraInicio', 'HoraFim', 'Local', 'CargaHoraria', 'Ministrante'],
  PresencaMinicurso: ['DataHora', 'MinicursoID', 'Titulo', 'Codigo', 'CPF', 'Nome']
};

/* ======================= Entrada HTTP ======================= */

function doGet() {
  return saida({ ok: true, evento: CONFIG.EVENTO, mensagem: 'API no ar' });
}

const props = () => PropertiesService.getScriptProperties();
const sorteioFeito = () => !!props().getProperty('SORTEIO_EM');

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return saida({ ok: false, erro: 'Requisição inválida.' });
  }
  const acoes = {
    buscarCpf, inscrever, painel, registrarMinicurso, certificado, responder,
    adminLogin, adminScan, adminTabela,
    adminResumo, adminSortear, adminPedirConfirmacao, adminChamarEspera, adminEnviarEmails
  };
  const fn = acoes[req.acao];
  if (!fn) return saida({ ok: false, erro: 'Ação desconhecida.' });
  try {
    return saida(fn(req));
  } catch (err) {
    return saida({ ok: false, erro: err.message || String(err) });
  }
}

function saida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================= Planilha ======================= */

const _abasChecadas = {};

function aba(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (!_abasChecadas[nome]) {
    // Abas criadas numa versão anterior: acrescenta no fim as colunas novas que faltarem.
    const ult = sh.getLastColumn();
    const cab = ult ? sh.getRange(1, 1, 1, ult).getDisplayValues()[0].map(c => String(c).trim()) : [];
    const faltam = ABAS[nome].filter(c => cab.indexOf(c) < 0);
    if (faltam.length) {
      sh.getRange(1, ult + 1, 1, faltam.length).setValues([faltam]).setFontWeight('bold');
      sh.getRange(1, ult + 1, sh.getMaxRows(), faltam.length).setNumberFormat('@');
    }
  }
  _abasChecadas[nome] = true;
  return sh;
}

function cabecalho(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(c => String(c).trim());
}

// Cada objeto lido traz _linha (número da linha na planilha) para poder ser atualizado depois.
function ler(nome) {
  const sh = aba(nome);
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  const cab = cabecalho(sh);
  const saidaLinhas = [];
  sh.getRange(2, 1, ult - 1, cab.length).getDisplayValues().forEach((l, i) => {
    if (l.join('') === '') return;
    const o = { _linha: i + 2 };
    cab.forEach((c, j) => { if (c) o[c] = String(l[j]).trim(); });
    saidaLinhas.push(o);
  });
  return saidaLinhas;
}

// O apóstrofo força texto: mantém os zeros do CPF e evita fórmulas digitadas.
const comoTexto = v => (v === undefined || v === null || v === '') ? '' : "'" + v;

// Grava pelo nome da coluna (não pela posição).
function gravar(nome, obj) {
  gravarVarios(nome, [obj]);
}

function gravarVarios(nome, objs) {
  if (!objs.length) return;
  const sh = aba(nome);
  const cab = cabecalho(sh);
  const linhas = objs.map(o => cab.map(c => comoTexto(o[c])));
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, cab.length).setValues(linhas);
}

// alteracoes: [{ _linha: 5, Status: 'Confirmado', ... }]. Atualiza só as colunas informadas.
function atualizar(nome, alteracoes) {
  if (!alteracoes.length) return;
  const sh = aba(nome);
  const cab = cabecalho(sh);
  const porColuna = {};
  alteracoes.forEach(a => Object.keys(a).forEach(k => {
    if (k === '_linha') return;
    const col = cab.indexOf(k);
    if (col < 0) throw new Error('Coluna ' + k + ' não existe na aba ' + nome + '.');
    (porColuna[col] = porColuna[col] || []).push([a._linha, a[k]]);
  }));
  Object.keys(porColuna).forEach(col => {
    const itens = porColuna[col];
    const linhas = itens.map(i => i[0]);
    const ini = Math.min.apply(null, linhas);
    const fim = Math.max.apply(null, linhas);
    const rng = sh.getRange(ini, Number(col) + 1, fim - ini + 1, 1);
    const valores = rng.getValues();
    itens.forEach(([l, v]) => { valores[l - ini][0] = comoTexto(v); });
    rng.setValues(valores);
  });
}

function comTrava(fn) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try { return fn(); } finally { trava.releaseLock(); }
}

/* ======================= Utilidades ======================= */

const soDigitos = v => String(v || '').replace(/\D/g, '');
const limpa = v => String(v || '').trim().replace(/\s+/g, ' ').slice(0, 200);

function normCpf(v) {
  const d = soDigitos(v);
  return d ? d.padStart(11, '0') : '';
}

function cpfValido(v) {
  const c = normCpf(v);
  if (c.length !== 11 || /^(\d)\1+$/.test(c)) return false;
  for (let t = 9; t < 11; t++) {
    let s = 0;
    for (let i = 0; i < t; i++) s += Number(c[i]) * (t + 1 - i);
    const d = ((s * 10) % 11) % 10;
    if (d !== Number(c[t])) return false;
  }
  return true;
}

function fmtCpf(v) {
  return normCpf(v).replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function mascaraCpf(v) {
  const c = normCpf(v);
  return '***.' + c.slice(3, 6) + '.' + c.slice(6, 9) + '-**';
}

const hoje = () => Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyy-MM-dd');
const agora = () => Utilities.formatDate(new Date(), CONFIG.FUSO, 'dd/MM/yyyy HH:mm:ss');
const horaAgora = () => Utilities.formatDate(new Date(), CONFIG.FUSO, 'HH:mm');

const TURNOS = { manha: 'Manhã', tarde: 'Tarde' };
const turnoAtual = () => (horaAgora() < CONFIG.HORA_TARDE ? 'manha' : 'tarde');

function normData(v) {
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  return s;
}

function dataBr(iso) {
  const p = String(iso || '').split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : String(iso || '');
}

function normHora(v) {
  const m = String(v || '').match(/(\d{1,2}):(\d{2})/);
  return m ? m[1].padStart(2, '0') + ':' + m[2] : '';
}

function minutos(h) {
  const p = h.split(':').map(Number);
  return p[0] * 60 + p[1];
}

function dataExtenso() {
  const meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho',
    'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  const p = Utilities.formatDate(new Date(), CONFIG.FUSO, 'd/M/yyyy').split('/');
  return p[0] + ' de ' + meses[Number(p[1]) - 1] + ' de ' + p[2];
}

function novoCodigo() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function codigoValidacao(texto) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto + '|' + CONFIG.PREFIXO_QR)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0'))
    .join('').slice(0, 12).toUpperCase();
}

// QR no formato PREFIXO|TIPO|VALOR  (P = presencial, L = lanche, M = minicurso)
function lerQr(texto) {
  const p = String(texto || '').trim().split('|');
  if (p.length === 3 && p[0] === CONFIG.PREFIXO_QR) return { tipo: p[1], valor: p[2].trim().toUpperCase() };
  return null;
}

function inscritoPorCpf(cpf) {
  const c = normCpf(cpf);
  return ler('Inscritos').find(r => normCpf(r.CPF) === c) || null;
}

function inscritoPorCodigo(codigo) {
  const c = String(codigo || '').trim().toUpperCase();
  return c ? (ler('Inscritos').find(r => r.Codigo === c) || null) : null;
}

/* ======================= Participante ======================= */

function buscarCpf(req) {
  const cpf = normCpf(req.cpf);
  if (!cpfValido(cpf)) throw new Error('CPF inválido. Confira os números.');
  const insc = inscritoPorCpf(cpf);
  return insc ? { ok: true, existe: true, painel: montarPainel(insc) } : { ok: true, existe: false };
}

function inscrever(req) {
  const d = req.dados || {};
  const cpf = normCpf(d.cpf);
  if (!cpfValido(cpf)) throw new Error('CPF inválido.');
  const nome = limpa(d.nome);
  const email = limpa(d.email).toLowerCase();
  if (nome.split(' ').length < 2) throw new Error('Informe o nome completo.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('E-mail inválido.');
  if (soDigitos(d.telefone).length < 10) throw new Error('Telefone inválido.');
  if (!d.lgpd) throw new Error('É preciso concordar com o uso dos dados para a inscrição.');

  return comTrava(() => {
    const existente = inscritoPorCpf(cpf);
    if (existente) return { ok: true, jaInscrito: true, painel: montarPainel(existente) };
    const insc = {
      Codigo: novoCodigo(), CPF: cpf, Nome: nome, Email: email,
      Telefone: soDigitos(d.telefone), Instituicao: limpa(d.instituicao),
      Identidade: limpa(d.identidade), Atuacao: limpa(d.atuacao),
      Cidade: limpa(d.cidade), UF: limpa(d.uf).toUpperCase().slice(0, 2),
      DataInscricao: agora(), Status: ST.INSCRITO, Token: novoToken()
    };
    // Depois do sorteio, quem se inscreve entra no fim da lista de espera.
    if (sorteioFeito()) {
      insc.Status = ST.ESPERA;
      insc.Origem = 'Inscrição após o sorteio';
      insc.PosicaoEspera = ultimaPosicaoEspera() + 1;
    }
    gravar('Inscritos', insc);
    if (insc.Status === ST.ESPERA && CONFIG.AVISAR_ESPERA) enfileirar([{ tipo: 'espera', insc: insc }]);
    return { ok: true, painel: montarPainel(insc) };
  });
}

function painel(req) {
  const insc = inscritoPorCodigo(req.codigo);
  if (!insc) throw new Error('Inscrição não encontrada. Entre novamente com o CPF.');
  return { ok: true, painel: montarPainel(insc) };
}

function montarPainel(insc) {
  const freq = frequencia(insc.Codigo);
  return {
    situacao: situacao(insc),
    codigo: insc.Codigo,
    nome: insc.Nome,
    primeiroNome: insc.Nome.split(' ')[0],
    cpf: mascaraCpf(insc.CPF),
    evento: CONFIG.EVENTO,
    frequencia: freq.percentual,
    diasPresentes: freq.dias,
    totalDias: CONFIG.DIAS_EVENTO.length,
    freqMinima: CONFIG.FREQ_MINIMA,
    certEvento: freq.percentual >= CONFIG.FREQ_MINIMA,
    kitRecebido: ler('Kit').some(r => r.Codigo === insc.Codigo),
    qrPresencial: CONFIG.PREFIXO_QR + '|P|' + insc.Codigo,
    qrLanche: CONFIG.PREFIXO_QR + '|L|' + insc.Codigo,
    minicursos: minicursosDo(insc.Codigo)
  };
}

function frequencia(codigo) {
  const dias = new Set(
    ler('Presenca')
      .filter(r => r.Codigo === codigo)
      .map(r => normData(r.Dia))
      .filter(d => CONFIG.DIAS_EVENTO.indexOf(d) >= 0)
  );
  const total = CONFIG.DIAS_EVENTO.length || 1;
  return { dias: dias.size, percentual: Math.round(dias.size / total * 100) };
}

function listaMinicursos() {
  return ler('Minicursos').filter(r => r.ID).map(r => ({
    id: r.ID.trim().toUpperCase(),
    titulo: r.Titulo,
    data: normData(r.Data),
    inicio: normHora(r.HoraInicio),
    fim: normHora(r.HoraFim),
    local: r.Local,
    ch: r.CargaHoraria,
    ministrante: r.Ministrante
  }));
}

function minicursosDo(codigo) {
  const mcs = listaMinicursos();
  return ler('PresencaMinicurso')
    .filter(r => r.Codigo === codigo)
    .map(r => {
      const mc = mcs.find(m => m.id === r.MinicursoID) || { id: r.MinicursoID, titulo: r.Titulo, data: '' };
      return Object.assign({}, mc, { dataBr: mc.data ? dataBr(mc.data) : '', registradoEm: r.DataHora });
    })
    .sort((a, b) => (a.data + a.inicio).localeCompare(b.data + b.inicio));
}

function registrarMinicurso(req) {
  const insc = inscritoPorCodigo(req.codigo);
  if (!insc) throw new Error('Inscrição não encontrada.');
  exigirConfirmado(insc);
  const qr = lerQr(req.qr);
  if (!qr || qr.tipo !== 'M') throw new Error('Este não é um QR Code de minicurso do ' + CONFIG.EVENTO + '.');
  const mc = listaMinicursos().find(m => m.id === qr.valor);
  if (!mc) throw new Error('Minicurso não encontrado.');

  if (CONFIG.VALIDAR_HORARIO_MINICURSO) {
    if (mc.data !== hoje()) throw new Error('Este minicurso acontece em ' + dataBr(mc.data) + '.');
    const ag = minutos(horaAgora());
    if (mc.inicio && ag < minutos(mc.inicio) - CONFIG.TOLERANCIA_MIN) {
      throw new Error('A presença abre ' + CONFIG.TOLERANCIA_MIN + ' minutos antes do início (' + mc.inicio + ').');
    }
    if (mc.fim && ag > minutos(mc.fim) + CONFIG.TOLERANCIA_MIN) {
      throw new Error('O registro de presença deste minicurso já foi encerrado.');
    }
  }

  return comTrava(() => {
    const ja = ler('PresencaMinicurso').some(r => r.Codigo === insc.Codigo && r.MinicursoID === mc.id);
    if (!ja) {
      gravar('PresencaMinicurso', {
        DataHora: agora(), MinicursoID: mc.id, Titulo: mc.titulo,
        Codigo: insc.Codigo, CPF: insc.CPF, Nome: insc.Nome
      });
    }
    return { ok: true, jaRegistrado: ja, minicurso: mc.titulo, minicursos: minicursosDo(insc.Codigo) };
  });
}

function certificado(req) {
  const insc = inscritoPorCodigo(req.codigo);
  if (!insc) throw new Error('Inscrição não encontrada.');
  const nome = insc.Nome.toUpperCase();
  const cpf = fmtCpf(insc.CPF);
  const base = {
    titulo: 'CERTIFICADO',
    evento: CONFIG.EVENTO,
    localData: CONFIG.CIDADE + ', ' + dataExtenso() + '.',
    assinatura: CONFIG.ASSINATURA
  };

  if (req.tipo === 'evento') {
    const f = frequencia(insc.Codigo);
    if (f.percentual < CONFIG.FREQ_MINIMA) {
      throw new Error('O certificado do evento exige ' + CONFIG.FREQ_MINIMA + '% de frequência. A sua está em ' + f.percentual + '%.');
    }
    return {
      ok: true,
      cert: Object.assign(base, {
        texto: 'Certificamos que ' + nome + ', CPF ' + cpf + ', participou do ' + CONFIG.EVENTO +
          ', promovido pelo ' + CONFIG.INSTITUICAO + ', realizado de ' + CONFIG.PERIODO_TEXTO +
          ', com frequência de ' + f.percentual + '% e carga horária total de ' + CONFIG.CARGA_HORARIA_EVENTO + ' horas.',
        codigoValidacao: codigoValidacao(insc.Codigo + '|EVENTO'),
        arquivo: 'certificado-coines-2027.pdf'
      })
    };
  }

  const id = String(req.id || '').trim().toUpperCase();
  const reg = ler('PresencaMinicurso').find(r => r.Codigo === insc.Codigo && r.MinicursoID === id);
  if (!reg) throw new Error('Não há presença registrada neste minicurso.');
  const mc = listaMinicursos().find(m => m.id === id) || { titulo: reg.Titulo, data: '' };
  let texto = 'Certificamos que ' + nome + ', CPF ' + cpf + ', participou do minicurso "' + mc.titulo + '"';
  if (mc.ministrante) texto += ', ministrado por ' + mc.ministrante;
  if (mc.data) texto += ', realizado em ' + dataBr(mc.data);
  texto += ' durante o ' + CONFIG.EVENTO + ', promovido pelo ' + CONFIG.INSTITUICAO;
  if (mc.ch) texto += ', com carga horária de ' + mc.ch + ' horas';
  texto += '.';
  return {
    ok: true,
    cert: Object.assign(base, {
      texto: texto,
      codigoValidacao: codigoValidacao(insc.Codigo + '|MC|' + id),
      arquivo: 'certificado-minicurso-' + id.toLowerCase() + '.pdf'
    })
  };
}

/* ======================= Administração ======================= */

function checarSenha(req) {
  const senha = PropertiesService.getScriptProperties().getProperty('ADMIN_SENHA');
  if (!senha) throw new Error('Senha de administração não configurada. Execute a função setup().');
  if (String(req.senha || '') !== senha) throw new Error('Senha incorreta.');
}

function adminLogin(req) {
  checarSenha(req);
  return {
    ok: true, evento: CONFIG.EVENTO, prefixo: CONFIG.PREFIXO_QR, abas: Object.keys(ABAS),
    turnoAtual: turnoAtual(), horaTarde: CONFIG.HORA_TARDE
  };
}

function adminScan(req) {
  checarSenha(req);
  const modos = {
    presenca: { aba: 'Presenca', tipo: 'P', porDia: true, feito: 'Presença registrada.' },
    kit: { aba: 'Kit', tipo: 'P', porDia: false, feito: 'Kit de material entregue.' },
    lanche: { aba: 'Lanche', tipo: 'L', porDia: true, porTurno: true, feito: 'Lanche liberado.' }
  };
  const cfg = modos[req.modo];
  if (!cfg) throw new Error('Modo inválido.');

  // Lanche: 1 de manhã e 1 à tarde por dia. O admin informa o turno; sem ele, usa o horário.
  const turno = cfg.porTurno ? (TURNOS[req.turno] ? req.turno : turnoAtual()) : '';
  const turnoNome = turno ? TURNOS[turno] : '';
  const daTurno = turno === 'manha' ? 'da manhã' : 'da tarde';

  const nomesQr = { P: 'QR Code presencial', L: 'QR Code do lanche', M: 'QR Code de minicurso' };
  const entrada = String(req.qr || '').trim();
  const qr = lerQr(entrada);
  let insc;
  if (qr) {
    if (qr.tipo !== cfg.tipo) {
      throw new Error('Este é o ' + (nomesQr[qr.tipo] || 'QR errado') + '. Aqui vale o ' + nomesQr[cfg.tipo] + '.');
    }
    insc = inscritoPorCodigo(qr.valor);
  } else if (soDigitos(entrada).length === 11) {
    insc = inscritoPorCpf(entrada);   // CPF digitado à mão
  } else {
    insc = inscritoPorCodigo(entrada); // código digitado à mão
  }
  if (!insc) throw new Error('Participante não encontrado.');
  exigirConfirmado(insc);

  const dia = hoje();
  return comTrava(() => {
    const regs = ler(cfg.aba);
    const considerados = cfg.porDia
      ? regs.filter(r => normData(r.Dia) === dia && (!turno || r.Turno === turnoNome))
      : regs;
    const anterior = considerados.find(r => r.Codigo === insc.Codigo);
    let status = 'ok';
    let mensagem;
    if (anterior) {
      status = 'repetido';
      const hora = anterior.DataHora.slice(11, 16);
      if (turno) mensagem = 'Já pegou o lanche ' + daTurno + ' hoje, às ' + hora + '.';
      else if (cfg.porDia) mensagem = 'Já registrado hoje às ' + hora + '.';
      else mensagem = 'Já recebeu em ' + anterior.DataHora.slice(0, 16) + '.';
    } else {
      const reg = { DataHora: agora(), Codigo: insc.Codigo, CPF: insc.CPF, Nome: insc.Nome };
      if (cfg.porDia) reg.Dia = dia;
      if (turno) reg.Turno = turnoNome;
      gravar(cfg.aba, reg);
      mensagem = turno ? 'Lanche ' + daTurno + ' liberado.' : cfg.feito;
      if (req.modo === 'presenca' && CONFIG.DIAS_EVENTO.indexOf(dia) < 0) {
        mensagem += ' Atenção: hoje não é dia do evento, não conta na frequência.';
      }
    }
    return {
      ok: true, status: status, nome: insc.Nome, cpf: mascaraCpf(insc.CPF), mensagem: mensagem,
      total: considerados.length + (status === 'ok' ? 1 : 0),
      rotuloTotal: turno ? turnoNome.toLowerCase() + ' de hoje' : (cfg.porDia ? 'hoje' : 'no total'),
      turno: turno
    };
  });
}

function adminTabela(req) {
  checarSenha(req);
  if (!ABAS[req.aba]) throw new Error('Aba inválida.');
  const valores = aba(req.aba).getDataRange().getDisplayValues();
  const cabecalho = valores.shift() || [];
  return { ok: true, aba: req.aba, cabecalho: cabecalho, linhas: valores.filter(l => l.join('') !== '') };
}

/* ======================= Configuração inicial ======================= */

function setup() {
  Object.keys(ABAS).forEach(nome => aba(nome).getRange('A:Z').setNumberFormat('@'));

  const mc = aba('Minicursos');
  if (mc.getLastRow() < 2) {
    mc.getRange(2, 1, 2, 8).setValues([
      ['MC01', 'Introdução à Libras', '22/09/2027', '09:00', '12:00', 'Sala 1', '3', 'Nome do ministrante'],
      ['MC02', 'Literatura surda', '23/09/2027', '14:00', '17:00', 'Sala 2', '3', 'Nome do ministrante']
    ]);
  }

  // Inscrições feitas antes desta versão: preenche situação e token que faltarem.
  const semDados = ler('Inscritos').filter(i => !i.Status || !i.Token);
  atualizar('Inscritos', semDados.map(i => ({
    _linha: i._linha, Status: i.Status || ST.INSCRITO, Token: i.Token || novoToken()
  })));

  instalarRotina();

  const p = props();
  if (!p.getProperty('ADMIN_SENHA')) p.setProperty('ADMIN_SENHA', 'coines2027');
  Logger.log('Pronto. Senha do admin: ' + p.getProperty('ADMIN_SENHA') +
    ' (troque em Configurações do projeto > Propriedades do script > ADMIN_SENHA). ' +
    'Cota de e-mails hoje: ' + MailApp.getRemainingDailyQuota());
}

// Roda de hora em hora (gatilho criado pelo setup).
function instalarRotina() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'rotina')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('rotina').timeBased().everyHours(1).create();
}

function rotina() {
  if (CONFIG.CHAMADA_AUTOMATICA && sorteioFeito()) {
    comTrava(() => { encerrarPrazos(); chamarEspera(); });
  }
  processarFila();
}

/**
 * SÓ PARA TESTES: desfaz o sorteio e volta todo mundo para "Inscrito".
 * Execute pelo editor. Não use depois de mandar os e-mails reais.
 */
function desfazerSorteio() {
  comTrava(() => {
    atualizar('Inscritos', ler('Inscritos').map(i => ({
      _linha: i._linha, Status: ST.INSCRITO, Origem: '', PosicaoEspera: '',
      PrazoConfirmacao: '', Resposta: '', RespostaEm: ''
    })));
    atualizar('Emails', ler('Emails').filter(e => e.Status === 'Pendente')
      .map(e => ({ _linha: e._linha, Status: 'Cancelado' })));
    props().deleteProperty('SORTEIO_EM');
  });
  Logger.log('Sorteio desfeito.');
}

/* ======================= Sorteio, confirmação e lista de espera ======================= */

function novoToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

const statusDe = i => i.Status || ST.INSCRITO;
const agoraIso = () => Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyy-MM-dd HH:mm');

// Prazo: fim do dia (23:59) daqui a N dias.
function prazoEmDias(dias) {
  const d = new Date(Date.now() + Number(dias) * 86400000);
  return Utilities.formatDate(d, CONFIG.FUSO, 'yyyy-MM-dd') + ' 23:59';
}

function prazoBr(iso) {
  if (!iso) return '';
  return dataBr(iso.slice(0, 10)) + ' às ' + iso.slice(11, 16);
}

const prazoVencido = i => !!i.PrazoConfirmacao && agoraIso() > i.PrazoConfirmacao;

function ultimaPosicaoEspera(lista) {
  return (lista || ler('Inscritos')).reduce((m, i) => Math.max(m, Number(i.PosicaoEspera) || 0), 0);
}

// Posição atual na fila (contando só quem ainda está esperando).
function posicaoAtual(insc, lista) {
  const minha = Number(insc.PosicaoEspera) || 0;
  return (lista || ler('Inscritos'))
    .filter(i => statusDe(i) === ST.ESPERA && (Number(i.PosicaoEspera) || 0) < minha).length + 1;
}

function exigirConfirmado(insc) {
  if (!CONFIG.EXIGIR_CONFIRMACAO || !sorteioFeito()) return;
  if (statusDe(insc) !== ST.CONFIRMADO) {
    throw new Error('Presença não confirmada. Situação da inscrição: ' + situacao(insc).titulo + '.');
  }
}

function situacao(insc) {
  const st = statusDe(insc);
  const s = { status: st, prazo: prazoBr(insc.PrazoConfirmacao), podeResponder: false, posicao: 0 };
  if (st === ST.INSCRITO) {
    s.titulo = 'Inscrição recebida';
    s.texto = 'Aguarde o sorteio das ' + CONFIG.VAGAS + ' vagas presenciais. O resultado chega por e-mail.';
  } else if (st === ST.SORTEADO) {
    s.podeResponder = true;
    if (insc.PrazoConfirmacao) {
      s.titulo = 'Confirme sua presença';
      s.texto = 'Você tem uma vaga no ' + CONFIG.EVENTO + '. Responda até ' + s.prazo +
        '. Se não responder, a vaga passa para a lista de espera.';
    } else {
      s.titulo = 'Você foi sorteado(a)!';
      s.texto = 'Em breve você vai receber por e-mail o pedido de confirmação. Se já sabe, pode responder agora.';
    }
  } else if (st === ST.CONFIRMADO) {
    s.titulo = 'Presença confirmada';
    s.texto = 'Sua vaga presencial está garantida. Mostre o QR Code presencial na entrada do evento.';
    s.podeResponder = true;
  } else if (st === ST.ESPERA) {
    s.posicao = posicaoAtual(insc);
    s.titulo = 'Lista de espera: posição ' + s.posicao;
    s.texto = 'Se abrir uma vaga, você recebe um e-mail para confirmar. Confira também a caixa de spam.';
  } else if (st === ST.RECUSOU) {
    s.titulo = 'Você liberou sua vaga';
    s.texto = 'Obrigado por avisar. A vaga foi passada para a próxima pessoa da lista de espera.';
  } else if (st === ST.EXPIRADO) {
    s.titulo = 'Prazo de confirmação encerrado';
    s.texto = 'Não recebemos sua resposta no prazo, e a vaga passou para a lista de espera.';
  }
  return s;
}

function responder(req) {
  const resposta = req.resposta;
  if (resposta !== 'sim' && resposta !== 'nao') throw new Error('Resposta inválida.');
  const r = comTrava(() => {
    const lista = ler('Inscritos');
    const token = String(req.token || '').trim();
    const codigo = String(req.codigo || '').trim().toUpperCase();
    const insc = token ? lista.find(i => i.Token === token) : (codigo ? lista.find(i => i.Codigo === codigo) : null);
    if (!insc) throw new Error('Link inválido ou inscrição não encontrada.');
    const st = statusDe(insc);
    let mensagem;

    if (st === ST.SORTEADO && prazoVencido(insc)) {
      atualizar('Inscritos', [{ _linha: insc._linha, Status: ST.EXPIRADO }]);
      chamarEspera();
      mensagem = 'O prazo para responder terminou em ' + prazoBr(insc.PrazoConfirmacao) + '.';
    } else if (st === ST.SORTEADO || st === ST.CONFIRMADO) {
      if (resposta === 'sim') {
        if (st !== ST.CONFIRMADO) {
          atualizar('Inscritos', [{ _linha: insc._linha, Status: ST.CONFIRMADO, Resposta: 'Sim', RespostaEm: agora() }]);
        }
        mensagem = 'Nos vemos no ' + CONFIG.EVENTO + '! Seu QR Code de entrada está na sua inscrição.';
      } else {
        atualizar('Inscritos', [{ _linha: insc._linha, Status: ST.RECUSOU, Resposta: 'Não', RespostaEm: agora() }]);
        chamarEspera();
        mensagem = 'Tudo certo, sua vaga foi liberada para a lista de espera. Obrigado por avisar.';
      }
    } else {
      mensagem = situacao(insc).titulo + '. ' + situacao(insc).texto;
    }
    return { mensagem: mensagem, codigo: insc.Codigo };
  });
  processarFila(); // manda logo o e-mail de quem foi chamado
  return { ok: true, mensagem: r.mensagem, painel: montarPainel(inscritoPorCodigo(r.codigo)) };
}

// Marca como Expirado quem não respondeu no prazo. Chamar dentro de comTrava.
function encerrarPrazos() {
  const vencidos = ler('Inscritos').filter(i => statusDe(i) === ST.SORTEADO && prazoVencido(i));
  atualizar('Inscritos', vencidos.map(i => ({ _linha: i._linha, Status: ST.EXPIRADO })));
  return vencidos.length;
}

// Ocupa as vagas livres com a lista de espera, na ordem. Chamar dentro de comTrava.
function chamarEspera() {
  if (!sorteioFeito()) return 0;
  const lista = ler('Inscritos');
  const ocupadas = lista.filter(i => statusDe(i) === ST.SORTEADO || statusDe(i) === ST.CONFIRMADO).length;
  const livres = CONFIG.VAGAS - ocupadas;
  if (livres <= 0) return 0;
  const chamados = lista
    .filter(i => statusDe(i) === ST.ESPERA)
    .sort((a, b) => (Number(a.PosicaoEspera) || 0) - (Number(b.PosicaoEspera) || 0))
    .slice(0, livres);
  const prazo = prazoEmDias(CONFIG.PRAZO_CHAMADA_DIAS);
  atualizar('Inscritos', chamados.map(i => ({
    _linha: i._linha, Status: ST.SORTEADO, Origem: 'Lista de espera', PrazoConfirmacao: prazo
  })));
  enfileirar(chamados.map(i => ({ tipo: 'chamada', insc: i })));
  return chamados.length;
}

function resumo() {
  const lista = ler('Inscritos');
  const conta = st => lista.filter(i => statusDe(i) === st).length;
  const sorteados = lista.filter(i => statusDe(i) === ST.SORTEADO);
  const ocupadas = sorteados.length + conta(ST.CONFIRMADO);
  const emails = ler('Emails');
  let cota = null;
  try { cota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  return {
    vagas: CONFIG.VAGAS,
    sorteioEm: props().getProperty('SORTEIO_EM') || '',
    total: lista.length,
    inscritos: conta(ST.INSCRITO),
    aguardando: sorteados.length,
    semPedido: sorteados.filter(i => !i.PrazoConfirmacao).length,
    vencidos: sorteados.filter(prazoVencido).length,
    confirmados: conta(ST.CONFIRMADO),
    recusaram: conta(ST.RECUSOU),
    expirados: conta(ST.EXPIRADO),
    espera: conta(ST.ESPERA),
    vagasLivres: Math.max(CONFIG.VAGAS - ocupadas, 0),
    emailsPendentes: emails.filter(e => e.Status === 'Pendente').length,
    emailsEnviados: emails.filter(e => e.Status === 'Enviado').length,
    emailsErro: emails.filter(e => e.Status === 'Erro').length,
    cotaEmail: cota,
    prazoConfirmacaoDias: CONFIG.PRAZO_CONFIRMACAO_DIAS,
    prazoChamadaDias: CONFIG.PRAZO_CHAMADA_DIAS,
    chamadaAutomatica: CONFIG.CHAMADA_AUTOMATICA
  };
}

function adminResumo(req) {
  checarSenha(req);
  return { ok: true, resumo: resumo() };
}

function adminSortear(req) {
  checarSenha(req);
  const r = comTrava(() => {
    if (sorteioFeito()) throw new Error('O sorteio já foi feito em ' + props().getProperty('SORTEIO_EM') + '.');
    const candidatos = ler('Inscritos').filter(i => statusDe(i) === ST.INSCRITO);
    if (!candidatos.length) throw new Error('Não há inscritos para sortear.');
    // Embaralhamento de Fisher-Yates
    for (let k = candidatos.length - 1; k > 0; k--) {
      const j = Math.floor(Math.random() * (k + 1));
      const t = candidatos[k]; candidatos[k] = candidatos[j]; candidatos[j] = t;
    }
    const sorteados = candidatos.slice(0, CONFIG.VAGAS);
    const espera = candidatos.slice(CONFIG.VAGAS);
    atualizar('Inscritos', sorteados.map(i => ({ _linha: i._linha, Status: ST.SORTEADO, Origem: 'Sorteio' }))
      .concat(espera.map((i, n) => ({ _linha: i._linha, Status: ST.ESPERA, Origem: 'Sorteio', PosicaoEspera: n + 1 }))));
    enfileirar(sorteados.map(i => ({ tipo: 'sorteado', insc: i }))
      .concat(CONFIG.AVISAR_ESPERA ? espera.map(i => ({ tipo: 'espera', insc: i })) : []));
    props().setProperty('SORTEIO_EM', agora());
    return { sorteados: sorteados.length, espera: espera.length };
  });
  processarFila();
  return { ok: true, mensagem: 'Sorteio feito: ' + r.sorteados + ' sorteados e ' + r.espera + ' na lista de espera.', resumo: resumo() };
}

function adminPedirConfirmacao(req) {
  checarSenha(req);
  const dias = Math.max(1, Math.min(60, Number(req.dias) || CONFIG.PRAZO_CONFIRMACAO_DIAS));
  const n = comTrava(() => {
    if (!sorteioFeito()) throw new Error('Faça o sorteio primeiro.');
    const alvo = ler('Inscritos').filter(i => statusDe(i) === ST.SORTEADO && !i.PrazoConfirmacao);
    const prazo = prazoEmDias(dias);
    atualizar('Inscritos', alvo.map(i => ({ _linha: i._linha, PrazoConfirmacao: prazo })));
    enfileirar(alvo.map(i => ({ tipo: 'confirmacao', insc: i })));
    return alvo.length;
  });
  processarFila();
  return { ok: true, mensagem: n ? 'Pedido de confirmação gerado para ' + n + ' pessoa(s).' : 'Todos os sorteados já receberam o pedido.', resumo: resumo() };
}

function adminChamarEspera(req) {
  checarSenha(req);
  const r = comTrava(() => {
    if (!sorteioFeito()) throw new Error('Faça o sorteio primeiro.');
    return { expirados: encerrarPrazos(), chamados: chamarEspera() };
  });
  processarFila();
  return {
    ok: true,
    mensagem: r.expirados + ' prazo(s) encerrado(s). ' + r.chamados + ' pessoa(s) chamada(s) da lista de espera.',
    resumo: resumo()
  };
}

function adminEnviarEmails(req) {
  checarSenha(req);
  const r = processarFila();
  return { ok: true, mensagem: r.enviados + ' e-mail(s) enviado(s). ' + r.restantes + ' na fila.', resumo: resumo() };
}

/* ======================= E-mails (fila) ======================= */

// itens: [{ tipo, insc }]
function enfileirar(itens) {
  gravarVarios('Emails', itens.filter(x => x.insc && x.insc.Email).map(x => ({
    Criado: agora(), Tipo: x.tipo, Codigo: x.insc.Codigo, Email: x.insc.Email, Nome: x.insc.Nome, Status: 'Pendente'
  })));
}

// Envia o que der dentro da cota diária do Gmail. O resto sai na próxima rodada da rotina.
function processarFila() {
  const trava = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!trava.tryLock(2000)) return { enviados: 0, restantes: -1 };
  try {
    const pendentes = ler('Emails').filter(e => e.Status === 'Pendente');
    if (!pendentes.length) return { enviados: 0, restantes: 0 };
    const lista = ler('Inscritos');
    const porCodigo = {};
    lista.forEach(i => { porCodigo[i.Codigo] = i; });
    let cota = MailApp.getRemainingDailyQuota();
    const inicio = Date.now();
    const alteracoes = [];
    let enviados = 0;
    for (const e of pendentes) {
      if (cota <= 0 || Date.now() - inicio > 4 * 60 * 1000) break;
      const insc = porCodigo[e.Codigo];
      if (!insc || !emailAindaFazSentido(e.Tipo, insc)) {
        alteracoes.push({ _linha: e._linha, Status: 'Ignorado', Erro: 'Situação mudou antes do envio' });
        continue;
      }
      try {
        const m = montarEmail(e.Tipo, insc, lista);
        const opcoes = { to: insc.Email, subject: m.assunto, htmlBody: m.html, body: m.texto, name: CONFIG.REMETENTE };
        if (CONFIG.RESPONDER_PARA) opcoes.replyTo = CONFIG.RESPONDER_PARA;
        MailApp.sendEmail(opcoes);
        alteracoes.push({ _linha: e._linha, Status: 'Enviado', Enviado: agora(), Erro: '' });
        enviados++; cota--;
      } catch (err) {
        alteracoes.push({ _linha: e._linha, Status: 'Erro', Erro: String(err.message || err).slice(0, 200) });
        if (/limit|quota|cota/i.test(String(err))) break;
      }
    }
    atualizar('Emails', alteracoes);
    return { enviados: enviados, restantes: pendentes.length - alteracoes.length };
  } finally {
    trava.releaseLock();
  }
}

function emailAindaFazSentido(tipo, insc) {
  const st = statusDe(insc);
  if (tipo === 'sorteado') return st === ST.SORTEADO || st === ST.CONFIRMADO;
  if (tipo === 'confirmacao' || tipo === 'chamada') return st === ST.SORTEADO;
  if (tipo === 'espera') return st === ST.ESPERA;
  return true;
}

function montarEmail(tipo, insc, lista) {
  const nome = insc.Nome.split(' ')[0];
  const ev = CONFIG.EVENTO;
  const linkSim = CONFIG.SITE_URL + '?t=' + insc.Token + '&r=sim';
  const linkNao = CONFIG.SITE_URL + '?t=' + insc.Token + '&r=nao';
  const prazo = prazoBr(insc.PrazoConfirmacao);
  const botoesResposta =
    botao(linkSim, 'Sim, vou ao evento presencial', '#1F9D55') + ' ' + botao(linkNao, 'Não vou poder ir', '#C62828');
  let assunto, paragrafos, extra = '';

  if (tipo === 'sorteado') {
    assunto = ev + ': você foi sorteado(a)!';
    paragrafos = [
      'Parabéns! Você foi sorteado(a) para uma das ' + CONFIG.VAGAS + ' vagas presenciais do ' + ev + ', de ' + CONFIG.PERIODO_TEXTO + '.',
      'Em breve você vai receber outro e-mail pedindo a confirmação da sua presença. Fique de olho na sua caixa de entrada e no spam.'
    ];
    extra = botao(CONFIG.SITE_URL, 'Ver minha inscrição', '#0B3D91');
  } else if (tipo === 'espera') {
    const pos = posicaoAtual(insc, lista);
    assunto = ev + ': você está na lista de espera';
    paragrafos = [
      'Obrigado pela sua inscrição no ' + ev + '. As ' + CONFIG.VAGAS + ' vagas presenciais foram preenchidas no sorteio, e você está na lista de espera, na posição ' + pos + '.',
      'Se alguém liberar a vaga, você recebe um e-mail para confirmar. Fique de olho na sua caixa de entrada e no spam.'
    ];
    extra = botao(CONFIG.SITE_URL, 'Acompanhar minha posição', '#0B3D91');
  } else if (tipo === 'confirmacao') {
    assunto = ev + ': confirme sua presença até ' + prazo;
    paragrafos = [
      'Você foi sorteado(a) para o ' + ev + ', de ' + CONFIG.PERIODO_TEXTO + '.',
      'Você vai participar do evento presencial? Responda até <b>' + prazo + '</b>. Se não responder, a vaga passa para a lista de espera.'
    ];
    extra = botoesResposta;
  } else if (tipo === 'chamada') {
    assunto = ev + ': abriu uma vaga para você!';
    paragrafos = [
      'Boa notícia: abriu uma vaga presencial no ' + ev + ', de ' + CONFIG.PERIODO_TEXTO + ', e ela é sua.',
      'Você vai participar do evento presencial? Responda até <b>' + prazo + '</b>. Se não responder, a vaga passa para a próxima pessoa da lista.'
    ];
    extra = botoesResposta;
  } else {
    throw new Error('Tipo de e-mail desconhecido: ' + tipo);
  }

  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1A2233">' +
    '<div style="background:#0B3D91;color:#fff;padding:18px 22px;border-radius:12px 12px 0 0;font-size:22px;font-weight:bold">' +
    'COINES <span style="color:#F5B700">2027</span></div>' +
    '<div style="border:1px solid #DDE3EE;border-top:0;padding:22px;border-radius:0 0 12px 12px;font-size:16px;line-height:1.5">' +
    '<p>Olá, ' + escHtml(nome) + '!</p>' +
    paragrafos.map(p => '<p>' + p + '</p>').join('') +
    '<p style="margin:24px 0">' + extra + '</p>' +
    '<p style="color:#5B6478;font-size:13px">' + CONFIG.ASSINATURA + '<br>' + CONFIG.INSTITUICAO + '</p>' +
    '</div></div>';

  const texto = 'Olá, ' + nome + '!\n\n' +
    paragrafos.map(p => p.replace(/<[^>]+>/g, '')).join('\n\n') + '\n\n' +
    (tipo === 'confirmacao' || tipo === 'chamada'
      ? 'Sim, vou: ' + linkSim + '\nNão vou poder ir: ' + linkNao
      : CONFIG.SITE_URL) +
    '\n\n' + CONFIG.ASSINATURA;

  return { assunto: assunto, html: html, texto: texto };
}

function botao(url, rotulo, cor) {
  return '<a href="' + url + '" style="display:inline-block;background:' + cor +
    ';color:#fff;text-decoration:none;padding:12px 18px;border-radius:10px;font-weight:bold;margin:4px 0">' + rotulo + '</a>';
}

function escHtml(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
