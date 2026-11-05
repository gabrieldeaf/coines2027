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
  DIAS_EVENTO: ['2027-10-20', '2027-10-21', '2027-10-22'], // formato AAAA-MM-DD
  PERIODO_TEXTO: '20 a 22 de outubro de 2027',
  FREQ_MINIMA: 75,              // % mínimo para liberar o certificado do evento
  CARGA_HORARIA_EVENTO: 24,     // horas
  VALIDAR_HORARIO_MINICURSO: true, // só aceita o QR do minicurso no dia e horário dele
  TOLERANCIA_MIN: 30,           // minutos antes do início e depois do fim
  HORA_TARDE: '12:00',          // a partir deste horário o admin sugere o lanche da tarde
  FUSO: 'America/Sao_Paulo',
  PREFIXO_QR: 'COINES27'
};

const ABAS = {
  Inscritos: ['Codigo', 'CPF', 'Nome', 'Email', 'Telefone', 'Instituicao', 'Identidade', 'Atuacao', 'Cidade', 'UF', 'DataInscricao'],
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

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return saida({ ok: false, erro: 'Requisição inválida.' });
  }
  const acoes = {
    buscarCpf, inscrever, painel, registrarMinicurso, certificado,
    adminLogin, adminScan, adminTabela
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

function ler(nome) {
  const sh = aba(nome);
  const ult = sh.getLastRow();
  if (ult < 2) return [];
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return sh.getRange(2, 1, ult - 1, cab.length).getDisplayValues()
    .filter(l => l.join('') !== '')
    .map(l => {
      const o = {};
      cab.forEach((c, i) => { o[String(c).trim()] = String(l[i]).trim(); });
      return o;
    });
}

function gravar(nome, obj) {
  // Grava pelo nome da coluna (não pela posição) e usa apóstrofo para forçar texto:
  // mantém os zeros do CPF e evita fórmulas digitadas.
  const sh = aba(nome);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(c => String(c).trim());
  const linha = cab.map(c => (obj[c] === undefined || obj[c] === '') ? '' : "'" + obj[c]);
  sh.appendRow(linha);
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
      DataInscricao: agora()
    };
    gravar('Inscritos', insc);
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
      ['MC01', 'Introdução à Libras', '20/10/2027', '09:00', '12:00', 'Sala 1', '3', 'Nome do ministrante'],
      ['MC02', 'Literatura surda', '21/10/2027', '14:00', '17:00', 'Sala 2', '3', 'Nome do ministrante']
    ]);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_SENHA')) props.setProperty('ADMIN_SENHA', 'coines2027');
  Logger.log('Pronto. Senha do admin: ' + props.getProperty('ADMIN_SENHA') +
    ' (troque em Configurações do projeto > Propriedades do script > ADMIN_SENHA)');
}
