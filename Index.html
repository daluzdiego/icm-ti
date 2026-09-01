/**
 * ============================================================================
 *  SISTEMA ICM 3.0 - Escola Cristo Rei (Rede ICM)
 *  Code.gs - Backend completo (Google Apps Script)
 * ============================================================================
 *  Estrutura: 14 abas na planilha vinculada (ver função setupSpreadsheet()).
 *  Autenticação: código de 4 dígitos, sessão guardada em CacheService
 *  (ScriptCache) e retornada ao cliente para uso nas chamadas protegidas.
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// CONFIGURAÇÃO GERAL
// ----------------------------------------------------------------------------

const SHEET_NAMES = {
  USUARIOS: 'USUARIOS',
  CHROMEBOOKS: 'CHROMEBOOKS',
  RECURSOS: 'RECURSOS',
  HORARIOS_TURMAS: 'HORARIOS_TURMAS',
  AGENDAMENTOS: 'AGENDAMENTOS',
  RESERVAS_LABORATORIO: 'RESERVAS_LABORATORIO',
  HORARIOS_LABORATORIO: 'HORARIOS_LABORATORIO',
  LOG_CHROMEBOOKS: 'LOG_CHROMEBOOKS',
  RETIRADAS_CHROMEBOOKS: 'RETIRADAS_CHROMEBOOKS',
  VINCULOS_CHROMEBOOKS: 'VINCULOS_CHROMEBOOKS',
  LOG_RECURSOS: 'LOG_RECURSOS',
  AVARIAS: 'AVARIAS',
  MANUTENCAO: 'MANUTENCAO',
  TOKENS_ALUNOS: 'TOKENS_ALUNOS',
  CONFIGURACOES: 'CONFIGURACOES',
  LOG_SISTEMA: 'LOG_SISTEMA',
  AUDITORIA_CHROMEBOOKS: 'AUDITORIA_CHROMEBOOKS',
  AUDITORIA_LABORATORIO: 'AUDITORIA_LABORATORIO'
};

const SHEET_HEADERS = {
  USUARIOS: ['ID','Nome','Email','Codigo_4_Digitos','Tipo','Turno','Ativo','Chromebook_Fixo_ID','Foto_URL','Data_Cadastro'],
  CHROMEBOOKS: ['ID','Tipo','Status','Responsavel_ID','Data_Retirada','Localizacao','Observacoes','Avaria_Desc','Avaria_Data','Avaria_Status'],
  RECURSOS: ['ID','Tipo','Nome','Status','Responsavel_ID','Data_Retirada','Localizacao','Observacoes'],
  HORARIOS_TURMAS: ['Horario_ID','Dia_Semana','Periodo','Hora_Inicio','Hora_Fim','Turma_ID','Turma','Professor_ID','Professor_Nome','Sala','Disciplina','Ativo'],
  AGENDAMENTOS: ['ID','Usuario_ID','Data','Horario','Tipo','Quantidade','Chromes_IDS','Turma','Observacao','Status','Data_Criacao'],
  RESERVAS_LABORATORIO: ['ID','Data','Usuario_ID','Professor','Email','Horario','Chromes','Chromes_IDS','Fones','VR','Projetor','Turma','Pauta','Status','Timestamp'],
  HORARIOS_LABORATORIO: ['Inicio','Fim','Ordem','Status','Bloqueado','Label'],
  LOG_CHROMEBOOKS: ['ID','Data_Hora','Acao','Usuario_ID','Usuario_Nome','Quantidade','Chromes_IDS','Tipo','Avaria_Desc'],
  RETIRADAS_CHROMEBOOKS: ['ID','Chromebook_ID','Retirado_Por_ID','Retirado_Por_Nome','Responsavel_ID','Responsavel_Nome','Data_Retirada','Data_Retorno_Prevista','Data_Devolucao_Real','Turma','Sala','Tipo_Retirada','Status'],
  VINCULOS_CHROMEBOOKS: ['Chromebook_ID','Turno','Professor_Nome','Professor_ID','Ativo','Observacoes'],
  LOG_RECURSOS: ['ID','Data_Hora','Acao','Usuario_ID','Usuario_Nome','Recurso_ID','Tipo'],
  AVARIAS: ['ID','Equipamento_ID','Tipo','Data','Usuario_ID','Descricao','Status','Data_Conserto','Tecnico','Observacoes'],
  MANUTENCAO: ['ID','Equipamento_ID','Tipo','Data_Entrada','Data_Saida','Status','Tecnico','Observacoes'],
  TOKENS_ALUNOS: ['Token','Professor_ID','Quantidade_Liberada','Data_Criacao','Expiracao','Retirados','Status'],
  CONFIGURACOES: ['Configuracao','Valor'],
  LOG_SISTEMA: ['ID','Timestamp','Nivel','Acao','Usuario_ID','Mensagem','Detalhes'],
  AUDITORIA_CHROMEBOOKS: ['AUDITORIA_ID','DATA_HORA','ACAO','CHROMEBOOK_ID','PROFESSOR_ID','PROFESSOR_NOME','TIPO_USO','CONTEXTO','TURNO','DIA_SEMANA','HORARIO_AULA','TURMA','DISCIPLINA','SALA','LOCAL_ATUAL','ORIGEM','STATUS','OBSERVACAO'],
  AUDITORIA_LABORATORIO: ['AUDITORIA_ID','DATA_HORA','ACAO','RESERVA_ID','PROFESSOR_ID','PROFESSOR_NOME','DATA_RESERVA','HORARIO_INICIO','HORARIO_FIM','TURMA','LABORATORIO','CHROMEBOOK','FONES','PROJETOR','ATIVIDADE','ORIGEM','STATUS','OBSERVACAO']
};

const DEFAULT_CONFIG = {
  ESTOQUE_MAXIMO_CHROMES_ALUNOS: 32,
  ESTOQUE_MAXIMO_FONES: 13,
  ESTOQUE_MAXIMO_VR: 13,
  ESTOQUE_MAXIMO_PROJETOR: 1,
  TOTAL_CHROMEBOOKS_PROFESSORES: 17,
  DOMINIO_INSTITUCIONAL: '@redeicm.org.br',
  TIMEZONE: 'America/Sao_Paulo',
  EMAIL_SUPORTE: 'suporte@redeicm.org.br',
  EMAIL_RELATORIO: 'diego.daluz@redeicm.org.br',
  ALERTA_ESTOQUE_BAIXO: 20,
  CALENDAR_ID: '',
  WEB_APP_URL: '',
  AUDITORIA_SPREADSHEET_ID: '1GTXzcRIgtT7XBiWhHFGKhAHXVNeojbB01nbNtVv6bqU',
  AUDITORIA_ATUALIZACAO_MINUTOS: 5
};

// ----------------------------------------------------------------------------
// ENTRY POINTS (Web App)
// ----------------------------------------------------------------------------

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ICM T.I - Gestão Inteligente de Recursos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------------------------------
// UTILITÁRIOS DE PLANILHA
// ----------------------------------------------------------------------------

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (SHEET_HEADERS[name]) sh.appendRow(SHEET_HEADERS[name]);
  }
  return sh;
}

/**
 * Converte dados do Apps Script para tipos aceitos pelo google.script.run.
 * Em especial, Date não pode atravessar a fronteira servidor -> navegador.
 * A conversão é feita recursivamente sem alterar o objeto original.
 */
function serializarParaCliente(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }

  if (Array.isArray(valor)) {
    return valor.map(serializarParaCliente);
  }

  if (valor && typeof valor === 'object') {
    const resultado = {};
    Object.keys(valor).forEach(chave => {
      resultado[chave] = serializarParaCliente(valor[chave]);
    });
    return resultado;
  }

  return valor;
}

let __SHEET_MEMO = Object.create(null);
function limparMemoPlanilha(name) { try { delete __SHEET_MEMO[name]; } catch (e) {} }
function sheetToObjects(name) {
  if (__SHEET_MEMO[name]) return __SHEET_MEMO[name];
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return (__SHEET_MEMO[name] = []);
  const headers = values[0];
  const result = values.slice(1).filter(r => r.join('') !== '').map((r, idx) => {
    const obj = { _row: idx + 2 };
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
  __SHEET_MEMO[name] = result;
  return result;
}

function findRowById(name, idField, idValue) {
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf(idField);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(idValue)) {
      return { rowIndex: i + 1, headers: headers, values: values[i] };
    }
  }
  return null;
}

function updateCell(name, rowIndex, headerName, value) {
  limparMemoPlanilha(name);
  const sh = getSheet(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf(headerName) + 1;
  if (col > 0) sh.getRange(rowIndex, col).setValue(value);
}

function updateRowByFields(name, rowIndex, fieldsObj) {
  limparMemoPlanilha(name);
  const sh = getSheet(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(fieldsObj).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sh.getRange(rowIndex, col).setValue(fieldsObj[key]);
  });
}

function appendObject(name, obj) {
  limparMemoPlanilha(name);
  const sh = getSheet(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
  sh.appendRow(row);
  return row;
}

function appendObjectsBatch(name, objects) {
  if (!Array.isArray(objects) || !objects.length) return;
  limparMemoPlanilha(name);
  const sh = getSheet(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = objects.map(obj => headers.map(h => (obj[h] !== undefined ? obj[h] : '')));
  sh.getRange(sh.getLastRow()+1, 1, rows.length, headers.length).setValues(rows);
}

function atualizarLinhasBatch(name, updates) {
  if (!Array.isArray(updates) || !updates.length) return;
  limparMemoPlanilha(name);
  const sh = getSheet(name);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const minRow = Math.max(2, Math.min(...updates.map(x=>x.rowIndex)));
  const maxRow = Math.min(lastRow, Math.max(...updates.map(x=>x.rowIndex)));
  const values = sh.getRange(minRow,1,maxRow-minRow+1,lastCol).getValues();
  const indexByRow = new Map();
  updates.forEach(u=>indexByRow.set(Number(u.rowIndex),u.fields||{}));
  for(let i=0;i<values.length;i++){
    const rowNumber=minRow+i, fields=indexByRow.get(rowNumber);
    if(!fields) continue;
    Object.keys(fields).forEach(key=>{const col=headers.indexOf(key);if(col>=0)values[i][col]=fields[key];});
  }
  sh.getRange(minRow,1,values.length,lastCol).setValues(values);
}

function ensureColumn(name, header) {
  const sh = getSheet(name);
  const lastColumn = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const existing = headers.indexOf(header) + 1;
  if (existing > 0) return existing;
  const col = lastColumn + 1;
  sh.getRange(1, col).setValue(header);
  return col;
}
function salvarCalendarEventId(sheetName, id, eventId) {
  if (!id || !eventId) return false;
  try { ensureColumn(sheetName, 'Calendar_Event_ID'); const row=findRowById(sheetName,'ID',id); if(!row)return false; updateCell(sheetName,row.rowIndex,'Calendar_Event_ID',eventId); return true; }
  catch(err){ logSistema('WARNING','CALENDAR_ID_SALVAR_ERRO','',err.message); return false; }
}
function obterCalendarEventId(row) {
  if(!row||!row.headers||!row.values)return '';
  const i=row.headers.indexOf('Calendar_Event_ID');
  return i>=0?String(row.values[i]||''):'';
}

function genId(prefix) {
  return prefix + '_' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function nowStr() {
  return Utilities.formatDate(new Date(), DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function logSistema(nivel, acao, usuarioId, mensagem, detalhes) {
  try {
    appendObject(SHEET_NAMES.LOG_SISTEMA, {
      ID: genId('LOG'),
      Timestamp: nowStr(),
      Nivel: nivel,
      Acao: acao,
      Usuario_ID: usuarioId || '',
      Mensagem: mensagem || '',
      Detalhes: detalhes ? JSON.stringify(detalhes) : ''
    });
  } catch (err) {
    // Evita quebra do fluxo principal por erro de log
  }
}

function getConfigValor(chave) {
  const configs = sheetToObjects(SHEET_NAMES.CONFIGURACOES);
  const found = configs.find(c => c.Configuracao === chave);
  return found ? found.Valor : DEFAULT_CONFIG[chave];
}

// ----------------------------------------------------------------------------
// SETUP INICIAL (rodar manualmente uma vez ao implantar)
// ----------------------------------------------------------------------------

function garantirColunasSistema(nome, colunas){
  const sh=getSheet(nome);
  const wanted=Array.isArray(colunas)?colunas:[];
  if(!wanted.length)return sh;
  const lastCol=Math.max(1,sh.getLastColumn());
  let headers=sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
  if(!headers.some(Boolean)){headers=[];sh.getRange(1,1,1,wanted.length).setValues([wanted]);return sh;}
  const faltantes=wanted.filter(c=>!headers.includes(c));
  if(faltantes.length){
    sh.insertColumnsAfter(Math.max(1,sh.getLastColumn()),faltantes.length);
    sh.getRange(1,headers.length+1,1,faltantes.length).setValues([faltantes]);
  }
  return sh;
}
function idsChromebooksSistema(valor){
  if(Array.isArray(valor))return valor.map(String).map(x=>x.trim()).filter(Boolean);
  const s=String(valor||'').trim();
  if(!s)return [];
  try{const a=JSON.parse(s);if(Array.isArray(a))return a.map(String).map(x=>x.trim()).filter(Boolean);}catch(e){}
  return s.split(/[,;\n|]+/).map(x=>x.trim()).filter(Boolean);
}
function idsChromebooksTextoSistema(ids){return JSON.stringify((ids||[]).map(String).map(x=>x.trim()).filter(Boolean));}
function idsComprometidosDoRegistroSistema(r){return idsChromebooksSistema(r && r.Chromes_IDS);}

function prepararControleIndividualChromebooks(){
  garantirColunasSistema(SHEET_NAMES.AGENDAMENTOS,['Chromes_IDS']);
  garantirColunasSistema(SHEET_NAMES.RESERVAS_LABORATORIO,['Chromes_IDS']);
  limparMemoPlanilha(SHEET_NAMES.AGENDAMENTOS); limparMemoPlanilha(SHEET_NAMES.RESERVAS_LABORATORIO);
  return {sucesso:true,mensagem:'Controle individual de Chromebooks preparado.'};
}
function revisarVinculosChromebooks(){
  const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS).filter(u=>String(u.Ativo||'').toUpperCase()==='SIM'&&String(u.Tipo||'').toLowerCase().startsWith('prof'));
  const vinculos=sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS).filter(v=>valorBooleanoSistema(v.Ativo,true));
  const grade=sheetToObjects(SHEET_NAMES.HORARIOS_TURMAS);
  const resultado=usuarios.map(u=>{
    const nome=normalizarNomeSistema(u.Nome), turnoUsuario=normalizarNomeSistema(u.Turno||'');
    const vs=vinculos.filter(v=>String(v.Professor_ID||'')===String(u.ID)||normalizarNomeSistema(v.Professor_Nome)===nome);
    const turnosGrade=new Set();
    grade.filter(r=>normalizarNomeSistema(r.Professor_Nome||r.Professor)===nome&&(r.Ativo===undefined||valorBooleanoSistema(r.Ativo,true))).forEach(r=>{
      const ini=minutosDoHorarioSistema(r.Hora_Inicio||r.Horario||'');
      if(ini!==null)turnosGrade.add(ini<12*60?'Manhã':'Tarde');
      else if(String(r.Periodo||'').trim()) turnosGrade.add(Number(r.Periodo)<=6?'Manhã':'Tarde');
    });
    const manha=vs.find(v=>normalizarNomeSistema(v.Turno)==='manha');
    const tarde=vs.find(v=>normalizarNomeSistema(v.Turno)==='tarde');
    const esperado=turnoUsuario==='ambos'?['manha','tarde']:turnoUsuario?[turnoUsuario]:[];
    const faltantes=esperado.filter(t=>!vs.some(v=>normalizarNomeSistema(v.Turno)===t));
    const vinculosSemGrade=vs.filter(v=>{const t=normalizarNomeSistema(v.Turno);return turnosGrade.size&&!turnosGrade.has(t==='manha'?'Manhã':'Tarde');});
    return {id:u.ID,nome:u.Nome,turno:u.Turno,chromebooks:vs.map(v=>({id:v.Chromebook_ID,turno:v.Turno,nome:v.Professor_Nome})),turnosNaGrade:[...turnosGrade],faltantes,vinculosSemGrade:vinculosSemGrade.map(v=>({id:v.Chromebook_ID,turno:v.Turno,nome:v.Professor_Nome})),status:(faltantes.length||vinculosSemGrade.length)?'REVISAR':'OK'};
  });
  return serializarParaCliente({sucesso:true,total:resultado.length,revisar:resultado.filter(x=>x.status!=='OK'),vinculos:resultado});
}

function setupSpreadsheet() {
  garantirColunasSistema(SHEET_NAMES.AGENDAMENTOS,['Chromes_IDS']);
  garantirColunasSistema(SHEET_NAMES.RESERVAS_LABORATORIO,['Chromes_IDS']);
  Object.keys(SHEET_HEADERS).forEach(name => {
    if (name === SHEET_NAMES.AUDITORIA_CHROMEBOOKS || name === SHEET_NAMES.AUDITORIA_LABORATORIO) return;
    const sh = getSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEET_HEADERS[name]);
    }
  });

  seedVinculosChromebooks();
  sincronizarProfessorIdsVinculos();

  // Popula CONFIGURACOES se vazio
  const confSheet = getSheet(SHEET_NAMES.CONFIGURACOES);
  if (confSheet.getLastRow() < 2) {
    Object.keys(DEFAULT_CONFIG).forEach(key => {
      confSheet.appendRow([key, DEFAULT_CONFIG[key]]);
    });
  }

  // Usuários de exemplo
  const usersSheet = getSheet(SHEET_NAMES.USUARIOS);
  if (usersSheet.getLastRow() < 2) {
    const exemplo = [
      ['USR_001', 'Diego da Luz', 'diego.daluz@redeicm.org.br', '5798', 'admin', 'Ambos', 'SIM', 'PROF001', '', nowStr()],
      ['USR_002', 'Felipe Vieira', 'juliana.boettge@redeicm.org.br', '5239', 'professor', 'Manhã', 'SIM', 'PROF002', '', nowStr()],
      ['USR_003', 'Murilo Santos', 'murilo.santos@redeicm.org.br', '2340', 'professor', 'Manhã', 'SIM', 'PROF003', '', nowStr()]
    ];
    exemplo.forEach(r => usersSheet.appendRow(r));
  }

  // Chromebooks de exemplo (17 professores + 32 alunos)
  const chSheet = getSheet(SHEET_NAMES.CHROMEBOOKS);
  if (chSheet.getLastRow() < 2) {
    for (let i = 1; i <= 17; i++) {
      chSheet.appendRow(['PROF' + String(i).padStart(3, '0'), 'PROF', 'DISPONIVEL', '', '', 'Sala ' + String(i).padStart(2, '0'), '', '', '', '']);
    }
    for (let i = 1; i <= 32; i++) {
      chSheet.appendRow(['ALUNO' + String(i).padStart(3, '0'), 'ALUNO', 'DISPONIVEL', '', '', 'Laboratório 1', '', '', '', '']);
    }
  }

  // Recursos de exemplo (fones, VR, projetor, controles)
  const recSheet = getSheet(SHEET_NAMES.RECURSOS);
  if (recSheet.getLastRow() < 2) {
    for (let i = 1; i <= 13; i++) recSheet.appendRow(['FONE_' + String(i).padStart(3, '0'), 'FONE', 'Fone ' + i, 'DISPONIVEL', '', '', 'Laboratório', '']);
    for (let i = 1; i <= 13; i++) recSheet.appendRow(['VR_' + String(i).padStart(3, '0'), 'VR', 'VR ' + i, 'DISPONIVEL', '', '', 'Laboratório', '']);
    recSheet.appendRow(['PROJ_001', 'PROJETOR', 'Projetor 1', 'DISPONIVEL', '', '', 'Laboratório', '']);
    for (let i = 1; i <= 5; i++) recSheet.appendRow(['CONTROLE_' + String(i).padStart(3, '0'), 'CONTROLE', 'Controle ' + i, 'DISPONIVEL', '', '', '', '']);
  }

  // Horários padrão do laboratório
  const horSheet = getSheet(SHEET_NAMES.HORARIOS_LABORATORIO);
  if (horSheet.getLastRow() < 2) {
    const blocos = [
      ['07:00', '07:50', 1], ['07:50', '08:40', 2], ['08:40', '09:30', 3],
      ['09:50', '10:40', 4], ['10:40', '11:30', 5], ['13:30', '14:20', 6],
      ['14:20', '15:10', 7], ['15:30', '16:20', 8], ['16:20', '17:10', 9]
    ];
    blocos.forEach(b => horSheet.appendRow([b[0], b[1], b[2], 'ATIVO', 'FALSE', b[0] + ' - ' + b[1]]));
  }

  logSistema('INFO', 'SETUP', '', 'Planilha inicializada com sucesso');
  return 'Setup concluído com sucesso.';
}

// ----------------------------------------------------------------------------
// AUTENTICAÇÃO
// ----------------------------------------------------------------------------

function normalizarCodigoLogin(codigo) {
  // A planilha pode guardar o código como número, texto, com espaços ou até
  // perder um zero à esquerda. Normalizamos os dois lados para 4 dígitos.
  const limpo = String(codigo == null ? '' : codigo).replace(/\D/g, '');
  if (!limpo) return '';
  return limpo.padStart(4, '0').slice(-4);
}

function loginComCodigo(codigo) {
  try {
    const codigoNormalizado = normalizarCodigoLogin(codigo);
    if (codigoNormalizado.length !== 4) {
      return { sucesso: false, mensagem: 'Código deve ter 4 dígitos.' };
    }

    const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
    const candidatos = usuarios.filter(u => {
      const ativo = String(u.Ativo == null ? '' : u.Ativo).trim().toUpperCase();
      return ativo === 'SIM' && normalizarCodigoLogin(u.Codigo_4_Digitos) === codigoNormalizado;
    });

    if (!candidatos.length) {
      logSistema('WARNING', 'LOGIN_FALHA', '', 'Tentativa de login com código inválido.', {codigo: codigoNormalizado});
      return { sucesso: false, mensagem: 'Código inválido ou usuário inativo.' };
    }

    // Nunca escolher silenciosamente um usuário quando o mesmo código estiver
    // cadastrado para mais de uma pessoa. Isso evita entrar na conta errada.
    if (candidatos.length > 1) {
      logSistema('ERROR', 'LOGIN_CODIGO_DUPLICADO', '', 'Código de acesso cadastrado para mais de um usuário.', {codigo: codigoNormalizado});
      return { sucesso: false, mensagem: 'Este código está cadastrado para mais de um usuário. Peça ao administrador para corrigir o cadastro.' };
    }

    // Cria uma cópia para não alterar o objeto mantido no cache de planilha.
    const usuario = Object.assign({}, candidatos[0]);
    delete usuario._row;

    // Mantém o código em formato consistente na sessão, inclusive quando
    // o valor original da planilha perdeu zero à esquerda.
    usuario.Codigo_4_Digitos = codigoNormalizado;

    const token = Utilities.getUuid();
    const cache = CacheService.getScriptCache();
    cache.put('session_' + token, JSON.stringify(usuario), 21600); // 6 horas
    logSistema('INFO', 'LOGIN', usuario.ID, 'Login realizado com sucesso');

    // google.script.run não aceita objetos Date no retorno.
    const usuarioCliente = serializarParaCliente(usuario);
    return { sucesso: true, usuario: usuarioCliente, token: token };
  } catch (err) {
    logSistema('ERROR', 'LOGIN_ERRO', '', err.message);
    return { sucesso: false, mensagem: 'Erro ao processar login: ' + err.message };
  }
}

function getUsuarioLogado(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const data = cache.get('session_' + token);
  return data ? JSON.parse(data) : null;
}

function restaurarSessao(token) {
  const usuario = getUsuarioLogado(token);
  if (!usuario) return { sucesso: false, mensagem: 'Sessão expirada.' };
  return serializarParaCliente({ sucesso: true, usuario: usuario });
}

/**
 * Valida a sessão no servidor. Toda operação protegida deve passar por aqui.
 */
function exigirSessao(token) {
  const usuario = getUsuarioLogado(token);
  return usuario || null;
}

/**
 * Valida sessão e perfil administrativo.
 */
function perfilPodeAdministrar(usuario) {
  if (!usuario) return false;
  const tipo = String(usuario.Tipo || '').trim().toLowerCase();
  return ['admin','administrador','gestao','gestão'].includes(tipo);
}

function exigirAdmin(token) {
  const usuario = exigirSessao(token);
  if (!usuario) return null;
  return ['admin','administrador'].includes(String(usuario.Tipo||'').trim().toLowerCase()) ? usuario : null;
}

function perfilPodeRetirarParaProfessor(usuario) {
  if (!usuario) return false;
  const tipo = String(usuario.Tipo || '').trim().toLowerCase();
  return perfilPodeAdministrar(usuario) || ['atendente','assistente','secretaria','secretário'].includes(tipo);
}
function exigirPermissaoComunicacaoRelatorios(token) {
  const usuario = exigirSessao(token);
  if (!usuario) return { usuario:null, resposta:respostaSessaoExpirada() };
  const tipo = String(usuario.Tipo || '').trim().toLowerCase();
  if (!['admin','administrador'].includes(tipo)) return { usuario, resposta:respostaSemPermissao() };
  return { usuario, resposta:null };
}

function perfilPodeOperarReserva(usuario){
  if(!usuario)return false;
  const tipo=String(usuario.Tipo||'').trim().toLowerCase();
  return ['admin','administrador','gestao','gestão','atendente','assistente','secretaria','secretário'].includes(tipo);
}

function exigirProprioOuOperador(token, usuarioId){
  const usuario=exigirSessao(token);
  if(!usuario)return null;
  return (String(usuario.ID)===String(usuarioId)||perfilPodeOperarReserva(usuario))?usuario:null;
}

/**
 * Permite a operação ao próprio usuário ou a um administrador.
 */
function exigirProprioOuAdmin(token, usuarioId) {
  const usuario = exigirSessao(token);
  if (!usuario) return null;
  const ehAdmin = String(usuario.Tipo || '').toLowerCase() === 'admin';
  const ehProprio = String(usuario.ID) === String(usuarioId);
  return (ehAdmin || ehProprio) ? usuario : null;
}

function respostaSessaoExpirada() {
  return { sucesso: false, mensagem: 'Sessão expirada. Faça login novamente.' };
}

function respostaSemPermissao() {
  return { sucesso: false, mensagem: 'Você não tem permissão para realizar esta operação.' };
}

function logout(token) {
  if (token) {
    const cache = CacheService.getScriptCache();
    cache.remove('session_' + token);
  }
  return { sucesso: true };
}

function getEmailSuporte() {
  return getConfigValor('EMAIL_SUPORTE');
}

// ----------------------------------------------------------------------------
// DASHBOARD
// ----------------------------------------------------------------------------

function getDashboardData(token) {
  const usuario = getUsuarioLogado(token);
  if (!usuario) return { sucesso: false, mensagem: 'Sessão expirada.' };

  // Dashboard somente leitura: não atualiza localização nem relê a planilha.
  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  const recursos = sheetToObjects(SHEET_NAMES.RECURSOS);
  const reservas = sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO);
  const hoje = Utilities.formatDate(new Date(), DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const resumoChrome = {
    total: chromes.length,
    totalAlunos: chromes.filter(c => String(c.Tipo).toUpperCase() === 'ALUNO').length,
    totalProfessores: chromes.filter(c => String(c.Tipo).toUpperCase() === 'PROF').length,
    disponiveis: chromes.filter(c => String(c.Status).toUpperCase() === 'DISPONIVEL').length,
    disponiveisAlunos: chromes.filter(c => String(c.Tipo).toUpperCase() === 'ALUNO' && String(c.Status).toUpperCase() === 'DISPONIVEL').length,
    disponiveisProfessores: chromes.filter(c => String(c.Tipo).toUpperCase() === 'PROF' && String(c.Status).toUpperCase() === 'DISPONIVEL').length,
    emUso: chromes.filter(c => String(c.Status).toUpperCase() === 'EM_USO').length,
    emUsoAlunos: chromes.filter(c => String(c.Tipo).toUpperCase() === 'ALUNO' && String(c.Status).toUpperCase() === 'EM_USO').length,
    manutencao: chromes.filter(c => String(c.Status).toUpperCase() === 'MANUTENCAO').length,
    manutencaoAlunos: chromes.filter(c => String(c.Tipo).toUpperCase() === 'ALUNO' && String(c.Status).toUpperCase() === 'MANUTENCAO').length
  };

  const statsRecursos = {};
  recursos.forEach(r => {
    const tipo = String(r.Tipo || '').toUpperCase();
    if (!statsRecursos[tipo]) statsRecursos[tipo] = { total: 0, disponivel: 0, emUso: 0, manutencao: 0 };
    statsRecursos[tipo].total++;
    const status = String(r.Status || '').toUpperCase();
    if (status === 'DISPONIVEL') statsRecursos[tipo].disponivel++;
    if (status === 'EM_USO') statsRecursos[tipo].emUso++;
    if (status === 'MANUTENCAO') statsRecursos[tipo].manutencao++;
  });

  const reservasAtivas = reservas.filter(r => String(r.Status).toUpperCase() === 'ATIVA');
  const reservasHoje = reservasAtivas.filter(r => normalizarDataSistema(r.Data) === hoje);
  const proximasReservas = reservasAtivas
    .map(r => {
      const h = intervaloHorarioLaboratorio(r.Data, r.Horario);
      return {
        id: r.ID,
        data: normalizarDataSistema(r.Data),
        horario: h ? h.inicio + ' - ' + h.fim : String(r.Horario || ''),
        professor: r.Professor || '',
        turma: r.Turma || '',
        chromes: quantidadeNumericaSistema(r.Chromes),
        status: 'ATIVA'
      };
    })
    .filter(r => r.data >= hoje)
    .sort((a,b) => (a.data + ' ' + a.horario).localeCompare(b.data + ' ' + b.horario))
    .slice(0, 6);

  return {
    sucesso: true,
    usuario: usuario,
    chromebooks: resumoChrome,
    chromebooksDisponiveis: resumoChrome.disponiveisAlunos,
    chromebooksEmUso: resumoChrome.emUso,
    agendamentosHoje: reservasHoje.length,
    reservasHoje: reservasHoje.length,
    recursosCadastrados: recursos.length,
    recursosDisponiveis: Object.keys(statsRecursos).reduce((s,k) => s + statsRecursos[k].disponivel, 0),
    recursosStats: statsRecursos,
    proximasReservas: proximasReservas
  };
}

// ----------------------------------------------------------------------------
// CHROMEBOOKS
// ----------------------------------------------------------------------------

/**
 * Analítico do Dashboard.
 * Usa os registros operacionais já existentes para não depender da planilha
 * separada de auditoria na abertura da tela. Cada retirada conta como uma
 * utilização registrada; o tipo do Chromebook diferencia fixo de aluno.
 */
function getDashboardAnalytics(token, professorId, dias) {
  const usuario = exigirSessao(token);
  if (!usuario) return respostaSessaoExpirada();

  const ehAdmin = perfilPodeAdministrar(usuario);
  const idSolicitado = String(professorId || '').trim();
  const professorSelecionado = ehAdmin ? idSolicitado : String(usuario.ID);
  const quantidadeDias = Math.min(90, Math.max(7, Number(dias) || 30));
  const cache=CacheService.getScriptCache();
  const cacheKey='ICM_DASH_ANALYTICS_'+String(usuario.ID)+'_'+String(professorSelecionado||'TODOS')+'_'+quantidadeDias;
  try { const cached=cache.get(cacheKey); if(cached) return JSON.parse(cached); } catch(e) {}

  const hojeDate = new Date();
  const hoje = Utilities.formatDate(hojeDate, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const inicioDate = new Date(hojeDate.getTime() - (quantidadeDias - 1) * 86400000);
  const inicio = Utilities.formatDate(inicioDate, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const profMap = {};
  usuarios.forEach(u => {
    if (String(u.Ativo).toUpperCase() === 'SIM' && ['professor','professora'].includes(String(u.Tipo || '').trim().toLowerCase())) {
      profMap[String(u.ID)] = String(u.Nome || u.ID);
    }
  });

  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  const tipoChrome = {};
  chromes.forEach(c => tipoChrome[String(c.ID)] = String(c.Tipo || '').toUpperCase());

  const retiradas = sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  const reservas = sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO);

  const inRange = d => d >= inicio && d <= hoje;
  const diasSerie = [];
  for (let i = 0; i < quantidadeDias; i++) {
    const d = new Date(inicioDate.getTime() + i * 86400000);
    diasSerie.push(Utilities.formatDate(d, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  }
  const mapaDia = {};
  diasSerie.forEach(d => mapaDia[d] = { laboratorio: 0, fixos: 0, alunos: 0 });

  const usoFixos = [];
  const usoAlunos = [];
  retiradas.forEach(r => {
    const d = normalizarDataSistema(r.Data_Retirada);
    if (!d || !inRange(d) || String(r.Status || '').toUpperCase() === 'CANCELADA') return;
    const responsavel = String(r.Responsavel_ID || '');
    if (professorSelecionado && responsavel !== professorSelecionado) return;
    const tipo = tipoChrome[String(r.Chromebook_ID)] || '';
    if (tipo === 'PROF') {
      usoFixos.push(r);
      if (mapaDia[d]) mapaDia[d].fixos++;
    } else if (tipo === 'ALUNO') {
      usoAlunos.push(r);
      if (mapaDia[d]) mapaDia[d].alunos++;
    }
  });

  const reservasUso = reservas.filter(r => {
    const d = normalizarDataSistema(r.Data);
    const status = String(r.Status || '').toUpperCase();
    const prof = String(r.Usuario_ID || '');
    const profNome = String(r.Professor || '');
    const pertence = professorSelecionado
      ? (prof === professorSelecionado || normalizarNomeSistema(profNome) === normalizarNomeSistema(profMap[professorSelecionado] || ''))
      : true;
    return d && inRange(d) && pertence && ['ATIVA','CONCLUIDA','REALIZADA','FINALIZADA'].includes(status);
  });
  reservasUso.forEach(r => {
    const d = normalizarDataSistema(r.Data);
    if (mapaDia[d]) mapaDia[d].laboratorio++;
  });

  const labels = diasSerie.map(d => {
    const p = d.split('-');
    return p[2] + '/' + p[1];
  });

  const rankingMap = {};
  Object.keys(profMap).forEach(id => rankingMap[id] = { id, nome: profMap[id], laboratorio: 0, fixos: 0, alunos: 0, total: 0 });
  retiradas.forEach(r => {
    const d = normalizarDataSistema(r.Data_Retirada);
    const pid = String(r.Responsavel_ID || '');
    if (!d || !inRange(d) || !rankingMap[pid]) return;
    const tipo = tipoChrome[String(r.Chromebook_ID)] || '';
    if (tipo === 'PROF') rankingMap[pid].fixos++;
    if (tipo === 'ALUNO') rankingMap[pid].alunos++;
  });
  reservas.filter(r => {
    const d = normalizarDataSistema(r.Data);
    const status = String(r.Status || '').toUpperCase();
    return d && inRange(d) && ['ATIVA','CONCLUIDA','REALIZADA','FINALIZADA'].includes(status);
  }).forEach(r => {
    const pid = String(r.Usuario_ID || '');
    if (rankingMap[pid]) rankingMap[pid].laboratorio++;
  });
  const ranking = Object.values(rankingMap).map(x => ({ ...x, total: x.laboratorio + x.fixos + x.alunos })).sort((a,b) => b.total - a.total || a.nome.localeCompare(b.nome,'pt-BR')).slice(0, 12);

  const recursos = {
    laboratorio: reservasUso.length,
    fixos: usoFixos.length,
    alunos: usoAlunos.length,
    fones: reservasUso.reduce((s,r) => s + quantidadeNumericaSistema(r.Fones), 0),
    vr: reservasUso.reduce((s,r) => s + quantidadeNumericaSistema(r.VR), 0),
    projetor: reservasUso.reduce((s,r) => s + quantidadeNumericaSistema(r.Projetor), 0)
  };

  const selectedName = professorSelecionado ? (profMap[professorSelecionado] || 'Professor selecionado') : 'Todos os professores';
  const totalAtividades = reservasUso.length + usoFixos.length + usoAlunos.length;

  const resposta=serializarParaCliente({
    sucesso: true,
    inicio,
    fim: hoje,
    dias: quantidadeDias,
    professorId: professorSelecionado,
    professorNome: selectedName,
    totalAtividades,
    indicadores: {
      laboratorio: reservasUso.length,
      fixos: usoFixos.length,
      alunos: usoAlunos.length,
      total: totalAtividades
    },
    serie: { labels, laboratorio: diasSerie.map(d => mapaDia[d].laboratorio), fixos: diasSerie.map(d => mapaDia[d].fixos), alunos: diasSerie.map(d => mapaDia[d].alunos) },
    recursos,
    ranking
  });
  try { cache.put(cacheKey,JSON.stringify(resposta),15); } catch(e) {}
  return resposta;
}

function getEstoqueChrome() {
  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  const fixos = chromes.filter(c => String(c.Tipo).toUpperCase() === 'PROF');
  const alunos = chromes.filter(c => String(c.Tipo).toUpperCase() === 'ALUNO');
  const alunosDisponiveis = alunos.filter(c => String(c.Status).toUpperCase() === 'DISPONIVEL').length;
  const alunosEmUso = alunos.filter(c => String(c.Status).toUpperCase() === 'EM_USO').length;
  const alunosManutencao = alunos.filter(c => String(c.Status).toUpperCase() === 'MANUTENCAO').length;
  return {
    fixosTotal: fixos.length,
    fixosDisponiveis: fixos.filter(c => String(c.Status).toUpperCase() === 'DISPONIVEL').length,
    alunosTotal: alunos.length,
    alunosDisponiveis: alunosDisponiveis,
    alunosEmUso: alunosEmUso,
    alunosManutencao: alunosManutencao,
    total: chromes.length,
    totalDisponivel: alunosDisponiveis
  };
}

function getInventarioChromebooks(token) {
  const usuario = exigirSessao(token);
  if (!usuario) return respostaSessaoExpirada();

  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const nomes = {};
  usuarios.forEach(u => nomes[String(u.ID)] = u.Nome || String(u.ID));

  const retiradasAtivas = getSheet(SHEET_NAMES.RETIRADAS_CHROMEBOOKS).getLastRow()>1
    ? sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS).filter(r=>String(r.Status||'').toUpperCase()==='ABERTA')
    : [];
  const retiradaMap={};
  retiradasAtivas.forEach(r=>retiradaMap[String(r.Chromebook_ID)]=r);

  const chromesBase = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  chromesBase.filter(c=>String(c.Tipo||'').toUpperCase()==='PROF'&&String(c.Status||'').toUpperCase()==='EM_USO')
    .forEach(c=>{ atualizarLocalizacaoChromebookProfessor(c.ID,'CONSULTA_INVENTARIO'); });
  const chromesAtualizados = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  const chromes = chromesAtualizados.map(c => {
    const ret=retiradaMap[String(c.ID)]||null;
    const contextoAtual=(String(c.Tipo||'').toUpperCase()==='PROF'&&String(c.Status||'').toUpperCase()==='EM_USO') ? (obterLocalizacaoAtualChromebook(c.ID,new Date())||null) : null;
    return {
      id: c.ID,
      tipo: String(c.Tipo || '').toUpperCase(),
      status: String(c.Status || '').toUpperCase(),
      responsavelId: c.Responsavel_ID || '',
      responsavel: c.Responsavel_ID ? (nomes[String(c.Responsavel_ID)] || String(c.Responsavel_ID)) : '',
      localizacao: c.Localizacao || '',
      observacoes: c.Observacoes || '',
      avaria: c.Avaria_Desc || '',
      retiradoPor: ret ? (ret.Retirado_Por_Nome||'') : '',
      turma: ret ? (ret.Turma||'') : '',
      sala: ret ? (ret.Sala||'') : '',
      retornoPrevisto: ret ? (ret.Data_Retorno_Prevista||'') : '',
      localizacaoAtual: contextoAtual ? (contextoAtual.localAtual || c.Localizacao || '') : (c.Localizacao || ''),
      contextoAtual: contextoAtual
    };
  }).sort((a,b) => String(a.id).localeCompare(String(b.id)));

  return {
    sucesso: true,
    resumo: {
      total: chromes.length,
      disponiveis: chromes.filter(c => c.status === 'DISPONIVEL').length,
      emUso: chromes.filter(c => c.status === 'EM_USO').length,
      manutencao: chromes.filter(c => c.status === 'MANUTENCAO').length,
      alunos: chromes.filter(c => c.tipo === 'ALUNO').length,
      professores: chromes.filter(c => c.tipo === 'PROF').length
    },
    chromebooks: chromes
  };
}

/**
 * Inventário rápido para a tela principal de Chromebooks.
 *
 * IMPORTANTE: esta função NÃO executa o motor de localização, auditoria ou
 * linha do tempo. Esses cálculos ficam para as consultas específicas.
 * Assim a tela principal responde rapidamente e uma falha na localização
 * jamais impede o inventário de aparecer.
 */
function getInventarioChromebooksRapido(token) {
  const usuario = exigirSessao(token);
  if (!usuario) return respostaSessaoExpirada();
  const cache=CacheService.getScriptCache();
  const cacheKey='ICM_INVENTARIO_RAPIDO_V2';
  try { const cached=cache.get(cacheKey); if(cached) return JSON.parse(cached); } catch(e) {}

  try {
    const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
    const nomes = {};
    usuarios.forEach(u => nomes[String(u.ID)] = u.Nome || String(u.ID));

    const retiradas = sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS)
      .filter(r => String(r.Status || '').toUpperCase() === 'ABERTA');
    const retiradaMap = {};
    retiradas.forEach(r => {
      retiradaMap[String(r.Chromebook_ID)] = r;
    });

    const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS)
      .map(c => {
        const ret = retiradaMap[String(c.ID)] || null;
        const retorno = ret ? ret.Data_Retorno_Prevista : '';
        const retornoTexto = retorno instanceof Date
          ? Utilities.formatDate(retorno, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm')
          : String(retorno || '');

        return {
          id: String(c.ID || ''),
          tipo: String(c.Tipo || '').toUpperCase(),
          status: String(c.Status || '').toUpperCase(),
          responsavelId: c.Responsavel_ID || '',
          responsavel: c.Responsavel_ID
            ? (nomes[String(c.Responsavel_ID)] || String(c.Responsavel_ID))
            : '',
          localizacao: c.Localizacao || '',
          localizacaoAtual: c.Localizacao || '',
          observacoes: c.Observacoes || '',
          avaria: c.Avaria_Desc || '',
          retiradoPor: ret ? (ret.Retirado_Por_Nome || '') : '',
          retiradoPorId: ret ? (ret.Retirado_Por_ID || '') : '',
          turma: ret ? (ret.Turma || '') : '',
          sala: ret ? (ret.Sala || '') : '',
          retornoPrevisto: retornoTexto,
          contextoAtual: null
        };
      })
      .sort((a,b) => String(a.id).localeCompare(String(b.id)));

    const resposta=serializarParaCliente({
      sucesso: true,
      resumo: {
        total: chromes.length,
        disponiveis: chromes.filter(c => c.status === 'DISPONIVEL').length,
        emUso: chromes.filter(c => c.status === 'EM_USO').length,
        manutencao: chromes.filter(c => c.status === 'MANUTENCAO').length,
        alunos: chromes.filter(c => c.tipo === 'ALUNO').length,
        professores: chromes.filter(c => c.tipo === 'PROF').length
      },
      chromebooks: chromes
    });
    try { cache.put(cacheKey,JSON.stringify(resposta),5); } catch(e) {}
    return resposta;
  } catch (err) {
    logSistema('ERROR', 'INVENTARIO_RAPIDO_ERRO', usuario.ID, err.message);
    return { sucesso: false, mensagem: 'Não foi possível carregar o inventário: ' + err.message };
  }
}



function getDetalhesUsoChromebooks(token) {
  const usuario = exigirSessao(token);
  if (!usuario) return respostaSessaoExpirada();

  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const nomes = {};
  usuarios.forEach(u => nomes[String(u.ID)] = u.Nome || String(u.ID));

  // Consulta isolada para enriquecer o inventário depois da carga principal.
  // Não interfere na montagem inicial dos Chromebooks.
  const sh = getSheet(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  const retiradas = sh.getLastRow() > 1
    ? sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS)
        .filter(r => String(r.Status || '').toUpperCase() === 'ABERTA')
    : [];

  const detalhes = {};
  retiradas.forEach(r => {
    detalhes[String(r.Chromebook_ID)] = {
      retiradoPor: r.Retirado_Por_Nome || '',
      retiradoPorId: r.Retirado_Por_ID || '',
      turma: r.Turma || '',
      sala: r.Sala || '',
      retornoPrevisto: r.Data_Retorno_Prevista || '',
      responsavel: r.Responsavel_Nome || nomes[String(r.Responsavel_ID)] || ''
    };
  });

  return serializarParaCliente({ sucesso: true, detalhes: detalhes });
}

// ============================================================================
// DISPONIBILIDADE CENTRAL DE CHROMEBOOKS DE ALUNOS
// Regra: somente Tipo=ALUNO entra no estoque operacional.
// Reservas de laboratório e agendamentos compartilham a mesma capacidade.
// ============================================================================

function normalizarDataSistema(valor) {
  if (!valor) return '';
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const texto = String(valor).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
  const m = texto.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
  const d = new Date(texto);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function horarioTextoSistema(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, DEFAULT_CONFIG.TIMEZONE, 'HH:mm');
  }
  if (typeof valor === 'number' && isFinite(valor)) {
    const minutos = Math.round(valor * 24 * 60);
    const h = Math.floor(minutos / 60) % 24;
    const m = minutos % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  const texto = String(valor).trim();
  let m = texto.match(/T(\d{1,2}):(\d{2})/);
  if (m) return String(Number(m[1])).padStart(2, '0') + ':' + m[2];
  m = texto.match(/(\d{1,2}):(\d{2})/);
  if (m) return String(Number(m[1])).padStart(2, '0') + ':' + m[2];
  const d = new Date(texto);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, DEFAULT_CONFIG.TIMEZONE, 'HH:mm');
  return texto;
}

function minutosDoHorarioSistema(horario) {
  const texto = horarioTextoSistema(horario);
  if (!texto) return null;
  const m = texto.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}

function valorBooleanoSistema(valor, padrao) {
  if (valor === null || valor === undefined || valor === '') return padrao;
  const t = String(valor).trim().toUpperCase();
  if (['TRUE','VERDADEIRO','SIM','S','1','YES'].includes(t)) return true;
  if (['FALSE','FALSO','NAO','NÃO','N','0','NO'].includes(t)) return false;
  return padrao;
}

function horarioLaboratorioAtivo(h) {
  const status = String(h && h.Status !== undefined ? h.Status : '').trim().toUpperCase();
  if (!status) return true;
  return ['ATIVO','ATIVA','SIM','TRUE','VERDADEIRO','1'].includes(status);
}

function obterHorariosLaboratorioConfigurados() {
  const cache=CacheService.getScriptCache();
  const cacheKey='ICM_HORARIOS_LAB_V2';
  try{
    const cached=cache.get(cacheKey);
    if(cached)return JSON.parse(cached);
  }catch(e){}

  const horarios=sheetToObjects(SHEET_NAMES.HORARIOS_LABORATORIO)
    .filter(h => horarioLaboratorioAtivo(h))
    .map((h, idx) => ({
      ...h,
      _ordemSistema: Number(h.Ordem || idx + 1),
      _inicioSistema: horarioTextoSistema(h.Inicio),
      _fimSistema: horarioTextoSistema(h.Fim)
    }))
    .filter(h => h._inicioSistema && h._fimSistema)
    .sort((a,b) => a._ordemSistema - b._ordemSistema);
  try{cache.put(cacheKey,JSON.stringify(horarios),30);}catch(e){}
  return horarios;
}

function invalidarCacheHorariosLaboratorio(){
  try{CacheService.getScriptCache().remove('ICM_HORARIOS_LAB_V2');}catch(e){}
}

function horarioBloqueadoSistema(h, todosHorarios) {
  if (!h) return true;
  const nota = String(h.Label || h.Observacao || h.Observações || h.Tipo || '').trim().toUpperCase();
  if (/ALMO[CÇ]O|INTERVALO/.test(nota)) return true;
  if (h.Disponivel !== undefined && h.Disponivel !== '') return !valorBooleanoSistema(h.Disponivel, true);
  // Compatibilidade com a planilha atual: TRUE = liberado, FALSE = bloqueado.
  if (h.Bloqueado !== undefined && h.Bloqueado !== '') return !valorBooleanoSistema(h.Bloqueado, false);
  return false;
}

function criarDataHoraSistema(data, horario) {
  const intervalo = intervaloHorarioLaboratorio(data, horario);
  return intervalo ? criarDataHoraPorMinutosSistema(intervalo.data, intervalo.inicioMin) : null;
}

function criarDataHoraPorMinutosSistema(data, minutos) {
  const dataNormalizada = normalizarDataSistema(data);
  if (!dataNormalizada || minutos === null || minutos === undefined) return null;
  const h = Math.floor(Number(minutos) / 60);
  const m = Number(minutos) % 60;
  return Utilities.parseDate(
    dataNormalizada + ' ' + String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0'),
    DEFAULT_CONFIG.TIMEZONE,
    'yyyy-MM-dd HH:mm'
  );
}

function intervaloHorarioLaboratorio(data, horario) {
  const dataNormalizada = normalizarDataSistema(data);
  if (!dataNormalizada) return null;
  const texto = horarioTextoSistema(horario);
  if (!texto) return null;
  const horarios = obterHorariosLaboratorioConfigurados();
  if (!horarios.length) return null;
  const pares = String(horario || '').match(/(\d{1,2}:\d{2})\s*(?:—|-|–|a|até)\s*(\d{1,2}:\d{2})/i);
  let inicio = null, fim = null, ordem = null;
  if (pares) {
    inicio = horarioTextoSistema(pares[1]);
    fim = horarioTextoSistema(pares[2]);
    const idxInicio = horarios.findIndex(x => x._inicioSistema === inicio);
    const idxFim = horarios.findIndex(x => x._fimSistema === fim);
    if (idxInicio >= 0 && idxFim >= idxInicio) {
      let consecutivo = true;
      for (let i = idxInicio + 1; i <= idxFim; i++) {
        if (Number(horarios[i]._ordemSistema) !== Number(horarios[i-1]._ordemSistema) + 1) { consecutivo = false; break; }
      }
      if (consecutivo) ordem = horarios[idxInicio]._ordemSistema; else inicio = null;
    } else inicio = null;
  }
  if (!inicio) {
    const h = horarios.find(x => x._inicioSistema === texto);
    if (h) { inicio = h._inicioSistema; fim = h._fimSistema; ordem = h._ordemSistema; }
  }
  if (!inicio || !fim) return null;
  const inicioMin = minutosDoHorarioSistema(inicio), fimMin = minutosDoHorarioSistema(fim);
  if (inicioMin === null || fimMin === null || fimMin <= inicioMin) return null;
  return { data:dataNormalizada, inicioMin, fimMin, inicio, fim, label:inicio+' — '+fim, ordem };
}

function intervaloHorarioLaboratorioComDuracao(data, horario, duracaoPeriodos) {
  const horarios = obterHorariosLaboratorioConfigurados();
  const base = intervaloHorarioLaboratorio(data, horario);
  const duracao = Math.max(1, Number(duracaoPeriodos) || 1);
  if (!base || duracao === 1) return base;
  const idx = horarios.findIndex(h => h._inicioSistema === base.inicio);
  if (idx < 0 || idx + duracao > horarios.length) return null;
  for (let i = idx + 1; i < idx + duracao; i++) {
    if (Number(horarios[i]._ordemSistema) !== Number(horarios[i-1]._ordemSistema) + 1) return null;
  }
  const fim = horarios[idx + duracao - 1];
  return { data:normalizarDataSistema(data), inicioMin:minutosDoHorarioSistema(horarios[idx]._inicioSistema), fimMin:minutosDoHorarioSistema(fim._fimSistema), inicio:horarios[idx]._inicioSistema, fim:fim._fimSistema, label:horarios[idx]._inicioSistema+' — '+fim._fimSistema, ordem:horarios[idx]._ordemSistema, duracaoPeriodos:duracao };
}

function horarioPassadoSistema(data, inicioMin) {
  const dataNormalizada = normalizarDataSistema(data);
  if (!dataNormalizada || inicioMin === null || inicioMin === undefined) return false;

  const hoje = Utilities.formatDate(new Date(), DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (dataNormalizada < hoje) return true;
  if (dataNormalizada > hoje) return false;

  const horaAtual = Utilities.formatDate(new Date(), DEFAULT_CONFIG.TIMEZONE, 'HH:mm');
  const partes = horaAtual.split(':').map(Number);
  const agoraMin = (partes[0] * 60) + partes[1];

  // O horário fica indisponível assim que seu início chega.
  // Ex.: às 17:14, 17:10–18:00 já não pode mais ser reservado.
  return Number(inicioMin) <= agoraMin;
}

function intervalosConflitamSistema(aInicio, aFim, bInicio, bFim) {
  return aInicio < bFim && bInicio < aFim;
}

function quantidadeNumericaSistema(valor) {
  const n = Number(valor);
  return isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function listarCompromissosChromebooksSistema(data, horario, opcoes) {
  opcoes = opcoes || {};
  const intervalo = intervaloHorarioLaboratorio(data, horario);
  if (!intervalo) return { intervalo: null, compromissos: [], quantidadeComprometida: 0, idsComprometidos: [] };
  const dataAlvo = intervalo.data;
  const compromissos = [];
  const idsComprometidos = new Set();
  const reservas = opcoes.reservasLab || sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO);
  reservas.forEach(r => {
    if (String(r.Status).toUpperCase() !== 'ATIVA') return;
    if (opcoes.excluirReservaId && String(r.ID) === String(opcoes.excluirReservaId)) return;
    if (normalizarDataSistema(r.Data) !== dataAlvo) return;
    const h = intervaloHorarioLaboratorio(dataAlvo, r.Horario);
    if (!h || !intervalosConflitamSistema(intervalo.inicioMin, intervalo.fimMin, h.inicioMin, h.fimMin)) return;
    const qtd = quantidadeNumericaSistema(r.Chromes);
    const ids = idsComprometidosDoRegistroSistema(r);
    ids.forEach(id=>idsComprometidos.add(String(id)));
    if(qtd>0) compromissos.push({ origem:'RESERVA_LABORATORIO', id:r.ID, quantidade:qtd, ids, inicio:h.inicio, fim:h.fim, turma:r.Turma||'' });
  });
  const agendamentos = opcoes.agendamentos || sheetToObjects(SHEET_NAMES.AGENDAMENTOS);
  agendamentos.forEach(a => {
    if (String(a.Status).toUpperCase() !== 'AGENDADO') return;
    if (opcoes.excluirAgendamentoId && String(a.ID) === String(opcoes.excluirAgendamentoId)) return;
    if (normalizarDataSistema(a.Data) !== dataAlvo) return;
    const h = intervaloHorarioLaboratorio(dataAlvo, a.Horario);
    if (!h || !intervalosConflitamSistema(intervalo.inicioMin, intervalo.fimMin, h.inicioMin, h.fimMin)) return;
    const qtd = quantidadeNumericaSistema(a.Quantidade);
    const ids = idsChromebooksSistema(a.Chromes_IDS);
    ids.forEach(id=>idsComprometidos.add(String(id)));
    if(qtd>0) compromissos.push({ origem:'AGENDAMENTO', id:a.ID, quantidade:qtd, ids, inicio:h.inicio, fim:h.fim, turma:a.Turma||'' });
  });
  return { intervalo, compromissos, quantidadeComprometida:compromissos.reduce((s,c)=>s+c.quantidade,0), idsComprometidos:[...idsComprometidos] };
}


// ============================================================================
// CONTROLE TEMPORAL DAS RETIRADAS DE CHROMEBOOKS
// Cada equipamento retirado passa a ter uma janela de uso própria.
// Isso permite que a disponibilidade varie durante o dia.
// ============================================================================

const SALAS_TURMAS = {
  '72':'201','32':'201',
  '91':'203','11':'203',
  '71':'204','31':'204',
  '62':'215','41':'215',
  '61':'216','12':'216',
  '81':'301','52':'301',
  '82':'305','51':'305',
  '22':'303','21':'304','42':'306'
};

function obterSalaDaTurma(turma) {
  const chave=String(turma||'').trim().toUpperCase()
    .replace(/^TURMA\s*/,'').replace(/^ANO\s*/,'')
    .replace(/[º°]/g,'');

  // Fonte oficial: HORARIOS_TURMAS. A sala deve acompanhar a grade atual.
  try {
    const linhas=sheetToObjects(SHEET_NAMES.HORARIOS_TURMAS);
    const alvo=normalizarNomeSistema(chave);
    const linha=linhas.find(r=>{
      const t=String(r.Turma||'').trim().replace(/\.0$/,'');
      const tid=String(r.Turma_ID||'').trim();
      return normalizarNomeSistema(t)===alvo || normalizarNomeSistema(tid)===alvo;
    });
    if(linha && String(linha.Sala||'').trim()) return String(linha.Sala).trim();
  }catch(e){}

  // Compatibilidade com dados antigos.
  return SALAS_TURMAS[chave] ? 'Sala '+SALAS_TURMAS[chave] : '';
}


// ============================================================================
// VÍNCULO TEMPORAL DOS CHROMEBOOKS FIXOS
// Um Chromebook físico pode ter responsáveis pedagógicos diferentes por turno.
// ============================================================================
function normalizarNomeSistema(valor){
  return String(valor||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
}
function seedVinculosChromebooks(){
  const sh=getSheet(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
  if(sh.getLastRow()>1)return;
  const mapa=[
    ['PROF001','Manhã','Felipe Vieira'],['PROF001','Tarde','Felipe Vieira'],['PROF002','Manhã','Murilo Santos'],['PROF002','Tarde','Murilo Santos'],
    ['PROF003','Manhã','João Abreu'],['PROF003','Tarde','Gabriel Maio'],['PROF004','Manhã','Angelina Ladeira'],['PROF004','Tarde','Bárbara Paim'],
    ['PROF005','Manhã','Cezar Alexandre'],['PROF005','Tarde','Caroline Correa'],['PROF006','Manhã','Tatiane Macedo'],['PROF006','Tarde','Katia Dorneles'],
    ['PROF007','Manhã','Juliane Prates'],['PROF007','Tarde','Gisele Avila'],['PROF008','Manhã','Patrícia Soares'],['PROF008','Tarde','Josiane Cardoso'],
    ['PROF009','Manhã','Rejane Carvalho'],['PROF009','Tarde','Juliane Lima'],['PROF010','Manhã','Tanise Flores'],['PROF010','Tarde','Karina Santos'],
    ['PROF011','Manhã','Tatiane Alvarenga'],['PROF011','Tarde','Maria Carolina'],['PROF012','Manhã','Valdire Caetano'],['PROF012','Tarde','Maria Lucia'],
    ['PROF013','Tarde','Cátia Rosa'],['PROF014','Tarde','Marcela Gonzalez'],['PROF015','Tarde','Renata Pardo'],['PROF016','Tarde','Marina Salun'],['PROF017','Tarde','Stela Carrasco']
  ];
  mapa.forEach(r=>appendObject(SHEET_NAMES.VINCULOS_CHROMEBOOKS,{Chromebook_ID:r[0],Turno:r[1],Professor_Nome:r[2],Professor_ID:'',Ativo:'SIM',Observacoes:''}));
}

/**
 * ============================================================================
 * ATUALIZAÇÃO DO PROFESSOR DE EDUCAÇÃO FÍSICA
 * Juliana Boettge -> Felipe Vieira
 *
 * Executar UMA vez após publicar esta versão.
 * Atualiza somente dados operacionais/atuais. Não altera auditorias históricas.
 * ============================================================================
 */
function atualizarProfessorEducacaoFisicaFelipe() {
  const NOME_ANTIGO = 'Juliana Boettge';
  const NOME_NOVO = 'Felipe Vieira';
  const antigoNorm = normalizarNomeSistema(NOME_ANTIGO);
  const novoNorm = normalizarNomeSistema(NOME_NOVO);
  const resultado = {
    sucesso: true,
    usuarioAtualizado: false,
    horariosAtualizados: 0,
    vinculosAtualizados: 0,
    retiradasAbertasAtualizadas: 0,
    reservasAtivasAtualizadas: 0,
    agendamentosAtivosAtualizados: 0,
    professorId: ''
  };

  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  let felipe = usuarios.find(u => normalizarNomeSistema(u.Nome) === novoNorm);
  let antigo = usuarios.find(u => normalizarNomeSistema(u.Nome) === antigoNorm);

  // Se Felipe ainda não foi cadastrado, reaproveita o cadastro antigo,
  // preservando ID, código e demais configurações já existentes.
  if (!felipe && antigo) {
    const row = findRowById(SHEET_NAMES.USUARIOS, 'ID', antigo.ID);
    if (row) {
      updateRowByFields(SHEET_NAMES.USUARIOS, row.rowIndex, {Nome: NOME_NOVO});
      felipe = Object.assign({}, antigo, {Nome: NOME_NOVO});
      resultado.usuarioAtualizado = true;
    }
  }

  if (felipe) resultado.professorId = String(felipe.ID || '');

  // Grade oficial: Educação Física ou nome antigo.
  const horarios = sheetToObjects(SHEET_NAMES.HORARIOS_TURMAS);
  horarios.forEach(r => {
    const pn = normalizarNomeSistema(r.Professor_Nome || r.Professor || '');
    const disc = normalizarNomeSistema(r.Disciplina || '');
    const ehAntigo = pn === antigoNorm;
    const ehEducFisica = disc.indexOf('educ') >= 0 && disc.indexOf('fisic') >= 0;
    if (!ehAntigo && !ehEducFisica) return;
    const row = findRowById(SHEET_NAMES.HORARIOS_TURMAS, 'Horario_ID', r.Horario_ID);
    if (!row) return;
    const campos = {Professor_Nome: NOME_NOVO};
    if (resultado.professorId) campos.Professor_ID = resultado.professorId;
    updateRowByFields(SHEET_NAMES.HORARIOS_TURMAS, row.rowIndex, campos);
    resultado.horariosAtualizados++;
  });

  // Vínculos atuais de Chromebooks.
  const vinculos = sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
  vinculos.forEach(r => {
    if (normalizarNomeSistema(r.Professor_Nome || '') !== antigoNorm) return;
    const row = findRowById(SHEET_NAMES.VINCULOS_CHROMEBOOKS, 'Chromebook_ID', r.Chromebook_ID);
    if (!row) return;
    const campos = {Professor_Nome: NOME_NOVO};
    if (resultado.professorId) campos.Professor_ID = resultado.professorId;
    updateRowByFields(SHEET_NAMES.VINCULOS_CHROMEBOOKS, row.rowIndex, campos);
    resultado.vinculosAtualizados++;
  });

  // Somente registros ainda abertos/agendados/reservados.
  const retiradas = sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  retiradas.forEach(r => {
    if (String(r.Status || '').toUpperCase() !== 'ABERTA') return;
    if (normalizarNomeSistema(r.Responsavel_Nome || '') !== antigoNorm) return;
    const row = findRowById(SHEET_NAMES.RETIRADAS_CHROMEBOOKS, 'ID', r.ID);
    if (!row) return;
    const campos = {Responsavel_Nome: NOME_NOVO};
    if (resultado.professorId) campos.Responsavel_ID = resultado.professorId;
    updateRowByFields(SHEET_NAMES.RETIRADAS_CHROMEBOOKS, row.rowIndex, campos);
    resultado.retiradasAbertasAtualizadas++;
  });

  const reservas = sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO);
  reservas.forEach(r => {
    if (String(r.Status || '').toUpperCase() !== 'ATIVA') return;
    if (normalizarNomeSistema(r.Professor || '') !== antigoNorm) return;
    const row = findRowById(SHEET_NAMES.RESERVAS_LABORATORIO, 'ID', r.ID);
    if (!row) return;
    updateRowByFields(SHEET_NAMES.RESERVAS_LABORATORIO, row.rowIndex, {Professor:NOME_NOVO});
    resultado.reservasAtivasAtualizadas++;
  });

  const agendamentos = sheetToObjects(SHEET_NAMES.AGENDAMENTOS);
  agendamentos.forEach(r => {
    if (String(r.Status || '').toUpperCase() !== 'AGENDADO') return;
    const nome = r.Professor_Nome || r.Professor || '';
    if (normalizarNomeSistema(nome) !== antigoNorm) return;
    const row = findRowById(SHEET_NAMES.AGENDAMENTOS, 'ID', r.ID);
    if (!row) return;
    const campos = {};
    if (r.Professor_Nome !== undefined) campos.Professor_Nome = NOME_NOVO;
    if (r.Professor !== undefined) campos.Professor = NOME_NOVO;
    updateRowByFields(SHEET_NAMES.AGENDAMENTOS, row.rowIndex, campos);
    resultado.agendamentosAtivosAtualizados++;
  });

  limparMemoPlanilha(SHEET_NAMES.USUARIOS);
  limparMemoPlanilha(SHEET_NAMES.HORARIOS_TURMAS);
  limparMemoPlanilha(SHEET_NAMES.VINCULOS_CHROMEBOOKS);

  resultado.mensagem =
    'Professor atualizado para Felipe Vieira. ' +
    resultado.horariosAtualizados + ' horários, ' +
    resultado.vinculosAtualizados + ' vínculos e registros operacionais ativos ajustados.';
  return resultado;
}

function sincronizarProfessorIdsVinculos(){
  try{
    const sh=getSheet(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
    const vinculos=sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
    const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS);
    const updates=[];
    vinculos.forEach(v=>{
      if(String(v.Professor_ID||'').trim()) return;
      const alvo=normalizarNomeSistema(v.Professor_Nome);
      if(!alvo) return;
      let u=usuarios.find(x=>normalizarNomeSistema(x.Nome)===alvo);
      if(!u){
        const primeiro=alvo.split(/\s+/)[0];
        if(primeiro && primeiro.length>=4){
          const candidatos=usuarios.filter(x=>
            ['professor','professora'].includes(String(x.Tipo||'').trim().toLowerCase()) &&
            normalizarNomeSistema(x.Nome).split(/\s+/)[0]===primeiro
          );
          if(candidatos.length===1) u=candidatos[0];
        }
      }
      if(u && v._row) updates.push({rowIndex:Number(v._row),fields:{Professor_ID:u.ID,Professor_Nome:u.Nome}});
    });
    if(updates.length) atualizarLinhasBatch(SHEET_NAMES.VINCULOS_CHROMEBOOKS,updates);
    return updates.length;
  }catch(err){
    logSistema('WARNING','VINCULOS_PROFESSOR_ID_SYNC_ERRO','',err.message);
    return 0;
  }
}

function obterVinculosChromebook(chromebookId){
  sincronizarProfessorIdsVinculos();
  const vinculos=sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS)
    .filter(v=>String(v.Chromebook_ID)===String(chromebookId)&&valorBooleanoSistema(v.Ativo,true));
  const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS);
  return vinculos.map(v=>{
    let u=v.Professor_ID?usuarios.find(x=>String(x.ID)===String(v.Professor_ID)):null;
    if(!u){
      const alvo=normalizarNomeSistema(v.Professor_Nome);
      u=usuarios.find(x=>normalizarNomeSistema(x.Nome)===alvo);
      if(!u){
        const primeiro=alvo.split(/\s+/)[0];
        const candidatos=usuarios.filter(x=>
          ['professor','professora'].includes(String(x.Tipo||'').trim().toLowerCase()) &&
          normalizarNomeSistema(x.Nome).split(/\s+/)[0]===primeiro
        );
        if(candidatos.length===1) u=candidatos[0];
      }
    }
    return {chromebookId:v.Chromebook_ID,turno:v.Turno||'',professorId:u?u.ID:(v.Professor_ID||''),professorNome:u?u.Nome:(v.Professor_Nome||''),ativo:true};
  });
}
function horarioGradeContemMomentoSistema(valor, alvoMin){
  if(alvoMin===null||alvoMin===undefined)return false;
  const p=String(valor||'').match(/(\d{1,2}:\d{2})\s*(?:—|-|–|a|até)\s*(\d{1,2}:\d{2})/i);
  if(p){const a=minutosDoHorarioSistema(p[1]),b=minutosDoHorarioSistema(p[2]);return a!==null&&b!==null&&alvoMin>=a&&alvoMin<b;}
  const inicio=minutosDoHorarioSistema(valor);
  return inicio!==null&&inicio===alvoMin;
}
function obterContextoGradeProfessorSistema(data,inicioMin,professorNome,professorId){
  const dia=diaSemanaNomeSistema(data);
  const alvo=normalizarNomeSistema(professorNome);
  const alvoId=String(professorId||'').trim();
  const linhas=sheetToObjects(SHEET_NAMES.HORARIOS_TURMAS).filter(r=>{
    const d=normalizarNomeSistema(r.Dia_Semana!==undefined?r.Dia_Semana:r.Dia);
    const diaNorm=normalizarNomeSistema(dia);
    const diaOk=d===diaNorm||d.startsWith(diaNorm.slice(0,3));
    let horaOk=false;
    if(r.Hora_Inicio!==undefined || r.Hora_Fim!==undefined){
      const a=minutosDoHorarioSistema(r.Hora_Inicio);
      const b=minutosDoHorarioSistema(r.Hora_Fim);
      horaOk=a!==null&&b!==null&&inicioMin!==null&&inicioMin>=a&&inicioMin<b;
    }else{
      horaOk=horarioGradeContemMomentoSistema(r.Horario,inicioMin);
    }
    const rid=String(r.Professor_ID||'').trim();
    let profOk=!!alvoId && rid===alvoId;
    if(!profOk){
      const pn=normalizarNomeSistema(r.Professor_Nome!==undefined?r.Professor_Nome:r.Professor);
      const primeiro=alvo.split(/\s+/)[0];
      profOk=pn===alvo || (primeiro&&primeiro.length>=5&&pn.split(/\s+/)[0]===primeiro) ||
        (pn.length>=5&&alvo.startsWith(pn.slice(0,5)));
    }
    const ativo=r.Ativo===undefined || valorBooleanoSistema(r.Ativo,true);
    return diaOk&&horaOk&&profOk&&ativo;
  });
  const r=linhas[0];
  if(!r)return null;
  return {
    horarioId:String(r.Horario_ID||''),
    turmaId:String(r.Turma_ID||''),
    turma:String(r.Turma||'').replace(/\.0$/,''),
    disciplina:String(r.Disciplina||''),
    sala:String(r.Sala||'').trim() || obterSalaDaTurma(r.Turma),
    professorId:String(r.Professor_ID||professorId||''),
    professorNome:String(r.Professor_Nome||r.Professor||professorNome||''),
    periodo:String(r.Periodo||''),
    horaInicio:horarioTextoSistema(r.Hora_Inicio||r.Horario||''),
    horaFim:horarioTextoSistema(r.Hora_Fim||'')
  };
}
function obterGradeProfessorDoDia(data, professorId, professorNome){
  const dia=normalizarNomeSistema(diaSemanaNomeSistema(data));
  const linhas=sheetToObjects(SHEET_NAMES.HORARIOS_TURMAS).filter(r=>{
    const rd=normalizarNomeSistema(r.Dia_Semana!==undefined?r.Dia_Semana:r.Dia);
    const diaOk=rd===dia || rd.startsWith(dia.slice(0,3));
    const rid=String(r.Professor_ID||'').trim();
    let profOk=professorId && rid===String(professorId);
    if(!profOk){
      const pn=normalizarNomeSistema(r.Professor_Nome!==undefined?r.Professor_Nome:r.Professor);
      const alvo=normalizarNomeSistema(professorNome);
      const primeiro=alvo.split(/\s+/)[0];
      profOk=pn===alvo || (primeiro&&primeiro.length>=5&&pn.split(/\s+/)[0]===primeiro);
    }
    return diaOk&&profOk&&(r.Ativo===undefined||valorBooleanoSistema(r.Ativo,true));
  }).map(r=>{
    let inicio=horarioTextoSistema(r.Hora_Inicio);
    let fim=horarioTextoSistema(r.Hora_Fim);
    if(!inicio){
      const p=String(r.Horario||'').match(/(\d{1,2}:\d{2})\s*(?:—|-|–|a|até)\s*(\d{1,2}:\d{2})/i);
      if(p){inicio=horarioTextoSistema(p[1]);fim=horarioTextoSistema(p[2]);}
      else inicio=horarioTextoSistema(r.Horario);
    }
    const ini=minutosDoHorarioSistema(inicio), fi=minutosDoHorarioSistema(fim);
    return {...r,_inicioSistema:inicio,_fimSistema:fim,_inicioMin:ini,_fimMin:fi};
  }).filter(r=>r._inicioMin!==null).sort((a,b)=>a._inicioMin-b._inicioMin);
  return linhas;
}


// ============================================================================
// AUDITORIA E MOTOR DE LOCALIZAÇÃO DOS CHROMEBOOKS DE PROFESSORES
// A auditoria fica em uma planilha separada para não pesar a planilha operacional.
// O estado atual permanece em CHROMEBOOKS.Localizacao.
// ============================================================================
function getSSAuditoria(){
  const id=String(getConfigValor('AUDITORIA_SPREADSHEET_ID')||'').trim();
  if(!id) throw new Error('AUDITORIA_SPREADSHEET_ID não configurado.');
  return SpreadsheetApp.openById(id);
}
function getSheetAuditoria(nome){
  const ss=getSSAuditoria();
  let sh=ss.getSheetByName(nome);
  if(!sh){
    sh=ss.insertSheet(nome);
    const headers=SHEET_HEADERS[nome]||[];
    if(headers.length)sh.getRange(1,1,1,headers.length).setValues([headers]);
  }else if(sh.getLastRow()===0 && SHEET_HEADERS[nome]){
    sh.getRange(1,1,1,SHEET_HEADERS[nome].length).setValues([SHEET_HEADERS[nome]]);
  }
  return sh;
}
function appendAuditoria(nome,obj){
  const sh=getSheetAuditoria(nome);
  const headers=sh.getRange(1,1,1,sh.getLastColumn()||SHEET_HEADERS[nome].length).getValues()[0];
  sh.appendRow(headers.map(h=>obj[h]!==undefined?obj[h]:''));
}
function registrarAuditoriaChromebook(dados){
  try{
    appendAuditoria(SHEET_NAMES.AUDITORIA_CHROMEBOOKS,{
      AUDITORIA_ID:genId('AUDCHR'), DATA_HORA:nowStr(), ACAO:dados.acao||'EVENTO',
      CHROMEBOOK_ID:dados.chromebookId||'', PROFESSOR_ID:dados.professorId||'', PROFESSOR_NOME:dados.professorNome||'',
      TIPO_USO:dados.tipoUso||'PROFESSOR', CONTEXTO:dados.contexto||'', TURNO:dados.turno||'',
      DIA_SEMANA:dados.diaSemana||'', HORARIO_AULA:dados.horarioAula||'', TURMA:dados.turma||'',
      DISCIPLINA:dados.disciplina||'', SALA:dados.sala||'', LOCAL_ATUAL:dados.localAtual||'',
      ORIGEM:dados.origem||'SISTEMA', STATUS:dados.status||'SUCESSO', OBSERVACAO:dados.observacao||''
    });
    return true;
  }catch(err){
    logSistema('WARNING','AUDITORIA_CHROME_ERRO',dados&&dados.professorId,err.message);
    return false;
  }
}
function registrarAuditoriaLaboratorio(dados){
  try{
    appendAuditoria(SHEET_NAMES.AUDITORIA_LABORATORIO,{
      AUDITORIA_ID:genId('AUDLAB'), DATA_HORA:nowStr(), ACAO:dados.acao||'EVENTO',
      RESERVA_ID:dados.reservaId||'', PROFESSOR_ID:dados.professorId||'', PROFESSOR_NOME:dados.professorNome||'',
      DATA_RESERVA:dados.dataReserva||'', HORARIO_INICIO:dados.horarioInicio||'', HORARIO_FIM:dados.horarioFim||'',
      TURMA:dados.turma||'', LABORATORIO:dados.laboratorio||'Laboratório de Informática',
      CHROMEBOOK:dados.chromebook||dados.chromes||0, FONES:dados.fones||0, PROJETOR:dados.projetor||0,
      ATIVIDADE:dados.atividade||'', ORIGEM:dados.origem||'SISTEMA', STATUS:dados.status||'SUCESSO',
      OBSERVACAO:dados.observacao||''
    });
    return true;
  }catch(err){
    logSistema('WARNING','AUDITORIA_LAB_ERRO',dados&&dados.professorId,err.message);
    return false;
  }
}

function obterLocalizacaoAtualChromebook(chromebookId,data){
  const chrome=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).find(c=>String(c.ID)===String(chromebookId));
  if(!chrome)return null;
  const abertas=sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS).filter(r=>
    String(r.Chromebook_ID)===String(chromebookId)&&String(r.Status||'').toUpperCase()==='ABERTA'
  );
  if(!abertas.length)return {localAtual:String(chrome.Localizacao||''),status:'DISPONIVEL',emUso:false};
  const ret=abertas[abertas.length-1];
  const professorId=String(ret.Responsavel_ID||chrome.Responsavel_ID||'');
  const professorNome=String(ret.Responsavel_Nome||'');
  const agora=new Date();
  const dataAlvo=normalizarDataSistema(data||agora);
  const horaAtual=Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'HH:mm');
  const agoraMin=minutosDoHorarioSistema(horaAtual);
  const grade=obterContextoGradeProfessorSistema(dataAlvo,agoraMin,professorNome,professorId);
  if(grade){
    return {localAtual:grade.sala||'',sala:grade.sala||'',turma:grade.turma||'',disciplina:grade.disciplina||'',
      horarioAula:(grade.horaInicio||'')+(grade.horaFim?' — '+grade.horaFim:''),professorId,professorNome,
      status:'EM_AULA',emUso:true,contexto:'AULA',retirada:ret.Data_Retirada,retornoPrevisto:ret.Data_Retorno_Prevista};
  }
  return {localAtual:String(chrome.Localizacao||ret.Sala||''),sala:String(chrome.Localizacao||ret.Sala||''),
    professorId,professorNome,status:'EM_USO_SEM_AULA',emUso:true,contexto:'FORA_DA_GRADE',
    retirada:ret.Data_Retirada,retornoPrevisto:ret.Data_Retorno_Prevista};
}

function atualizarLocalizacaoChromebookProfessor(chromebookId, motivo){
  try{
    const chrome=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).find(c=>String(c.ID)===String(chromebookId));
    if(!chrome || String(chrome.Tipo||'').toUpperCase()!=='PROF' || String(chrome.Status||'').toUpperCase()!=='EM_USO') return null;
    const loc=obterLocalizacaoAtualChromebook(chromebookId,new Date());
    if(!loc || !loc.emUso) return loc;
    const assinatura=[loc.contexto||'',loc.turma||'',loc.disciplina||'',loc.sala||'',loc.horarioAula||''].join('|');
    const cache=CacheService.getScriptCache();
    const cacheKey='ICM_LOC_'+String(chromebookId);
    const anterior=cache.get(cacheKey)||'';
    const mudou=assinatura!==anterior;
    const novaLocalizacao=String(loc.localAtual||'').trim();
    const localAnterior=String(chrome.Localizacao||'').trim();
    if(novaLocalizacao && novaLocalizacao!==localAnterior){
      updateCell(SHEET_NAMES.CHROMEBOOKS,Number(chrome._row),'Localizacao',novaLocalizacao);
    }
    if(mudou){
      const vinculos=obterVinculosChromebook(chromebookId);
      const vinculo=vinculos.find(v=>String(v.professorId)===String(loc.professorId))||vinculos[0]||{};
      registrarAuditoriaChromebook({
        acao:'ATUALIZACAO_LOCALIZACAO',chromebookId,professorId:loc.professorId,professorNome:loc.professorNome,
        tipoUso:'PROFESSOR',contexto:loc.contexto,turno:vinculo.turno||'',diaSemana:diaSemanaNomeSistema(new Date()),
        horarioAula:loc.horarioAula||'',turma:loc.turma||'',disciplina:loc.disciplina||'',sala:loc.sala||'',
        localAtual:novaLocalizacao,origem:motivo||'MOTOR_HORARIOS',status:'SUCESSO',
        observacao:(localAnterior&&localAnterior!==novaLocalizacao?'Local anterior: '+localAnterior+'. ':'')+
          (loc.contexto==='FORA_DA_GRADE'?'Professor fora de uma aula cadastrada na grade.':'')
      });
      try{cache.put(cacheKey,assinatura,21600);}catch(e){}
    }
    return {...loc,atualizado:mudou,localAnterior};
  }catch(err){
    logSistema('WARNING','ATUALIZAR_LOCALIZACAO_CHROME_ERRO','',err.message);
    return null;
  }
}

function atualizarLocalizacoesChromebooksProfessores(){
  try{
    const chromes=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).filter(c=>String(c.Tipo||'').toUpperCase()==='PROF'&&String(c.Status||'').toUpperCase()==='EM_USO');
    const resultados=chromes.map(c=>atualizarLocalizacaoChromebookProfessor(c.ID,'TRIGGER_5_MINUTOS')).filter(Boolean);
    return {sucesso:true,atualizados:resultados.filter(r=>r.atualizado).length,verificados:chromes.length};
  }catch(err){
    logSistema('ERROR','MOTOR_LOCALIZACAO_ERRO','',err.message);
    return {sucesso:false,mensagem:err.message};
  }
}

function configurarTriggerLocalizacaoChromebooks(){
  const nome='atualizarLocalizacoesChromebooksProfessores';
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()===nome)ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(nome).timeBased().everyMinutes(5).create();
  return {sucesso:true,mensagem:'Atualização automática configurada a cada 5 minutos.'};
}

function getLinhaTempoChromebook(token,chromebookId,data){
  const sessao=exigirSessao(token); if(!sessao)return respostaSessaoExpirada();
  atualizarLocalizacaoChromebookProfessor(chromebookId,'CONSULTA_LINHA_TEMPO');
  const chromeRow=findRowById(SHEET_NAMES.CHROMEBOOKS,'ID',chromebookId);
  if(!chromeRow)return{sucesso:false,mensagem:'Chromebook não encontrado.'};
  const d=normalizarDataSistema(data||new Date()); if(!d)return{sucesso:false,mensagem:'Data inválida.'};
  const vinculos=obterVinculosChromebook(chromebookId);
  const retiradas=sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS).filter(r=>String(r.Chromebook_ID)===String(chromebookId));
  const timeline=[];
  vinculos.forEach(v=>{
    obterGradeProfessorDoDia(d,v.professorId,v.professorNome).forEach(g=>{
      const atual=retiradas.find(r=>String(r.Status||'').toUpperCase()==='ABERTA' && retiradaTemporalConflitaIntervalo(r,d,g._inicioMin,g._fimMin));
      timeline.push({
        inicio:g._inicioSistema,
        fim:g._fimSistema||'',
        label:g._inicioSistema+(g._fimSistema?' — '+g._fimSistema:''),
        bloqueado:false,
        passado:horarioPassadoSistema(d,g._inicioMin),
        esperado:{...v,horarioId:g.Horario_ID||'',periodo:g.Periodo||'',turmaId:g.Turma_ID||'',turma:String(g.Turma||'').replace(/\.0$/,''),disciplina:g.Disciplina||'',sala:g.Sala||obterSalaDaTurma(g.Turma)},
        atual:atual?{retiradoPor:atual.Retirado_Por_Nome,responsavel:atual.Responsavel_Nome,turma:atual.Turma,sala:atual.Sala,retirada:atual.Data_Retirada,retorno:atual.Data_Retorno_Prevista,status:atual.Status}:null,
        situacao:atual?'EM_USO':'GRADE'
      });
    });
  });
  const chrome=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).find(c=>String(c.ID)===String(chromebookId));
  const current=obterLocalizacaoAtualChromebook(chromebookId,d);
  return serializarParaCliente({
    sucesso:true,
    chromebookId,
    data:d,
    turnoVinculos:vinculos,
    atual:current,
    localizacaoAtual:current?current.localAtual:(chrome?chrome.Localizacao||'':''),
    timeline:timeline.sort((a,b)=>minutosDoHorarioSistema(a.inicio)-minutosDoHorarioSistema(b.inicio))
  });
}
function diaSemanaNomeSistema(data) {
  const d=normalizarDataSistema(data);
  if(!d)return '';
  const dt=Utilities.parseDate(d+' 12:00',DEFAULT_CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm');
  return ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'][dt.getDay()] || '';
}

function listarProfessoresParaRetirada(token) {
  const sessao=exigirSessao(token);
  if(!sessao)return respostaSessaoExpirada();
  const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS).filter(u=>
    String(u.Ativo||'').toUpperCase()==='SIM' &&
    ['professor','professora'].includes(String(u.Tipo||'').trim().toLowerCase())
  );
  return serializarParaCliente({
    sucesso:true,
    professores:usuarios.map(u=>({id:u.ID,nome:u.Nome,email:u.Email||'',tipo:u.Tipo||''}))
      .sort((a,b)=>String(a.nome).localeCompare(String(b.nome),'pt-BR'))
  });
}

function obterContextoRetirada(token, data, horario, responsavelId, turmaInformada) {
  const sessao=exigirSessao(token);
  if(!sessao)return respostaSessaoExpirada();
  const d=normalizarDataSistema(data);
  const h=horarioTextoSistema(horario);
  if(!d||!h)return {sucesso:false,mensagem:'Informe a data e o horário da retirada.'};
  const professor=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(responsavelId));
  if(!professor)return {sucesso:false,mensagem:'Professor responsável não encontrado.'};
  const alvoMin=minutosDoHorarioSistema(h);
  const contexto=obterContextoGradeProfessorSistema(d,alvoMin,professor.Nome,professor.ID);
  const turma=String(turmaInformada|| (contexto?contexto.turma:'')).trim();
  const sala=(contexto&&contexto.turma===turma&&contexto.sala)?contexto.sala:obterSalaDaTurma(turma);
  return serializarParaCliente({
    sucesso:true,data:d,horario:contexto&&contexto.horaInicio?contexto.horaInicio:h,dia:diaSemanaNomeSistema(d),
    professor:{id:professor.ID,nome:professor.Nome},
    turmas:contexto&&contexto.turma?[contexto.turma]:turma?[turma]:[],
    turma,sala,encontradoHorario:!!contexto,disciplina:contexto?contexto.disciplina:'',
    horarioId:contexto?contexto.horarioId:'',turmaId:contexto?contexto.turmaId:'',periodo:contexto?contexto.periodo:''
  });
}
function criarDataHoraRetiradaSistema(data, horario) {
  const textoData=String(data||'').trim();
  const textoHora=String(horario||'').trim();
  const dh=(textoData+' '+textoHora).match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}:\d{2})$/);
  if(dh) return Utilities.parseDate(dh[1]+' '+horarioTextoSistema(dh[2]),DEFAULT_CONFIG.TIMEZONE,'yyyy-MM-dd HH:mm');
  const d=normalizarDataSistema(data);
  const m=minutosDoHorarioSistema(horario);
  if(!d||m===null)return null;
  return criarDataHoraPorMinutosSistema(d,m);
}

function fecharRetiradaChromeSistema(chromeId, dataReal, usuarioId) {
  const sh=getSheet(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  const rows=sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  const abertas=rows.filter(r=>
    String(r.Chromebook_ID)===String(chromeId) &&
    String(r.Status||'').toUpperCase()==='ABERTA'
  );
  abertas.forEach(r=>{
    const row=findRowById(SHEET_NAMES.RETIRADAS_CHROMEBOOKS,'ID',r.ID);
    if(row) updateRowByFields(SHEET_NAMES.RETIRADAS_CHROMEBOOKS,row.rowIndex,{
      Data_Devolucao_Real:dataReal||nowStr(),
      Status:'DEVOLVIDA'
    });
  });
}

function registrarRetiradaTemporalChrome(chromeId, retiradoPor, responsavel, dataRetirada, dataRetorno, turma, tipoRetirada) {
  appendObject(SHEET_NAMES.RETIRADAS_CHROMEBOOKS,{
    ID:genId('RETCHR'),
    Chromebook_ID:chromeId,
    Retirado_Por_ID:retiradoPor.ID,
    Retirado_Por_Nome:retiradoPor.Nome,
    Responsavel_ID:responsavel.ID,
    Responsavel_Nome:responsavel.Nome,
    Data_Retirada:dataRetirada,
    Data_Retorno_Prevista:dataRetorno,
    Data_Devolucao_Real:'',
    Turma:turma||'',
    Sala:obterSalaDaTurma(turma),
    Tipo_Retirada:tipoRetirada||'PROFESSOR',
    Status:'ABERTA'
  });
}

function listarRetiradasTemporaisAtivas(opcoes) {
  const fonte = opcoes && opcoes.retiradasTemporais ? opcoes.retiradasTemporais : sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
  return fonte
    .filter(r=>String(r.Status||'').toUpperCase()==='ABERTA');
}

function retiradaTemporalConflitaIntervalo(r, data, inicioMin, fimMin) {
  const inicioData=normalizarDataSistema(r.Data_Retirada);
  const retornoData=normalizarDataSistema(r.Data_Retorno_Prevista);
  if(!inicioData)return false;
  const alvo=normalizarDataSistema(data);
  if(!alvo)return false;
  const inicio=criarDataHoraRetiradaSistema(inicioData, horarioTextoSistema(r.Data_Retirada));
  const retorno=criarDataHoraRetiradaSistema(retornoData||inicioData, horarioTextoSistema(r.Data_Retorno_Prevista));
  const alvoInicio=criarDataHoraPorMinutosSistema(alvo,inicioMin);
  const alvoFim=criarDataHoraPorMinutosSistema(alvo,fimMin);
  if(!inicio||!alvoInicio||!alvoFim)return false;
  const retornoEfetivo=retorno||new Date(8640000000000000);
  return inicio < alvoFim && alvoInicio < retornoEfetivo;
}

function quantidadeRetiradasTemporaisNoIntervalo(data, inicioMin, fimMin) {
  const abertas=listarRetiradasTemporaisAtivas();
  return abertas.filter(r=>retiradaTemporalConflitaIntervalo(r,data,inicioMin,fimMin)).length;
}

function calcularDisponibilidadeChromebooks(data, horario, opcoes) {
  opcoes = opcoes || {};
  const chromes = opcoes.chromes || sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  const alunos = chromes.filter(c=>String(c.Tipo||'').toUpperCase()==='ALUNO');
  const total=alunos.length;
  const manutencao=alunos.filter(c=>String(c.Status||'').toUpperCase()==='MANUTENCAO').length;
  const intervalo=intervaloHorarioLaboratorio(data,horario);
  if(!intervalo)return{sucesso:false,data:normalizarDataSistema(data),horario:String(horario||''),total,emUso:0,manutencao,fisicamenteDisponiveis:alunos.filter(c=>String(c.Status||'').toUpperCase()==='DISPONIVEL').length,quantidadeComprometida:0,disponivel:0,idsDisponiveis:[],idsComprometidos:[],compromissos:[]};

  const compromissos=listarCompromissosChromebooksSistema(data,horario,opcoes);
  const retiradasTemporais=listarRetiradasTemporaisAtivas(opcoes).filter(r=>retiradaTemporalConflitaIntervalo(r,intervalo.data,intervalo.inicioMin,intervalo.fimMin));
  const idsRetirada=new Set(retiradasTemporais.map(r=>String(r.Chromebook_ID)));
  const idsComprometidos=new Set((compromissos.idsComprometidos||[]).map(String));
  const idsFisicamenteDisponiveis=alunos.filter(c=>String(c.Status||'').toUpperCase()==='DISPONIVEL').map(c=>String(c.ID));
  const idsExatosIndisponiveis=new Set([...idsComprometidos,...idsRetirada]);
  let idsDisponiveis=idsFisicamenteDisponiveis.filter(id=>!idsExatosIndisponiveis.has(id));

  // Registros antigos guardam apenas a quantidade. Eles continuam valendo,
  // mas não sabemos quais equipamentos físicos foram usados. Reservamos uma
  // parcela conservadora dos disponíveis até que a reserva seja vinculada a IDs.
  const quantidadeLegada=compromissos.compromissos.filter(c=>!c.ids||!c.ids.length).reduce((s,c)=>s+c.quantidade,0);
  const idsLegados=idsDisponiveis.slice(0,quantidadeLegada);
  idsDisponiveis=idsDisponiveis.slice(quantidadeLegada);
  idsLegados.forEach(id=>idsExatosIndisponiveis.add(id));

  const emUso=alunos.filter(c=>String(c.Status||'').toUpperCase()==='EM_USO').length;
  const quantidadeRetiradaTemporal=retiradasTemporais.length;
  const quantidadeComprometida=compromissos.quantidadeComprometida;
  return{
    sucesso:true,data:intervalo.data,horario:intervalo.inicio+' - '+intervalo.fim,total,emUso,manutencao,
    fisicamenteDisponiveis:idsFisicamenteDisponiveis.length,quantidadeComprometida,quantidadeRetiradaTemporal,
    quantidadeLegada,disponivel:idsDisponiveis.length,idsDisponiveis,idsComprometidos:[...idsExatosIndisponiveis],
    idsRetirada:[...idsRetirada],idsLegados,
    compromissos:compromissos.compromissos,
    retiradasTemporais:retiradasTemporais.map(r=>({id:r.ID,chromebookId:r.Chromebook_ID,responsavel:r.Responsavel_Nome,retiradoPor:r.Retirado_Por_Nome,turma:r.Turma,sala:r.Sala,retornoPrevisto:r.Data_Retorno_Prevista}))
  };
}


function validarQuantidadeChromebooksSistema(quantidade) {
  const qtd = quantidadeNumericaSistema(quantidade);
  if (qtd < 1) return { ok: false, quantidade: 0, mensagem: 'A quantidade de Chromebooks deve ser maior que zero.' };
  return { ok: true, quantidade: qtd };
}

function getChromesDisponiveis(tipo) {
  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  return chromes.filter(c => c.Status === 'DISPONIVEL' && (!tipo || c.Tipo === tipo));
}

function getChromesEmUso(usuarioId) {
  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
  return chromes.filter(c => c.Status === 'EM_USO' && String(c.Responsavel_ID) === String(usuarioId));
}

function registrarLogChrome(usuarioId, acao, quantidade, ids, tipo, avariaDesc) {
  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const usuario = usuarios.find(u => String(u.ID) === String(usuarioId));
  appendObject(SHEET_NAMES.LOG_CHROMEBOOKS, {
    ID: genId('LOGCHR'),
    Data_Hora: nowStr(),
    Acao: acao,
    Usuario_ID: usuarioId,
    Usuario_Nome: usuario ? usuario.Nome : '',
    Quantidade: quantidade,
    Chromes_IDS: Array.isArray(ids) ? ids.join(',') : ids,
    Tipo: tipo,
    Avaria_Desc: avariaDesc || ''
  });
}


function chromeRetirarComResponsavel(token, retiradoPorId, responsavelId, ids, retornoData, retornoHora, turma, horarioRetirada) {
  const sessao=exigirSessao(token);
  if(!sessao)return respostaSessaoExpirada();
  if(String(sessao.ID)!==String(retiradoPorId) && !perfilPodeAdministrar(sessao))return respostaSemPermissao();
  if(!Array.isArray(ids)||!ids.length)return{sucesso:false,mensagem:'Nenhum Chromebook selecionado.'};

  const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS);
  const retiradoPor=usuarios.find(u=>String(u.ID)===String(retiradoPorId));
  const responsavel=usuarios.find(u=>String(u.ID)===String(responsavelId));
  if(!retiradoPor||!responsavel)return{sucesso:false,mensagem:'Retirante ou professor responsável não encontrado.'};
  if(!['professor','professora'].includes(String(responsavel.Tipo||'').trim().toLowerCase()))return{sucesso:false,mensagem:'O responsável precisa ser um professor.'};
  if(!perfilPodeRetirarParaProfessor(sessao) && String(sessao.ID)!==String(responsavel.ID))return respostaSemPermissao();

  const agora=new Date();
  const dataAgora=Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'yyyy-MM-dd');
  const horaAgora=Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'HH:mm');
  const dRetorno=normalizarDataSistema(retornoData||dataAgora);
  const hRetorno=horarioTextoSistema(retornoHora);
  if(!hRetorno)return{sucesso:false,mensagem:'Informe o horário previsto de retorno.'};

  const inicio=criarDataHoraRetiradaSistema(dataAgora,horarioRetirada||horaAgora);
  const retorno=criarDataHoraRetiradaSistema(dRetorno,hRetorno);
  if(!inicio||!retorno||retorno<=inicio)return{sucesso:false,mensagem:'O retorno previsto deve ser posterior ao horário da retirada.'};

  const lock=LockService.getScriptLock();
  try{
    // Não prende a operação por 15 segundos. Se houver outra gravação simultânea,
    // falhamos rapidamente e o usuário pode tentar novamente.
    lock.waitLock(3000);
    const chromes=sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
    const vinculos=sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
    const retirados=[],recusados=[],updatesChrome=[],retiradaRows=[];
    const nomeResponsavel=normalizarNomeSistema(responsavel.Nome);
    const tipoRetirada=String(retiradoPor.Tipo||'').toLowerCase().includes('prof')?'PROFESSOR':'ATENDENTE';
    const retiradaData=dataAgora+' '+horaAgora, retornoDataHora=dRetorno+' '+hRetorno;

    ids.map(String).forEach(id=>{
      const chrome=chromes.find(c=>String(c.ID)===id);
      if(!chrome){recusados.push(id);return;}
      const tipo=String(chrome.Tipo||'').toUpperCase();
      const status=String(chrome.Status||'').toUpperCase();
      const fixoDoResponsavel=String(chrome.ID)===String(responsavel.Chromebook_Fixo_ID||'') || vinculos.some(v=>String(v.Chromebook_ID)===id && valorBooleanoSistema(v.Ativo,true) && (String(v.Professor_ID||'')===String(responsavel.ID) || normalizarNomeSistema(v.Professor_Nome)===nomeResponsavel));
      const podeRetirarFixo=tipo==='PROF'&&fixoDoResponsavel&&perfilPodeRetirarParaProfessor(sessao);
      if(status!=='DISPONIVEL'||(tipo!=='ALUNO'&&!podeRetirarFixo)){recusados.push(id);return;}
      const rowIndex=Number(chrome._row)||0;
      if(!rowIndex){recusados.push(id);return;}
      updatesChrome.push({rowIndex,fields:{Status:'EM_USO',Responsavel_ID:responsavel.ID,Data_Retirada:nowStr()}});
      retiradaRows.push({ID:genId('RETCHR'),Chromebook_ID:id,Retirado_Por_ID:retiradoPor.ID,Retirado_Por_Nome:retiradoPor.Nome,Responsavel_ID:responsavel.ID,Responsavel_Nome:responsavel.Nome,Data_Retirada:retiradaData,Data_Retorno_Prevista:retornoDataHora,Data_Devolucao_Real:'',Turma:turma||'',Sala:obterSalaDaTurma(turma),Tipo_Retirada:tipoRetirada,Status:'ABERTA'});
      retirados.push(id);
    });

    if(updatesChrome.length) atualizarLinhasBatch(SHEET_NAMES.CHROMEBOOKS,updatesChrome);
    if(retiradaRows.length) appendObjectsBatch(SHEET_NAMES.RETIRADAS_CHROMEBOOKS,retiradaRows);
    if(retirados.length)registrarLogChrome(retiradoPor.ID,'RETIRADA',retirados.length,retirados,'RESPONSAVEL');
    if(retirados.length){
      const contexto=obterContextoGradeProfessorSistema(dataAgora,minutosDoHorarioSistema(horaAgora),responsavel.Nome,responsavel.ID);
      retirados.forEach(id=>registrarAuditoriaChromebook({
        acao:'RETIRADA',chromebookId:id,professorId:responsavel.ID,professorNome:responsavel.Nome,tipoUso:'PROFESSOR',
        contexto:'RETIRADA',turno:responsavel.Turno||'',diaSemana:diaSemanaNomeSistema(dataAgora),
        horarioAula:contexto&&contexto.horaInicio?(contexto.horaInicio+(contexto.horaFim?' — '+contexto.horaFim:'')):horaAgora,
        turma:turma|| (contexto?contexto.turma:''),disciplina:contexto?contexto.disciplina:'',sala:obterSalaDaTurma(turma|| (contexto?contexto.turma:'')),
        localAtual:obterSalaDaTurma(turma|| (contexto?contexto.turma:'')),origem:'RETIRADA_RESPONSAVEL',status:'SUCESSO',
        observacao:'Retirada realizada por '+retiradoPor.Nome+'. Professor responsável: '+responsavel.Nome+'.'
      }));
    }
    return {sucesso:retirados.length>0,mensagem:retirados.length+' Chromebook(s) retirado(s). Retorno previsto: '+hRetorno+'.',retirados,recusados,responsavelId:responsavel.ID,responsavelNome:responsavel.Nome,turma:turma||'',sala:obterSalaDaTurma(turma),retornoPrevisto:hRetorno};
  }catch(err){
    logSistema('ERROR','CHROME_RETIRAR_RESPONSAVEL',retiradoPorId,err.message);
    return{sucesso:false,mensagem:'Erro: '+err.message};
  }finally{try{lock.releaseLock();}catch(e){}}
}


function obterTodosVinculosProfessor(professorId, professorNome){
  sincronizarProfessorIdsVinculos();
  const alvoId=String(professorId||'').trim();
  const alvoNome=normalizarNomeSistema(professorNome);
  return sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS).filter(v=>{
    if(!valorBooleanoSistema(v.Ativo,true))return false;
    if(alvoId && String(v.Professor_ID||'')===alvoId)return true;
    const pn=normalizarNomeSistema(v.Professor_Nome);
    if(!pn||!alvoNome)return false;
    if(pn===alvoNome)return true;
    return pn.split(/\s+/)[0]===alvoNome.split(/\s+/)[0] && pn.split(/\s+/)[0].length>=5;
  });
}


/**
 * V8 — consulta rápida do Chromebook fixo do próprio professor.
 * Não executa sincronização, localização, auditoria ou inventário completo.
 */
function getMeuChromebookFixoRapido(token) {
  const sessao=exigirSessao(token);
  if(!sessao)return respostaSessaoExpirada();

  const u=findRowById(SHEET_NAMES.USUARIOS,'ID',sessao.ID);
  if(!u)return{sucesso:false,mensagem:'Usuário não encontrado.'};

  const idx=n=>u.headers.indexOf(n);
  const nome=String(u.values[idx('Nome')]||sessao.Nome||'');
  const turno=String(u.values[idx('Turno')]||sessao.Turno||'').trim();
  let fixoId=String(u.values[idx('Chromebook_Fixo_ID')]||'').trim();

  if(!fixoId||fixoId.toLowerCase()==='professor'){
    try{
      const vs=sheetToObjects(SHEET_NAMES.VINCULOS_CHROMEBOOKS);
      const id=String(sessao.ID||''),nn=normalizarNomeSistema(nome),tn=normalizarNomeSistema(turno);
      const candidatos=vs.filter(x=>valorBooleanoSistema(x.Ativo,true)&&
        ((id&&String(x.Professor_ID||'')===id)||(nn&&normalizarNomeSistema(x.Professor_Nome)===nn)));
      const v=candidatos.find(x=>!tn||tn==='ambos'||normalizarNomeSistema(x.Turno)===tn)||candidatos[0];
      if(v)fixoId=String(v.Chromebook_ID||'').trim();
    }catch(e){}
  }

  if(!fixoId||fixoId.toLowerCase()==='professor')
    return{sucesso:false,mensagem:'Seu Chromebook fixo ainda não está vinculado ao seu perfil.'};

  const c=findRowById(SHEET_NAMES.CHROMEBOOKS,'ID',fixoId);
  if(!c)return{sucesso:false,mensagem:'O Chromebook '+fixoId+' não foi encontrado no inventário.'};

  const get=n=>String(c.values[c.headers.indexOf(n)]||'');
  return{
    sucesso:true,
    professor:{id:sessao.ID,nome,turno},
    chromebook:{
      id:fixoId,
      status:(get('Status')||'DISPONIVEL').toUpperCase(),
      retiradoEm:get('Data_Retirada')||get('Retirada_Em')||'',
      retornoPrevisto:get('Retorno_Previsto')||''
    }
  };
}


function chromeRetirarFixo(token, idProfessor) {
  try {
    const usuarioSessao = exigirSessao(token);
    if (!usuarioSessao) return respostaSessaoExpirada();
    if (String(usuarioSessao.ID)!==String(idProfessor)) return respostaSemPermissao();
    const usuario = findRowById(SHEET_NAMES.USUARIOS, 'ID', idProfessor);
    if (!usuario) return { sucesso: false, mensagem: 'Usuário não encontrado.' };
    const idxFixo = usuario.headers.indexOf('Chromebook_Fixo_ID');
    let fixoId = String(usuario.values[idxFixo]||'').trim();
    const professorId=String(usuario.values[usuario.headers.indexOf('ID')]||'');
    const professorNome=String(usuario.values[usuario.headers.indexOf('Nome')]||'');
    const turnoUsuario=String(usuario.values[usuario.headers.indexOf('Turno')]||'').trim();
    const vs=obterTodosVinculosProfessor(professorId,professorNome);
    const turnoNorm=normalizarNomeSistema(turnoUsuario);
    const vinculoDoFixo=fixoId ? vs.find(v=>String(v.Chromebook_ID)===fixoId) : null;
    // Nunca confiar em um ID antigo/errado gravado em USUARIOS. O vínculo por
    // professor + turno é a fonte operacional, enquanto o ID físico continua
    // sendo o patrimônio do equipamento.
    if(!vinculoDoFixo || String(vinculoDoFixo.Turno||'').trim()==='' ||
       (turnoNorm && normalizarNomeSistema(vinculoDoFixo.Turno)!==turnoNorm && turnoNorm!=='ambos') ){
      const v=vs.find(x=>!turnoNorm || turnoNorm==='ambos' || normalizarNomeSistema(x.Turno)===turnoNorm)||vs[0];
      if(v)fixoId=String(v.Chromebook_ID||'').trim();
    }
    if(!fixoId || fixoId.toLowerCase()==='professor')return {sucesso:false,mensagem:'Usuário não possui Chromebook fixo corretamente vinculado ao turno '+(turnoUsuario||'do usuário')+'.'};
    if (!fixoId) return { sucesso: false, mensagem: 'Usuário não possui Chromebook fixo cadastrado.' };

    const chrome = findRowById(SHEET_NAMES.CHROMEBOOKS, 'ID', fixoId);
    if (!chrome) return { sucesso: false, mensagem: 'Chromebook fixo não encontrado no estoque.' };
    const statusIdx = chrome.headers.indexOf('Status');
    if (chrome.values[statusIdx] === 'EM_USO') return { sucesso: false, mensagem: 'Chromebook fixo já está em uso.' };

    const agora=new Date();
    const dataAgora=Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'yyyy-MM-dd');
    const horaAgora=Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'HH:mm');
    const contexto=obterContextoGradeProfessorSistema(dataAgora, minutosDoHorarioSistema(horaAgora), usuario.values[usuario.headers.indexOf('Nome')]);
    const horarios=obterHorariosLaboratorioConfigurados();
    const ultimo=horarios.slice().sort((a,b)=>b._ordemSistema-a._ordemSistema).find(h=>!horarioBloqueadoSistema(h,horarios))||horarios[horarios.length-1];
    const retornoHora=ultimo?ultimo._fimSistema:'17:10';
    updateRowByFields(SHEET_NAMES.CHROMEBOOKS, chrome.rowIndex, {
      Status: 'EM_USO', Responsavel_ID: idProfessor, Data_Retirada: nowStr()
    });
    registrarRetiradaTemporalChrome(fixoId,usuario,usuario,dataAgora+' '+horaAgora,dataAgora+' '+retornoHora,contexto?contexto.turma:'', 'PROFESSOR');
    registrarLogChrome(idProfessor, 'RETIRADA', 1, [fixoId], 'PROF');
    registrarAuditoriaChromebook({
      acao:'RETIRADA',chromebookId:fixoId,professorId:usuario.ID,professorNome:usuario.Nome,tipoUso:'PROFESSOR',
      contexto:'RETIRADA',turno:usuario.Turno||'',diaSemana:diaSemanaNomeSistema(dataAgora),
      horarioAula:contexto&&contexto.horaInicio?(contexto.horaInicio+(contexto.horaFim?' — '+contexto.horaFim:'')):horaAgora,
      turma:contexto?contexto.turma:'',disciplina:contexto?contexto.disciplina:'',sala:contexto?contexto.sala:'',
      localAtual:contexto?contexto.sala:'',origem:'RETIRADA_FIXA',status:'SUCESSO',
      observacao:'Retirada realizada pelo próprio professor.'
    });
    return { sucesso: true, mensagem: 'Chromebook fixo retirado com sucesso. Retorno previsto: '+retornoHora+'.', chromeId: fixoId, retornoPrevisto: retornoHora, turma:contexto?contexto.turma:'', sala:contexto?contexto.sala:'' };
  } catch (err) {
    logSistema('ERROR', 'CHROME_RETIRAR_FIXO', idProfessor, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function chromeRetirarAlunos(token, idProfessor, ids) {
  const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  if (!ids || !ids.length) return { sucesso: false, mensagem: 'Nenhum Chromebook selecionado.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000);
    const retirados = [];
    const recusados = [];
    ids.forEach(id => {
      const chrome = findRowById(SHEET_NAMES.CHROMEBOOKS, 'ID', id);
      if (!chrome) { recusados.push(id); return; }
      const tipo = String(chrome.values[chrome.headers.indexOf('Tipo')]).toUpperCase();
      const status = String(chrome.values[chrome.headers.indexOf('Status')]).toUpperCase();
      if (tipo !== 'ALUNO' || status !== 'DISPONIVEL') { recusados.push(id); return; }
      updateRowByFields(SHEET_NAMES.CHROMEBOOKS, chrome.rowIndex, {
        Status: 'EM_USO', Responsavel_ID: idProfessor, Data_Retirada: nowStr()
      });
      retirados.push(id);
    });
    if (retirados.length) {
      registrarLogChrome(idProfessor, 'RETIRADA', retirados.length, retirados, 'ALUNO');
      const prof=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(idProfessor));
      const agora=new Date();
      const ctx=prof?obterContextoGradeProfessorSistema(Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'yyyy-MM-dd'),minutosDoHorarioSistema(Utilities.formatDate(agora,DEFAULT_CONFIG.TIMEZONE,'HH:mm')),prof.Nome,prof.ID):null;
      retirados.forEach(id=>registrarAuditoriaChromebook({acao:'RETIRADA',chromebookId:id,professorId:idProfessor,professorNome:prof?prof.Nome:'',tipoUso:'ALUNO',contexto:'RETIRADA_ALUNOS',turno:prof?prof.Turno:'',diaSemana:diaSemanaNomeSistema(agora),horarioAula:ctx&&ctx.horaInicio?(ctx.horaInicio+(ctx.horaFim?' — '+ctx.horaFim:'')):'',turma:ctx?ctx.turma:'',disciplina:ctx?ctx.disciplina:'',sala:ctx?ctx.sala:'',localAtual:ctx?ctx.sala:'',origem:'RETIRADA_ALUNOS',status:'SUCESSO'}));
    }
    return {
      sucesso: retirados.length > 0,
      mensagem: retirados.length + ' Chromebook(s) retirado(s) com sucesso.' + (recusados.length ? ' ' + recusados.length + ' seleção(ões) não estava(m) disponível(eis) para empréstimo de aluno.' : ''),
      retirados: retirados, recusados: recusados
    };
  } catch (err) {
    logSistema('ERROR', 'CHROME_RETIRAR_ALUNOS', idProfessor, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function chromeRetirarFixoAlunos(token, idProfessor, ids) {
  const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  const r1 = chromeRetirarFixo(token, idProfessor);
  const r2 = chromeRetirarAlunos(token, idProfessor, ids);
  return {
    sucesso: r1.sucesso || r2.sucesso,
    mensagem: (r1.sucesso ? 'Fixo retirado. ' : '') + (r2.sucesso ? r2.retirados.length + ' selecionado(s) retirado(s).' : ''),
    detalhes: { fixo: r1, alunos: r2 }
  };
}

function chromeDevolver(token, idProfessor, avariaDesc) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const chromes = getChromesEmUso(idProfessor);
    if (!chromes.length) return { sucesso: false, mensagem: 'Nenhum Chromebook em uso por este usuário.' };
    const ids = [];
    chromes.forEach(c => {
      const row = findRowById(SHEET_NAMES.CHROMEBOOKS, 'ID', c.ID);
      const fields = { Status: 'DISPONIVEL', Responsavel_ID: '', Data_Retirada: '' };
      if (avariaDesc) {
        fields.Status = 'MANUTENCAO';
        fields.Avaria_Desc = avariaDesc;
        fields.Avaria_Data = nowStr();
        fields.Avaria_Status = 'PENDENTE';
        registrarAvaria(c.ID, 'CHROMEBOOK', idProfessor, avariaDesc);
      }
      updateRowByFields(SHEET_NAMES.CHROMEBOOKS, row.rowIndex, fields);
      fecharRetiradaChromeSistema(c.ID, nowStr(), idProfessor);
      ids.push(c.ID);
    });
    registrarLogChrome(idProfessor, 'DEVOLUCAO', ids.length, ids, 'TODOS', avariaDesc);
    const profDev=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(idProfessor));
    ids.forEach(id=>registrarAuditoriaChromebook({acao:'DEVOLUCAO',chromebookId:id,professorId:idProfessor,professorNome:profDev?profDev.Nome:'',tipoUso:String(id).toUpperCase().indexOf('ALUNO')===0?'ALUNO':'PROFESSOR',contexto:'DEVOLUCAO',turno:profDev?profDev.Turno:'',diaSemana:diaSemanaNomeSistema(new Date()),origem:'DEVOLUCAO',status:'SUCESSO',observacao:avariaDesc?'Avaria registrada: '+avariaDesc:'Equipamento devolvido ao estoque.'}));
    return { sucesso: true, mensagem: ids.length + ' Chromebook(s) devolvido(s).' };
  } catch (err) {
    logSistema('ERROR', 'CHROME_DEVOLVER', idProfessor, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function chromeDevolverUm(token, idProfessor, id, avariaDesc) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const row = findRowById(SHEET_NAMES.CHROMEBOOKS, 'ID', id);
    if (!row) return { sucesso: false, mensagem: 'Chromebook não encontrado.' };
    const fields = { Status: 'DISPONIVEL', Responsavel_ID: '', Data_Retirada: '' };
    if (avariaDesc) {
      fields.Status = 'MANUTENCAO';
      fields.Avaria_Desc = avariaDesc;
      fields.Avaria_Data = nowStr();
      fields.Avaria_Status = 'PENDENTE';
      registrarAvaria(id, 'CHROMEBOOK', idProfessor, avariaDesc);
    }
    updateRowByFields(SHEET_NAMES.CHROMEBOOKS, row.rowIndex, fields);
    fecharRetiradaChromeSistema(id, nowStr(), idProfessor);
    registrarLogChrome(idProfessor, 'DEVOLUCAO', 1, [id], 'UNITARIO', avariaDesc);
    const profDev1=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(idProfessor));
    registrarAuditoriaChromebook({acao:'DEVOLUCAO',chromebookId:id,professorId:idProfessor,professorNome:profDev1?profDev1.Nome:'',tipoUso:String(id).toUpperCase().indexOf('ALUNO')===0?'ALUNO':'PROFESSOR',contexto:'DEVOLUCAO',turno:profDev1?profDev1.Turno:'',diaSemana:diaSemanaNomeSistema(new Date()),origem:'DEVOLUCAO_UNITARIA',status:'SUCESSO',observacao:avariaDesc?'Avaria registrada: '+avariaDesc:'Equipamento devolvido ao estoque.'});
    return { sucesso: true, mensagem: 'Chromebook ' + id + ' devolvido.' };
  } catch (err) {
    logSistema('ERROR', 'CHROME_DEVOLVER_UM', idProfessor, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function chromeDevolverSelecionados(token, idProfessor, ids, avariaDesc) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    if (!Array.isArray(ids) || !ids.length) return { sucesso:false, mensagem:'Selecione pelo menos um Chromebook para devolver.' };

    const lock=LockService.getScriptLock();
    try{
      lock.waitLock(3000);
      const chromes=sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
      const selecionados=ids.map(String).filter(id=>chromes.some(c=>String(c.ID)===id && String(c.Status||'').toUpperCase()==='EM_USO' && String(c.Responsavel_ID)===String(idProfessor)));
      if(!selecionados.length)return{sucesso:false,mensagem:'Nenhum dos equipamentos selecionados está em seu uso.'};

      const agora=nowStr(), updates=[], devolvidos=[], retiradaRows=sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS);
      selecionados.forEach(id=>{
        const c=chromes.find(x=>String(x.ID)===id);
        if(!c)return;
        const fields={Status:'DISPONIVEL',Responsavel_ID:'',Data_Retirada:''};
        if(avariaDesc){fields.Status='MANUTENCAO';fields.Avaria_Desc=avariaDesc;fields.Avaria_Data=agora;fields.Avaria_Status='PENDENTE';registrarAvaria(id,'CHROMEBOOK',idProfessor,avariaDesc);}
        updates.push({rowIndex:Number(c._row),fields});
        devolvidos.push(id);
      });

      if(updates.length)atualizarLinhasBatch(SHEET_NAMES.CHROMEBOOKS,updates);
      const closeUpdates=[];
      retiradaRows.forEach(r=>{if(selecionados.includes(String(r.Chromebook_ID))&&String(r.Status||'').toUpperCase()==='ABERTA'){closeUpdates.push({rowIndex:Number(r._row),fields:{Data_Devolucao_Real:agora,Status:'DEVOLVIDA'}});}});
      if(closeUpdates.length)atualizarLinhasBatch(SHEET_NAMES.RETIRADAS_CHROMEBOOKS,closeUpdates);
      if(devolvidos.length){
        registrarLogChrome(idProfessor,'DEVOLUCAO',devolvidos.length,devolvidos,'SELECIONADOS',avariaDesc);
        const profDevSel=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(idProfessor));
        devolvidos.forEach(id=>registrarAuditoriaChromebook({acao:'DEVOLUCAO',chromebookId:id,professorId:idProfessor,professorNome:profDevSel?profDevSel.Nome:'',tipoUso:String(id).toUpperCase().indexOf('ALUNO')===0?'ALUNO':'PROFESSOR',contexto:'DEVOLUCAO',turno:profDevSel?profDevSel.Turno:'',diaSemana:diaSemanaNomeSistema(new Date()),origem:'DEVOLUCAO_SELECIONADA',status:'SUCESSO',observacao:avariaDesc?'Avaria registrada: '+avariaDesc:'Equipamento devolvido ao estoque.'}));
      }
      return{sucesso:devolvidos.length>0,mensagem:devolvidos.length+' Chromebook(s) selecionado(s) devolvido(s).',devolvidos};
    }finally{try{lock.releaseLock();}catch(e){}}
  } catch(err) {
    logSistema('ERROR','CHROME_DEVOLVER_SELECIONADOS',idProfessor,err.message);
    return{sucesso:false,mensagem:'Erro: '+err.message};
  }
}

function chromeEncerrarAula(token, idProfessor) {
  const usuarioSessao = exigirProprioOuAdmin(token, idProfessor);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  const resultado = chromeDevolver(token, idProfessor, '');
  invalidarTokensProfessor(idProfessor);
  return resultado;
}

function getCalendarSistema() {
  const calendarId = String(getConfigValor('CALENDAR_ID') || '').trim();
  if (calendarId) {
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error('O CALENDAR_ID configurado não foi encontrado ou não está acessível.');
    return cal;
  }
  return CalendarApp.getDefaultCalendar();
}

function getWebAppUrlSistema(){const c=String(getConfigValor('WEB_APP_URL')||'').trim();if(c)return c.replace(/\/$/,'');try{return ScriptApp.getService().getUrl()||'';}catch(e){return '';}}
function montarAcaoSistemaUrl(tipo,id){const b=getWebAppUrlSistema();return b?b+'?acao='+encodeURIComponent(tipo)+'&id='+encodeURIComponent(id):'';}
function escHtmlEmail(v){return String(v??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]));}
function emailBaseSistema(titulo,subtitulo,conteudo,botoes){const bs=(botoes||[]).map(b=>b.url?'<a href="'+escHtmlEmail(b.url)+'" style="display:inline-block;margin:4px 6px 4px 0;padding:11px 16px;border-radius:9px;background:'+(b.danger?'#E65353':'#1769FF')+';color:#fff;text-decoration:none;font-weight:700;font-size:13px">'+escHtmlEmail(b.label)+'</a>':'').join('');return '<div style="font-family:Arial,sans-serif;background:#F4F7FB;padding:24px;color:#10233F"><div style="max-width:620px;margin:0 auto;background:#fff;border:1px solid #E5EAF1;border-radius:16px;overflow:hidden"><div style="background:#061B36;padding:20px 24px;color:#fff"><div style="font-size:12px;letter-spacing:.08em;color:#00A9C7;font-weight:700">ICM T.I.</div><div style="font-size:20px;font-weight:800;margin-top:4px">Escola Cristo Rei</div></div><div style="padding:24px"><h2 style="margin:0 0 6px;font-size:22px">'+titulo+'</h2><div style="color:#66758A;font-size:13px;margin-bottom:18px">'+subtitulo+'</div>'+conteudo+'<div style="margin-top:22px">'+bs+'</div><div style="margin-top:22px;color:#8896A8;font-size:11px">Mensagem automática do ICM T.I. • Gestão Inteligente de Recursos.</div></div></div></div>';}

// ----------------------------------------------------------------------------
// AGENDAMENTOS DE CHROMEBOOKS
// ----------------------------------------------------------------------------

function chromeAgendar(token, dados) {
  const usuarioSessao = exigirProprioOuOperador(token, dados && dados.usuarioId);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  if (!dados || !dados.data || !dados.horario || !dados.turma) return { sucesso:false, mensagem:'Preencha data, horário e turma.' };

  const validacaoQtd = validarQuantidadeChromebooksSistema(dados.quantidade);
  if (!validacaoQtd.ok) return { sucesso:false, mensagem:validacaoQtd.mensagem };

  const quantidade = validacaoQtd.quantidade;
  const idsSelecionados=[...new Set(idsChromebooksSistema(dados.chromebooksIds))];
  if(idsSelecionados.length!==quantidade)return{sucesso:false,mensagem:'Selecione exatamente '+quantidade+' Chromebook(s).'};
  const intervalo = intervaloHorarioLaboratorio(dados.data, dados.horario);
  if (!intervalo) return { sucesso:false, mensagem:'Horário inválido ou não configurado para o laboratório.' };

  const horarioConfig = obterHorariosLaboratorioConfigurados().find(h => h._inicioSistema === intervalo.inicio);
  if (!horarioConfig || horarioBloqueadoSistema(horarioConfig)) {
    return { sucesso:false, mensagem:'Este horário está indisponível para agendamento.' };
  }
  if (horarioPassadoSistema(dados.data, intervalo.inicioMin)) {
    return { sucesso:false, mensagem:'Este horário já passou e não pode mais ser agendado.' };
  }

  const lock=LockService.getScriptLock();
  try {
    lock.waitLock(3000);

    const disponibilidade=calcularDisponibilidadeChromebooks(dados.data,intervalo.inicio);
    if(idsSelecionados.length){
      const disponiveisSet=new Set((disponibilidade.idsDisponiveis||[]).map(String));
      const invalidos=idsSelecionados.filter(id=>!disponiveisSet.has(String(id)));
      if(invalidos.length)return{sucesso:false,mensagem:'Um ou mais Chromebooks selecionados já estão comprometidos neste horário. Atualize a disponibilidade e selecione novamente.',disponibilidade,invalidos};
    }else if (quantidade>disponibilidade.disponivel) {
      return { sucesso:false, mensagem:'Não há Chromebooks de alunos suficientes para este horário. Disponível: '+disponibilidade.disponivel+'. Solicitado: '+quantidade+'.', disponibilidade };
    }

    const id=genId('AGD');
    const email=usuarioSessao.Email||'';

    appendObject(SHEET_NAMES.AGENDAMENTOS,{
      ID:id, Usuario_ID:dados.usuarioId, Data:intervalo.data, Horario:intervalo.inicio,
      Tipo:dados.tipo||'CHROMEBOOK', Quantidade:quantidade, Chromes_IDS:idsChromebooksTextoSistema(idsSelecionados), Turma:dados.turma,
      Observacao:dados.observacao||'', Status:'AGENDADO', Data_Criacao:nowStr()
    });
    registrarAuditoriaChromebook({
      acao:'AGENDAMENTO',chromebookId:'',professorId:dados.usuarioId,professorNome:usuarioSessao.Nome,tipoUso:'ALUNO',
      contexto:'AGENDAMENTO',turno:usuarioSessao.Turno||'',diaSemana:diaSemanaNomeSistema(intervalo.data),horarioAula:intervalo.label,
      turma:dados.turma,disciplina:'',sala:obterSalaDaTurma(dados.turma),localAtual:'',origem:'AGENDAMENTO_CHROMEBOOK',status:'SUCESSO',
      observacao:'Quantidade agendada: '+quantidade+'. ID: '+id+'. '+(dados.observacao||'')
    });

    let calendarCriado=false, emailEnviado=false;
    const avisos=[];

    try {
      const cal=getCalendarSistema();
      const inicio=criarDataHoraPorMinutosSistema(intervalo.data,intervalo.inicioMin);
      const fim=criarDataHoraPorMinutosSistema(intervalo.data,intervalo.fimMin);
      if(!inicio||!fim) throw new Error('Não foi possível montar a data/hora do evento.');

      const evento=cal.createEvent('Chromebooks — '+dados.turma,inicio,fim,{
        description:'Agendamento ID: '+id+
          '\nProfessor: '+(usuarioSessao.Nome||'')+
          '\nData: '+intervalo.data+
          '\nHorário: '+intervalo.label+
          '\nQuantidade: '+quantidade+
          '\nTurma: '+dados.turma+
          '\nObservação: '+(dados.observacao||'')
      });
      calendarCriado=!!evento;
      if (evento) salvarCalendarEventId(SHEET_NAMES.AGENDAMENTOS, id, evento.getId());
    } catch(calErr) {
      avisos.push('O agendamento foi salvo, mas o evento do Google Calendar não pôde ser criado.');
      logSistema('WARNING','CALENDAR_CHROME_ERRO',dados.usuarioId,calErr.message);
    }

    if(email){
      try{
        const dataFormatada=Utilities.formatDate(
          criarDataHoraPorMinutosSistema(intervalo.data,720),
          DEFAULT_CONFIG.TIMEZONE,'dd/MM/yyyy'
        );
        const abrirUrl=getWebAppUrlSistema();
        const cancelarUrl=montarAcaoSistemaUrl('cancelarAgendamentoChrome',id);
        const botoes=[];if(cancelarUrl)botoes.push({label:'Cancelar agendamento',url:cancelarUrl,danger:true});if(abrirUrl)botoes.push({label:'Abrir sistema',url:abrirUrl});
        const html=emailBaseSistema('💻 Agendamento confirmado','Sua retirada de Chromebooks foi agendada.','<div style="background:#F7F9FC;border:1px solid #E5EAF1;border-radius:12px;padding:16px;line-height:1.8;font-size:13px"><strong>📅 Data:</strong> '+escHtmlEmail(dataFormatada)+'<br><strong>🕘 Horário:</strong> '+escHtmlEmail(intervalo.label)+'<br><strong>👥 Turma:</strong> '+escHtmlEmail(dados.turma)+'<br><strong>💻 Quantidade:</strong> '+escHtmlEmail(quantidade)+'</div>',botoes);
        MailApp.sendEmail({to:email,subject:'[ICM T.I] 💻 Agendamento de Chromebooks confirmado',htmlBody:html});
        emailEnviado=true;
      }catch(mailErr){
        avisos.push('O agendamento foi criado, mas o e-mail não pôde ser enviado.');
        logSistema('WARNING','EMAIL_CHROME_ERRO',dados.usuarioId,mailErr.message);
      }
    }else avisos.push('O usuário não possui e-mail cadastrado.');

    return { sucesso:true, mensagem:'Agendamento criado com sucesso.'+(avisos.length?' '+avisos.join(' '):''), id, calendarCriado, emailEnviado, disponibilidade };
  }catch(err){
    logSistema('ERROR','CHROME_AGENDAR',dados&&dados.usuarioId,err.message);
    return { sucesso:false, mensagem:'Erro: '+err.message };
  }finally{
    try{lock.releaseLock();}catch(e){}
  }
}


function getAgendamentosChrome(token, usuarioId) {
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();
  const ehAdmin=perfilPodeAdministrar(usuarioSessao);
  const alvo=ehAdmin&&usuarioId?usuarioId:usuarioSessao.ID;
  const horarios=obterHorariosLaboratorioConfigurados(), porInicio={}, porLabel={};
  horarios.forEach(h=>{porInicio[h._inicioSistema]=h;porLabel[h._inicioSistema+' — '+h._fimSistema]=h;});
  const agendamentos=sheetToObjects(SHEET_NAMES.AGENDAMENTOS).filter(a=>String(a.Usuario_ID)===String(alvo)).reverse();
  return serializarParaCliente({sucesso:true,agendamentos:agendamentos.map(a=>{const raw=String(a.Horario||'');const h=porInicio[horarioTextoSistema(raw)]||porLabel[raw];return {...a,HorarioLabel:h?h._inicioSistema+' — '+h._fimSistema:raw};})});
}


function cancelarAgendamentoChrome(token,id){
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();

  const row=findRowById(SHEET_NAMES.AGENDAMENTOS,'ID',id);
  if(!row)return{sucesso:false,mensagem:'Agendamento não encontrado.'};

  const usuarioAgendamento=row.values[row.headers.indexOf('Usuario_ID')];
  if(!perfilPodeAdministrar(usuarioSessao)&&String(usuarioAgendamento)!==String(usuarioSessao.ID))return respostaSemPermissao();

  updateCell(SHEET_NAMES.AGENDAMENTOS,row.rowIndex,'Status','CANCELADO');
  const profCancelChrome=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(usuarioAgendamento));
  registrarAuditoriaChromebook({
    acao:'AGENDAMENTO_CANCELADO',chromebookId:'',professorId:usuarioAgendamento,professorNome:profCancelChrome?profCancelChrome.Nome:'',tipoUso:'ALUNO',
    contexto:'CANCELAMENTO_AGENDAMENTO',turno:profCancelChrome?profCancelChrome.Turno:'',diaSemana:diaSemanaNomeSistema(row.values[row.headers.indexOf('Data')]),
    horarioAula:horarioTextoSistema(row.values[row.headers.indexOf('Horario')]),turma:row.values[row.headers.indexOf('Turma')]||'',sala:obterSalaDaTurma(row.values[row.headers.indexOf('Turma')]||''),
    origem:'CANCELAMENTO_AGENDAMENTO',status:'SUCESSO',observacao:'Quantidade: '+(row.values[row.headers.indexOf('Quantidade')]||0)+'. ID: '+id
  });

  let calendarRemovido=false,emailEnviado=false;
  const data=row.values[row.headers.indexOf('Data')];
  const horario=row.values[row.headers.indexOf('Horario')];
  const turma=row.values[row.headers.indexOf('Turma')];
  const email=usuarioSessao.Email||'';

  try{
    const cal=getCalendarSistema();
    const eventId=obterCalendarEventId(row);
    if(eventId){const ev=cal.getEventById(eventId);if(ev){ev.deleteEvent();calendarRemovido=true;}}
    if(!calendarRemovido){
      const dataNorm=normalizarDataSistema(data);
      const inicioBusca=criarDataHoraPorMinutosSistema(dataNorm,0);
      const fimBusca=new Date(inicioBusca.getTime());fimBusca.setDate(fimBusca.getDate()+1);
      cal.getEvents(inicioBusca,fimBusca).forEach(ev=>{if((ev.getDescription()||'').indexOf('Agendamento ID: '+id)!==-1){ev.deleteEvent();calendarRemovido=true;}});
    }
  }catch(err){logSistema('WARNING','CALENDAR_CANCELAMENTO_CHROME_ERRO',usuarioSessao.ID,err.message)}

  if(email){
    try{
      const dataFormatada=Utilities.formatDate(
        criarDataHoraPorMinutosSistema(normalizarDataSistema(data),720),
        DEFAULT_CONFIG.TIMEZONE,'dd/MM/yyyy'
      );
      const abrirUrl=getWebAppUrlSistema();
      const html=emailBaseSistema('❌ Agendamento cancelado','O agendamento de Chromebooks foi cancelado.','<div style="background:#FFF7F7;border:1px solid #F4D5D5;border-radius:12px;padding:16px;line-height:1.8;font-size:13px"><strong>📅 Data:</strong> '+escHtmlEmail(dataFormatada)+'<br><strong>🕘 Horário:</strong> '+escHtmlEmail(horario||'')+'<br><strong>👥 Turma:</strong> '+escHtmlEmail(turma||'')+'</div>',abrirUrl?[{label:'Abrir sistema',url:abrirUrl}]:[]);
      MailApp.sendEmail({to:email,subject:'[ICM T.I] ❌ Agendamento de Chromebooks cancelado',htmlBody:html});
      emailEnviado=true;
    }catch(err){logSistema('WARNING','EMAIL_CANCELAMENTO_CHROME_ERRO',usuarioSessao.ID,err.message)}
  }

  return{sucesso:true,mensagem:'Agendamento cancelado.',calendarRemovido,emailEnviado};
}


// ----------------------------------------------------------------------------
// RECURSOS (FONE / VR / PROJETOR / CONTROLE)
// ----------------------------------------------------------------------------

function listarRecursos() {
  const recursos = sheetToObjects(SHEET_NAMES.RECURSOS);
  const stats = {};
  recursos.forEach(r => {
    stats[r.Tipo] = stats[r.Tipo] || { total: 0, disponivel: 0, emUso: 0, manutencao: 0 };
    stats[r.Tipo].total++;
    if (r.Status === 'DISPONIVEL') stats[r.Tipo].disponivel++;
    else if (r.Status === 'EM_USO') stats[r.Tipo].emUso++;
    else if (r.Status === 'MANUTENCAO') stats[r.Tipo].manutencao++;
  });
  return { recursos: recursos, stats: stats, total: recursos.length };
}

function registrarLogRecurso(usuarioId, acao, recursoId, tipo) {
  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const usuario = usuarios.find(u => String(u.ID) === String(usuarioId));
  appendObject(SHEET_NAMES.LOG_RECURSOS, {
    ID: genId('LOGREC'),
    Data_Hora: nowStr(),
    Acao: acao,
    Usuario_ID: usuarioId,
    Usuario_Nome: usuario ? usuario.Nome : '',
    Recurso_ID: recursoId,
    Tipo: tipo
  });
}

function recursoRetirar(token, id, usuarioId, localizacao) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, usuarioId);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const row = findRowById(SHEET_NAMES.RECURSOS, 'ID', id);
    if (!row) return { sucesso: false, mensagem: 'Recurso não encontrado.' };
    const statusIdx = row.headers.indexOf('Status');
    if (row.values[statusIdx] !== 'DISPONIVEL') return { sucesso: false, mensagem: 'Recurso não está disponível.' };
    const tipoIdx = row.headers.indexOf('Tipo');
    updateRowByFields(SHEET_NAMES.RECURSOS, row.rowIndex, {
      Status: 'EM_USO', Responsavel_ID: usuarioId, Data_Retirada: nowStr(),
      Localizacao: localizacao || row.values[row.headers.indexOf('Localizacao')]
    });
    registrarLogRecurso(usuarioId, 'RETIRADA', id, row.values[tipoIdx]);
    return { sucesso: true, mensagem: 'Recurso retirado com sucesso.' };
  } catch (err) {
    logSistema('ERROR', 'RECURSO_RETIRAR', usuarioId, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function recursoDevolver(token, id, usuarioId) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, usuarioId);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const row = findRowById(SHEET_NAMES.RECURSOS, 'ID', id);
    if (!row) return { sucesso: false, mensagem: 'Recurso não encontrado.' };
    const ehAdmin = String(usuarioSessao.Tipo || '').toLowerCase() === 'admin';
    const responsavel = row.values[row.headers.indexOf('Responsavel_ID')];
    if (!ehAdmin && String(responsavel) !== String(usuarioId)) return respostaSemPermissao();
    const tipoIdx = row.headers.indexOf('Tipo');
    updateRowByFields(SHEET_NAMES.RECURSOS, row.rowIndex, {
      Status: 'DISPONIVEL', Responsavel_ID: '', Data_Retirada: ''
    });
    registrarLogRecurso(usuarioId, 'DEVOLUCAO', id, row.values[tipoIdx]);
    return { sucesso: true, mensagem: 'Recurso devolvido com sucesso.' };
  } catch (err) {
    logSistema('ERROR', 'RECURSO_DEVOLVER', usuarioId, err.message);
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

// ----------------------------------------------------------------------------
// CONTROLES DE PROJETORES
// ----------------------------------------------------------------------------

function listarControles() {
  const recursos = sheetToObjects(SHEET_NAMES.RECURSOS);
  return recursos.filter(r => r.Tipo === 'CONTROLE');
}

function controleRetirar(token, dados) {
  try {
    const usuarioSessao = exigirProprioOuAdmin(token, dados && dados.usuarioId);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const row = findRowById(SHEET_NAMES.RECURSOS, 'ID', dados.controleId);
    if (!row) return { sucesso: false, mensagem: 'Controle não encontrado.' };
    const statusIdx = row.headers.indexOf('Status');
    if (row.values[statusIdx] !== 'DISPONIVEL') return { sucesso: false, mensagem: 'Controle não está disponível.' };
    updateRowByFields(SHEET_NAMES.RECURSOS, row.rowIndex, {
      Status: 'EM_USO', Responsavel_ID: dados.usuarioId, Data_Retirada: nowStr(), Localizacao: dados.sala || ''
    });
    registrarLogRecurso(dados.usuarioId, 'RETIRADA', dados.controleId, 'CONTROLE');
    return { sucesso: true, mensagem: 'Controle retirado com sucesso.' };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function controleDevolver(token, controleId, usuarioId) {
  return recursoDevolver(token, controleId, usuarioId);
}

// ----------------------------------------------------------------------------
// LABORATÓRIO
// ----------------------------------------------------------------------------

function getHorariosLaboratorio(token, data) {
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();
  const dataNormalizada=normalizarDataSistema(data);
  if(!dataNormalizada)return{sucesso:false,mensagem:'Data inválida para consultar a agenda.'};

  const horarios=obterHorariosLaboratorioConfigurados();
  const reservasLab=sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO).filter(r=>String(r.Status||'').toUpperCase()==='ATIVA'&&normalizarDataSistema(r.Data)===dataNormalizada);
  const legados=sheetToObjects(SHEET_NAMES.AGENDAMENTOS).filter(a=>String(a.Status||'').toUpperCase()==='AGENDADO'&&normalizarDataSistema(a.Data)===dataNormalizada);
  const chromes=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).filter(c=>String(c.Tipo||'').toUpperCase()==='ALUNO');
  const retiradasTemporais=sheetToObjects(SHEET_NAMES.RETIRADAS_CHROMEBOOKS).filter(r=>String(r.Status||'').toUpperCase()==='ABERTA');
  const total=chromes.length;
  const manutencao=chromes.filter(c=>String(c.Status||'').toUpperCase()==='MANUTENCAO').length;
  const intervalos=horarios.map(h=>({h,inicioMin:minutosDoHorarioSistema(h._inicioSistema),fimMin:minutosDoHorarioSistema(h._fimSistema)}));
  const porInicio={}; intervalos.forEach(x=>porInicio[x.h._inicioSistema]=x);
  function localIntervalo(valor){
    const raw=String(valor||'');
    const p=raw.match(/(\d{1,2}:\d{2})\s*(?:—|-|–|a|até)\s*(\d{1,2}:\d{2})/i);
    if(p){const ini=horarioTextoSistema(p[1]),fim=horarioTextoSistema(p[2]);const x=intervalos.find(y=>y.h._inicioSistema===ini&&y.h._fimSistema===fim);if(x)return x;}
    return porInicio[horarioTextoSistema(raw)]||null;
  }
  const compromissos=[];
  reservasLab.forEach(r=>{const x=localIntervalo(r.Horario);if(x)compromissos.push({origem:'RESERVA_LABORATORIO',id:r.ID,quantidade:quantidadeNumericaSistema(r.Chromes),inicioMin:x.inicioMin,fimMin:x.fimMin,inicio:x.h._inicioSistema,fim:x.h._fimSistema,turma:r.Turma||'',professor:r.Professor||'',registro:r});});
  legados.forEach(a=>{const x=localIntervalo(a.Horario);if(x)compromissos.push({origem:'AGENDAMENTO',id:a.ID,quantidade:quantidadeNumericaSistema(a.Quantidade),inicioMin:x.inicioMin,fimMin:x.fimMin,inicio:x.h._inicioSistema,fim:x.h._fimSistema,turma:a.Turma||'',professor:a.Professor||'',registro:a});});
  const resultado=intervalos.map(x=>{
    const h=x.h;
    const label=h._inicioSistema+' — '+h._fimSistema;
    const bloqueadoConfiguracao=horarioBloqueadoSistema(h,horarios);
    const passado=horarioPassadoSistema(dataNormalizada,x.inicioMin);
    const bloqueado=bloqueadoConfiguracao || passado;
    const cs=compromissos.filter(c=>intervalosConflitamSistema(x.inicioMin,x.fimMin,c.inicioMin,c.fimMin));
    const reserva=cs.find(c=>c.origem==='RESERVA_LABORATORIO');
    const legado=!reserva?cs.find(c=>c.origem==='AGENDAMENTO'):null;
    const comprometido=cs.reduce((sum,c)=>sum+c.quantidade,0);
    const disponibilidadeChrome=calcularDisponibilidadeChromebooks(dataNormalizada,h._inicioSistema,{chromes,reservasLab,agendamentos:legados,retiradasTemporais});
    const disponivel=disponibilidadeChrome.sucesso?disponibilidadeChrome.disponivel:Math.max(0,total-manutencao-comprometido);
    const emUsoEfetivo=disponibilidadeChrome.sucesso?disponibilidadeChrome.emUso:0;
    const ocupante=reserva||legado;
    const statusOcupacao=passado?'INDISPONIVEL':(bloqueado?'INDISPONIVEL':(ocupante?'OCUPADO':'LIVRE'));
    const labOcupado=!!reserva;
    return {inicio:h._inicioSistema,fim:h._fimSistema,ordem:h._ordemSistema,label,bloqueado,passado,motivoBloqueio:passado?'HORARIO_PASSADO':(bloqueadoConfiguracao?'BLOQUEADO_CONFIGURACAO':''),statusOcupacao,labOcupado,chromebookAgendamentoPermitido:!bloqueado, reserva:reserva?reserva.registro:null,legado:legado?legado.registro:null,professor:ocupante?(ocupante.professor||''):'',turma:ocupante?(ocupante.turma||''):'',chromebooksTotal:total,chromebooksEmUso:emUsoEfetivo,chromebooksComprometidos:comprometido,chromebooksDisponiveis:disponivel,chromebooksDisponiveisIds:disponibilidadeChrome.idsDisponiveis||[],chromebooksComprometidosIds:disponibilidadeChrome.idsComprometidos||[],chromebooksRetiradaIds:disponibilidadeChrome.idsRetirada||[],chromebooksLegadosIds:disponibilidadeChrome.idsLegados||[],compromissos:cs.map(c=>({origem:c.origem,id:c.id,quantidade:c.quantidade,ids:c.ids||[],inicio:c.inicio,fim:c.fim,turma:c.turma}))};
  });
  return serializarParaCliente({sucesso:true,data:dataNormalizada,horarios:resultado});
}

function getChromebooksParaReserva(token,reservaId){
  const sessao=exigirSessao(token); if(!sessao)return respostaSessaoExpirada();
  const row=findRowById(SHEET_NAMES.RESERVAS_LABORATORIO,'ID',reservaId); if(!row)return{sucesso:false,mensagem:'Reserva não encontrada.'};
  const obj={}; row.headers.forEach((h,i)=>obj[h]=row.values[i]);
  const pode=perfilPodeAdministrar(sessao)||String(obj.Usuario_ID)===String(sessao.ID)||perfilPodeRetirarParaProfessor(sessao);
  if(!pode)return respostaSemPermissao();
  const data=normalizarDataSistema(obj.Data), intervalo=intervaloHorarioLaboratorio(data,obj.Horario);
  if(!intervalo)return{sucesso:false,mensagem:'Horário da reserva inválido.'};
  const atuais=new Set(idsChromebooksSistema(obj.Chromes_IDS));
  const chromes=sheetToObjects(SHEET_NAMES.CHROMEBOOKS).filter(c=>String(c.Tipo||'').toUpperCase()==='ALUNO');
  const disp=calcularDisponibilidadeChromebooks(data,intervalo.inicio,{excluirReservaId:String(reservaId)});
  const disponiveis=new Set((disp.idsDisponiveis||[]).map(String));
  return serializarParaCliente({sucesso:true,reserva:obj,chromebooks:chromes.map(c=>({id:String(c.ID),status:String(c.Status||'').toUpperCase(),selecionado:atuais.has(String(c.ID)),disponivel:atuais.has(String(c.ID))||disponiveis.has(String(c.ID))})),quantidade:quantidadeNumericaSistema(obj.Chromes),idsSelecionados:[...atuais]});
}
function atualizarChromebooksReserva(token,reservaId,ids){
  const sessao=exigirSessao(token); if(!sessao)return respostaSessaoExpirada();
  const row=findRowById(SHEET_NAMES.RESERVAS_LABORATORIO,'ID',reservaId); if(!row)return{sucesso:false,mensagem:'Reserva não encontrada.'};
  const usuarioReserva=row.values[row.headers.indexOf('Usuario_ID')];
  if(!(perfilPodeAdministrar(sessao)||String(usuarioReserva)===String(sessao.ID)||perfilPodeRetirarParaProfessor(sessao)))return respostaSemPermissao();
  const selecionados=[...new Set(idsChromebooksSistema(ids))];
  const qtd=quantidadeNumericaSistema(row.values[row.headers.indexOf('Chromes')]);
  if(selecionados.length!==qtd)return{sucesso:false,mensagem:'Selecione exatamente '+qtd+' Chromebook(s).'};
  const data=normalizarDataSistema(row.values[row.headers.indexOf('Data')]), intervalo=intervaloHorarioLaboratorio(data,row.values[row.headers.indexOf('Horario')]);
  const disp=calcularDisponibilidadeChromebooks(data,intervalo.inicio,{excluirReservaId:String(reservaId)});
  const ok=new Set((disp.idsDisponiveis||[]).map(String));
  const invalidos=selecionados.filter(id=>!ok.has(String(id)));
  if(invalidos.length)return{sucesso:false,mensagem:'Há Chromebooks indisponíveis na seleção: '+invalidos.join(', '),invalidos,disponibilidade:disp};
  updateRowByFields(SHEET_NAMES.RESERVAS_LABORATORIO,row.rowIndex,{Chromes_IDS:idsChromebooksTextoSistema(selecionados),Chromes:qtd});
  return{sucesso:true,mensagem:'Chromebooks da reserva atualizados.',idsSelecionados:selecionados,quantidade:qtd};
}

function reservarLaboratorio(token,dados){
  const usuarioSessao=exigirProprioOuOperador(token,dados&&dados.usuarioId);
  if(!usuarioSessao)return getUsuarioLogado(token)?respostaSemPermissao():respostaSessaoExpirada();
  if(!dados||!dados.data||!dados.horario||!dados.turma)return{sucesso:false,mensagem:'Preencha data, horário e turma.'};

  const horariosConfigurados=obterHorariosLaboratorioConfigurados();
  const recebido=String(dados.horario||'').trim();
  const horarioSelecionado=horariosConfigurados.find(h=>{
    const label=h._inicioSistema+' — '+h._fimSistema;
    return h._inicioSistema===horarioTextoSistema(recebido)||label===recebido||String(h.Label||'').trim()===recebido;
  });

  if(!horarioSelecionado)return{sucesso:false,mensagem:'Selecione um horário disponível da lista.'};
  if(horarioBloqueadoSistema(horarioSelecionado,horariosConfigurados))return{sucesso:false,mensagem:'Este horário está indisponível.'};

  const inicioSelecionado=minutosDoHorarioSistema(horarioSelecionado._inicioSistema);
  if(horarioPassadoSistema(dados.data,inicioSelecionado))return{sucesso:false,mensagem:'Este horário já passou e não pode mais ser reservado.'};

  const intervalo=intervaloHorarioLaboratorio(dados.data,horarioSelecionado._inicioSistema);
  if(!intervalo)return{sucesso:false,mensagem:'Horário inválido ou não configurado para o laboratório.'};

  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(3000);

    const reservasAtivas=sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO).filter(r=>
      String(r.Status||'').toUpperCase()==='ATIVA'&&normalizarDataSistema(r.Data)===intervalo.data
    );

    const ocupado=reservasAtivas.some(r=>{
      const rh=intervaloHorarioLaboratorio(intervalo.data,r.Horario);
      return rh&&intervalosConflitamSistema(rh.inicioMin,rh.fimMin,intervalo.inicioMin,intervalo.fimMin);
    });
    if(ocupado)return{sucesso:false,mensagem:'Este horário já está reservado no laboratório. Escolha outro horário.'};

    const validacaoQtd=(dados.chromes===''||dados.chromes===null||dados.chromes===undefined)
      ?{ok:true,quantidade:0}:validarQuantidadeChromebooksSistema(dados.chromes);
    const idsSelecionados=[...new Set(idsChromebooksSistema(dados.chromebooksIds))];
    const quantidadeChromes=quantidadeNumericaSistema(dados.chromes);
    if(!validacaoQtd.ok&&quantidadeChromes!==0)return{sucesso:false,mensagem:validacaoQtd.mensagem};
    if(idsSelecionados.length && idsSelecionados.length!==quantidadeChromes)return{sucesso:false,mensagem:'A quantidade de Chromebooks deve ser igual à seleção.'};
    if(quantidadeChromes>0 && !idsSelecionados.length)return{sucesso:false,mensagem:'Selecione quais Chromebooks serão usados no laboratório.'};

    const disponibilidade=calcularDisponibilidadeChromebooks(dados.data,intervalo.inicio);
    if(idsSelecionados.length){
      const ds=new Set((disponibilidade.idsDisponiveis||[]).map(String));
      const invalidos=idsSelecionados.filter(id=>!ds.has(String(id)));
      if(invalidos.length)return{sucesso:false,mensagem:'Um ou mais Chromebooks selecionados já estão comprometidos neste horário. Atualize a seleção.',disponibilidade,invalidos};
    } else if(quantidadeChromes>disponibilidade.disponivel){
      return{sucesso:false,mensagem:'Não há Chromebooks de alunos suficientes para este horário. Disponível: '+disponibilidade.disponivel+'. Solicitado: '+quantidadeChromes+'.',disponibilidade};
    }

    const id=genId('RES');
    const email=dados.email||usuarioSessao.Email||'';

    appendObject(SHEET_NAMES.RESERVAS_LABORATORIO,{
      ID:id,Data:intervalo.data,Usuario_ID:dados.usuarioId,
      Professor:dados.professor||usuarioSessao.Nome,Email:email,Horario:intervalo.label,
      Chromes:quantidadeChromes,Chromes_IDS:idsChromebooksTextoSistema(idsSelecionados),Fones:quantidadeNumericaSistema(dados.fones),
      VR:quantidadeNumericaSistema(dados.vr),Projetor:quantidadeNumericaSistema(dados.projetor),
      Turma:dados.turma,Pauta:dados.pauta||'',Status:'ATIVA',Timestamp:nowStr()
    });
    registrarAuditoriaLaboratorio({
      acao:'RESERVA_CRIADA',reservaId:id,professorId:dados.usuarioId,professorNome:dados.professor||usuarioSessao.Nome,
      dataReserva:intervalo.data,horarioInicio:intervalo.inicio,horarioFim:intervalo.fim,turma:dados.turma,
      laboratorio:'Laboratório de Informática',chromebook:quantidadeChromes,fones:quantidadeNumericaSistema(dados.fones),
      projetor:quantidadeNumericaSistema(dados.projetor),atividade:dados.pauta||'',origem:'RESERVA_LABORATORIO',status:'SUCESSO'
    });

    let calendarCriado=false,emailEnviado=false;
    const avisos=[];

    try{
      const cal=getCalendarSistema();
      const inicio=criarDataHoraPorMinutosSistema(intervalo.data,intervalo.inicioMin);
      const fim=criarDataHoraPorMinutosSistema(intervalo.data,intervalo.fimMin);
      if(!inicio||!fim)throw new Error('Não foi possível montar a data/hora do evento.');

      const evento=cal.createEvent('Laboratório — '+dados.turma,inicio,fim,{
        description:'Reserva ID: '+id+
          '\nProfessor: '+(dados.professor||usuarioSessao.Nome)+
          '\nData: '+intervalo.data+
          '\nHorário: '+intervalo.label+
          '\nChromebooks: '+quantidadeChromes+
          '\nFones: '+quantidadeNumericaSistema(dados.fones)+
          '\nPauta: '+(dados.pauta||'')
      });
      calendarCriado=!!evento;
      if (evento) salvarCalendarEventId(SHEET_NAMES.RESERVAS_LABORATORIO, id, evento.getId());
    }catch(calErr){
      avisos.push('A reserva foi salva, mas o evento do Google Calendar não pôde ser criado.');
      logSistema('WARNING','CALENDAR_RESERVA_ERRO',dados.usuarioId,calErr.message);
    }

    if(email){
      try{
        const dataFormatada=Utilities.formatDate(
          criarDataHoraPorMinutosSistema(intervalo.data,720),
          DEFAULT_CONFIG.TIMEZONE,'dd/MM/yyyy'
        );
        const abrirUrl=getWebAppUrlSistema();
        const cancelarUrl=montarAcaoSistemaUrl('cancelarReservaLaboratorio',id);
        const botoes=[];if(cancelarUrl)botoes.push({label:'Cancelar reserva',url:cancelarUrl,danger:true});if(abrirUrl)botoes.push({label:'Abrir sistema',url:abrirUrl});
        const html=emailBaseSistema('✅ Reserva confirmada','Seu laboratório foi reservado com sucesso.','<div style="background:#F7F9FC;border:1px solid #E5EAF1;border-radius:12px;padding:16px;line-height:1.8;font-size:13px"><strong>📅 Data:</strong> '+escHtmlEmail(dataFormatada)+'<br><strong>🕘 Horário:</strong> '+escHtmlEmail(intervalo.label)+'<br><strong>👥 Turma:</strong> '+escHtmlEmail(dados.turma)+'<br><strong>💻 Chromebooks:</strong> '+escHtmlEmail(quantidadeChromes)+'<br><strong>👤 Professor:</strong> '+escHtmlEmail(dados.professor||usuarioSessao.Nome)+'</div>',botoes);
        MailApp.sendEmail({to:email,subject:'[ICM T.I] ✅ Reserva de laboratório confirmada',htmlBody:html});
        emailEnviado=true;
      }catch(mailErr){
        avisos.push('A reserva foi criada, mas o e-mail de confirmação não pôde ser enviado.');
        logSistema('WARNING','EMAIL_RESERVA_ERRO',dados.usuarioId,mailErr.message);
      }
    }else avisos.push('O usuário não possui e-mail cadastrado; nenhuma confirmação por e-mail foi enviada.');

    return{
      sucesso:true,
      mensagem:avisos.length?'Reserva de laboratório criada com sucesso. '+avisos.join(' '):'Reserva de laboratório criada com sucesso.',
      id,horario:intervalo.label,calendarCriado,emailEnviado,disponibilidade
    };
  }catch(err){
    logSistema('ERROR','RESERVAR_LAB',dados&&dados.usuarioId,err.message);
    return{sucesso:false,mensagem:'Erro: '+err.message};
  }finally{
    try{lock.releaseLock()}catch(e){}
  }
}


function getReservasLaboratorio(token, usuarioId) {
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();
  const ehAdmin=perfilPodeAdministrar(usuarioSessao);
  const alvo=ehAdmin&&usuarioId?usuarioId:usuarioSessao.ID;
  const filtradas=sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO)
    .filter(r=>String(r.Status||'').toUpperCase()==='ATIVA')
    .filter(r=>ehAdmin?(!usuarioId||String(r.Usuario_ID)===String(alvo)):String(r.Usuario_ID)===String(alvo));
  return serializarParaCliente({sucesso:true,reservas:filtradas.reverse()});
}

function cancelarReservaLaboratorio(token,id){
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();

  const row=findRowById(SHEET_NAMES.RESERVAS_LABORATORIO,'ID',id);
  if(!row)return{sucesso:false,mensagem:'Reserva não encontrada.'};

  const usuarioReserva=row.values[row.headers.indexOf('Usuario_ID')];
  if(!perfilPodeAdministrar(usuarioSessao)&&String(usuarioReserva)!==String(usuarioSessao.ID))return respostaSemPermissao();

  updateCell(SHEET_NAMES.RESERVAS_LABORATORIO,row.rowIndex,'Status','CANCELADA');
  registrarAuditoriaLaboratorio({
    acao:'RESERVA_CANCELADA',reservaId:id,professorId:row.values[row.headers.indexOf('Usuario_ID')]||'',
    professorNome:row.values[row.headers.indexOf('Professor')]||usuarioSessao.Nome,dataReserva:normalizarDataSistema(row.values[row.headers.indexOf('Data')]),
    horarioInicio:horarioTextoSistema(row.values[row.headers.indexOf('Horario')]),horarioFim:'',turma:row.values[row.headers.indexOf('Turma')]||'',
    laboratorio:'Laboratório de Informática',chromebook:row.values[row.headers.indexOf('Chromes')]||0,fones:row.values[row.headers.indexOf('Fones')]||0,
    projetor:row.values[row.headers.indexOf('Projetor')]||0,atividade:row.values[row.headers.indexOf('Pauta')]||'',origem:'CANCELAMENTO_RESERVA',status:'SUCESSO'
  });

  let calendarRemovido=false,emailEnviado=false;
  const data=row.values[row.headers.indexOf('Data')];
  const horario=row.values[row.headers.indexOf('Horario')];
  const turma=row.values[row.headers.indexOf('Turma')];
  const email=row.values[row.headers.indexOf('Email')]||usuarioSessao.Email||'';

  try{
    const cal=getCalendarSistema();
    const eventId=obterCalendarEventId(row);
    if(eventId){const ev=cal.getEventById(eventId);if(ev){ev.deleteEvent();calendarRemovido=true;}}
    if(!calendarRemovido){
      const dataNorm=normalizarDataSistema(data);
      const inicioBusca=criarDataHoraPorMinutosSistema(dataNorm,0);
      const fimBusca=new Date(inicioBusca.getTime());fimBusca.setDate(fimBusca.getDate()+1);
      cal.getEvents(inicioBusca,fimBusca).forEach(ev=>{if((ev.getDescription()||'').indexOf('Reserva ID: '+id)!==-1){ev.deleteEvent();calendarRemovido=true;}});
    }
  }catch(err){logSistema('WARNING','CALENDAR_CANCELAMENTO_ERRO',usuarioSessao.ID,err.message)}

  if(email){
    try{
      const dataFormatada=Utilities.formatDate(
        criarDataHoraPorMinutosSistema(normalizarDataSistema(data),720),
        DEFAULT_CONFIG.TIMEZONE,'dd/MM/yyyy'
      );
      const abrirUrl=getWebAppUrlSistema();
      const html=emailBaseSistema('❌ Reserva cancelada','A reserva do laboratório foi cancelada e o horário foi liberado.','<div style="background:#FFF7F7;border:1px solid #F4D5D5;border-radius:12px;padding:16px;line-height:1.8;font-size:13px"><strong>📅 Data:</strong> '+escHtmlEmail(dataFormatada)+'<br><strong>🕘 Horário:</strong> '+escHtmlEmail(horario||'')+'<br><strong>👥 Turma:</strong> '+escHtmlEmail(turma||'')+'</div>',abrirUrl?[{label:'Abrir sistema',url:abrirUrl}]:[]);
      MailApp.sendEmail({to:email,subject:'[ICM T.I] ❌ Reserva de laboratório cancelada',htmlBody:html});
      emailEnviado=true;
    }catch(err){logSistema('WARNING','EMAIL_CANCELAMENTO_ERRO',usuarioSessao.ID,err.message)}
  }

  return{sucesso:true,mensagem:'Reserva cancelada.',calendarRemovido,emailEnviado};
}


// ----------------------------------------------------------------------------
// LOCALIZAÇÃO
// ----------------------------------------------------------------------------

function getLocalizacao() {
  sheetToObjects(SHEET_NAMES.CHROMEBOOKS)
    .filter(c=>String(c.Tipo||'').toUpperCase()==='PROF'&&String(c.Status||'').toUpperCase()==='EM_USO')
    .forEach(c=>atualizarLocalizacaoChromebookProfessor(c.ID,'CONSULTA_LOCALIZACAO'));
  const chromes = sheetToObjects(SHEET_NAMES.CHROMEBOOKS).filter(c => c.Status === 'EM_USO');
  const recursos = sheetToObjects(SHEET_NAMES.RECURSOS).filter(r => r.Status === 'EM_USO');
  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);

  const mapResp = (id) => {
    const u = usuarios.find(us => String(us.ID) === String(id));
    return u ? u.Nome : 'Desconhecido';
  };

  const chromesLoc = chromes.map(c => ({
    id: c.ID, tipo: 'Chromebook', responsavel: mapResp(c.Responsavel_ID),
    local: c.Localizacao, desde: c.Data_Retirada, status: 'Em uso'
  }));
  const recursosLoc = recursos.map(r => ({
    id: r.ID, tipo: r.Tipo, responsavel: mapResp(r.Responsavel_ID),
    local: r.Localizacao, desde: r.Data_Retirada, status: 'Em uso'
  }));

  return { chromebooks: chromesLoc, recursos: recursosLoc, total: chromesLoc.length + recursosLoc.length };
}

// ----------------------------------------------------------------------------
// COMUNICAÇÃO
// ----------------------------------------------------------------------------

function getComunicacaoDados(token) {
  const auth=exigirPermissaoComunicacaoRelatorios(token);
  if(auth.resposta)return auth.resposta;
  const usuarioSessao=auth.usuario;
  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS).filter(u => String(u.Ativo).toUpperCase() === 'SIM');
  return usuarios.map(u => ({ id: u.ID, nome: u.Nome, email: u.Email, tipo: u.Tipo }));
}

function comEnviarMensagem(token,destinatarioId, tipo, assunto, mensagem) {
  const auth=exigirPermissaoComunicacaoRelatorios(token);
  if(auth.resposta)return auth.resposta;
  const usuarioSessao=auth.usuario;
  try {
    const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
    const usuario = usuarios.find(u => String(u.ID) === String(destinatarioId));
    if (!usuario || !usuario.Email) return { sucesso: false, mensagem: 'Destinatário sem e-mail cadastrado.' };
    MailApp.sendEmail({ to: usuario.Email, subject: '[ICM T.I] ' + (assunto || tipo), htmlBody: mensagem });
    logSistema('INFO', 'ENVIO_EMAIL', destinatarioId, 'Mensagem enviada: ' + assunto);
    return { sucesso: true, mensagem: 'Mensagem enviada para ' + usuario.Nome };
  } catch (err) {
    logSistema('ERROR', 'ENVIO_EMAIL_ERRO', destinatarioId, err.message);
    return { sucesso: false, mensagem: 'Erro ao enviar: ' + err.message };
  }
}

function comEnviarMassa(token,tipo, assunto, mensagem) {
  const auth=exigirPermissaoComunicacaoRelatorios(token);
  if(auth.resposta)return auth.resposta;
  const usuarioSessao=auth.usuario;
  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS).filter(u => String(u.Ativo).toUpperCase() === 'SIM');
  let enviados = 0, falhas = 0, semEmail = 0;
  usuarios.forEach(u => {
    if (!u.Email) { semEmail++; return; }
    try { MailApp.sendEmail({ to: u.Email, subject: '[ICM T.I] ' + (assunto || tipo), htmlBody: mensagem }); enviados++; }
    catch (err) { falhas++; logSistema('WARNING', 'ENVIO_MASSA_ERRO', u.ID, err.message); }
  });
  logSistema('INFO', 'ENVIO_MASSA', '', 'Enviado para ' + enviados + ' usuários. Falhas: ' + falhas + '. Sem e-mail: ' + semEmail + '. Assunto: ' + assunto);
  return { sucesso: falhas === 0, mensagem: 'Mensagem enviada para ' + enviados + ' destinatário(s).' + (falhas ? ' Falhas: ' + falhas + '.' : '') + (semEmail ? ' Sem e-mail: ' + semEmail + '.' : ''), enviados, falhas, semEmail };
}

// ----------------------------------------------------------------------------
// RELATÓRIOS
// ----------------------------------------------------------------------------

function getRelatoriosDados() {
  return sheetToObjects(SHEET_NAMES.USUARIOS).map(u => ({ id: u.ID, nome: u.Nome, tipo: u.Tipo }));
}

function gerarRelatorio(token,dados) {
  const auth=exigirPermissaoComunicacaoRelatorios(token);
  if(auth.resposta)return auth.resposta;
  const usuarioSessao=auth.usuario;
  const tipo = dados.tipo || 'CHROMEBOOK';
  if (tipo === 'CHROMEBOOK') return gerarRelatorioChromebooks(dados);
  if (tipo === 'LABORATORIO') return gerarRelatorioLaboratorio(dados);
  if (tipo === 'RECURSO') return gerarRelatorioRecursos(dados);
  return { sucesso: false, mensagem: 'Tipo de relatório inválido.' };
}

function gerarRelatorioChromebooks(dados) {
  const logs = sheetToObjects(SHEET_NAMES.LOG_CHROMEBOOKS);
  const filtrados = filtrarPorPeriodo(logs, 'Data_Hora', dados.inicio, dados.fim);
  const retiradas = filtrados.filter(l => l.Acao === 'RETIRADA').length;
  const devolucoes = filtrados.filter(l => l.Acao === 'DEVOLUCAO').length;
  const porUsuario = {};
  filtrados.forEach(l => { porUsuario[l.Usuario_Nome] = (porUsuario[l.Usuario_Nome] || 0) + 1; });
  return {
    sucesso: true, tipo: 'CHROMEBOOK', totalOperacoes: filtrados.length,
    retiradas: retiradas, devolucoes: devolucoes,
    topUsuarios: Object.entries(porUsuario).sort((a, b) => b[1] - a[1]).slice(0, 5)
  };
}

function gerarRelatorioLaboratorio(dados) {
  const reservas = sheetToObjects(SHEET_NAMES.RESERVAS_LABORATORIO);
  const filtradas = filtrarPorPeriodo(reservas, 'Data', dados.inicio, dados.fim);
  const porTurma = {};
  filtradas.forEach(r => { porTurma[r.Turma] = (porTurma[r.Turma] || 0) + 1; });
  return {
    sucesso: true, tipo: 'LABORATORIO', totalReservas: filtradas.length,
    ativas: filtradas.filter(r => r.Status === 'ATIVA').length,
    canceladas: filtradas.filter(r => r.Status === 'CANCELADA').length,
    porTurma: porTurma
  };
}

function gerarRelatorioRecursos(dados) {
  const logs = sheetToObjects(SHEET_NAMES.LOG_RECURSOS);
  const filtrados = filtrarPorPeriodo(logs, 'Data_Hora', dados.inicio, dados.fim);
  const porTipo = {};
  filtrados.forEach(l => { porTipo[l.Tipo] = (porTipo[l.Tipo] || 0) + 1; });
  return { sucesso: true, tipo: 'RECURSO', totalOperacoes: filtrados.length, porTipo: porTipo };
}

function filtrarPorPeriodo(lista,campoData,inicio,fim){
  if(!inicio||!fim)return lista;
  const inicioNorm=normalizarDataSistema(inicio);
  const fimNorm=normalizarDataSistema(fim);
  if(!inicioNorm||!fimNorm)return lista;
  return lista.filter(item=>{
    const d=normalizarDataSistema(item[campoData]);
    return d&&d>=inicioNorm&&d<=fimNorm;
  });
}


// ----------------------------------------------------------------------------
// PERFIL
// ----------------------------------------------------------------------------

function getPerfil(token,usuarioId){
  const usuarioSessao=exigirSessao(token);
  if(!usuarioSessao)return respostaSessaoExpirada();

  const alvo=usuarioId||usuarioSessao.ID;
  if(!perfilPodeAdministrar(usuarioSessao)&&String(alvo)!==String(usuarioSessao.ID))return respostaSemPermissao();

  const usuarios=sheetToObjects(SHEET_NAMES.USUARIOS);
  const usuario=usuarios.find(u=>String(u.ID)===String(alvo));
  if(!usuario)return{sucesso:false,mensagem:'Usuário não encontrado.'};

  const logsChrome=sheetToObjects(SHEET_NAMES.LOG_CHROMEBOOKS).filter(l=>String(l.Usuario_ID)===String(alvo));
  const agendamentos=sheetToObjects(SHEET_NAMES.AGENDAMENTOS).filter(a=>String(a.Usuario_ID)===String(alvo));

  return serializarParaCliente({
    sucesso:true,usuario,
    estatisticas:{
      emprestimos:logsChrome.filter(l=>l.Acao==='RETIRADA').length,
      agendamentos:agendamentos.length,
      avisosEnviados:0,visualizacoes:0
    },
    historico:logsChrome.slice(-20).reverse()
  });
}



function salvarFotoPerfil(usuarioId, url) {
  const row = findRowById(SHEET_NAMES.USUARIOS, 'ID', usuarioId);
  if (!row) return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  updateCell(SHEET_NAMES.USUARIOS, row.rowIndex, 'Foto_URL', url);
  return { sucesso: true, mensagem: 'Foto atualizada.' };
}

function extrairIdArquivoDrive(ref){
  const texto=String(ref||'').trim();
  if(!texto)return'';
  const m=texto.match(/[-\w]{20,}/);
  return m?m[0]:texto;
}

function urlFotoDrivePorId(fileId){
  return 'https://drive.google.com/thumbnail?id='+encodeURIComponent(fileId)+'&sz=w500-h500';
}

function salvarFotoPerfilDrive(token,usuarioId,ref){
  const usuarioSessao=exigirProprioOuAdmin(token,usuarioId);
  if(!usuarioSessao)return getUsuarioLogado(token)?respostaSemPermissao():respostaSessaoExpirada();

  try{
    const fileId=extrairIdArquivoDrive(ref);
    if(!fileId)return{sucesso:false,mensagem:'Informe um link ou ID válido do Google Drive.'};
    const file=DriveApp.getFileById(fileId);
    const mime=String(file.getMimeType()||'').toLowerCase();
    if(!mime.startsWith('image/'))return{sucesso:false,mensagem:'O arquivo selecionado no Drive não é uma imagem.'};

    const url=urlFotoDrivePorId(fileId);
    const row=findRowById(SHEET_NAMES.USUARIOS,'ID',usuarioId);
    if(!row)return{sucesso:false,mensagem:'Usuário não encontrado.'};
    updateCell(SHEET_NAMES.USUARIOS,row.rowIndex,'Foto_URL',url);
    return{sucesso:true,mensagem:'Foto do Google Drive vinculada ao perfil.',url};
  }catch(err){
    logSistema('ERROR','FOTO_DRIVE_ERRO',usuarioId,err.message);
    return{sucesso:false,mensagem:'Não foi possível acessar essa imagem no Google Drive: '+err.message};
  }
}

function salvarFotoPerfilBase64(token,usuarioId,nomeArquivo,mimeType,base64){
  const usuarioSessao=exigirProprioOuAdmin(token,usuarioId);
  if(!usuarioSessao)return getUsuarioLogado(token)?respostaSemPermissao():respostaSessaoExpirada();

  try{
    if(!base64)return{sucesso:false,mensagem:'Imagem não recebida.'};
    const mime=String(mimeType||'image/jpeg').toLowerCase();
    if(!mime.startsWith('image/'))return{sucesso:false,mensagem:'O arquivo enviado não é uma imagem.'};

    const bytes=Utilities.base64Decode(base64);
    if(bytes.length>2*1024*1024)return{sucesso:false,mensagem:'A imagem excede o limite de 2 MB.'};

    const nomeSeguro=String(nomeArquivo||'foto-perfil.jpg').replace(/[\\\/:*?"<>|]/g,'_').slice(0,100);
    const pastas=DriveApp.getFoldersByName('ICM T.I - Fotos de Perfil');
    const pasta=pastas.hasNext()?pastas.next():DriveApp.createFolder('ICM T.I - Fotos de Perfil');
    const arquivo=pasta.createFile(Utilities.newBlob(bytes,mime,nomeSeguro));
    const url=urlFotoDrivePorId(arquivo.getId());

    const row=findRowById(SHEET_NAMES.USUARIOS,'ID',usuarioId);
    if(!row)return{sucesso:false,mensagem:'Usuário não encontrado.'};
    updateCell(SHEET_NAMES.USUARIOS,row.rowIndex,'Foto_URL',url);

    return{sucesso:true,mensagem:'Foto atualizada com sucesso.',url,fileId:arquivo.getId()};
  }catch(err){
    logSistema('ERROR','FOTO_UPLOAD_ERRO',usuarioId,err.message);
    return{sucesso:false,mensagem:'Não foi possível salvar a foto: '+err.message};
  }
}

function atualizarPerfil(token, dados) {
  const usuarioSessao = exigirProprioOuAdmin(token, dados && dados.id);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  const row = findRowById(SHEET_NAMES.USUARIOS, 'ID', dados.id);
  if (!row) return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  updateRowByFields(SHEET_NAMES.USUARIOS, row.rowIndex, {
    Nome: dados.nome, Email: dados.email
  });
  return { sucesso: true, mensagem: 'Perfil atualizado com sucesso.' };
}

// ----------------------------------------------------------------------------
// ADMINISTRAÇÃO
// ----------------------------------------------------------------------------

function listarUsuariosAdmin(token) {
  const usuarioSessao = exigirAdmin(token);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();

  const usuarios = sheetToObjects(SHEET_NAMES.USUARIOS);
  const administradores = usuarios.filter(u =>
    ['admin','administrador'].includes(String(u.Tipo || '').trim().toLowerCase())
  );

  return serializarParaCliente({
    sucesso:true,
    usuarios,
    resumo:{
      total:usuarios.length,
      ativos:usuarios.filter(u=>String(u.Ativo||'').toUpperCase()==='SIM').length,
      administradores:administradores.length
    }
  });
}

function cadastrarUsuario(token, dados) {
  try {
    const usuarioSessao = exigirAdmin(token);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const id = genId('USR');
    let codigo = dados.codigo;
    if (!codigo) {
      codigo = String(Math.floor(1000 + Math.random() * 9000));
    }
    appendObject(SHEET_NAMES.USUARIOS, {
      ID: id, Nome: dados.nome, Email: dados.email, Codigo_4_Digitos: codigo,
      Tipo: dados.tipo || 'professor', Turno: dados.turno || 'Ambos', Ativo: 'SIM',
      Chromebook_Fixo_ID: dados.chromebookFixoId || '', Foto_URL: '', Data_Cadastro: nowStr()
    });
    return { sucesso: true, mensagem: 'Usuário cadastrado com sucesso.', id: id, codigo: codigo };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function editarUsuario(token, dados) {
  const usuarioSessao = exigirAdmin(token);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  const row = findRowById(SHEET_NAMES.USUARIOS, 'ID', dados.id);
  if (!row) return { sucesso: false, mensagem: 'Usuário não encontrado.' };
  const fields = {};
  ['Nome', 'Email', 'Codigo_4_Digitos', 'Tipo', 'Turno', 'Ativo', 'Chromebook_Fixo_ID'].forEach(f => {
    if (dados[f] !== undefined) fields[f] = dados[f];
  });
  updateRowByFields(SHEET_NAMES.USUARIOS, row.rowIndex, fields);
  return { sucesso: true, mensagem: 'Usuário atualizado com sucesso.' };
}

function listarChromebooksAdmin(token) {
  const usuarioSessao = exigirAdmin(token);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  return sheetToObjects(SHEET_NAMES.CHROMEBOOKS);
}

function cadastrarChromebook(token, dados) {
  try {
    const usuarioSessao = exigirAdmin(token);
    if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
    const id = dados.id || genId('CHR');
    appendObject(SHEET_NAMES.CHROMEBOOKS, {
      ID: id, Tipo: dados.tipo || 'ALUNO', Status: 'DISPONIVEL', Responsavel_ID: '',
      Data_Retirada: '', Localizacao: dados.localizacao || '', Observacoes: dados.observacoes || '',
      Avaria_Desc: '', Avaria_Data: '', Avaria_Status: ''
    });
    return { sucesso: true, mensagem: 'Chromebook cadastrado com sucesso.', id: id };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function getConfiguracoesAdmin(token) {
  const usuarioSessao = exigirAdmin(token);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  return sheetToObjects(SHEET_NAMES.CONFIGURACOES);
}

function salvarConfiguracao(token, chave, valor) {
  const usuarioSessao = exigirAdmin(token);
  if (!usuarioSessao) return getUsuarioLogado(token) ? respostaSemPermissao() : respostaSessaoExpirada();
  const row = findRowById(SHEET_NAMES.CONFIGURACOES, 'Configuracao', chave);
  if (row) {
    updateCell(SHEET_NAMES.CONFIGURACOES, row.rowIndex, 'Valor', valor);
  } else {
    appendObject(SHEET_NAMES.CONFIGURACOES, { Configuracao: chave, Valor: valor });
  }
  return { sucesso: true, mensagem: 'Configuração salva.' };
}

// ----------------------------------------------------------------------------
// MANUTENÇÃO / AVARIAS
// ----------------------------------------------------------------------------

function registrarAvaria(equipamentoId, tipo, usuarioId, descricao) {
  appendObject(SHEET_NAMES.AVARIAS, {
    ID: genId('AVR'), Equipamento_ID: equipamentoId, Tipo: tipo, Data: nowStr(),
    Usuario_ID: usuarioId, Descricao: descricao, Status: 'PENDENTE',
    Data_Conserto: '', Tecnico: '', Observacoes: ''
  });
}

function listarManutencoes() {
  return sheetToObjects(SHEET_NAMES.MANUTENCAO).filter(m => m.Status !== 'CONCLUIDO');
}

function registrarManutencao(dados) {
  try {
    const id = genId('MNT');
    appendObject(SHEET_NAMES.MANUTENCAO, {
      ID: id, Equipamento_ID: dados.equipamentoId, Tipo: dados.tipo,
      Data_Entrada: nowStr(), Data_Saida: '', Status: 'AGUARDANDO',
      Tecnico: dados.tecnico || '', Observacoes: dados.observacoes || ''
    });
    // Atualiza status na origem (Chromebooks ou Recursos)
    const sheetOrigem = dados.tipo === 'CHROMEBOOK' ? SHEET_NAMES.CHROMEBOOKS : SHEET_NAMES.RECURSOS;
    const row = findRowById(sheetOrigem, 'ID', dados.equipamentoId);
    if (row) updateCell(sheetOrigem, row.rowIndex, 'Status', 'MANUTENCAO');
    return { sucesso: true, mensagem: 'Manutenção registrada.', id: id };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function concluirManutencao(id, dataSaida, observacoes) {
  try {
    const row = findRowById(SHEET_NAMES.MANUTENCAO, 'ID', id);
    if (!row) return { sucesso: false, mensagem: 'Manutenção não encontrada.' };
    const equipIdx = row.headers.indexOf('Equipamento_ID');
    const tipoIdx = row.headers.indexOf('Tipo');
    const equipamentoId = row.values[equipIdx];
    const tipo = row.values[tipoIdx];

    updateRowByFields(SHEET_NAMES.MANUTENCAO, row.rowIndex, {
      Status: 'CONCLUIDO', Data_Saida: dataSaida || nowStr(), Observacoes: observacoes || ''
    });

    const sheetOrigem = tipo === 'CHROMEBOOK' ? SHEET_NAMES.CHROMEBOOKS : SHEET_NAMES.RECURSOS;
    const rowOrigem = findRowById(sheetOrigem, 'ID', equipamentoId);
    if (rowOrigem) updateCell(sheetOrigem, rowOrigem.rowIndex, 'Status', 'DISPONIVEL');

    return { sucesso: true, mensagem: 'Manutenção concluída, equipamento disponível novamente.' };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

// ----------------------------------------------------------------------------
// TOKENS PARA ALUNOS (retirada sem código do professor)
// ----------------------------------------------------------------------------

function gerarTokenAula(professorId, qtdLiberada) {
  try {
    const token = Utilities.getUuid().substring(0, 4).toUpperCase();
    const agora = new Date();
    const expiracao = new Date(agora.getTime() + 20 * 60000);
    appendObject(SHEET_NAMES.TOKENS_ALUNOS, {
      Token: token, Professor_ID: professorId, Quantidade_Liberada: qtdLiberada,
      Data_Criacao: nowStr(), Expiracao: Utilities.formatDate(expiracao, DEFAULT_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      Retirados: 0, Status: 'ATIVO'
    });
    return { sucesso: true, token: token, expiracao: expiracao };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function validarTokenAluno(token) {
  const row = findRowById(SHEET_NAMES.TOKENS_ALUNOS, 'Token', token);
  if (!row) return { sucesso: false, mensagem: 'Token inválido.' };
  const statusIdx = row.headers.indexOf('Status');
  const expIdx = row.headers.indexOf('Expiracao');
  if (row.values[statusIdx] !== 'ATIVO') return { sucesso: false, mensagem: 'Token expirado ou já utilizado.' };
  const expiracao = new Date(row.values[expIdx]);
  if (new Date() > expiracao) {
    updateCell(SHEET_NAMES.TOKENS_ALUNOS, row.rowIndex, 'Status', 'EXPIRADO');
    return { sucesso: false, mensagem: 'Token expirado.' };
  }
  const qtdIdx = row.headers.indexOf('Quantidade_Liberada');
  const retIdx = row.headers.indexOf('Retirados');
  return {
    sucesso: true, quantidadeLiberada: row.values[qtdIdx], retirados: row.values[retIdx],
    disponivel: row.values[qtdIdx] - row.values[retIdx]
  };
}

function alunoRetirarToken(token, quantidade) {
  try {
    const validacao = validarTokenAluno(token);
    if (!validacao.sucesso) return validacao;
    if (quantidade > validacao.disponivel) {
      return { sucesso: false, mensagem: 'Quantidade solicitada excede o limite liberado pelo professor.' };
    }

    const disponiveis = getChromesDisponiveis('ALUNO').slice(0, quantidade);
    if (disponiveis.length < quantidade) {
      return { sucesso: false, mensagem: 'Não há Chromebooks suficientes disponíveis.' };
    }

    const row = findRowById(SHEET_NAMES.TOKENS_ALUNOS, 'Token', token);
    const professorId = row.values[row.headers.indexOf('Professor_ID')];
    const ids = disponiveis.map(c => c.ID);
    ids.forEach(id => {
      const chromeRow = findRowById(SHEET_NAMES.CHROMEBOOKS, 'ID', id);
      updateRowByFields(SHEET_NAMES.CHROMEBOOKS, chromeRow.rowIndex, {
        Status: 'EM_USO', Responsavel_ID: professorId, Data_Retirada: nowStr()
      });
    });

    updateCell(SHEET_NAMES.TOKENS_ALUNOS, row.rowIndex, 'Retirados', Number(validacao.retirados) + quantidade);
    registrarLogChrome(professorId, 'RETIRADA', quantidade, ids, 'ALUNO_TOKEN');
    const profToken=sheetToObjects(SHEET_NAMES.USUARIOS).find(u=>String(u.ID)===String(professorId));
    ids.forEach(id=>registrarAuditoriaChromebook({acao:'RETIRADA',chromebookId:id,professorId:professorId,professorNome:profToken?profToken.Nome:'',tipoUso:'ALUNO',contexto:'RETIRADA_TOKEN',turno:profToken?profToken.Turno:'',diaSemana:diaSemanaNomeSistema(new Date()),origem:'TOKEN_ALUNO',status:'SUCESSO',observacao:'Retirada realizada por token da aula.'}));

    return { sucesso: true, mensagem: quantidade + ' Chromebook(s) retirado(s) via token.', ids: ids };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro: ' + err.message };
  }
}

function invalidarTokensProfessor(professorId) {
  const sh = getSheet(SHEET_NAMES.TOKENS_ALUNOS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const profCol = headers.indexOf('Professor_ID');
  const statusCol = headers.indexOf('Status');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][profCol]) === String(professorId) && values[i][statusCol] === 'ATIVO') {
      sh.getRange(i + 1, statusCol + 1).setValue('EXPIRADO');
    }
  }
}

// ----------------------------------------------------------------------------
// FIM DO ARQUIVO
// ----------------------------------------------------------------------------
