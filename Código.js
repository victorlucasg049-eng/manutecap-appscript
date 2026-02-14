// ==================================================
//   SISTEMA DE GESTÃO DE MANUTENÇÃO HOTELEIRA
//   Web App – Google Apps Script (Versão 2.0)
// ==================================================

const CONFIG = {
  SHEET_ID: '1DOGQZg6AxQyen4_P-ZjWc17owU2S7OEM--Am8Qmp1lU',
  SHEET_NAME: 'OrdensServico',
  PREVENTIVA_SHEET_NAME: 'Preventivas',
  DRIVE_FOLDER_NAME: 'ManutencaoHotel',
  STATUS: {
    ABERTO: 'Aberto',
    EM_ANDAMENTO: 'Em Andamento',
    CONCLUIDO: 'Concluído',
    CANCELADO: 'Cancelado'
  },
  FREQUENCIAS: ['Semanal', 'Quinzenal', 'Mensal', 'Bimestral', 'Semestral', 'Anual']
};

// ---------- WEB APP ----------
function doGet(e) {
  const page = e?.parameter?.page || 'dashboard';
  const baseUrl = ScriptApp.getService().getUrl();

  let template;
  switch (page) {
    case 'recepcao':    template = HtmlService.createTemplateFromFile('PainelRecepcao'); break;
    case 'manutencao':  template = HtmlService.createTemplateFromFile('PainelManutencao'); break;
    case 'preventiva':  template = HtmlService.createTemplateFromFile('PainelPreventiva'); break;
    case 'relatorios':  template = HtmlService.createTemplateFromFile('Relatorios'); break;
    case 'configuracoes': template = HtmlService.createTemplateFromFile('Configuracoes'); break;
    case 'diagnostico': template = HtmlService.createTemplateFromFile('Diagnostico'); break;
    case 'teste':       template = HtmlService.createTemplateFromFile('TesteComunicacao'); break;
    default:            template = HtmlService.createTemplateFromFile('index'); break;
  }
  template.baseUrl = baseUrl;
  template.CONFIG = CONFIG;
  return template.evaluate()
    .setTitle('Gestão de Manutenção Hoteleira')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- ACESSO À PLANILHA ----------
function getSpreadsheetId() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('SPREADSHEET_ID') || CONFIG.SHEET_ID;
}

function setSpreadsheetId(id) {
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  return { success: true };
}

function getSpreadsheet() {
  const id = getSpreadsheetId();
  if (!id) throw new Error('ID da planilha não configurado. Acesse Configurações.');
  return SpreadsheetApp.openById(id);
}

// ---------- CRIAÇÃO DAS ABAS (PADRÃO) ----------
function criarPlanilhaOrdens(ss) {
  const sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  const headers = [
    'ID', 'Data Abertura', 'Setor', 'Local', 'Descrição do Problema',
    'Prioridade', 'Solicitante', 'URL Foto Avaria', 'Status',
    'Data Conclusão', 'Técnico Responsável', 'Descrição do Reparo',
    'URL Foto Conclusão', 'Materiais Utilizados', 'Tempo Gasto (horas)',
    'Custo Estimado (R$)', 'Checklist (JSON)', 'Observações', 'Última Atualização'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#1e40af').setFontColor('white').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(12, 280);
  return sheet;
}

function criarPlanilhaPreventivas(ss) {
  const sheet = ss.insertSheet(CONFIG.PREVENTIVA_SHEET_NAME);
  const headers = [
    'ID', 'Setor', 'Local', 'Descrição', 'Prioridade', 'Frequência',
    'Próxima Data', 'Última Execução', 'Checklist Modelo (JSON)',
    'Ativo', 'Criado em', 'Observações'
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#c2410c').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

// ---------- CLASSE BASE (SHEET) ----------
class SheetBase {
  constructor(sheetName, expectedHeaders) {
    this.sheetName = sheetName;
    this.expectedHeaders = expectedHeaders;
    this._headerRow = null;
    this._headerIndex = {};
    this._allData = null;
    this._lastLoad = 0;
    this.cacheTTL = 5000;
  }

  getSheet() {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(this.sheetName);
    if (!sheet) sheet = this.criarAba(ss);
    return sheet;
  }

  criarAba(ss) {
    throw new Error('Método criarAba deve ser implementado na subclasse');
  }

  verificarEstrutura() {
    const sheet = this.getSheet();
    const lastCol = sheet.getLastColumn();
    // Aba existe mas está completamente vazia (sem nenhuma coluna) → só adiciona cabeçalho
    if (lastCol === 0) {
      const expected = this.expectedHeaders;
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
      sheet.setFrozenRows(1);
      console.log(`[verificarEstrutura] Cabeçalhos adicionados à aba vazia: ${this.sheetName}`);
      this.invalidateCache();
    }
    // NUNCA apaga dados existentes — apenas loga aviso se cabeçalhos divergirem
    else {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const expected = this.expectedHeaders;
      const ok = expected.every((h, idx) => headers[idx] === h);
      if (!ok) {
        console.warn(`[verificarEstrutura] Aviso: cabeçalhos da aba "${this.sheetName}" diferem do esperado. Os dados serão lidos assim mesmo.`);
      }
    }
  }

  // Só cria a aba se ela não existir — nunca apaga dados
  recriarAba(motivo) {
    console.log(`[recriarAba] ${this.sheetName}: ${motivo}`);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(this.sheetName);
    if (!sheet) {
      this.criarAba(ss);
      this.invalidateCache();
    } else {
      console.warn(`[recriarAba] Aba "${this.sheetName}" já existe — não será apagada.`);
    }
  }

  loadHeaders() {
    if (this._headerRow) return;
    this.verificarEstrutura();
    const sheet = this.getSheet();
    const lastCol = sheet.getLastColumn();
    this._headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    this._headerIndex = {};
    this._headerRow.forEach((name, idx) => {
      if (name) this._headerIndex[name] = idx + 1;
    });
  }

  col(name) {
    this.loadHeaders();
    const col = this._headerIndex[name];
    if (!col) throw new Error(`Coluna '${name}' não encontrada na aba ${this.sheetName}.`);
    return col;
  }

  getAllData(forceRefresh = false) {
    const now = Date.now();
    if (!forceRefresh && this._allData && (now - this._lastLoad < this.cacheTTL)) return this._allData;
    this.loadHeaders();
    const sheet = this.getSheet();
    const numRows = sheet.getLastRow();
    const numCols = sheet.getLastColumn();
    if (numRows <= 1) {
      this._allData = { headers: this._headerRow, rows: [] };
    } else {
      const values = sheet.getRange(1, 1, numRows, numCols).getValues();
      this._allData = { headers: values[0], rows: values.slice(1) };
    }
    this._lastLoad = now;
    return this._allData;
  }

  invalidateCache() {
    this._allData = null;
    this._lastLoad = 0;
    this._headerRow = null;
    this._headerIndex = {};
  }

  rowToObject(row, headers = null) {
    const obj = {};
    const h = headers || this._headerRow;
    h.forEach((name, idx) => {
      if (!name) return;
      let val = row[idx];
      // Converte Date → string formatada (Apps Script retorna Date para células de data)
      if (val instanceof Date) {
        if (isNaN(val.getTime()) || val.getFullYear() < 1970) {
          val = '';
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        }
      }
      // Garante que o valor é primitivo serializável
      obj[name] = (val === null || val === undefined) ? '' : val;
    });
    return obj;
  }

  objectToRow(obj) {
    this.loadHeaders();
    return this._headerRow.map(name => (obj[name] !== undefined ? obj[name] : ''));
  }

  findRowById(id) {
    const data = this.getAllData();
    for (let i = 0; i < data.rows.length; i++) {
      if (String(data.rows[i][0]) === String(id)) {
        return { rowIndex: i + 2, rowData: data.rows[i] };
      }
    }
    return null;
  }

  insertRow(rowObject) {
    const sheet = this.getSheet();
    const rowArray = this.objectToRow(rowObject);
    sheet.appendRow(rowArray);
    this.invalidateCache();
    const newId = rowArray[0];
    return this.findRowById(newId);
  }

  updateFields(rowIndex, fields) {
    this.loadHeaders();
    const sheet = this.getSheet();
    for (const [name, value] of Object.entries(fields)) {
      const col = this._headerIndex[name];
      if (col) sheet.getRange(rowIndex, col).setValue(value);
    }
    this.invalidateCache();
    return true;
  }
}

// ---------- ORDENS DE SERVIÇO ----------
class OrdensSheet extends SheetBase {
  constructor() {
    super(CONFIG.SHEET_NAME, [
      'ID', 'Data Abertura', 'Setor', 'Local', 'Descrição do Problema',
      'Prioridade', 'Solicitante', 'URL Foto Avaria', 'Status',
      'Data Conclusão', 'Técnico Responsável', 'Descrição do Reparo',
      'URL Foto Conclusão', 'Materiais Utilizados', 'Tempo Gasto (horas)',
      'Custo Estimado (R$)', 'Checklist (JSON)', 'Observações', 'Última Atualização'
    ]);
  }

  criarAba(ss) { return criarPlanilhaOrdens(ss); }

  generateId() {
    return `OS-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}-${Math.floor(Math.random() * 9000 + 1000)}`;
  }

  getPendentes() {
    try {
      const data = this.getAllData();
      const statusCol = this.col('Status') - 1;
      const prioridadePeso = { 'Urgente': 1, 'Alta': 2, 'Média': 3, 'Baixa': 4 };
      return data.rows
        .filter(row => {
          const s = String(row[statusCol] || '').trim();
          return s === CONFIG.STATUS.ABERTO || s === CONFIG.STATUS.EM_ANDAMENTO;
        })
        .map(row => this.rowToObject(row, data.headers))
        .sort((a, b) => {
          const pa = prioridadePeso[a.Prioridade] || 3;
          const pb = prioridadePeso[b.Prioridade] || 3;
          if (pa !== pb) return pa - pb;
          const da = parseDateBR(a['Data Abertura']);
          const db = parseDateBR(b['Data Abertura']);
          return db - da;
        });
    } catch (e) {
      console.error('Erro em getPendentes:', e);
      return [];
    }
  }

  getHistorico() {
    try {
      const data = this.getAllData();
      const statusCol = this.col('Status') - 1;
      return data.rows
        .filter(row => {
          const s = String(row[statusCol] || '').trim();
          return s === CONFIG.STATUS.CONCLUIDO || s === CONFIG.STATUS.CANCELADO;
        })
        .map(row => this.rowToObject(row, data.headers))
        .sort((a, b) => {
          const da = parseDateBR(a['Data Conclusão'] || a['Data Abertura']);
          const db = parseDateBR(b['Data Conclusão'] || b['Data Abertura']);
          return db - da;
        });
    } catch (e) {
      console.error('Erro em getHistorico:', e);
      return [];
    }
  }

  getAll() {
    try {
      const data = this.getAllData();
      return data.rows.map(row => this.rowToObject(row, data.headers));
    } catch (e) {
      console.error('Erro em getAll:', e);
      return [];
    }
  }

  create(dados) {
    const agora = new Date();
    const dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const id = this.generateId();
    const novaOS = {
      'ID': id,
      'Data Abertura': dataFormatada,
      'Setor': dados.setor || '',
      'Local': dados.local || '',
      'Descrição do Problema': dados.descricao || '',
      'Prioridade': dados.prioridade || 'Média',
      'Solicitante': dados.solicitante || '',
      'URL Foto Avaria': dados.fotoUrl || '',
      'Status': CONFIG.STATUS.ABERTO,
      'Data Conclusão': '',
      'Técnico Responsável': '',
      'Descrição do Reparo': '',
      'URL Foto Conclusão': '',
      'Materiais Utilizados': '',
      'Tempo Gasto (horas)': '',
      'Custo Estimado (R$)': '',
      'Checklist (JSON)': dados.checklistJSON || '[]',
      'Observações': dados.observacoes || '',
      'Última Atualização': dataFormatada
    };
    this.insertRow(novaOS);
    return { success: true, id, message: `Ordem ${id} criada com sucesso!` };
  }

  update(ordemId, campos) {
    const found = this.findRowById(ordemId);
    if (!found) throw new Error(`Ordem '${ordemId}' não encontrada`);
    campos['Última Atualização'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    this.updateFields(found.rowIndex, campos);
    return { success: true, message: `Ordem ${ordemId} atualizada.` };
  }
}

// ---------- PREVENTIVAS ----------
class PreventivasSheet extends SheetBase {
  constructor() {
    super(CONFIG.PREVENTIVA_SHEET_NAME, [
      'ID', 'Setor', 'Local', 'Descrição', 'Prioridade', 'Frequência',
      'Próxima Data', 'Última Execução', 'Checklist Modelo (JSON)',
      'Ativo', 'Criado em', 'Observações'
    ]);
  }

  criarAba(ss) { return criarPlanilhaPreventivas(ss); }

  generateId() {
    return `PREV-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}-${Math.floor(Math.random() * 9000 + 1000)}`;
  }

  create(dados) {
    const id = this.generateId();
    const dataCriacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const proximaData = calcularProximaData(dados.frequencia, dados.dataInicio);
    const nova = {
      'ID': id,
      'Setor': dados.setor,
      'Local': dados.local,
      'Descrição': dados.descricao,
      'Prioridade': dados.prioridade || 'Média',
      'Frequência': dados.frequencia,
      'Próxima Data': proximaData,
      'Última Execução': '',
      'Checklist Modelo (JSON)': JSON.stringify(dados.checklist || []),
      'Ativo': 'Sim',
      'Criado em': dataCriacao,
      'Observações': dados.observacoes || ''
    };
    this.insertRow(nova);
    return { success: true, id, message: 'Preventiva agendada com sucesso!' };
  }

  getAtivas() {
    try {
      const data = this.getAllData();
      const ativoCol = this.col('Ativo') - 1;
      return data.rows
        .filter(row => row[ativoCol] === 'Sim')
        .map(row => this.rowToObject(row, data.headers));
    } catch (e) {
      console.error('Erro em getAtivas:', e);
      return [];
    }
  }

  getById(id) {
    try {
      const found = this.findRowById(id);
      if (!found) return null;
      const data = this.getAllData();
      return this.rowToObject(found.rowData, data.headers);
    } catch (e) {
      console.error('Erro em getById:', e);
      return null;
    }
  }

  desativar(id) {
    const found = this.findRowById(id);
    if (!found) throw new Error('Preventiva não encontrada');
    this.updateFields(found.rowIndex, { 'Ativo': 'Não' });
    return { success: true, message: 'Preventiva desativada.' };
  }

  registrarExecucao(prevId, proximaData) {
    const found = this.findRowById(prevId);
    if (!found) throw new Error('Preventiva não encontrada');
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    this.updateFields(found.rowIndex, {
      'Última Execução': hoje,
      'Próxima Data': proximaData
    });
  }
}

// ---------- HELPER DE DATA ----------
function parseDateBR(dateStr) {
  if (!dateStr) return new Date(0);
  try {
    const partes = String(dateStr).split(' ')[0].split('/');
    if (partes.length !== 3) return new Date(0);
    return new Date(partes[2], partes[1] - 1, partes[0]);
  } catch { return new Date(0); }
}

// ---------- Instâncias globais ----------
const ordensSheet = new OrdensSheet();
const preventivasSheet = new PreventivasSheet();

// ---------- FUNÇÕES PÚBLICAS ----------
function getOrdensPendentes() {
  try {
    const pendentes = ordensSheet.getPendentes();
    console.log(`[Backend] getOrdensPendentes: ${pendentes.length} registros`);
    return pendentes;
  } catch (e) {
    console.error('[Backend] ERRO em getOrdensPendentes:', e.message);
    return [];
  }
}

function getOrdensHistorico() {
  try {
    const historico = ordensSheet.getHistorico();
    console.log(`[Backend] getOrdensHistorico: ${historico.length} registros`);
    return historico;
  } catch (e) {
    console.error('[Backend] Erro em getOrdensHistorico:', e.message);
    return [];
  }
}

function listarPreventivasAtivas() {
  try {
    const ativas = preventivasSheet.getAtivas();
    console.log(`[Backend] listarPreventivasAtivas: ${ativas.length} registros`);
    return ativas;
  } catch (e) {
    console.error('[Backend] Erro em listarPreventivasAtivas:', e.message);
    return [];
  }
}

function getTodasOrdens() {
  try { return ordensSheet.getAll(); }
  catch (e) { console.error(e); return []; }
}

function getOrdemById(id) {
  try {
    const found = ordensSheet.findRowById(id);
    if (!found) return null;
    const data = ordensSheet.getAllData();
    return ordensSheet.rowToObject(found.rowData, data.headers);
  } catch (e) { console.error(e); return null; }
}

function criarOrdemServico(dados) {
  try {
    if (dados.fotoBase64 && dados.fotoBase64.length > 100) {
      try { dados.fotoUrl = salvarFoto(dados.fotoBase64, 'temp', 'avaria'); }
      catch (e) { console.warn('Erro ao salvar foto (continuando sem foto):', e.message); }
    }
    const result = ordensSheet.create(dados);
    if (result.success) {
      try { enviarNotificacaoNovaOrdem(result.id, dados); } catch(e) {}
    }
    return result;
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function atualizarOrdem(dados) {
  try {
    const campos = {
      'Status': dados.status,
      'Técnico Responsável': dados.tecnico
    };
    if (dados.status === CONFIG.STATUS.CONCLUIDO) {
      campos['Data Conclusão'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      campos['Descrição do Reparo'] = dados.descricaoReparo || '';
      campos['Materiais Utilizados'] = dados.materiais || '';
      campos['Tempo Gasto (horas)'] = dados.tempoGasto || '';
      campos['Custo Estimado (R$)'] = dados.custo || '';
      if (dados.fotoConclusaoBase64 && dados.fotoConclusaoBase64.length > 100) {
        try { campos['URL Foto Conclusão'] = salvarFoto(dados.fotoConclusaoBase64, dados.id, 'conclusao'); }
        catch (e) { console.warn('Erro ao salvar foto conclusão:', e.message); }
      }
    } else if (dados.status === CONFIG.STATUS.CANCELADO) {
      campos['Data Conclusão'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    }
    if (dados.checklistJSON) campos['Checklist (JSON)'] = dados.checklistJSON;
    return ordensSheet.update(dados.id, campos);
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function salvarPreventiva(dados) {
  try { return preventivasSheet.create(dados); }
  catch (error) { return { success: false, message: error.toString() }; }
}

function getPreventivaById(id) {
  try { return preventivasSheet.getById(id); }
  catch (e) { console.error(e); return null; }
}

function desativarPreventiva(id) {
  try { return preventivasSheet.desativar(id); }
  catch (e) { return { success: false, message: e.toString() }; }
}

function gerarOSPreventiva(preventivaId) {
  try {
    const found = preventivasSheet.findRowById(preventivaId);
    if (!found) throw new Error('Preventiva não encontrada');
    const data = preventivasSheet.getAllData();
    const prevObj = preventivasSheet.rowToObject(found.rowData, data.headers);
    const checklistModelo = JSON.parse(prevObj['Checklist Modelo (JSON)'] || '[]');
    const checklist = checklistModelo.map(item => ({ descricao: item, concluido: false }));
    const dadosOS = {
      setor: prevObj.Setor,
      local: prevObj.Local,
      prioridade: prevObj.Prioridade || 'Média',
      descricao: `[PREVENTIVA] ${prevObj.Descrição}`,
      solicitante: 'Sistema Automático',
      observacoes: `Gerado da preventiva ${preventivaId}. Frequência: ${prevObj.Frequência}`,
      checklistJSON: JSON.stringify(checklist)
    };
    const result = criarOrdemServico(dadosOS);
    if (!result.success) throw new Error(result.message);
    const proximaData = calcularProximaData(prevObj.Frequência, null);
    preventivasSheet.registrarExecucao(preventivaId, proximaData);
    return { success: true, ordemId: result.id };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ---------- DIAGNÓSTICO ----------
function diagnosticarPlanilha() {
  const resultado = {
    spreadsheetId: getSpreadsheetId(),
    spreadsheetNome: '',
    abas: [],
    ordens: { existe: false, headers: [], linhas: 0, ok: false },
    preventivas: { existe: false, headers: [], linhas: 0, ok: false },
    erros: []
  };
  try {
    const ss = getSpreadsheet();
    resultado.spreadsheetNome = ss.getName();
    resultado.abas = ss.getSheets().map(s => s.getName());
    const ordensSheetObj = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (ordensSheetObj) {
      resultado.ordens.existe = true;
      resultado.ordens.linhas = ordensSheetObj.getLastRow();
      if (resultado.ordens.linhas > 0) {
        const headers = ordensSheetObj.getRange(1, 1, 1, ordensSheetObj.getLastColumn()).getValues()[0];
        resultado.ordens.headers = headers;
        resultado.ordens.ok = ['ID', 'Status', 'Prioridade'].every(e => headers.includes(e));
      }
    }
    const prevSheet = ss.getSheetByName(CONFIG.PREVENTIVA_SHEET_NAME);
    if (prevSheet) {
      resultado.preventivas.existe = true;
      resultado.preventivas.linhas = prevSheet.getLastRow();
      if (resultado.preventivas.linhas > 0) {
        const headers = prevSheet.getRange(1, 1, 1, prevSheet.getLastColumn()).getValues()[0];
        resultado.preventivas.headers = headers;
        resultado.preventivas.ok = ['ID', 'Prioridade', 'Frequência'].every(e => headers.includes(e));
      }
    }
  } catch (e) { resultado.erros.push(e.toString()); }
  return resultado;
}

// ---------- GEMINI IA ----------
function gerarChecklistComGemini(setor, local, descricao) {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return {
        success: true,
        checklist: [
          'Inspecionar visualmente o equipamento/área',
          'Verificar funcionamento básico dos componentes',
          'Limpar e higienizar superfícies e componentes',
          'Testar funcionamento após manutenção',
          'Verificar conexões e fixações',
          'Registrar no sistema com foto'
        ],
        message: 'Usando checklist padrão (chave Gemini não configurada)'
      };
    }
    const MODELO = 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODELO}:generateContent?key=${apiKey}`;
    const prompt = `Você é um especialista em manutenção hoteleira. Crie um checklist detalhado de manutenção preventiva.
Setor: ${setor}
Local: ${local}  
Tarefa: ${descricao}

Retorne APENAS uma lista de itens de verificação práticos e específicos, um por linha, iniciando com "- ".
Inclua 6-10 itens relevantes, com verbos de ação claros. Não inclua títulos ou numeração.
Exemplo de formato:
- Verificar funcionamento da TV: ligar e testar todos os canais
- Inspecionar controle remoto: checar pilhas e funcionamento dos botões`;

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, maxOutputTokens: 1000 }
    };
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error(`API Gemini retornou código ${response.getResponseCode()}`);
    }
    const result = JSON.parse(response.getContentText());
    const text = result.candidates[0].content.parts[0].text;
    const linhas = text.split('\n')
      .filter(l => l.trim().startsWith('-'))
      .map(l => l.trim().substring(2).trim())
      .filter(l => l.length > 3);
    const checklist = linhas.length >= 3 ? linhas : ['Inspecionar visualmente', 'Verificar funcionamento', 'Limpar componentes', 'Testar', 'Registrar'];
    return { success: true, checklist };
  } catch (error) {
    console.error('Erro ao chamar Gemini:', error);
    return {
      success: true,
      checklist: ['Inspecionar visualmente', 'Verificar funcionamento', 'Limpar componentes', 'Testar operação', 'Fotografar e registrar'],
      message: 'Usando checklist padrão (IA indisponível)'
    };
  }
}

// ---------- DRIVE ----------
function salvarFoto(base64Data, ordemId, tipo) {
  const pasta = getOrCreateDriveFolder();
  const subPasta = getOrCreateSubFolder(pasta, ordemId);
  const conteudo = base64Data.split(',')[1] || base64Data;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(conteudo),
    'image/jpeg',
    `${tipo}_${ordemId}_${Date.now()}.jpg`
  );
  const arquivo = subPasta.createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return arquivo.getUrl();
}

function getOrCreateDriveFolder() {
  const pastas = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (pastas.hasNext()) return pastas.next();
  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

function getOrCreateSubFolder(pastaPai, nome) {
  const pastas = pastaPai.getFoldersByName(nome);
  if (pastas.hasNext()) return pastas.next();
  return pastaPai.createFolder(nome);
}

// ---------- CONFIGURAÇÕES ----------
function salvarConfiguracoes(config) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (config.spreadsheetId && config.spreadsheetId.trim()) {
      props.setProperty('SPREADSHEET_ID', config.spreadsheetId.trim());
    }
    props.setProperty('EMAIL_NOTIFICATIONS', config.emailNotifications ? 'true' : 'false');
    props.setProperty('EMAILS_EQUIPE', config.emailsEquipe || '');
    if (config.geminiApiKey && config.geminiApiKey.trim()) {
      props.setProperty('GEMINI_API_KEY', config.geminiApiKey.trim());
    }
    return { success: true, message: '✅ Configurações salvas com sucesso!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function carregarConfiguracoes() {
  try {
    const props = PropertiesService.getScriptProperties();
    return {
      spreadsheetId: props.getProperty('SPREADSHEET_ID') || CONFIG.SHEET_ID,
      emailNotifications: props.getProperty('EMAIL_NOTIFICATIONS') === 'true',
      emailsEquipe: props.getProperty('EMAILS_EQUIPE') || '',
      geminiApiKey: props.getProperty('GEMINI_API_KEY') || ''
    };
  } catch (e) { return {}; }
}

// ---------- RELATÓRIOS ----------
function getEstatisticas() {
  try {
    const ordens = getTodasOrdens();
    const stats = {
      total: ordens.length,
      abertos: 0, emAndamento: 0, concluidos: 0, cancelados: 0,
      porSetor: {},
      porPrioridade: { 'Urgente': 0, 'Alta': 0, 'Média': 0, 'Baixa': 0 },
      porMes: {},
      tempoMedioResolucao: 0,
      custoTotal: 0
    };
    let somaTempo = 0, countTempo = 0;
    ordens.forEach(o => {
      if (o.Status === CONFIG.STATUS.ABERTO) stats.abertos++;
      else if (o.Status === CONFIG.STATUS.EM_ANDAMENTO) stats.emAndamento++;
      else if (o.Status === CONFIG.STATUS.CONCLUIDO) stats.concluidos++;
      else if (o.Status === CONFIG.STATUS.CANCELADO) stats.cancelados++;

      const setor = o.Setor || 'Não especificado';
      stats.porSetor[setor] = (stats.porSetor[setor] || 0) + 1;

      const prioridade = o.Prioridade || 'Média';
      if (stats.porPrioridade[prioridade] !== undefined) stats.porPrioridade[prioridade]++;

      const custo = parseFloat(o['Custo Estimado (R$)']) || 0;
      stats.custoTotal += custo;

      // Agrupamento por mês
      if (o['Data Abertura']) {
        const partes = String(o['Data Abertura']).split('/');
        if (partes.length >= 2) {
          const mesAno = `${partes[1]}/${partes[2] ? partes[2].substring(0,4) : ''}`;
          stats.porMes[mesAno] = (stats.porMes[mesAno] || 0) + 1;
        }
      }

      if (o.Status === CONFIG.STATUS.CONCLUIDO && o['Tempo Gasto (horas)']) {
        const tempo = parseFloat(o['Tempo Gasto (horas)']);
        if (!isNaN(tempo)) { somaTempo += tempo; countTempo++; }
      }
    });
    stats.tempoMedioResolucao = countTempo ? parseFloat((somaTempo / countTempo).toFixed(2)) : 0;
    return stats;
  } catch (e) {
    console.error('Erro em getEstatisticas:', e);
    return { total: 0, abertos: 0, emAndamento: 0, concluidos: 0, cancelados: 0, porSetor: {}, porPrioridade: {}, porMes: {}, tempoMedioResolucao: 0, custoTotal: 0 };
  }
}

function exportarParaCSV() {
  try {
    const ordens = getTodasOrdens();
    if (ordens.length === 0) throw new Error('Nenhuma ordem para exportar');
    const cabecalhos = Object.keys(ordens[0]);
    let csv = '\uFEFF' + cabecalhos.join(';') + '\n'; // BOM para UTF-8
    ordens.forEach(o => {
      const linha = cabecalhos.map(h => {
        let val = String(o[h] || '').replace(/"/g, '""');
        if (val.includes(';') || val.includes('\n') || val.includes('"')) val = `"${val}"`;
        return val;
      }).join(';');
      csv += linha + '\n';
    });
    const pasta = getOrCreateDriveFolder();
    const nome = `Ordens_Manutencao_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm')}.csv`;
    const arquivo = pasta.createFile(nome, csv, MimeType.CSV);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: arquivo.getUrl(), message: `Exportado: ${nome}` };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ---------- PREVENTIVAS (AUXILIAR) ----------
function calcularProximaData(frequencia, dataInicioStr) {
  const hoje = new Date();
  let base = dataInicioStr ? new Date(dataInicioStr) : hoje;
  if (isNaN(base.getTime()) || base < hoje) base = hoje;
  const nova = new Date(base);
  switch(frequencia) {
    case 'Semanal':    nova.setDate(nova.getDate() + 7); break;
    case 'Quinzenal':  nova.setDate(nova.getDate() + 14); break;
    case 'Mensal':     nova.setMonth(nova.getMonth() + 1); break;
    case 'Bimestral':  nova.setMonth(nova.getMonth() + 2); break;
    case 'Semestral':  nova.setMonth(nova.getMonth() + 6); break;
    case 'Anual':      nova.setFullYear(nova.getFullYear() + 1); break;
    default:           nova.setMonth(nova.getMonth() + 1); break;
  }
  return Utilities.formatDate(nova, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function verificarPreventivasVencidas() {
  try {
    const ativas = preventivasSheet.getAtivas();
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    ativas.forEach(prev => {
      const proxStr = prev['Próxima Data'];
      if (!proxStr) return;
      const prox = parseDateBR(proxStr);
      if (prox <= hoje) gerarOSPreventiva(prev.ID);
    });
  } catch (e) { console.error('Erro em verificarPreventivasVencidas:', e); }
}

function configurarTriggerDiario() {
  // Remove triggers existentes primeiro
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'verificarPreventivasVencidas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('verificarPreventivasVencidas')
    .timeBased().everyDays(1).atHour(6).create();
  return { success: true, message: 'Trigger diário configurado para 06:00.' };
}

// ---------- HISTÓRICO DE ORDENS (compatibilidade) ----------
function getOrdensHistoricoLegacy() {
  return getOrdensHistorico();
}

// ---------- NOTIFICAÇÕES ----------
function enviarNotificacaoNovaOrdem(ordemId, dados) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('EMAIL_NOTIFICATIONS') !== 'true') return;
  const emails = props.getProperty('EMAILS_EQUIPE') || Session.getActiveUser().getEmail();
  if (!emails) return;
  const lista = emails.split(',').map(e => e.trim()).filter(e => e);
  const corPrioridade = { 'Urgente': '#dc2626', 'Alta': '#f97316', 'Média': '#2563eb', 'Baixa': '#16a34a' };
  const cor = corPrioridade[dados.prioridade] || '#2563eb';
  const assunto = `[Hotel] Nova OS ${dados.prioridade === 'Urgente' ? '🚨 URGENTE' : ''}: ${ordemId}`;
  const corpo = `
    <div style="font-family:sans-serif; max-width:600px; margin:0 auto;">
      <div style="background:${cor}; color:white; padding:20px; border-radius:8px 8px 0 0;">
        <h2 style="margin:0;">🔧 Nova Ordem de Serviço</h2>
        <p style="margin:4px 0 0; opacity:0.9;">ID: ${ordemId}</p>
      </div>
      <div style="background:#f8fafc; padding:20px; border:1px solid #e2e8f0; border-top:none; border-radius:0 0 8px 8px;">
        <table style="width:100%; border-collapse:collapse;">
          <tr><td style="padding:8px; font-weight:600; color:#475569;">Setor</td><td style="padding:8px;">${dados.setor}</td></tr>
          <tr style="background:white;"><td style="padding:8px; font-weight:600; color:#475569;">Local</td><td style="padding:8px;">${dados.local}</td></tr>
          <tr><td style="padding:8px; font-weight:600; color:#475569;">Prioridade</td><td style="padding:8px;"><strong style="color:${cor};">${dados.prioridade}</strong></td></tr>
          <tr style="background:white;"><td style="padding:8px; font-weight:600; color:#475569;">Solicitante</td><td style="padding:8px;">${dados.solicitante}</td></tr>
          <tr><td style="padding:8px; font-weight:600; color:#475569;" colspan="2">Descrição</td></tr>
          <tr style="background:white;"><td colspan="2" style="padding:8px;">${dados.descricao}</td></tr>
        </table>
      </div>
    </div>`;
  try {
    MailApp.sendEmail({ to: lista.join(','), subject: assunto, htmlBody: corpo });
  } catch(e) { console.warn('Erro ao enviar email:', e.message); }
}

// ---------- TESTE / DIAGNÓSTICO ----------

/**
 * Cole esta função no Apps Script e rode manualmente para diagnosticar.
 * Ela NÃO modifica nada — apenas lê e loga.
 */
function diagnosticarSeguro() {
  try {
    const ss = getSpreadsheet();
    console.log('✅ Planilha:', ss.getName());
    console.log('📋 Abas:', ss.getSheets().map(s => s.getName()).join(', '));

    const ordSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!ordSheet) { console.error('❌ Aba OrdensServico NÃO encontrada!'); return; }

    const lastRow = ordSheet.getLastRow();
    const lastCol = ordSheet.getLastColumn();
    console.log(`📊 OrdensServico: ${lastRow} linhas, ${lastCol} colunas`);

    if (lastRow < 1) { console.warn('⚠️ Aba vazia'); return; }

    const headers = ordSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    console.log('📌 Cabeçalhos:', JSON.stringify(headers));

    const statusIdx = headers.indexOf('Status');
    console.log('🔍 Coluna "Status" está na posição:', statusIdx, '(esperado: 8)');

    if (lastRow > 1) {
      const amostra = ordSheet.getRange(2, 1, Math.min(3, lastRow - 1), lastCol).getValues();
      amostra.forEach((row, i) => {
        const statusVal = row[statusIdx];
        // Detecta objetos Date (a causa mais comum de falha na serialização)
        const tiposEspeciais = row.map((v, ci) => v instanceof Date ? `col${ci}=Date` : null).filter(Boolean);
        console.log(`Linha ${i+2}: ID="${row[0]}" | Status="${statusVal}" | tipo=${typeof statusVal}`);
        if (tiposEspeciais.length > 0) {
          console.warn(`  ⚠️ Colunas com objeto Date (causam erro de serialização): ${tiposEspeciais.join(', ')}`);
        }
      });

      // Testa se getOrdensPendentes retorna algo
      const pendentes = getOrdensPendentes();
      console.log(`✅ getOrdensPendentes() retornou: ${pendentes.length} registros`);
      if (pendentes.length > 0) {
        console.log('Primeiro registro:', JSON.stringify(pendentes[0]).substring(0, 300));
      }
    }
  } catch (e) {
    console.error('ERRO no diagnóstico:', e.message, e.stack);
  }
}

function testarConexao() {
  return {
    sucesso: true,
    mensagem: 'Conexão OK!',
    timestamp: new Date().toISOString(),
    versao: '2.0',
    planilhaId: getSpreadsheetId()
  };
}

function autorizarWebApp() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId());
  const nome = ss.getName();
  console.log('✅ Autorização OK. Planilha:', nome);
  return nome;
}