// ===========================
// CONFIGURAÇÕES GERAIS
// ===========================
const SPREADSHEET_ID = '1ajlFoT0kkAwOYVFymFMs5ipiEkmvQdNZzwVJx6B98FM';
const DRIVE_FOLDER_ID = '1g_8lgL55WAb32E6XQ4dl2KMX-AfhvXMY';

const COLUNAS_POCOS = [
  'ID',
  'Estado',
  'Município',
  'Comunidade',
  'Região',
  'Latitude',
  'Longitude',
  'Beneficiários',
  'Investimento',
  'Vazão (L/H)',
  'Profundidade (m)',
  'Status',
  'ResumoStatus',
  'Solicitante',
  'ContatoSolicitante',
  'DataSolicitacao',
  'DataOrcamentoPrevisto',
  'DataInstalacao',
  'DataConclusao',
  'DataPagamento',
  'OrcamentoPrevisto',
  'OrcamentoAprovado',
  'OrcamentoExecutado',
  'Valor Previsto Perfuração',
  'Valor Previsto Instalação',
  'Valor Realizado',
  'TermoAutorizacaoURL',
  'NotaFiscalURL',
  'ContatosJSON',
  'EvidenciasJSON',
  'LinhaDoTempoJSON',
  'Doadores',
  'Empresa Responsável',
  'Observações',
  'DataCadastro',
  'DataUltimaAtualizacao',
  'ResponsavelContato',
  'ContatoInstalacao',
  'TelefoneContato',
  'TelefoneContatoNormalizado',
  'StatusContato',
  'ProximaAcao',
  'UltimoContato',
  'ImpactoNoStatus',
  'TipoPoco',
  'SituacaoHidrica',
  'AcoesPosInstalacao',
  'UsoAguaComunitario',
  'GeocodificacaoFonte',
  'GeocodificacaoPrecisao',
  'GeocodificacaoStatus',
  'GeocodificacaoTimestamp'
];

// ===========================
// RENDER HTML
// ===========================
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistema de Poços Missionários')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizarNumero(valor) {
  if (valor instanceof Date) return 0;
  if (typeof valor === 'number') return valor;
  if (!valor) return 0;
  return Number(String(valor).replace(/[^0-9,-]+/g, '').replace(',', '.')) || 0;
}

function normalizarCoordenadaGS(valor, limite, nomeCampo) {
  if (valor === undefined || valor === null) return '';
  if (valor instanceof Date) return '';
  const texto = typeof valor === 'number' ? valor.toString() : String(valor).trim();
  if (!texto) return '';
  const numero = Number(texto.replace(/\s+/g, '').replace(',', '.'));
  if (!Number.isFinite(numero)) {
    throw new Error(`Valor de ${nomeCampo} inválido.`);
  }
  if (numero < -limite || numero > limite) {
    const nome = nomeCampo.charAt(0).toUpperCase() + nomeCampo.slice(1);
    throw new Error(`${nome} deve estar entre ${-limite} e ${limite}.`);
  }
  return Number(numero.toFixed(6));
}

function prepararGeocodificacaoRegistro(geocodificacao, coordenadasInformadas, agora) {
  if (!coordenadasInformadas) {
    return { fonte: '', precisao: '', status: 'Sem coordenadas', timestamp: '' };
  }

  const dados = geocodificacao || {};
  const statusInformado = dados.status || dados.Status || '';
  const sucesso = dados.sucesso === true || String(statusInformado).toUpperCase() === 'OK';
  if (!sucesso) {
    throw new Error('A geocodificação das coordenadas é obrigatória antes de salvar.');
  }

  let timestamp = dados.timestamp || dados.Timestamp || '';
  if (timestamp) {
    if (!(timestamp instanceof Date)) {
      const data = new Date(timestamp);
      timestamp = Number.isNaN(data.getTime()) ? agora : data;
    }
  } else {
    timestamp = agora;
  }

  return {
    fonte: dados.fonte || dados.Fonte || '',
    precisao: dados.precisao || dados.Precisao || '',
    status: statusInformado || 'OK',
    timestamp
  };
}

function extrairDataDeTexto(texto) {
  if (!texto || typeof texto !== 'string') return null;
  const match = texto.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (!match) return null;
  const [, dia, mes, ano] = match;
  return new Date(Number(ano), Number(mes) - 1, Number(dia));
}

function extrairStatusDaEtapa(texto) {
  if (!texto || typeof texto !== 'string') return 'Sem registro';
  const lower = texto.toLowerCase();
  if (lower.includes('conclu')) return 'Concluída';
  if (lower.includes('andamento')) return 'Em andamento';
  if (lower.includes('previst')) return 'Prevista';
  if (lower.includes('licen') || lower.includes('document')) return 'Documentação';
  return 'Planejado';
}

const PROCESS_STAGES = [
  {
    id: 'solicitado',
    label: 'Solicitação registrada',
    status: 'Solicitado',
    field: 'DataSolicitacao',
    description: 'Dados iniciais do poço enviados pela equipe de campo.'
  },
  {
    id: 'orcamento',
    label: 'Orçamento previsto',
    status: 'Orçamento previsto',
    field: 'DataOrcamentoPrevisto',
    description: 'Análise orçamentária concluída e vinculada aos doadores.'
  },
  {
    id: 'instalacao',
    label: 'Instalação em andamento',
    status: 'Instalação',
    field: 'DataInstalacao',
    description: 'Equipe técnica mobilizada para executar a instalação do poço.'
  },
  {
    id: 'conclusao',
    label: 'Poço em operação',
    status: 'Concluído',
    field: 'DataConclusao',
    description: 'Infraestrutura entregue e validada junto à comunidade.'
  },
  {
    id: 'pagamento',
    label: 'Pagamento finalizado',
    status: 'Pago',
    field: 'DataPagamento',
    description: 'Prestação de contas encerrada com pagamento do fornecedor.'
  }
];

function normalizarTextoStatus(valor) {
  if (valor == null) return '';
  return String(valor)
    .normalize('NFD')
    .replace(/[^\w\s-]/g, '')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

const STATUS_ALIAS_MAP = (() => {
  const mapa = {};
  const registrar = (texto, etapaId) => {
    const chave = normalizarTextoStatus(texto);
    if (chave) {
      mapa[chave] = etapaId;
    }
  };
  PROCESS_STAGES.forEach(stage => {
    registrar(stage.status, stage.id);
    registrar(stage.label, stage.id);
    registrar(stage.id, stage.id);
  });
  registrar('planejado', 'orcamento');
  registrar('planejamento', 'orcamento');
  registrar('previsto', 'orcamento');
  registrar('orcamento', 'orcamento');
  registrar('analise orcamentaria', 'orcamento');
  registrar('em andamento', 'instalacao');
  registrar('instalando', 'instalacao');
  registrar('execucao', 'instalacao');
  registrar('executando', 'instalacao');
  registrar('realizado', 'conclusao');
  registrar('finalizado', 'conclusao');
  registrar('concluido', 'conclusao');
  registrar('concluido pago', 'pagamento');
  registrar('concluido/pago', 'pagamento');
  registrar('pagamento', 'pagamento');
  registrar('quitado', 'pagamento');
  registrar('solicitacao', 'solicitado');
  registrar('solicitacao registrada', 'solicitado');
  registrar('pedido', 'solicitado');
  registrar('solicitado', 'solicitado');
  registrar('orçamento previsto', 'orcamento');
  registrar('instalação', 'instalacao');
  registrar('concluído', 'conclusao');
  registrar('pago', 'pagamento');
  return mapa;
})();

function obterEtapaPorStatusInformado(status) {
  const chave = normalizarTextoStatus(status);
  if (!chave) return null;
  const etapaId = STATUS_ALIAS_MAP[chave];
  if (!etapaId) return null;
  return PROCESS_STAGES.find(stage => stage.id === etapaId) || null;
}

function padronizarStatusProcessual(status) {
  const texto = status == null ? '' : String(status).trim();
  const etapa = obterEtapaPorStatusInformado(texto);
  return etapa ? etapa.status : texto;
}

function obterIdEtapaPorStatus(status) {
  const etapa = obterEtapaPorStatusInformado(status);
  return etapa ? etapa.id : '';
}

function converterParaData(valor) {
  if (!valor) return '';
  if (valor instanceof Date) {
    return isNaN(valor.getTime()) ? '' : valor;
  }
  if (typeof valor === 'number') {
    const d = new Date(valor);
    return isNaN(d.getTime()) ? '' : d;
  }
  if (typeof valor === 'string') {
    const texto = valor.trim();
    if (!texto) return '';
    const dataTexto = extrairDataDeTexto(texto);
    if (dataTexto) return dataTexto;
    const tentativa = new Date(texto);
    return isNaN(tentativa.getTime()) ? '' : tentativa;
  }
  return '';
}

function formatarDataISO(valor) {
  const data = converterParaData(valor);
  if (!data) return '';
  return new Date(data.getTime() - data.getTimezoneOffset() * 60000).toISOString();
}

function parseJSONSeguro(texto, padrao) {
  if (!texto) return padrao;
  try {
    if (typeof texto === 'object') return texto;
    return JSON.parse(texto);
  } catch (err) {
    return padrao;
  }
}

function determinarStatusProcessual(registro) {
  if (!registro) return 'Solicitado';
  for (let i = PROCESS_STAGES.length - 1; i >= 0; i--) {
    const etapa = PROCESS_STAGES[i];
    const data = registro[etapa.field];
    if (data) {
      const dataConvertida = converterParaData(data);
      if (dataConvertida) {
        return etapa.status;
      }
    }
  }
  return 'Solicitado';
}

function gerarLinhaDoTempoBase(registro, extras = []) {
  const statusAtual = registro.Status || determinarStatusProcessual(registro);
  const indiceAtual = Math.max(0, PROCESS_STAGES.findIndex(e => e.status === statusAtual));
  const linhaBase = PROCESS_STAGES.map((etapa, index) => {
    const data = converterParaData(registro[etapa.field]);
    let situacao = 'pendente';
    if (index < indiceAtual) situacao = 'concluido';
    else if (index === indiceAtual) situacao = data ? 'concluido' : 'andamento';
    return {
      id: etapa.id,
      titulo: etapa.label,
      descricao: etapa.description,
      status: situacao,
      data: data ? data : ''
    };
  });

  const extrasValidos = Array.isArray(extras) ? extras.filter(item => item && item.titulo) : [];
  return linhaBase.concat(extrasValidos.map(item => ({
    id: item.id || Utilities.getUuid(),
    titulo: item.titulo,
    descricao: item.descricao || '',
    status: item.status || 'pendente',
    data: converterParaData(item.data) || ''
  })));
}

function mapearRegistroPoco(dados) {
  const registro = Object.assign({}, dados);
  registro.Contatos = parseJSONSeguro(registro.ContatosJSON, []);
  registro.Evidencias = parseJSONSeguro(registro.EvidenciasJSON, []);
  const extrasTimeline = parseJSONSeguro(registro.LinhaDoTempoJSON, []);
  const statusPadrao = padronizarStatusProcessual(registro.Status);
  registro.Status = statusPadrao || determinarStatusProcessual(registro);
  registro.LinhaDoTempo = gerarLinhaDoTempoBase(registro, extrasTimeline);
  return registro;
}

function garantirColunas(sheet, colunasDesejadas) {
  if (!sheet) return [];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(colunasDesejadas);
    return colunasDesejadas.slice();
  }
  const ultimaColuna = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, Math.max(ultimaColuna, colunasDesejadas.length)).getValues()[0];
  let alterado = false;
  colunasDesejadas.forEach(coluna => {
    if (!headers.includes(coluna)) {
      headers.push(coluna);
      alterado = true;
    }
  });
  if (alterado) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return headers;
}

// ===========================
// INICIALIZAÇÃO DAS GUIAS
// ===========================
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const names = ['Poços', 'Doadores', 'PrestaçãoContas'];

  names.forEach(name => {
    if (!ss.getSheetByName(name)) {
      const sh = ss.insertSheet(name);
      if (name === 'Poços') {
        sh.appendRow(COLUNAS_POCOS);
      } else if (name === 'Doadores') {
        sh.appendRow(['ID', 'Nome', 'Email', 'Telefone', 'ValorDoado', 'DataDoacao', 'PoçosVinculados']);
      } else if (name === 'PrestaçãoContas') {
        sh.appendRow(['PoçoID', 'Data', 'Descrição', 'Valor', 'ComprovanteURL', 'Categoria', 'RegistradoPor']);
      }
    }
  });
  return 'Guias verificadas/criadas com sucesso.';
}


// Listar poços
function listarPocos() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values.shift();
  return values.map(row => {
    const bruto = Object.fromEntries(headers.map((h, i) => [h, row[i]]));
    const registro = mapearRegistroPoco(bruto);
    const statusAtual = registro.Status || determinarStatusProcessual(registro);
    let indiceEtapa = PROCESS_STAGES.findIndex(stage => stage.status === statusAtual);
    if (indiceEtapa < 0) indiceEtapa = 0;
    const etapa = PROCESS_STAGES[indiceEtapa];

    return Object.assign({}, registro, {
      Status: statusAtual,
      EtapaId: etapa.id,
      EtapaIndice: indiceEtapa,
      EtapaNome: etapa.label,
      StatusProcessual: etapa.status
    });
  });
}

function obterKanbanProcesso() {
  return listarPocos().map(p => ({
    id: p.ID,
    titulo: p['Comunidade'] || p['Município'] || p['Estado'] || 'Poço sem identificação',
    local: [p['Município'], p['Estado']].filter(Boolean).join(' • '),
    status: p.Status,
    previsto: Number(p.OrcamentoPrevisto || p['Valor Previsto Perfuração'] || 0),
    executado: Number(p.OrcamentoExecutado || p['Valor Realizado'] || 0),
    beneficiarios: Number(p['Beneficiários'] || 0),
    responsavel: p.ResponsavelContato || ''
  }));
}

function obterCronogramaPocos() {
  const mapaCampos = {
    DataSolicitacao: 'solicitacao',
    DataOrcamentoPrevisto: 'orcamento',
    DataInstalacao: 'instalacao',
    DataConclusao: 'conclusao',
    DataPagamento: 'pagamento'
  };

  return listarPocos().map(p => {
    const datas = {};
    PROCESS_STAGES.forEach(stage => {
      const chave = mapaCampos[stage.field];
      if (!chave) return;
      const dataConvertida = converterParaData(p[stage.field]);
      datas[chave] = dataConvertida ? formatarDataISO(dataConvertida) : '';
    });

    const inicio = converterParaData(p.DataSolicitacao) || converterParaData(p.DataCadastro) || new Date();
    const fim = converterParaData(p.DataPagamento) || converterParaData(p.DataConclusao) || converterParaData(p.DataInstalacao) || inicio;
    const previsto = Number(p.OrcamentoPrevisto || p['Valor Previsto Perfuração'] || 0);
    const executado = Number(p.OrcamentoExecutado || p['Valor Realizado'] || 0);
    const progresso = previsto > 0 ? Math.min(executado / previsto, 1) : 0;

    return {
      id: p.ID,
      nome: p['Comunidade'] || p['Município'] || p['Estado'] || 'Poço sem identificação',
      estado: p['Estado'] || '',
      status: p.Status || 'Solicitado',
      datas,
      inicio: formatarDataISO(inicio),
      fim: formatarDataISO(fim),
      previsto,
      executado,
      progresso,
      beneficiarios: Number(p['Beneficiários'] || 0)
    };
  });
}

// Salvar novo poço

function salvarPoco(poco) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  garantirColunas(sh, COLUNAS_POCOS);
  const id = Utilities.getUuid();
  const agora = new Date();

  const contatos = Array.isArray(poco.contatos) ? poco.contatos : [];
  const evidencias = Array.isArray(poco.evidencias) ? poco.evidencias : [];
  const extrasTimeline = Array.isArray(poco.timeline) ? poco.timeline : [];

  const latitudeNormalizada = normalizarCoordenadaGS(poco.latitude, 90, 'latitude');
  const longitudeNormalizada = normalizarCoordenadaGS(poco.longitude, 180, 'longitude');
  const coordenadasInformadas = latitudeNormalizada !== '' && longitudeNormalizada !== '';
  const geocodificacaoDados = prepararGeocodificacaoRegistro(poco.geocodificacao, coordenadasInformadas, agora);

  const registro = {
    ID: id,
    Estado: poco.estado || '',
    Município: poco.municipio || '',
    Comunidade: poco.comunidade || '',
    Região: poco.regiao || '',
    Latitude: coordenadasInformadas ? latitudeNormalizada : '',
    Longitude: coordenadasInformadas ? longitudeNormalizada : '',
    Beneficiários: Number(poco.beneficiarios) || 0,
    Investimento: Number(poco.investimento) || 0,
    'Vazão (L/H)': poco.vazao || '',
    'Profundidade (m)': poco.profundidade || '',
    Status: poco.status || '',
    ResumoStatus: poco.resumoStatus || '',
    Solicitante: poco.solicitante || '',
    ContatoSolicitante: poco.contatoSolicitante || '',
    DataSolicitacao: converterParaData(poco.dataSolicitacao) || agora,
    DataOrcamentoPrevisto: converterParaData(poco.dataOrcamentoPrevisto) || '',
    DataInstalacao: converterParaData(poco.dataInstalacao) || '',
    DataConclusao: converterParaData(poco.dataConclusao) || '',
    DataPagamento: converterParaData(poco.dataPagamento) || '',
    OrcamentoPrevisto: Number(poco.orcamentoPrevisto) || 0,
    OrcamentoAprovado: Number(poco.orcamentoAprovado) || 0,
    OrcamentoExecutado: Number(poco.orcamentoExecutado) || 0,
    'Valor Previsto Perfuração': Number(poco.valorPrevPerf) || Number(poco.orcamentoPrevisto) || 0,
    'Valor Previsto Instalação': Number(poco.valorPrevInst) || 0,
    'Valor Realizado': Number(poco.orcamentoExecutado) || 0,
    TermoAutorizacaoURL: poco.termoAutorizacaoURL || '',
    NotaFiscalURL: poco.notaFiscalURL || '',
    ContatosJSON: JSON.stringify(contatos),
    EvidenciasJSON: JSON.stringify(evidencias),
    LinhaDoTempoJSON: JSON.stringify(extrasTimeline),
    Doadores: poco.doadores || '',
    'Empresa Responsável': poco.empresaResponsavel || '',
    Observações: poco.observacoes || '',
    DataCadastro: agora,
    DataUltimaAtualizacao: agora,
    ResponsavelContato: poco.responsavelContato || '',
    ContatoInstalacao: poco.contatoInstalacao || '',
    TelefoneContato: poco.telefoneContato || '',
    TelefoneContatoNormalizado: poco.telefoneContatoNormalizado || '',
    StatusContato: poco.statusContato || '',
    ProximaAcao: poco.proximaAcao || '',
    UltimoContato: converterParaData(poco.ultimoContato) || '',
    ImpactoNoStatus: poco.impactoNoStatus || '',
    TipoPoco: poco.tipoPoco || '',
    SituacaoHidrica: poco.situacaoHidrica || '',
    AcoesPosInstalacao: poco.acoesPosInstalacao || '',
    UsoAguaComunitario: poco.usoAguaComunitario || '',
    GeocodificacaoFonte: geocodificacaoDados.fonte || '',
    GeocodificacaoPrecisao: geocodificacaoDados.precisao || '',
    GeocodificacaoStatus: geocodificacaoDados.status || (coordenadasInformadas ? 'OK' : 'Sem coordenadas'),
    GeocodificacaoTimestamp: geocodificacaoDados.timestamp || ''
  };

  registro.Status = determinarStatusProcessual(registro);
  const linhaTempoCompleta = gerarLinhaDoTempoBase(registro, extrasTimeline);
  registro.LinhaDoTempoJSON = JSON.stringify(linhaTempoCompleta.slice(PROCESS_STAGES.length));

  const data = COLUNAS_POCOS.map(coluna => {
    const valor = registro[coluna];
    return valor === undefined ? '' : valor;
  });
  sh.appendRow(data);
  return { success: true, id };
}

function atualizarPoco(poco) {
  if (!poco || !poco.id) throw new Error('ID do poço não informado.');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Poços');
  garantirColunas(sh, COLUNAS_POCOS);

  const valores = sh.getDataRange().getValues();
  if (valores.length <= 1) throw new Error('Nenhum poço cadastrado.');
  const headers = valores.shift();
  const idIndex = headers.indexOf('ID');
  if (idIndex === -1) throw new Error('Estrutura da planilha de poços inválida.');

  let rowIndex = -1;
  let registroAtual = null;
  for (let i = 0; i < valores.length; i++) {
    if (valores[i][idIndex] === poco.id) {
      rowIndex = i + 2; // considerar cabeçalho
      registroAtual = {};
      headers.forEach((h, idx) => registroAtual[h] = valores[i][idx]);
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Poço não encontrado para atualização.');

  const atualMapeado = mapearRegistroPoco(registroAtual);
  const agora = new Date();

  const contatos = Array.isArray(poco.contatos) ? poco.contatos : atualMapeado.Contatos;
  const evidencias = Array.isArray(poco.evidencias) ? poco.evidencias : atualMapeado.Evidencias;
  const extrasTimeline = Array.isArray(poco.timeline) ? poco.timeline : (registroAtual.LinhaDoTempoJSON ? parseJSONSeguro(registroAtual.LinhaDoTempoJSON, []) : []);

  const latitudeAtual = normalizarCoordenadaGS(registroAtual.Latitude, 90, 'latitude');
  const longitudeAtual = normalizarCoordenadaGS(registroAtual.Longitude, 180, 'longitude');
  const latitudeNova = poco.latitude !== undefined ? normalizarCoordenadaGS(poco.latitude, 90, 'latitude') : latitudeAtual;
  const longitudeNova = poco.longitude !== undefined ? normalizarCoordenadaGS(poco.longitude, 180, 'longitude') : longitudeAtual;
  const coordenadasInformadas = latitudeNova !== '' && longitudeNova !== '';
  const coordenadasAlteradas = (poco.latitude !== undefined && latitudeNova !== latitudeAtual) || (poco.longitude !== undefined && longitudeNova !== longitudeAtual);
  const geocodeAnteriorOk = String(registroAtual.GeocodificacaoStatus || '').toUpperCase() === 'OK';

  let geocodificacaoDados;
  if (coordenadasInformadas) {
    if (poco.geocodificacao) {
      geocodificacaoDados = prepararGeocodificacaoRegistro(poco.geocodificacao, true, agora);
    } else if (coordenadasAlteradas) {
      throw new Error('As coordenadas informadas precisam ser validadas com geocodificação antes de salvar.');
    } else {
      let timestampGeo = registroAtual.GeocodificacaoTimestamp || '';
      if (timestampGeo && !(timestampGeo instanceof Date)) {
        const dataGeo = new Date(timestampGeo);
        timestampGeo = Number.isNaN(dataGeo.getTime()) ? agora : dataGeo;
      }
      const statusAnterior = registroAtual.GeocodificacaoStatus || '';
      geocodificacaoDados = {
        fonte: registroAtual.GeocodificacaoFonte || '',
        precisao: registroAtual.GeocodificacaoPrecisao || '',
        status: statusAnterior || (geocodeAnteriorOk ? 'OK' : 'Pendente'),
        timestamp: timestampGeo || ''
      };
    }
  } else {
    geocodificacaoDados = { fonte: '', precisao: '', status: 'Sem coordenadas', timestamp: '' };
  }

  const registro = {
    ID: registroAtual.ID,
    Estado: poco.estado !== undefined ? poco.estado : registroAtual.Estado,
    Município: poco.municipio !== undefined ? poco.municipio : registroAtual['Município'],
    Comunidade: poco.comunidade !== undefined ? poco.comunidade : registroAtual.Comunidade,
    Região: poco.regiao !== undefined ? poco.regiao : (registroAtual['Região'] || ''),
    Latitude: latitudeNova,
    Longitude: longitudeNova,
    Beneficiários: poco.beneficiarios !== undefined ? Number(poco.beneficiarios) || 0 : (Number(registroAtual['Beneficiários']) || 0),
    Investimento: poco.investimento !== undefined ? Number(poco.investimento) || 0 : (Number(registroAtual.Investimento) || 0),
    'Vazão (L/H)': poco.vazao !== undefined ? poco.vazao : registroAtual['Vazão (L/H)'],
    'Profundidade (m)': poco.profundidade !== undefined ? poco.profundidade : registroAtual['Profundidade (m)'],
    Status: poco.status !== undefined ? poco.status : (registroAtual.Status || ''),
    ResumoStatus: poco.resumoStatus !== undefined ? poco.resumoStatus : (registroAtual.ResumoStatus || ''),
    Solicitante: poco.solicitante !== undefined ? poco.solicitante : (registroAtual.Solicitante || ''),
    ContatoSolicitante: poco.contatoSolicitante !== undefined ? poco.contatoSolicitante : (registroAtual.ContatoSolicitante || ''),
    DataSolicitacao: poco.dataSolicitacao !== undefined ? converterParaData(poco.dataSolicitacao) || '' : converterParaData(registroAtual.DataSolicitacao) || '',
    DataOrcamentoPrevisto: poco.dataOrcamentoPrevisto !== undefined ? converterParaData(poco.dataOrcamentoPrevisto) || '' : converterParaData(registroAtual.DataOrcamentoPrevisto) || '',
    DataInstalacao: poco.dataInstalacao !== undefined ? converterParaData(poco.dataInstalacao) || '' : converterParaData(registroAtual.DataInstalacao) || '',
    DataConclusao: poco.dataConclusao !== undefined ? converterParaData(poco.dataConclusao) || '' : converterParaData(registroAtual.DataConclusao) || '',
    DataPagamento: poco.dataPagamento !== undefined ? converterParaData(poco.dataPagamento) || '' : converterParaData(registroAtual.DataPagamento) || '',
    OrcamentoPrevisto: poco.orcamentoPrevisto !== undefined ? Number(poco.orcamentoPrevisto) || 0 : (Number(registroAtual.OrcamentoPrevisto) || 0),
    OrcamentoAprovado: poco.orcamentoAprovado !== undefined ? Number(poco.orcamentoAprovado) || 0 : (Number(registroAtual.OrcamentoAprovado) || 0),
    OrcamentoExecutado: poco.orcamentoExecutado !== undefined ? Number(poco.orcamentoExecutado) || 0 : (Number(registroAtual.OrcamentoExecutado) || 0),
    'Valor Previsto Perfuração': poco.valorPrevPerf !== undefined ? Number(poco.valorPrevPerf) || 0 : (Number(registroAtual['Valor Previsto Perfuração']) || 0),
    'Valor Previsto Instalação': poco.valorPrevInst !== undefined ? Number(poco.valorPrevInst) || 0 : (Number(registroAtual['Valor Previsto Instalação']) || 0),
    'Valor Realizado': poco.orcamentoExecutado !== undefined ? Number(poco.orcamentoExecutado) || 0 : (Number(registroAtual['Valor Realizado']) || 0),
    TermoAutorizacaoURL: poco.termoAutorizacaoURL !== undefined ? poco.termoAutorizacaoURL : (registroAtual.TermoAutorizacaoURL || ''),
    NotaFiscalURL: poco.notaFiscalURL !== undefined ? poco.notaFiscalURL : (registroAtual.NotaFiscalURL || ''),
    ContatosJSON: JSON.stringify(contatos),
    EvidenciasJSON: JSON.stringify(evidencias),
    LinhaDoTempoJSON: JSON.stringify(extrasTimeline),
    Doadores: poco.doadores !== undefined ? poco.doadores : (registroAtual.Doadores || ''),
    'Empresa Responsável': poco.empresaResponsavel !== undefined ? poco.empresaResponsavel : (registroAtual['Empresa Responsável'] || ''),
    Observações: poco.observacoes !== undefined ? poco.observacoes : (registroAtual.Observações || ''),
    DataCadastro: converterParaData(registroAtual.DataCadastro) || converterParaData(registroAtual['DataCadastro']) || new Date(),
    DataUltimaAtualizacao: agora,
    ResponsavelContato: poco.responsavelContato !== undefined ? poco.responsavelContato : (registroAtual.ResponsavelContato || ''),
    ContatoInstalacao: poco.contatoInstalacao !== undefined ? poco.contatoInstalacao : (registroAtual.ContatoInstalacao || ''),
    TelefoneContato: poco.telefoneContato !== undefined ? poco.telefoneContato : (registroAtual.TelefoneContato || ''),
    TelefoneContatoNormalizado: poco.telefoneContatoNormalizado !== undefined ? poco.telefoneContatoNormalizado : (registroAtual.TelefoneContatoNormalizado || ''),
    StatusContato: poco.statusContato !== undefined ? poco.statusContato : (registroAtual.StatusContato || ''),
    ProximaAcao: poco.proximaAcao !== undefined ? poco.proximaAcao : (registroAtual.ProximaAcao || ''),
    UltimoContato: poco.ultimoContato !== undefined ? converterParaData(poco.ultimoContato) || '' : converterParaData(registroAtual.UltimoContato) || '',
    ImpactoNoStatus: poco.impactoNoStatus !== undefined ? poco.impactoNoStatus : (registroAtual.ImpactoNoStatus || ''),
    TipoPoco: poco.tipoPoco !== undefined ? poco.tipoPoco : (registroAtual.TipoPoco || ''),
    SituacaoHidrica: poco.situacaoHidrica !== undefined ? poco.situacaoHidrica : (registroAtual.SituacaoHidrica || ''),
    AcoesPosInstalacao: poco.acoesPosInstalacao !== undefined ? poco.acoesPosInstalacao : (registroAtual.AcoesPosInstalacao || ''),
    UsoAguaComunitario: poco.usoAguaComunitario !== undefined ? poco.usoAguaComunitario : (registroAtual.UsoAguaComunitario || ''),
    GeocodificacaoFonte: geocodificacaoDados.fonte || '',
    GeocodificacaoPrecisao: geocodificacaoDados.precisao || '',
    GeocodificacaoStatus: geocodificacaoDados.status || (coordenadasInformadas ? 'OK' : 'Sem coordenadas'),
    GeocodificacaoTimestamp: geocodificacaoDados.timestamp || ''
  };

  registro.Status = determinarStatusProcessual(registro);
  const linhaTempoCompleta = gerarLinhaDoTempoBase(registro, extrasTimeline);
  registro.LinhaDoTempoJSON = JSON.stringify(linhaTempoCompleta.slice(PROCESS_STAGES.length));

  const data = COLUNAS_POCOS.map(coluna => registro[coluna] === undefined ? '' : registro[coluna]);
  sh.getRange(rowIndex, 1, 1, COLUNAS_POCOS.length).setValues([data]);
  return { success: true, id: poco.id };
}

function geocodificarCoordenadas(lat, lon) {
  try {
    const latitude = normalizarCoordenadaGS(lat, 90, 'latitude');
    const longitude = normalizarCoordenadaGS(lon, 180, 'longitude');
    if (latitude === '' || longitude === '') {
      return { sucesso: false, mensagem: 'Informe latitude e longitude válidas.' };
    }

    const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${encodeURIComponent(latitude)}&lon=${encodeURIComponent(longitude)}&zoom=10&accept-language=pt-BR&addressdetails=1`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'EmpSocialGeocoder/1.0 (contato@empsocial.org)'
      }
    });

    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      return { sucesso: false, mensagem: `Serviço de geocodificação indisponível (${statusCode}).` };
    }

    const payload = JSON.parse(response.getContentText() || '{}');
    const address = payload.address || {};
    const estado = address.state || address.region || '';
    const municipio = address.city || address.town || address.village || address.municipality || address.county || '';
    if (!estado && !municipio) {
      return { sucesso: false, mensagem: 'Não foi possível identificar estado ou município para essas coordenadas.' };
    }

    return {
      sucesso: true,
      estado,
      municipio,
      cidade: municipio,
      fonte: 'Nominatim',
      precisao: payload.type || '',
      status: 'OK',
      timestamp: new Date().toISOString(),
      displayName: payload.display_name || ''
    };
  } catch (error) {
    return { sucesso: false, mensagem: error.message || 'Falha ao processar a geocodificação.' };
  }
}

// Upload de arquivo (imagem, pdf etc.)
function uploadFile(base64, nomeArquivo, mimeType) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1]), mimeType, nomeArquivo);
  const file = folder.createFile(blob);
  return file.getUrl();
}


// ===========================
// FUNÇÕES DE DOADORES
// ===========================

// Listar doadores
function listarDoadores() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Doadores');
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

// Salvar novo doador
function salvarDoador(doador) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Doadores');
  const id = Utilities.getUuid();
  const data = [
    id,
    doador.nome || '',
    doador.email || '',
    doador.telefone || '',
    Number(doador.valorDoado) || 0,
    new Date(),
    ''
  ];
  sh.appendRow(data);
  return { success: true, id };
}

// Vincular doador a poços
function vincularDoadorAPocos(doadorId, pocosIds) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Poços');
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idIndex = headers.indexOf('ID');
  const doadoresIndex = headers.indexOf('Doadores');

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (pocosIds.includes(row[idIndex])) {
      const atuais = row[doadoresIndex] ? row[doadoresIndex].split(',') : [];
      if (!atuais.includes(doadorId)) atuais.push(doadorId);
      sh.getRange(i + 2, doadoresIndex + 1).setValue(atuais.join(','));
    }
  }

  const shDoadores = ss.getSheetByName('Doadores');
  if (shDoadores) {
    const valoresDoadores = shDoadores.getDataRange().getValues();
    const headersDoadores = valoresDoadores.shift();
    const idIndex = headersDoadores.indexOf('ID');
    const vinculadosIndex = headersDoadores.indexOf('PoçosVinculados');
    for (let i = 0; i < valoresDoadores.length; i++) {
      if (valoresDoadores[i][idIndex] === doadorId) {
        const atuais = valoresDoadores[i][vinculadosIndex] ? valoresDoadores[i][vinculadosIndex].split(',') : [];
        pocosIds.forEach(id => {
          if (!atuais.includes(id)) atuais.push(id);
        });
        shDoadores.getRange(i + 2, vinculadosIndex + 1).setValue(atuais.join(','));
        break;
      }
    }
  }
  return 'Doador vinculado aos poços com sucesso.';
}

// ===========================
// FUNÇÕES DE PRESTAÇÃO DE CONTAS
// ===========================

// Listar prestações (todas ou filtradas por poço)
function listarPrestacoes(pocoId) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PrestaçãoContas');
  const [headers, ...rows] = sh.getDataRange().getValues();
  let registros = rows.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
  if (pocoId) registros = registros.filter(r => r.PoçoID === pocoId);
  return registros;
}

// Salvar nova despesa
function salvarPrestacao(despesa) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('PrestaçãoContas');
  const row = [
    despesa.pocoId,
    despesa.data ? new Date(despesa.data) : new Date(),
    despesa.descricao || '',
    Number(despesa.valor) || 0,
    despesa.comprovanteURL || '',
    despesa.categoria || '',
    despesa.registradoPor || ''
  ];
  sh.appendRow(row);

  // Atualizar valor realizado na planilha de Poços
  const shPocos = ss.getSheetByName('Poços');
  const values = shPocos.getDataRange().getValues();
  const headers = values.shift();
  const idIndex = headers.indexOf('ID');
  const orcamentoExecutadoIndex = headers.indexOf('OrcamentoExecutado');

  for (let i = 0; i < values.length; i++) {
    if (values[i][idIndex] === despesa.pocoId) {
      const executadoAtual = orcamentoExecutadoIndex !== -1 ? Number(values[i][orcamentoExecutadoIndex]) || 0 : 0;
      const novoExecutado = executadoAtual + (Number(despesa.valor) || 0);
      const categoria = (despesa.categoria || '').toLowerCase();
      const dataPagamento = categoria.includes('pag') ? (despesa.data ? new Date(despesa.data) : new Date()) : undefined;
      atualizarPoco({
        id: despesa.pocoId,
        orcamentoExecutado: novoExecutado,
        dataPagamento: dataPagamento
      });
      break;
    }
  }

  return { success: true };
}

// ===========================
// CONTATOS E ENGAJAMENTO DE CAMPO
// ===========================

function atualizarContatoPoco(registro) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Poços');
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idIndex = headers.indexOf('ID');
  if (idIndex === -1) return 'Planilha sem coluna ID';

  const campos = {
    ResponsavelContato: headers.indexOf('ResponsavelContato'),
    ContatoInstalacao: headers.indexOf('ContatoInstalacao'),
    TelefoneContato: headers.indexOf('TelefoneContato'),
    StatusContato: headers.indexOf('StatusContato'),
    ProximaAcao: headers.indexOf('ProximaAcao'),
    UltimoContato: headers.indexOf('UltimoContato'),
    ImpactoNoStatus: headers.indexOf('ImpactoNoStatus')
  };

  for (let i = 0; i < values.length; i++) {
    if (values[i][idIndex] === registro.pocoId) {
      const rowIndex = i + 2;
      if (registro.responsavelContato !== undefined && campos.ResponsavelContato !== -1) {
        sh.getRange(rowIndex, campos.ResponsavelContato + 1).setValue(registro.responsavelContato);
      }
      if (registro.contatoInstalacao !== undefined && campos.ContatoInstalacao !== -1) {
        sh.getRange(rowIndex, campos.ContatoInstalacao + 1).setValue(registro.contatoInstalacao);
      }
      if (registro.telefoneContato !== undefined && campos.TelefoneContato !== -1) {
        sh.getRange(rowIndex, campos.TelefoneContato + 1).setValue(registro.telefoneContato);
      }
      if (registro.statusContato !== undefined && campos.StatusContato !== -1) {
        sh.getRange(rowIndex, campos.StatusContato + 1).setValue(registro.statusContato);
      }
      if (registro.proximaAcao !== undefined && campos.ProximaAcao !== -1) {
        sh.getRange(rowIndex, campos.ProximaAcao + 1).setValue(registro.proximaAcao);
      }
      if (registro.ultimoContato !== undefined && campos.UltimoContato !== -1) {
        sh.getRange(rowIndex, campos.UltimoContato + 1).setValue(registro.ultimoContato ? new Date(registro.ultimoContato) : '');
      }
      if (registro.impactoNoStatus !== undefined && campos.ImpactoNoStatus !== -1) {
        sh.getRange(rowIndex, campos.ImpactoNoStatus + 1).setValue(registro.impactoNoStatus);
      }
      break;
    }
  }
  return 'Contato atualizado com sucesso.';
}

function registrarContatoPoco(contato) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Contatos');
  if (!sh) throw new Error('A aba "Contatos" não foi encontrada. Execute criarBaseDeDados().');

  const id = Utilities.getUuid();
  const row = [
    id,
    contato.pocoId,
    contato.responsavelContato || '',
    contato.contatoExterno || '',
    contato.organizacaoContato || '',
    contato.dataContato ? new Date(contato.dataContato) : new Date(),
    contato.resumo || '',
    contato.proximaAcao || '',
    contato.statusContato || '',
    contato.impactoPrevisto || '',
    contato.registradoPor || ''
  ];
  sh.appendRow(row);

  atualizarContatoPoco({
    pocoId: contato.pocoId,
    responsavelContato: contato.responsavelContato,
    contatoInstalacao: contato.contatoExterno,
    statusContato: contato.statusContato,
    proximaAcao: contato.proximaAcao,
    ultimoContato: contato.dataContato,
    impactoNoStatus: contato.impactoPrevisto
  });

  return { success: true, id };
}

function listarContatosPorPoco(pocoId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Contatos');
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values.shift();
  return values
    .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])))
    .filter(r => !pocoId || r['PoçoID'] === pocoId);
}

// ===================================================
// FUNÇÕES DE RELATÓRIO / ANÁLISE
// ===================================================

// Obter dados completos de um poço (detalhes + despesas)
function obterRelatorioPoco(pocoId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPocos = ss.getSheetByName('Poços');
  const shPrest = ss.getSheetByName('PrestaçãoContas');

  const valoresPocos = shPocos.getDataRange().getValues();
  if (valoresPocos.length <= 1) return { poco: null, despesas: [], timeline: [], evidencias: [], contatos: [] };
  const headersPocos = valoresPocos.shift();
  const registrosPocos = valoresPocos.map(r => Object.fromEntries(headersPocos.map((h, i) => [h, r[i]])));
  const bruto = registrosPocos.find(p => p.ID === pocoId);
  if (!bruto) {
    return { poco: null, despesas: [], timeline: [], evidencias: [], contatos: [] };
  }

  const poco = mapearRegistroPoco(bruto);

  const prestacoes = shPrest.getDataRange().getValues();
  const headersPrest = prestacoes.shift();
  const despesas = prestacoes.map(r => Object.fromEntries(headersPrest.map((h, i) => [h, r[i]])))
                             .filter(d => d.PoçoID === pocoId)
                             .map(d => Object.assign({}, d, {
                               DataISO: formatarDataISO(d.Data)
                             }));

  return {
    poco,
    despesas,
    timeline: poco.LinhaDoTempo,
    evidencias: poco.Evidencias,
    contatos: poco.Contatos
  };
}

function atualizarStatusPoco(id, novoStatus) {
  const etapa = PROCESS_STAGES.find(stage => stage.status === novoStatus);
  if (!etapa) return 'Status inválido';

  const campoParaPropriedade = {
    DataSolicitacao: 'dataSolicitacao',
    DataOrcamentoPrevisto: 'dataOrcamentoPrevisto',
    DataInstalacao: 'dataInstalacao',
    DataConclusao: 'dataConclusao',
    DataPagamento: 'dataPagamento'
  };

  const payload = { id };
  const indiceAlvo = PROCESS_STAGES.findIndex(stage => stage.status === novoStatus);
  PROCESS_STAGES.forEach((stage, indice) => {
    const propriedade = campoParaPropriedade[stage.field];
    if (!propriedade) return;
    if (indice === indiceAlvo) {
      payload[propriedade] = new Date();
    } else if (indice > indiceAlvo) {
      payload[propriedade] = '';
    }
  });

  atualizarPoco(payload);
  return 'Status atualizado';
}

function obterDashboardAnalitico() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPocos = ss.getSheetByName('Poços');
  const shPrest = ss.getSheetByName('PrestaçãoContas');
  const shDoadores = ss.getSheetByName('Doadores');
  const shContatos = ss.getSheetByName('Contatos');

  const valoresPocos = shPocos.getDataRange().getValues();
  const headersPocos = valoresPocos.shift();
  const pocos = valoresPocos.map(r => Object.fromEntries(headersPocos.map((h, i) => [h, r[i]])));

  const contagemEtapas = PROCESS_STAGES.reduce((acc, stage) => {
    acc[stage.id] = 0;
    return acc;
  }, {});
  const contagemOutros = {};
  const mapIdParaNome = {};
  const mapIdParaPoco = {};
  pocos.forEach(p => {
    const statusPadrao = padronizarStatusProcessual(p['Status']);
    const statusFinal = statusPadrao || determinarStatusProcessual(p);
    const etapa = obterEtapaPorStatusInformado(statusFinal);
    p['Status'] = statusFinal;
    p.__etapaId = etapa ? etapa.id : '';
    if (etapa) {
      contagemEtapas[etapa.id] += 1;
    } else {
      const chave = statusFinal && statusFinal !== '' ? statusFinal : 'Sem status';
      contagemOutros[chave] = (contagemOutros[chave] || 0) + 1;
    }
    const nome = p['Comunidade'] || p['Município'] || p['Estado'] || p.ID;
    mapIdParaNome[p.ID] = nome;
    mapIdParaPoco[p.ID] = p;
  });

  const valoresPrest = shPrest.getDataRange().getValues();
  let prestacoes = [];
  if (valoresPrest.length > 1) {
    const headersPrest = valoresPrest.shift();
    prestacoes = valoresPrest.map(r => Object.fromEntries(headersPrest.map((h, i) => [h, r[i]])));
  }

  const valoresDoadores = shDoadores.getDataRange().getValues();
  let doadores = [];
  if (valoresDoadores.length > 1) {
    const headersDoadores = valoresDoadores.shift();
    doadores = valoresDoadores.map(r => Object.fromEntries(headersDoadores.map((h, i) => [h, r[i]])));
  }

  let contatos = [];
  if (shContatos) {
    const valoresContatos = shContatos.getDataRange().getValues();
    if (valoresContatos.length > 1) {
      const headersContatos = valoresContatos.shift();
      contatos = valoresContatos.map(r => Object.fromEntries(headersContatos.map((h, i) => [h, r[i]])));
    }
  }

  const numero = normalizarNumero;

  const totalPocos = pocos.length;
  const concluidos = (contagemEtapas.conclusao || 0) + (contagemEtapas.pagamento || 0);
  const emExecucao = contagemEtapas.instalacao || 0;
  const planejados = (contagemEtapas.solicitado || 0) + (contagemEtapas.orcamento || 0);

  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']), 0);
  const investimentoPlanejado = pocos.reduce((acc, p) => acc + numero(p.OrcamentoAprovado || p['Investimento']), 0);
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p.OrcamentoExecutado || p['Valor Realizado']), 0);
  const beneficiariosTotal = pocos.reduce((acc, p) => acc + Number(p['Beneficiários'] || 0), 0);

  const porEstadoMapa = {};
  const pipelineContatosMapa = {};
  const dadosPorAno = {};
  const acoesPosInstalacao = [];
  const alertas = [];
  const alertasSet = new Set();
  const adicionarAlerta = alerta => {
    const chave = `${alerta.poco}__${alerta.motivo}`;
    if (alertasSet.has(chave)) return;
    alertasSet.add(chave);
    alertas.push(alerta);
  };

  let metrosPerfuradosTotal = 0;
  let vazaoHoraTotal = 0;
  let pocosSecos = 0;
  let pocosArtesianos = 0;

  const parseData = valor => {
    if (!valor) return null;
    if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
    if (typeof valor === 'string') {
      const dataTexto = extrairDataDeTexto(valor);
      if (dataTexto) return dataTexto;
      const dataLivre = new Date(valor);
      if (!isNaN(dataLivre.getTime())) return dataLivre;
    }
    return null;
  };

  const obterDataReferencia = poco => {
    return parseData(poco['DataConclusao'])
      || parseData(poco['DataInstalacao'])
      || parseData(poco['DataSolicitacao'])
      || parseData(poco['DataCadastro']);
  };

  const nomeDoPoco = poco => poco['Comunidade'] ? `${poco['Comunidade']} - ${poco['Município']}` : poco['Município'] || poco['Estado'] || 'Sem identificação';

  pocos.forEach(p => {
    const estado = p['Estado'] || 'Não informado';
    if (!porEstadoMapa[estado]) {
      porEstadoMapa[estado] = { estado, pocos: 0, beneficiarios: 0 };
    }
    porEstadoMapa[estado].pocos += 1;
    porEstadoMapa[estado].beneficiarios += Number(p['Beneficiários'] || 0);

    const statusContato = p['StatusContato'] || 'Sem registro';
    pipelineContatosMapa[statusContato] = (pipelineContatosMapa[statusContato] || 0) + 1;

    const metros = numero(p['Profundidade (m)']);
    const vazaoHora = numero(p['Vazão (L/H)']);
    metrosPerfuradosTotal += metros;
    vazaoHoraTotal += vazaoHora;

    const situacaoLower = (p['SituacaoHidrica'] || '').toString().toLowerCase();
    if (situacaoLower.includes('seco')) {
      pocosSecos += 1;
      adicionarAlerta({
        poco: nomeDoPoco(p),
        motivo: 'Poço identificado como seco (sem vazão produtiva)',
        responsavel: p['ResponsavelContato'] || '-',
        proximaAcao: p['AcoesPosInstalacao'] || p['ProximaAcao'] || '-',
        status: p['Status'] || '-'
      });
    }

    const tipoLower = (p['TipoPoco'] || '').toString().toLowerCase();
    if (tipoLower.includes('artesian')) {
      pocosArtesianos += 1;
    }

    const dataReferencia = obterDataReferencia(p);
    const anoReferencia = dataReferencia ? dataReferencia.getFullYear() : null;
    const etapaId = p.__etapaId || obterIdEtapaPorStatus(p['Status']);
    if (anoReferencia !== null && (etapaId === 'conclusao' || etapaId === 'pagamento')) {
      if (!dadosPorAno[anoReferencia]) {
        dadosPorAno[anoReferencia] = {
          totalInstalacoes: 0,
          investimento: 0,
          beneficiarios: 0,
          metros: 0,
          vazaoDia: 0
        };
      }
      const referencia = dadosPorAno[anoReferencia];
      const valorExecutado = numero(p.OrcamentoExecutado || p['Valor Realizado']);
      referencia.totalInstalacoes += 1;
      referencia.investimento += valorExecutado || numero(p.OrcamentoAprovado || p['Investimento']);
      referencia.beneficiarios += Number(p['Beneficiários'] || 0);
      referencia.metros += metros;
      referencia.vazaoDia += vazaoHora * 24;
    }

    if ((p['AcoesPosInstalacao'] && String(p['AcoesPosInstalacao']).trim()) || (p['UsoAguaComunitario'] && String(p['UsoAguaComunitario']).trim())) {
      acoesPosInstalacao.push({
        poco: nomeDoPoco(p),
        estado: p['Estado'] || '-',
        situacaoHidrica: p['SituacaoHidrica'] || 'Não informada',
        status: p['Status'] || '-',
        acoes: p['AcoesPosInstalacao'] || '',
        usos: p['UsoAguaComunitario'] || ''
      });
    }
  });

  const proximasAcoes = pocos
    .filter(p => p['ProximaAcao'])
    .map(p => {
      const ultimoContato = p['UltimoContato'] ? new Date(p['UltimoContato']) : null;
      const diasSemContato = ultimoContato ? Math.max(Math.floor((new Date().getTime() - ultimoContato.getTime()) / 86400000), 0) : null;
      const valorPrevisto = numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']);
      const valorExecutado = numero(p.OrcamentoExecutado || p['Valor Realizado']);
      const gapFinanceiro = valorPrevisto - valorExecutado;
      const etapaId = p.__etapaId || obterIdEtapaPorStatus(p['Status']);
      const statusFinalizado = etapaId === 'conclusao' || etapaId === 'pagamento';
      if ((diasSemContato != null && diasSemContato > 12) || (!statusFinalizado && gapFinanceiro > 40000)) {
        adicionarAlerta({
          poco: nomeDoPoco(p),
          motivo: diasSemContato != null && diasSemContato > 12
            ? `Sem contato há ${diasSemContato} dias`
            : `Gap financeiro de ${gapFinanceiro.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`,
          responsavel: p['ResponsavelContato'] || '-',
          proximaAcao: p['ProximaAcao'] || '-',
          status: p['Status'] || '-'
        });
      }
      return {
        poco: nomeDoPoco(p),
        responsavel: p['ResponsavelContato'] || '-',
        contato: p['ContatoInstalacao'] || '-',
        proximaAcao: p['ProximaAcao'],
        statusContato: p['StatusContato'] || 'Sem registro',
        ultimoContato: ultimoContato ? ultimoContato.toISOString() : '',
        diasSemContato,
        impacto: p['ImpactoNoStatus'] || '',
        situacaoHidrica: p['SituacaoHidrica'] || 'Não informada'
      };
    })
    .sort((a, b) => {
      const dataA = a.ultimoContato ? new Date(a.ultimoContato) : null;
      const dataB = b.ultimoContato ? new Date(b.ultimoContato) : null;
      if (!dataA && !dataB) return a.poco.localeCompare(b.poco);
      if (!dataA) return 1;
      if (!dataB) return -1;
      return dataA.getTime() - dataB.getTime();
    })
    .slice(0, 6);

  let ultimosContatos = contatos
    .map(c => ({
      poco: mapIdParaNome[c['PoçoID']] || c['PoçoID'],
      responsavel: c['ResponsavelContato'] || '-',
      contato: c['ContatoExterno'] || '-',
      data: c['DataContato'] ? new Date(c['DataContato']).toISOString() : '',
      resumo: c['Resumo'] || '',
      status: c['StatusContato'] || 'Sem registro'
    }))
    .sort((a, b) => new Date(b.data) - new Date(a.data))
    .slice(0, 6);

  const gastosPorCategoria = {};
  prestacoes.forEach(d => {
    const cat = d['Categoria'] || 'Outros';
    gastosPorCategoria[cat] = (gastosPorCategoria[cat] || 0) + numero(d['Valor']);
  });

  const pipelineContatos = Object.keys(pipelineContatosMapa).map(status => ({
    status,
    total: pipelineContatosMapa[status],
    percentual: totalPocos ? (pipelineContatosMapa[status] / totalPocos) * 100 : 0
  })).sort((a, b) => b.total - a.total);

  const historicoInstalacoes = Object.keys(dadosPorAno).map(ano => {
    const registro = dadosPorAno[ano];
    return {
      ano: Number(ano),
      total: registro.totalInstalacoes,
      investimento: registro.investimento,
      beneficiarios: registro.beneficiarios,
      metros: registro.metros,
      vazaoDia: registro.vazaoDia
    };
  }).sort((a, b) => b.ano - a.ano);

  const anoAtual = new Date().getFullYear();
  const dadosAnoAtual = dadosPorAno[anoAtual] || {
    totalInstalacoes: 0,
    investimento: 0,
    beneficiarios: 0,
    metros: 0,
    vazaoDia: 0
  };

  const hidricos = {
    totalMonitorados: totalPocos,
    pocosSecos,
    pocosArtesianos,
    totalMetrosPerfurados: metrosPerfuradosTotal,
    totalVazaoDia: vazaoHoraTotal * 24,
    mediaVazaoDia: totalPocos ? (vazaoHoraTotal * 24) / totalPocos : 0
  };

  const acoesPosInstalacaoOrdenadas = acoesPosInstalacao
    .slice()
    .sort((a, b) => {
      const estadoA = (a.estado || '').toString();
      const estadoB = (b.estado || '').toString();
      const comparacaoEstado = estadoA.localeCompare(estadoB);
      if (comparacaoEstado !== 0) return comparacaoEstado;
      return a.poco.localeCompare(b.poco);
    });

  const distribuicaoStatus = PROCESS_STAGES.map(stage => ({
    status: stage.status,
    total: contagemEtapas[stage.id] || 0
  }));
  Object.keys(contagemOutros).forEach(status => {
    distribuicaoStatus.push({ status, total: contagemOutros[status] });
  });

  const doacoesTotais = doadores.reduce((acc, d) => acc + numero(d['ValorDoado']), 0);

  const doadoresDestaque = doadores.map(d => {
    const pocoses = (d['PoçosVinculados'] || '').split(',').map(id => id.trim()).filter(Boolean);
    const beneficiariosApoiados = pocoses.reduce((acc, id) => {
      const poco = mapIdParaPoco[id];
      return acc + (poco ? Number(poco['Beneficiários'] || 0) : 0);
    }, 0);
    return {
      nome: d['Nome'] || 'Sem identificação',
      valor: numero(d['ValorDoado']),
      quantidadePocos: pocoses.length,
      beneficiariosApoiados
    };
  }).sort((a, b) => b.valor - a.valor).slice(0, 5);

  return {
    totais: {
      totalPocos,
      concluidos,
      taxaConclusao: totalPocos ? (concluidos / totalPocos) * 100 : 0,
      beneficiariosTotal,
      mediaBeneficiarios: totalPocos ? beneficiariosTotal / totalPocos : 0,
      doadoresAtivos: doadores.length,
      doacoesTotais,
      investimentoPrevisto,
      investimentoPlanejado,
      investimentoRealizado,
      gapFinanceiro: investimentoPrevisto - investimentoRealizado
    },
    distribuicaoStatus,
    pipelineContatos,
    porEstado: Object.values(porEstadoMapa).sort((a, b) => b.pocos - a.pocos),
    proximasAcoes,
    ultimosContatos,
    gastosPorCategoria: Object.keys(gastosPorCategoria)
      .map(cat => ({ categoria: cat, valor: gastosPorCategoria[cat] }))
      .sort((a, b) => b.valor - a.valor),
    doadoresDestaque,
    alertas,
    hidricos,
    indicadoresLegado: {
      anoReferencia: anoAtual,
      totalInstaladoAno: dadosAnoAtual.totalInstalacoes || 0,
      investimentoAno: dadosAnoAtual.investimento || 0,
      beneficiariosAno: dadosAnoAtual.beneficiarios || 0,
      metrosPerfuradosAno: dadosAnoAtual.metros || 0,
      vazaoDiaAno: dadosAnoAtual.vazaoDia || 0,
      historico: historicoInstalacoes
    },
    acoesPosInstalacao: acoesPosInstalacaoOrdenadas
  };
}

function obterResumoGestao() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPocos = ss.getSheetByName('Poços');
  const shEmpresas = ss.getSheetByName('Empresas');
  const shContatos = ss.getSheetByName('Contatos');

  const valoresPocos = shPocos.getDataRange().getValues();
  const headersPocos = valoresPocos.shift();
  const pocos = valoresPocos.map(r => Object.fromEntries(headersPocos.map((h, i) => [h, r[i]])));

  const numero = normalizarNumero;

  const contagemEtapas = PROCESS_STAGES.reduce((acc, stage) => {
    acc[stage.id] = 0;
    return acc;
  }, {});
  pocos.forEach(p => {
    const statusPadrao = padronizarStatusProcessual(p['Status']);
    const statusFinal = statusPadrao || determinarStatusProcessual(p);
    const etapa = obterEtapaPorStatusInformado(statusFinal);
    p['Status'] = statusFinal;
    p.__etapaId = etapa ? etapa.id : '';
    if (etapa) {
      contagemEtapas[etapa.id] += 1;
    }
  });

  let contatos = [];
  if (shContatos) {
    const valoresContatos = shContatos.getDataRange().getValues();
    if (valoresContatos.length > 1) {
      const headersContatos = valoresContatos.shift();
      contatos = valoresContatos.map(r => Object.fromEntries(headersContatos.map((h, i) => [h, r[i]])));
    }
  }

  const contatosPorPoco = {};
  contatos.forEach(c => {
    const id = c['PoçoID'];
    if (!contatosPorPoco[id]) contatosPorPoco[id] = [];
    contatosPorPoco[id].push(c);
  });

  const totalPocos = pocos.length;
  const concluidos = (contagemEtapas.conclusao || 0) + (contagemEtapas.pagamento || 0);
  const emExecucao = contagemEtapas.instalacao || 0;
  const planejados = (contagemEtapas.solicitado || 0) + (contagemEtapas.orcamento || 0);
  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']), 0);
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p.OrcamentoExecutado || p['Valor Realizado']), 0);

  const alertas = [];
  const andamento = pocos.map(p => {
    const valorPrevisto = numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']);
    const valorExecutado = numero(p.OrcamentoExecutado || p['Valor Realizado']);
    const gapFinanceiro = valorPrevisto - valorExecutado;
    const ultimoContato = p['UltimoContato'] ? new Date(p['UltimoContato']) : null;
    const diasSemContato = ultimoContato ? Math.max(Math.floor((new Date().getTime() - ultimoContato.getTime()) / 86400000), 0) : null;
    const etapaId = p.__etapaId || obterIdEtapaPorStatus(p['Status']);
    const statusFinalizado = etapaId === 'conclusao' || etapaId === 'pagamento';
    if ((diasSemContato != null && diasSemContato > 12) || (gapFinanceiro > 40000 && !statusFinalizado)) {
      alertas.push({
        poco: p['Comunidade'] ? `${p['Comunidade']} - ${p['Município']}` : p['Município'] || p['Estado'] || 'Sem identificação',
        motivo: diasSemContato != null && diasSemContato > 12
          ? `Sem contato há ${diasSemContato} dias`
          : `Gap financeiro de ${gapFinanceiro.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`,
        responsavel: p['ResponsavelContato'] || '-',
        status: p['Status'] || '-',
        proximaAcao: p['ProximaAcao'] || '-'
      });
    }

    const listaContatos = (contatosPorPoco[p.ID] || []).sort((a, b) => new Date(b['DataContato']) - new Date(a['DataContato']));
    const ultimoRegistro = listaContatos[0];

    return {
      id: p.ID,
      nome: p['Comunidade'] || p['Município'] || p['Estado'] || 'Sem identificação',
      local: `${p['Município'] || 'Sem município'} - ${p['Estado'] || 'Sem estado'}`,
      status: p['Status'] || '-',
      responsavel: p['ResponsavelContato'] || '-',
      empresa: p['Empresa Responsável'] || '-',
      proximaAcao: p['ProximaAcao'] || '-',
      statusContato: p['StatusContato'] || 'Sem registro',
      ultimoContato: ultimoRegistro ? new Date(ultimoRegistro['DataContato']).toISOString() : (ultimoContato ? ultimoContato.toISOString() : ''),
      diasSemContato,
      perfuracao: p['DataSolicitacao'] ? new Date(p['DataSolicitacao']).toISOString() : '-',
      instalacao: p['DataInstalacao'] ? new Date(p['DataInstalacao']).toISOString() : '-',
      valorPrevisto,
      valorExecutado,
      gapFinanceiro,
      impacto: p['ImpactoNoStatus'] || ''
    };
  });

  const fornecedoresMapa = {};
  andamento.forEach(item => {
    const nome = item.empresa || 'Sem fornecedor atribuído';
    if (!fornecedoresMapa[nome]) {
      fornecedoresMapa[nome] = {
        fornecedor: nome,
        pocosAtendidos: 0,
        valorPrevisto: 0,
        valorExecutado: 0,
        status: new Set()
      };
    }
    fornecedoresMapa[nome].pocosAtendidos += 1;
    fornecedoresMapa[nome].valorPrevisto += item.valorPrevisto;
    fornecedoresMapa[nome].valorExecutado += item.valorExecutado;
    fornecedoresMapa[nome].status.add(item.status);
  });

  let fornecedores = Object.values(fornecedoresMapa).map(f => ({
    fornecedor: f.fornecedor,
    pocosAtendidos: f.pocosAtendidos,
    valorPrevisto: f.valorPrevisto,
    valorExecutado: f.valorExecutado,
    status: Array.from(f.status).join(', ')
  }));

  if (shEmpresas) {
    const valoresEmpresas = shEmpresas.getDataRange().getValues();
    if (valoresEmpresas.length > 1) {
      const headersEmpresas = valoresEmpresas.shift();
      const empresas = valoresEmpresas.map(r => Object.fromEntries(headersEmpresas.map((h, i) => [h, r[i]])));
      fornecedores = fornecedores.map(f => {
        const empresaInfo = empresas.find(e => (e['NomeEmpresa'] || '').toLowerCase() === (f.fornecedor || '').toLowerCase());
        return empresaInfo
          ? Object.assign({}, f, { contato: empresaInfo['Contato'] || '', observacoes: empresaInfo['Observações'] || '' })
          : Object.assign({}, f, { contato: '', observacoes: '' });
      });
    }
  }

  const cronograma = [];
  andamento.forEach(item => {
    if (item.perfuracao && item.perfuracao !== '-') {
      const data = new Date(item.perfuracao);
      cronograma.push({
        poco: item.nome,
        etapa: 'Solicitação',
        descricao: 'Solicitação registrada',
        data: !isNaN(data.getTime()) ? data.toISOString() : '',
        status: 'Concluída'
      });
    }
    if (item.instalacao && item.instalacao !== '-') {
      const data = new Date(item.instalacao);
      cronograma.push({
        poco: item.nome,
        etapa: 'Instalação',
        descricao: 'Execução do sistema de abastecimento',
        data: !isNaN(data.getTime()) ? data.toISOString() : '',
        status: 'Concluída'
      });
    }
  });

  cronograma.sort((a, b) => {
    if (!a.data && !b.data) return 0;
    if (!a.data) return 1;
    if (!b.data) return -1;
    return new Date(a.data) - new Date(b.data);
  });

  return {
    resumo: {
      totalPocos,
      planejados,
      emExecucao,
      concluidos,
      investimentoPrevisto,
      investimentoRealizado,
      gapFinanceiro: investimentoPrevisto - investimentoRealizado
    },
    andamento,
    alertas,
    fornecedores,
    cronograma
  };
}

function obterAnaliseImpacto() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPocos = ss.getSheetByName('Poços');
  const shDoadores = ss.getSheetByName('Doadores');
  const shContatos = ss.getSheetByName('Contatos');

  const valoresPocos = shPocos.getDataRange().getValues();
  const headersPocos = valoresPocos.shift();
  const pocos = valoresPocos.map(r => Object.fromEntries(headersPocos.map((h, i) => [h, r[i]])));

  const contagemEtapas = PROCESS_STAGES.reduce((acc, stage) => {
    acc[stage.id] = 0;
    return acc;
  }, {});
  const contagemOutros = {};
  pocos.forEach(p => {
    const statusPadrao = padronizarStatusProcessual(p['Status']);
    const statusFinal = statusPadrao || determinarStatusProcessual(p);
    const etapa = obterEtapaPorStatusInformado(statusFinal);
    p['Status'] = statusFinal;
    p.__etapaId = etapa ? etapa.id : '';
    if (etapa) {
      contagemEtapas[etapa.id] += 1;
    } else {
      const chave = statusFinal && statusFinal !== '' ? statusFinal : 'Sem status';
      contagemOutros[chave] = (contagemOutros[chave] || 0) + 1;
    }
  });

  let doadores = [];
  if (shDoadores) {
    const valoresDoadores = shDoadores.getDataRange().getValues();
    if (valoresDoadores.length > 1) {
      const headersDoadores = valoresDoadores.shift();
      doadores = valoresDoadores.map(r => Object.fromEntries(headersDoadores.map((h, i) => [h, r[i]])));
    }
  }

  let contatos = [];
  if (shContatos) {
    const valoresContatos = shContatos.getDataRange().getValues();
    if (valoresContatos.length > 1) {
      const headersContatos = valoresContatos.shift();
      contatos = valoresContatos.map(r => Object.fromEntries(headersContatos.map((h, i) => [h, r[i]])));
    }
  }

  const numero = normalizarNumero;
  const totalBeneficiarios = pocos.reduce((acc, p) => acc + Number(p['Beneficiários'] || 0), 0);
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p.OrcamentoExecutado || p['Valor Realizado']), 0);
  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']), 0);
  const volumeDiario = pocos.reduce((acc, p) => acc + Number(p['Vazão (L/H)'] || 0) * 12, 0);

  const mapDoador = {};
  doadores.forEach(d => {
    const valor = numero(d['ValorDoado']);
    const pocoses = (d['PoçosVinculados'] || '').split(',').map(id => id.trim()).filter(Boolean);
    const beneficiariosApoiados = pocoses.reduce((acc, id) => {
      const poco = pocos.find(p => p.ID === id);
      return acc + (poco ? Number(poco['Beneficiários'] || 0) : 0);
    }, 0);
    mapDoador[d.ID] = Object.assign({}, d, { valor, pocoses, beneficiariosApoiados });
  });

  const doadoresImpacto = Object.values(mapDoador)
    .map(d => ({
      nome: d['Nome'] || 'Sem identificação',
      valor: d.valor,
      beneficiariosApoiados: d.beneficiariosApoiados,
      quantidadePocos: d.pocoses.length
    }))
    .sort((a, b) => b.valor - a.valor);

  const pocoImpacto = pocos.map(p => {
    const valorPrevisto = numero(p.OrcamentoPrevisto || p['Valor Previsto Perfuração']);
    const valorExecutado = numero(p.OrcamentoExecutado || p['Valor Realizado']);
    const doadoresIds = (p['Doadores'] || '').split(',').map(id => id.trim()).filter(Boolean);
    const doadoresNomes = doadoresIds.map(id => (mapDoador[id] ? mapDoador[id]['Nome'] : '')).filter(Boolean);
    return {
      id: p.ID,
      nome: p['Comunidade'] || p['Município'] || p['Estado'] || 'Sem identificação',
      local: `${p['Município'] || 'Sem município'} - ${p['Estado'] || 'Sem estado'}`,
      status: p['Status'] || '-',
      beneficiarios: Number(p['Beneficiários'] || 0),
      doadores: doadoresNomes.join(', ') || 'Sem doador vinculado',
      valorPrevisto,
      valorExecutado,
      gapFinanceiro: valorPrevisto - valorExecutado,
      vazao: Number(p['Vazão (L/H)'] || 0)
    };
  });

  const timeline = contatos
    .map(c => ({
      data: c['DataContato'] ? new Date(c['DataContato']).toISOString() : '',
      poco: pocos.find(p => p.ID === c['PoçoID'])?.['Comunidade'] || c['PoçoID'],
      resumo: c['Resumo'] || '-',
      responsavel: c['ResponsavelContato'] || '-',
      statusContato: c['StatusContato'] || 'Sem registro'
    }))
    .sort((a, b) => new Date(b.data) - new Date(a.data))
    .slice(0, 8);

  const distribuicaoStatus = PROCESS_STAGES.map(stage => ({
    status: stage.status,
    total: contagemEtapas[stage.id] || 0
  }));
  Object.keys(contagemOutros).forEach(status => {
    distribuicaoStatus.push({ status, total: contagemOutros[status] });
  });

  const regioesMapa = {};
  pocos.forEach(p => {
    const estado = p['Estado'] || 'Não informado';
    if (!regioesMapa[estado]) {
      regioesMapa[estado] = { estado, beneficiarios: 0, pocos: 0 };
    }
    regioesMapa[estado].beneficiarios += Number(p['Beneficiários'] || 0);
    regioesMapa[estado].pocos += 1;
  });

  const metricas = {
    beneficiariosTotais: totalBeneficiarios,
    familiasEstimadas: totalBeneficiarios ? Math.round(totalBeneficiarios / 4) : 0,
    volumeAguaDiario: volumeDiario,
    investimentoRealizado,
    investimentoPrevisto,
    custoPorPessoa: totalBeneficiarios ? investimentoRealizado / totalBeneficiarios : 0
  };

  return {
    metricas,
    doadores: doadoresImpacto,
    pocos: pocoImpacto,
    timeline,
    distribuicaoStatus,
    regioes: Object.values(regioesMapa)
  };
}

