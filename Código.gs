// ===========================
// CONFIGURAÇÕES GERAIS
// ===========================
const SPREADSHEET_ID = '1ajlFoT0kkAwOYVFymFMs5ipiEkmvQdNZzwVJx6B98FM';
const DRIVE_FOLDER_ID = '1g_8lgL55WAb32E6XQ4dl2KMX-AfhvXMY';

const COLUNAS_POCOS = [
  'ID', 'Estado', 'Município', 'Comunidade', 'Latitude', 'Longitude',
  'Beneficiários', 'Investimento', 'Vazão (L/H)', 'Profundidade (m)',
  'Perfuração', 'Instalação', 'Doador', 'Status',
  'Valor Previsto Perfuração', 'Valor Previsto Instalação',
  'Empresa Responsável', 'Observações', 'Valor Realizado', 'Doadores', 'DataCadastro',
  'ResponsavelContato', 'ContatoInstalacao', 'TelefoneContato', 'StatusContato',
  'ProximaAcao', 'UltimoContato', 'ImpactoNoStatus',
  'TipoPoco', 'SituacaoHidrica', 'AcoesPosInstalacao', 'UsoAguaComunitario'
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
  return values.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

// Salvar novo poço

function salvarPoco(poco) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  garantirColunas(sh, COLUNAS_POCOS);
  const id = Utilities.getUuid();
  const registro = {
    ID: id,
    Estado: poco.estado,
    Município: poco.municipio,
    Comunidade: poco.comunidade,
    Latitude: poco.latitude,
    Longitude: poco.longitude,
    Beneficiários: poco.beneficiarios,
    Investimento: poco.investimento,
    'Vazão (L/H)': poco.vazao,
    'Profundidade (m)': poco.profundidade,
    Perfuração: poco.perfuracao,
    Instalação: poco.instalacao,
    Doador: poco.doador,
    Status: poco.status,
    'Valor Previsto Perfuração': poco.valorPerf,
    'Valor Previsto Instalação': poco.valorInst,
    'Empresa Responsável': poco.empresa,
    Observações: poco.obs,
    'Valor Realizado': Number(poco.valorRealizado) || 0,
    Doadores: poco.doadores || '',
    DataCadastro: new Date(),
    ResponsavelContato: poco.responsavelContato || '',
    ContatoInstalacao: poco.contatoInstalacao || '',
    TelefoneContato: poco.telefoneContato || '',
    StatusContato: poco.statusContato || '',
    ProximaAcao: poco.proximaAcao || '',
    UltimoContato: poco.ultimoContato ? new Date(poco.ultimoContato) : '',
    ImpactoNoStatus: poco.impactoNoStatus || '',
    TipoPoco: poco.tipoPoco || '',
    SituacaoHidrica: poco.situacaoHidrica || '',
    AcoesPosInstalacao: poco.acoesPosInstalacao || '',
    UsoAguaComunitario: poco.usoAguaComunitario || ''
  };

  const data = COLUNAS_POCOS.map(coluna => {
    const valor = registro[coluna];
    if (valor === undefined) return '';
    return valor;
  });
  sh.appendRow(data);
  return { success: true, id };
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
  const valRealIndex = headers.indexOf('Valor Realizado');

  for (let i = 0; i < values.length; i++) {
    if (values[i][idIndex] === despesa.pocoId) {
      const atual = Number(values[i][valRealIndex]) || 0;
      shPocos.getRange(i + 2, valRealIndex + 1).setValue(atual + Number(despesa.valor));
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

  const pocos = shPocos.getDataRange().getValues();
  const headersPocos = pocos.shift();
  const poco = pocos.map(r => Object.fromEntries(headersPocos.map((h, i) => [h, r[i]])))
                    .find(p => p.ID === pocoId);

  const prestacoes = shPrest.getDataRange().getValues();
  const headersPrest = prestacoes.shift();
  const despesas = prestacoes.map(r => Object.fromEntries(headersPrest.map((h, i) => [h, r[i]])))
                             .filter(d => d.PoçoID === pocoId);

  return { poco, despesas };
}

function atualizarStatusPoco(id, novoStatus) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idIndex = headers.indexOf('ID');
  const statusIndex = headers.indexOf('Status');
  for (let i = 0; i < values.length; i++) {
    if (values[i][idIndex] === id) {
      sh.getRange(i + 2, statusIndex + 1).setValue(novoStatus);
      break;
    }
  }
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

  const mapIdParaNome = {};
  const mapIdParaPoco = {};
  pocos.forEach(p => {
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
  const concluidos = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'concluído').length;
  const emExecucao = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'em execução').length;
  const planejados = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'planejado').length;
  const outros = totalPocos - (concluidos + emExecucao + planejados);

  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']), 0);
  const investimentoPlanejado = pocos.reduce((acc, p) => acc + numero(p['Investimento']), 0);
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p['Valor Realizado']), 0);
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
    return parseData(poco['Instalação'])
      || parseData(poco['Perfuração'])
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
    const statusLower = (p['Status'] || '').toLowerCase();
    if (anoReferencia !== null && statusLower === 'concluído') {
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
      const valorExecutado = numero(p['Valor Realizado']);
      referencia.totalInstalacoes += 1;
      referencia.investimento += valorExecutado || numero(p['Investimento']);
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
      const valorPrevisto = numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']);
      const valorExecutado = numero(p['Valor Realizado']);
      const gapFinanceiro = valorPrevisto - valorExecutado;
      const statusLower = (p['Status'] || '').toLowerCase();
      if ((diasSemContato != null && diasSemContato > 12) || (statusLower !== 'concluído' && gapFinanceiro > 40000)) {
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

  const distribuicaoStatus = [
    { status: 'Planejado', total: planejados },
    { status: 'Em execução', total: emExecucao },
    { status: 'Concluído', total: concluidos }
  ];
  if (outros > 0) distribuicaoStatus.push({ status: 'Outros', total: outros });

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
  const concluidos = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'concluído').length;
  const emExecucao = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'em execução').length;
  const planejados = pocos.filter(p => (p['Status'] || '').toLowerCase() === 'planejado').length;
  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']), 0);
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p['Valor Realizado']), 0);

  const alertas = [];
  const andamento = pocos.map(p => {
    const valorPrevisto = numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']);
    const valorExecutado = numero(p['Valor Realizado']);
    const gapFinanceiro = valorPrevisto - valorExecutado;
    const ultimoContato = p['UltimoContato'] ? new Date(p['UltimoContato']) : null;
    const diasSemContato = ultimoContato ? Math.max(Math.floor((new Date().getTime() - ultimoContato.getTime()) / 86400000), 0) : null;
    if ((diasSemContato != null && diasSemContato > 12) || (gapFinanceiro > 40000 && (p['Status'] || '').toLowerCase() !== 'concluído')) {
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
      perfuracao: p['Perfuração'] || '-',
      instalacao: p['Instalação'] || '-',
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
    if (item.perfuracao) {
      const data = extrairDataDeTexto(item.perfuracao);
      cronograma.push({
        poco: item.nome,
        etapa: 'Perfuração',
        descricao: item.perfuracao,
        data: data ? data.toISOString() : '',
        status: extrairStatusDaEtapa(item.perfuracao)
      });
    }
    if (item.instalacao) {
      const data = extrairDataDeTexto(item.instalacao);
      cronograma.push({
        poco: item.nome,
        etapa: 'Instalação',
        descricao: item.instalacao,
        data: data ? data.toISOString() : '',
        status: extrairStatusDaEtapa(item.instalacao)
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
  const investimentoRealizado = pocos.reduce((acc, p) => acc + numero(p['Valor Realizado']), 0);
  const investimentoPrevisto = pocos.reduce((acc, p) => acc + numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']), 0);
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
    const valorPrevisto = numero(p['Valor Previsto Perfuração']) + numero(p['Valor Previsto Instalação']);
    const valorExecutado = numero(p['Valor Realizado']);
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

  const distribuicaoStatusMapa = {};
  pocos.forEach(p => {
    const status = p['Status'] || 'Sem status';
    distribuicaoStatusMapa[status] = (distribuicaoStatusMapa[status] || 0) + 1;
  });

  const distribuicaoStatus = Object.keys(distribuicaoStatusMapa).map(status => ({
    status,
    total: distribuicaoStatusMapa[status]
  }));

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

