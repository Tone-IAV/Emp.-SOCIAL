// ===========================
// CONFIGURAÇÕES GERAIS
// ===========================
const SPREADSHEET_ID = '1ajlFoT0kkAwOYVFymFMs5ipiEkmvQdNZzwVJx6B98FM';
const DRIVE_FOLDER_ID = '1g_8lgL55WAb32E6XQ4dl2KMX-AfhvXMY';

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
        sh.appendRow([
          'ID', 'Estado', 'Município', 'Comunidade', 'Latitude', 'Longitude',
          'Beneficiários', 'Investimento', 'Vazão (L/H)', 'Profundidade (m)',
          'Perfuração', 'Instalação', 'Doador', 'Status',
          'Valor Previsto Perfuração', 'Valor Previsto Instalação',
          'Empresa Responsável', 'Observações', 'Valor Realizado', 'Doadores', 'DataCadastro',
          'ResponsavelContato', 'ContatoInstalacao', 'TelefoneContato', 'StatusContato',
          'ProximaAcao', 'UltimoContato', 'ImpactoNoStatus'
        ]);
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
  const id = Utilities.getUuid();
  const data = [
    id, poco.estado, poco.municipio, poco.comunidade, poco.latitude, poco.longitude,
    poco.beneficiarios, poco.investimento, poco.vazao, poco.profundidade,
    poco.perfuracao, poco.instalacao, poco.doador, poco.status,
    poco.valorPerf, poco.valorInst, poco.empresa, poco.obs,
    Number(poco.valorRealizado) || 0, poco.doadores || '', new Date(),
    poco.responsavelContato || '', poco.contatoInstalacao || '', poco.telefoneContato || '',
    poco.statusContato || '', poco.proximaAcao || '', poco.ultimoContato ? new Date(poco.ultimoContato) : '',
    poco.impactoNoStatus || ''
  ];
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
    doador.valorDoado || '',
    new Date()
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
    despesa.data,
    despesa.descricao,
    despesa.valor,
    despesa.comprovanteURL
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
  pocos.forEach(p => {
    const nome = p['Comunidade'] || p['Município'] || p['Estado'] || p.ID;
    mapIdParaNome[p.ID] = nome;
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

  const numero = valor => {
    if (valor instanceof Date) return 0;
    if (typeof valor === 'number') return valor;
    if (!valor) return 0;
    return Number(String(valor).replace(/[^0-9,-]+/g, '').replace(',', '.')) || 0;
  };

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
  pocos.forEach(p => {
    const estado = p['Estado'] || 'Não informado';
    if (!porEstadoMapa[estado]) {
      porEstadoMapa[estado] = { estado, pocos: 0, beneficiarios: 0 };
    }
    porEstadoMapa[estado].pocos += 1;
    porEstadoMapa[estado].beneficiarios += Number(p['Beneficiários'] || 0);
  });

  const pipelineContatosMapa = {};
  pocos.forEach(p => {
    const status = p['StatusContato'] || 'Sem registro';
    pipelineContatosMapa[status] = (pipelineContatosMapa[status] || 0) + 1;
  });

  const proximasAcoes = pocos
    .filter(p => p['ProximaAcao'])
    .map(p => {
      const ultimoContato = p['UltimoContato'] ? new Date(p['UltimoContato']) : null;
      const diasSemContato = ultimoContato ? Math.max(Math.floor((new Date().getTime() - ultimoContato.getTime()) / 86400000), 0) : null;
      return {
        poco: p['Comunidade'] ? `${p['Comunidade']} - ${p['Município']}` : p['Município'] || p['Estado'] || 'Sem identificação',
        responsavel: p['ResponsavelContato'] || '-',
        contato: p['ContatoInstalacao'] || '-',
        proximaAcao: p['ProximaAcao'],
        statusContato: p['StatusContato'] || 'Sem registro',
        ultimoContato: ultimoContato ? ultimoContato.toISOString() : '',
        diasSemContato,
        impacto: p['ImpactoNoStatus'] || ''
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

  const distribuicaoStatus = [
    { status: 'Planejado', total: planejados },
    { status: 'Em execução', total: emExecucao },
    { status: 'Concluído', total: concluidos }
  ];
  if (outros > 0) distribuicaoStatus.push({ status: 'Outros', total: outros });

  const doacoesTotais = doadores.reduce((acc, d) => acc + numero(d['ValorDoado']), 0);

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
      .sort((a, b) => b.valor - a.valor)
  };
}

