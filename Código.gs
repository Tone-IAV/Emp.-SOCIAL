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
          'Empresa Responsável', 'Observações', 'DataCadastro'
        ]);
      } else if (name === 'Doadores') {
        sh.appendRow(['ID', 'Nome', 'Email', 'Telefone', 'ValorDoado', 'DataDoacao']);
      } else if (name === 'PrestaçãoContas') {
        sh.appendRow(['PoçoID', 'Data', 'Descrição', 'Valor', 'ComprovanteURL']);
      }
    }
  });
  return 'Guias verificadas/criadas com sucesso.';
}


// Listar poços
function listarPocos() {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

// Salvar novo poço

function salvarPoco(poco) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Poços');
  const id = Utilities.getUuid();
  const data = [
    id, poco.estado, poco.municipio, poco.comunidade, poco.latitude, poco.longitude,
    poco.beneficiarios, poco.investimento, poco.vazao, poco.profundidade,
    poco.perfuracao, poco.instalacao, poco.doador, poco.status,
    poco.valorPerf, poco.valorInst, poco.empresa, poco.obs, new Date()
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
  const valRealIndex = headers.indexOf('ValorRealizado');

  for (let i = 0; i < values.length; i++) {
    if (values[i][idIndex] === despesa.pocoId) {
      const atual = Number(values[i][valRealIndex]) || 0;
      shPocos.getRange(i + 2, valRealIndex + 1).setValue(atual + Number(despesa.valor));
      break;
    }
  }

  return { success: true };
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

