/**
 * Entry point for web app. Returns the main layout template that
 * contains the navigation menu shared across all views.
 *
 * @param {GoogleAppsScript.Events.DoGet} e Request context.
 */
var VIEW_DEFINITIONS = [
  {
    id: 'solicitacoes',
    label: 'Solicitações',
    sheet: 'Solicitações',
    fetch: 'getSolicitacoes',
    save: 'saveSolicitacao',
    partial: 'Solicitacoes'
  },
  {
    id: 'pocos',
    label: 'Poços',
    sheet: 'Poços',
    fetch: 'getPocos',
    save: 'savePoco',
    partial: 'Pocos'
  },
  {
    id: 'projetos',
    label: 'Projetos',
    sheet: 'Projetos',
    fetch: 'getProjetos',
    save: 'saveProjeto',
    partial: 'Projetos'
  },
  {
    id: 'doadores',
    label: 'Doadores',
    sheet: 'Doadores',
    fetch: 'getDoadores',
    save: 'saveDoador',
    partial: 'Doadores'
  }
];

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;

  if (page === 'detalhes-projeto') {
    var detailTemplate = HtmlService.createTemplateFromFile('ProjetoDetalhes');
    detailTemplate.pageParams = {
      view: (e.parameter && e.parameter.view) || 'projetos',
      key: (e.parameter && e.parameter.key) || ''
    };

    return detailTemplate
      .evaluate()
      .setTitle('Emp. Social - Detalhes do Projeto')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var views = VIEW_DEFINITIONS;

  var template = HtmlService.createTemplateFromFile('index');
  template.views = views;

  var requestedView = e && e.parameter && e.parameter.view;
  var availableIds = views.map(function(view) {
    return view.id;
  });
  var defaultView = availableIds[0];
  var selectedView = requestedView && availableIds.indexOf(requestedView) !== -1 ? requestedView : defaultView;

  template.currentView = selectedView;

  return template
    .evaluate()
    .setTitle('Emp. Social - Gestão de Projetos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Name of the sheet that contains the structural configuration for
 * the remaining tabs (guides).
 */
var STRUCTURE_SHEET_NAME = 'Estrutura';

/**
 * Loads the structural definition of the sheets from the configuration tab.
 * The "Estrutura" tab must contain the following columns:
 *   - Guia: name of the sheet/tab that stores the data.
 *   - Coluna: column name as it appears in the sheet header.
 *   - Rotulo (optional): human friendly label for the column.
 *   - Chave (optional): when set to "TRUE" (string), marks the primary key column.
 * Additional columns can be used for metadata and are simply ignored.
 *
 * @return {Object<string, Object>} Mapping between sheet names and their structure.
 */
function getStructure() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('structure');
  if (cached) {
    return JSON.parse(cached);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STRUCTURE_SHEET_NAME);
  if (!sheet) {
    throw new Error('A guia "' + STRUCTURE_SHEET_NAME + '" não foi encontrada.');
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) {
    throw new Error('A guia "' + STRUCTURE_SHEET_NAME + '" precisa conter cabeçalho e linhas de configuração.');
  }

  var header = values.shift();
  var indexes = header.reduce(function(acc, value, idx) {
    acc[value.toLowerCase()] = idx;
    return acc;
  }, {});

  if (indexes.guia === undefined || indexes.coluna === undefined) {
    throw new Error('As colunas "Guia" e "Coluna" são obrigatórias na guia de estrutura.');
  }

  var structure = {};

  values.forEach(function(row) {
    var sheetName = row[indexes.guia];
    var columnName = row[indexes.coluna];
    if (!sheetName || !columnName) {
      return;
    }

    sheetName = sheetName.trim();
    columnName = columnName.trim();

    if (!structure[sheetName]) {
      structure[sheetName] = {
        columns: [],
        labels: {},
        key: null
      };
    }

    structure[sheetName].columns.push(columnName);

    if (indexes.rotulo !== undefined) {
      var label = row[indexes.rotulo];
      if (label) {
        structure[sheetName].labels[columnName] = label;
      }
    }

    if (indexes.chave !== undefined) {
      var isKey = String(row[indexes.chave]).toLowerCase() === 'true';
      if (isKey) {
        structure[sheetName].key = columnName;
      }
    }
  });

  Object.keys(structure).forEach(function(name) {
    if (!structure[name].key && structure[name].columns.length > 0) {
      structure[name].key = structure[name].columns[0];
    }
  });

  cache.put('structure', JSON.stringify(structure), 300); // cache for 5 minutes
  return structure;
}

/**
 * Retrieves all data from a sheet using the column order defined in the structure.
 *
 * @param {string} sheetName - The name of the sheet to be read.
 * @return {Array<Object>} An array of objects representing the rows.
 */
function getSheetData(sheetName) {
  var structure = getStructure();
  var config = structure[sheetName];
  if (!config) {
    throw new Error('Não há estrutura configurada para a guia "' + sheetName + '".');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('A guia "' + sheetName + '" não foi encontrada.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headerValues = headerRange.getDisplayValues()[0];
  var headerIndex = headerValues.reduce(function(acc, value, idx) {
    acc[value] = idx;
    return acc;
  }, {});

  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var rows = dataRange.getDisplayValues();

  var filtered = rows
    .filter(function(row) {
      return row.join('').trim() !== '';
    })
    .map(function(row) {
      var result = {};
      config.columns.forEach(function(column) {
        var idx = headerIndex[column];
        result[column] = idx !== undefined ? row[idx] : '';
      });
      return result;
    });

  return filtered;
}

/**
 * Creates or updates a record in the provided sheet.
 *
 * @param {string} sheetName - Name of the sheet to modify.
 * @param {Object} record - Object containing the column/value pairs.
 * @return {Object} The stored record including the key value.
 */
function saveRecord(sheetName, record) {
  var structure = getStructure();
  var config = structure[sheetName];
  if (!config) {
    throw new Error('Não há estrutura configurada para a guia "' + sheetName + '".');
  }

  var keyColumn = config.key;
  if (!keyColumn) {
    throw new Error('Nenhuma coluna chave definida para a guia "' + sheetName + '".');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('A guia "' + sheetName + '" não foi encontrada.');
  }

  var headerValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var headerIndex = headerValues.reduce(function(acc, value, idx) {
    acc[value] = idx + 1; // 1-based index for Range operations
    return acc;
  }, {});

  var keyValue = record[keyColumn];
  var targetRow = null;

  if (keyValue) {
    var data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).getDisplayValues();
    data.some(function(row, idx) {
      if (row[headerIndex[keyColumn] - 1] === keyValue) {
        targetRow = idx + 2; // offset header
        return true;
      }
      return false;
    });
  }

  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
    if (targetRow === 1) {
      targetRow = 2; // ensure there's space for headers
    }
  }

  config.columns.forEach(function(column) {
    var columnIndex = headerIndex[column];
    if (!columnIndex) {
      throw new Error('A coluna "' + column + '" não existe na guia "' + sheetName + '".');
    }
    var value = record[column];
    if (column === keyColumn && !value) {
      value = generateIdentifier(sheetName);
      record[column] = value;
    }
    sheet.getRange(targetRow, columnIndex).setValue(value || '');
  });

  return record;
}

/**
 * Generates a unique identifier for the given sheet combining a prefix
 * with the current timestamp.
 *
 * @param {string} sheetName
 * @return {string}
 */
function generateIdentifier(sheetName) {
  var prefix = sheetName.replace(/[^A-Z0-9]/gi, '').substring(0, 3).toUpperCase();
  return prefix + '-' + new Date().getTime();
}

/** Helper methods exposed to the front-end **/
function getSolicitacoes() {
  return getSheetData('Solicitações');
}

function getPocos() {
  return getSheetData('Poços');
}

function getProjetos() {
  return getSheetData('Projetos');
}

function getDoadores() {
  return getSheetData('Doadores');
}

function saveSolicitacao(record) {
  return saveRecord('Solicitações', record);
}

function savePoco(record) {
  if (record.solicitacaoId) {
    record.solicitacaoId = String(record.solicitacaoId);
  }
  return saveRecord('Poços', record);
}

function saveProjeto(record) {
  if (record.solicitacaoId) {
    record.solicitacaoId = String(record.solicitacaoId);
  }
  return saveRecord('Projetos', record);
}

function saveDoador(record) {
  return saveRecord('Doadores', record);
}

function getViewConfig(viewId) {
  if (!viewId) {
    throw new Error('Uma visualização precisa ser informada.');
  }

  var definition = VIEW_DEFINITIONS.filter(function(view) {
    return view.id === viewId;
  })[0];

  if (!definition) {
    throw new Error('A visualização "' + viewId + '" não está configurada.');
  }

  var structure = getStructure();
  var sheetConfig = structure[definition.sheet];

  if (!sheetConfig) {
    throw new Error('Não há estrutura configurada para a guia "' + definition.sheet + '".');
  }

  return {
    id: definition.id,
    label: definition.label,
    sheet: definition.sheet,
    fetch: definition.fetch,
    save: definition.save,
    key: sheetConfig.key,
    columns: sheetConfig.columns.map(function(column) {
      return {
        name: column,
        label: sheetConfig.labels && sheetConfig.labels[column] ? sheetConfig.labels[column] : column,
        isKey: sheetConfig.key === column
      };
    })
  };
}

function getRecordByKey(viewId, keyValue) {
  if (!viewId) {
    throw new Error('Uma visualização precisa ser informada.');
  }

  var definition = VIEW_DEFINITIONS.filter(function(view) {
    return view.id === viewId;
  })[0];

  if (!definition) {
    throw new Error('A visualização "' + viewId + '" não está configurada.');
  }

  var structure = getStructure();
  var sheetConfig = structure[definition.sheet];

  if (!sheetConfig) {
    throw new Error('Não há estrutura configurada para a guia "' + definition.sheet + '".');
  }

  if (keyValue == null || keyValue === '') {
    return null;
  }

  var data = getSheetData(definition.sheet);
  var keyName = sheetConfig.key;

  var match = data.filter(function(row) {
    return String(row[keyName]) === String(keyValue);
  })[0];

  return match || null;
}

/**
 * Utility to include partial HTML files from the server.
 *
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
