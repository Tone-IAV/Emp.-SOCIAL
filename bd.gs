function criarBaseDeDados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abas = {
    "Poços": [
      'ID','Estado','Município','Comunidade','Latitude','Longitude','Beneficiários','Investimento',
      'Vazão (L/H)','Profundidade (m)','Perfuração','Instalação','Doador','Status',
      'Valor Previsto Perfuração','Valor Previsto Instalação','Empresa Responsável',
      'Observações','Valor Realizado','Doadores','DataCadastro'
    ],
    "Doadores": [
      'ID','Nome','Email','Telefone','ValorDoado','DataDoacao','PoçosVinculados'
    ],
    "PrestaçãoContas": [
      'PoçoID','Data','Descrição','Valor','ComprovanteURL','Categoria','RegistradoPor'
    ],
    "HistóricoStatus": [
      'PoçoID','StatusAnterior','NovoStatus','DataAlteracao','AlteradoPor'
    ],
    "Empresas": [
      'ID','NomeEmpresa','CNPJ','Tipo','Contato','Observações'
    ],
    "Configurações": [
      'Chave','Valor'
    ]
  };

  Object.entries(abas).forEach(([nome, colunas]) => {
    let sh = ss.getSheetByName(nome);
    if (!sh) {
      sh = ss.insertSheet(nome);
      sh.appendRow(colunas);
      sh.setFrozenRows(1);
    }
  });

  const shConfig = ss.getSheetByName("Configurações");
  const configs = [
    ['STATUS','Planejado,Em execução,Concluído'],
    ['TIPOS_DESPESA','Perfuração,Instalação,Manutenção'],
    ['PERFIS','Administrador,Doador,Visualizador']
  ];
  configs.forEach(c => {
    const range = shConfig.createTextFinder(c[0]).findNext();
    if (!range) shConfig.appendRow(c);
  });

  return '✅ Base de dados criada e configurada com sucesso.';
}
