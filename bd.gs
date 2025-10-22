function criarBaseDeDados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abas = {
    "Poços": [
      'ID','Estado','Município','Comunidade','Latitude','Longitude','Beneficiários','Investimento',
      'Vazão (L/H)','Profundidade (m)','Perfuração','Instalação','Doador','Status',
      'Valor Previsto Perfuração','Valor Previsto Instalação','Empresa Responsável',
      'Observações','Valor Realizado','Doadores','DataCadastro',
      'ResponsavelContato','ContatoInstalacao','TelefoneContato','StatusContato',
      'ProximaAcao','UltimoContato','ImpactoNoStatus'
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
    "Contatos": [
      'ID','PoçoID','ResponsavelContato','ContatoExterno','OrganizacaoContato','DataContato',
      'Resumo','ProximaAcao','StatusContato','ImpactoPrevisto','RegistradoPor'
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

function popularDadosDeExemplo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPocos = ss.getSheetByName('Poços');
  const shDoadores = ss.getSheetByName('Doadores');
  const shPrest = ss.getSheetByName('PrestaçãoContas');
  const shContatos = ss.getSheetByName('Contatos');

  shPocos.clearContents();
  shDoadores.clearContents();
  shPrest.clearContents();
  shContatos.clearContents();

  shPocos.appendRow([
    'ID','Estado','Município','Comunidade','Latitude','Longitude','Beneficiários','Investimento',
    'Vazão (L/H)','Profundidade (m)','Perfuração','Instalação','Doador','Status',
    'Valor Previsto Perfuração','Valor Previsto Instalação','Empresa Responsável',
    'Observações','Valor Realizado','Doadores','DataCadastro',
    'ResponsavelContato','ContatoInstalacao','TelefoneContato','StatusContato',
    'ProximaAcao','UltimoContato','ImpactoNoStatus'
  ]);

  const idPoco1 = Utilities.getUuid();
  const idPoco2 = Utilities.getUuid();
  const hoje = new Date();

  shPocos.appendRow([
    idPoco1,'CE','Crateús','Comunidade Vida','-5.167','-40.656',250,180000,
    3200,140,'Perfuração concluída','Bomba instalada','Fundação Esperança','Concluído',
    95000,85000,'Águas do Sertão','Poço entregue e em monitoramento',175000,'',hoje,
    'Maria Silva','João Pereira','(88) 99999-0000','Concluído','Visita de verificação',new Date('2024-05-22'),
    'Contato confirmou operação estável'
  ]);

  shPocos.appendRow([
    idPoco2,'PI','Picos','Assentamento Paz','-7.067','-41.467',180,150000,
    2800,120,'Perfuração agendada','Instalação prevista','Igreja Luz Viva','Em execução',
    90000,60000,'Poços Brasil','Aguardando orçamento final',45000,'',hoje,
    'Carlos Nunes','Eng. Ana Ramos','(86) 98888-1234','Em negociação','Enviar cronograma revisado',new Date('2024-05-28'),
    'Instalação depende da confirmação do fornecedor'
  ]);

  shDoadores.appendRow(['ID','Nome','Email','Telefone','ValorDoado','DataDoacao','PoçosVinculados']);
  const doadorId = Utilities.getUuid();
  shDoadores.appendRow([doadorId,'Fundação Esperança','contato@fundesperanca.org','(11) 3000-0000',250000,hoje,idPoco1]);

  shPrest.appendRow(['PoçoID','Data','Descrição','Valor','ComprovanteURL','Categoria','RegistradoPor']);
  shPrest.appendRow([idPoco1,new Date('2024-05-01'),'Compra de tubos',35000,'', 'Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-05'),'Serviço de instalação',42000,'', 'Instalação','Maria Silva']);
  shPrest.appendRow([idPoco2,new Date('2024-05-20'),'Adiantamento fornecedor',15000,'', 'Instalação','Carlos Nunes']);

  shContatos.appendRow(['ID','PoçoID','ResponsavelContato','ContatoExterno','OrganizacaoContato','DataContato','Resumo','ProximaAcao','StatusContato','ImpactoPrevisto','RegistradoPor']);
  shContatos.appendRow([Utilities.getUuid(), idPoco2, 'Carlos Nunes', 'Eng. Ana Ramos', 'Poços Brasil', new Date(), 'Reunião para alinhar logística', 'Enviar cronograma revisado', 'Em negociação', 'Prazo de instalação depende do novo cronograma', 'Carlos Nunes']);

  return '✅ Dados de exemplo carregados com sucesso. Ajuste conforme necessário.';
}
