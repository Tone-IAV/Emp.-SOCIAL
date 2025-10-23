function criarBaseDeDados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abas = {
    "Poços": COLUNAS_POCOS,
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
  const shEmpresas = ss.getSheetByName('Empresas');

  [shPocos, shDoadores, shPrest, shContatos, shEmpresas].forEach(sh => {
    if (sh) sh.clearContents();
  });

  shPocos.appendRow(COLUNAS_POCOS);

  shDoadores.appendRow(['ID','Nome','Email','Telefone','ValorDoado','DataDoacao','PoçosVinculados']);
  shPrest.appendRow(['PoçoID','Data','Descrição','Valor','ComprovanteURL','Categoria','RegistradoPor']);
  shContatos.appendRow(['ID','PoçoID','ResponsavelContato','ContatoExterno','OrganizacaoContato','DataContato','Resumo','ProximaAcao','StatusContato','ImpactoPrevisto','RegistradoPor']);
  if (shEmpresas) {
    shEmpresas.appendRow(['ID','NomeEmpresa','CNPJ','Tipo','Contato','Observações']);
  }

  const hoje = new Date();
  const idPoco1 = Utilities.getUuid();
  const idPoco2 = Utilities.getUuid();
  const idPoco3 = Utilities.getUuid();
  const idPoco4 = Utilities.getUuid();

  const doador1 = Utilities.getUuid();
  const doador2 = Utilities.getUuid();
  const doador3 = Utilities.getUuid();
  const doador4 = Utilities.getUuid();

  shPocos.appendRow([
    idPoco1,'CE','Crateús','Comunidade Vida','-5.167','-40.656',250,195000,
    3200,140,'Concluída em 18/05/2024','Instalada e testada em 22/05/2024','Fundação Esperança','Concluído',
    95000,85000,'Águas do Sertão','Poço entregue e em monitoramento',185000,`${doador1},${doador4}`,new Date('2024-04-02'),
    'Maria Silva','João Pereira','(88) 99999-0000','Monitoramento','Organizar treinamento da comunidade',new Date('2024-06-02'),
    'Operação estável, comunidade abastecida diariamente','Artesiano','Produzindo normalmente',
    'Realizar capacitação para gestão comunitária e irrigação de hortas','Consumo humano, irrigação comunitária'
  ]);

  shPocos.appendRow([
    idPoco2,'PI','Picos','Assentamento Paz','-7.067','-41.467',180,168000,
    2800,120,'Perfuração em andamento - término 08/06/2024','Instalação prevista para 15/06/2024','Igreja Luz Viva','Em execução',
    90000,72000,'Poços Brasil','Equipe em campo ajustando cronograma',68000,`${doador2}`,new Date('2024-04-18'),
    'Carlos Nunes','Eng. Ana Ramos','(86) 98888-1234','Em negociação','Validar plano de instalação revisado',new Date('2024-05-28'),
    'Instalação depende da confirmação do fornecedor local','Semiartesiano','Vazão reduzida',
    'Engajar comunidade para proteção da infraestrutura e capacitar operadores locais','Consumo doméstico e apoio à escola'
  ]);

  shPocos.appendRow([
    idPoco3,'BA','Juazeiro','Vila Esperança','-9.430','-40.507',320,210000,
    3500,155,'Estudo hidrogeológico concluído','Instalação aguardando viabilização logística','Instituto Água Viva','Planejado',
    110000,85000,'Fonte Limpa Engenharia','Comunidade mobilizada aguardando início',25000,`${doador3}`,new Date('2024-05-05'),
    'Luciana Prado','Carlos Menezes','(74) 97777-4567','Aguardando liberação','Confirmar disponibilidade do perfuratriz',new Date('2024-05-20'),
    'Risco de atraso por janela curta de perfuração','Artesiano','Em análise técnica',
    'Preparar treinamento de governança da água e mapeamento de usos produtivos','Produção agrícola familiar e consumo humano'
  ]);

  shPocos.appendRow([
    idPoco4,'MA','Codó','Serra Azul','-4.455','-43.890',140,132000,
    2500,110,'Licenciamento protocolado em 27/05/2024','Instalação condicionada ao licenciamento','Cooperativa Sementes do Bem','Planejado',
    70000,54000,'Nordeste Perfurações','Documentação complementar em análise',0,`${doador4}`,new Date('2024-05-25'),
    'Roberto Lima','Eng. Paula Duarte','(98) 95555-9988','Documentação','Enviar certidões atualizadas ao órgão ambiental',new Date('2024-05-27'),
    'Licença pendente pode atrasar início da perfuração','Raso','Em análise técnica',
    'Planejar ações educativas de uso responsável da água após liberação','Consumo humano e apoio a pequenos criadores'
  ]);

  shDoadores.appendRow([doador1,'Fundação Esperança','contato@fundesperanca.org','(11) 3000-0000',250000,new Date('2024-04-10'),idPoco1]);
  shDoadores.appendRow([doador2,'Igreja Luz Viva','doacoes@igrejaluzi.org','(31) 3555-5555',180000,new Date('2024-04-22'),`${idPoco1},${idPoco2}`]);
  shDoadores.appendRow([doador3,'Instituto Água Viva','parcerias@aguaviva.org','(21) 3666-6677',150000,new Date('2024-05-12'),idPoco3]);
  shDoadores.appendRow([doador4,'Cooperativa Sementes do Bem','relacionamento@sementesdobem.coop','(62) 3777-8899',95000,new Date('2024-05-18'),`${idPoco1},${idPoco4}`]);

  shPrest.appendRow([idPoco1,new Date('2024-04-25'),'Topografia e mobilização inicial',22000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-01'),'Compra de tubos',35000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-05'),'Serviço de instalação',42000,'','Instalação','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-18'),'Treinamento da comunidade',12000,'','Instalação','João Pereira']);
  shPrest.appendRow([idPoco2,new Date('2024-05-14'),'Adiantamento fornecedor perfuração',35000,'','Perfuração','Carlos Nunes']);
  shPrest.appendRow([idPoco2,new Date('2024-05-20'),'Logística de equipamentos',18000,'','Perfuração','Carlos Nunes']);
  shPrest.appendRow([idPoco3,new Date('2024-05-22'),'Consultoria hidrogeológica',25000,'','Perfuração','Luciana Prado']);
  shPrest.appendRow([idPoco3,new Date('2024-05-28'),'Reserva de materiais elétricos',8000,'','Instalação','Carlos Menezes']);

  shContatos.appendRow([Utilities.getUuid(), idPoco1, 'Maria Silva', 'João Pereira', 'Comunidade Vida', new Date('2024-06-02'), 'Monitoramento pós-instalação', 'Organizar treinamento da comunidade', 'Monitoramento', 'Indicadores positivos mantidos', 'Maria Silva']);
  shContatos.appendRow([Utilities.getUuid(), idPoco2, 'Carlos Nunes', 'Eng. Ana Ramos', 'Poços Brasil', new Date('2024-05-28'), 'Reunião para alinhar logística', 'Validar plano de instalação revisado', 'Em negociação', 'Prazo de instalação depende do novo cronograma', 'Carlos Nunes']);
  shContatos.appendRow([Utilities.getUuid(), idPoco3, 'Luciana Prado', 'Carlos Menezes', 'Fonte Limpa Engenharia', new Date('2024-05-20'), 'Análise de disponibilidade do equipamento', 'Confirmar disponibilidade do perfuratriz', 'Aguardando liberação', 'Atraso pode impactar início do projeto', 'Luciana Prado']);
  shContatos.appendRow([Utilities.getUuid(), idPoco4, 'Roberto Lima', 'Eng. Paula Duarte', 'Nordeste Perfurações', new Date('2024-05-27'), 'Envio de documentação complementar', 'Enviar certidões atualizadas ao órgão ambiental', 'Documentação', 'Licença pendente pode atrasar início da perfuração', 'Roberto Lima']);

  if (shEmpresas) {
    shEmpresas.appendRow([Utilities.getUuid(),'Águas do Sertão','12.345.678/0001-90','Perfuração','contato@aguass.com','Responsável pelo poço Comunidade Vida']);
    shEmpresas.appendRow([Utilities.getUuid(),'Poços Brasil','98.765.432/0001-10','Perfuração e instalação','ana.ramos@pocosbr.com','Equipe destacada para Assentamento Paz']);
    shEmpresas.appendRow([Utilities.getUuid(),'Fonte Limpa Engenharia','54.321.678/0001-55','Instalação','contato@fontelimpa.eng','Parceria técnica para Vila Esperança']);
    shEmpresas.appendRow([Utilities.getUuid(),'Nordeste Perfurações','43.210.987/0001-32','Perfuração','paula.duarte@nordesteperf.com','Responsável pela perfuração em Serra Azul']);
  }

  return '✅ Dados de exemplo carregados com sucesso. Ajuste conforme necessário.';
}
