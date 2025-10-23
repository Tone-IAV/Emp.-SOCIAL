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
    ['STATUS','Solicitado,Orçamento previsto,Instalação,Concluído,Pago'],
    ['TIPOS_DESPESA','Perfuração,Instalação,Manutenção,Logística,Pagamento'],
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

  const resetSheet = (sheet, headers) => {
    if (!sheet) return;
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  };

  resetSheet(shPocos, COLUNAS_POCOS);
  resetSheet(shDoadores, ['ID','Nome','Email','Telefone','ValorDoado','DataDoacao','PoçosVinculados']);
  resetSheet(shPrest, ['PoçoID','Data','Descrição','Valor','ComprovanteURL','Categoria','RegistradoPor']);
  resetSheet(shContatos, ['ID','PoçoID','ResponsavelContato','ContatoExterno','OrganizacaoContato','DataContato','Resumo','ProximaAcao','StatusContato','ImpactoPrevisto','RegistradoPor']);
  resetSheet(shEmpresas, ['ID','NomeEmpresa','CNPJ','Tipo','Contato','Observações']);

  const idPoco1 = Utilities.getUuid();
  const idPoco2 = Utilities.getUuid();
  const idPoco3 = Utilities.getUuid();
  const idPoco4 = Utilities.getUuid();

  const doador1 = Utilities.getUuid();
  const doador2 = Utilities.getUuid();
  const doador3 = Utilities.getUuid();
  const doador4 = Utilities.getUuid();

  const contatosPoco1 = JSON.stringify([
    { nome: 'Maria Silva', telefone: '(88) 99999-0000', email: 'maria@poços.org', papel: 'Articulação local' },
    { nome: 'João Pereira', telefone: '(88) 99888-2222', email: 'joao@comunidadevida.org', papel: 'Liderança comunitária' }
  ]);
  const evidenciasPoco1 = JSON.stringify([
    { tipo: 'imagem', titulo: 'Entrega do poço', url: 'https://drive.google.com/example-poco1-foto1', data: '2024-05-18' },
    { tipo: 'documento', titulo: 'Relatório técnico final', url: 'https://drive.google.com/example-poco1-relatorio', data: '2024-05-20' }
  ]);

  const contatosPoco2 = JSON.stringify([
    { nome: 'Carlos Nunes', telefone: '(86) 98888-1234', email: 'carlos@missao.org', papel: 'Gestor de campo' },
    { nome: 'Eng. Ana Ramos', telefone: '(86) 97777-5432', email: 'ana@pocosbr.com', papel: 'Engenharia' }
  ]);
  const evidenciasPoco2 = JSON.stringify([
    { tipo: 'imagem', titulo: 'Equipe em instalação', url: 'https://drive.google.com/example-poco2-foto1', data: '2024-05-22' }
  ]);

  const contatosPoco3 = JSON.stringify([
    { nome: 'Luciana Prado', telefone: '(74) 97777-4567', email: 'luciana@institutoaguaviva.org', papel: 'Mobilização regional' }
  ]);
  const evidenciasPoco3 = JSON.stringify([
    { tipo: 'documento', titulo: 'Estudo hidrogeológico', url: 'https://drive.google.com/example-poco3-estudo', data: '2024-05-02' }
  ]);

  const contatosPoco4 = JSON.stringify([
    { nome: 'Roberto Lima', telefone: '(98) 95555-9988', email: 'roberto@cooperativasertao.org', papel: 'Articulação local' }
  ]);

  const pocosExemplo = [
    {
      ID: idPoco1,
      Estado: 'CE',
      'Município': 'Crateús',
      Comunidade: 'Comunidade Vida',
      'Região': 'Sertão dos Inhamuns',
      Latitude: '-5.167',
      Longitude: '-40.656',
      'Beneficiários': 250,
      Investimento: 185000,
      'Vazão (L/H)': 3200,
      'Profundidade (m)': 140,
      Status: 'Pago',
      ResumoStatus: 'Poço entregue, monitorado e com pagamento finalizado.',
      Solicitante: 'Maria Silva',
      ContatoSolicitante: 'maria@poços.org | (88) 99999-0000',
      DataSolicitacao: new Date('2024-03-28'),
      DataOrcamentoPrevisto: new Date('2024-04-05'),
      DataInstalacao: new Date('2024-04-28'),
      DataConclusao: new Date('2024-05-18'),
      DataPagamento: new Date('2024-05-25'),
      OrcamentoPrevisto: 190000,
      OrcamentoAprovado: 185000,
      OrcamentoExecutado: 185000,
      'Valor Previsto Perfuração': 190000,
      'Valor Previsto Instalação': 0,
      'Valor Realizado': 185000,
      TermoAutorizacaoURL: 'https://drive.google.com/example-poco1-termo',
      NotaFiscalURL: 'https://drive.google.com/example-poco1-notafiscal',
      ContatosJSON: contatosPoco1,
      EvidenciasJSON: evidenciasPoco1,
      LinhaDoTempoJSON: JSON.stringify([]),
      Doadores: `${doador1},${doador4}`,
      'Empresa Responsável': 'Águas do Sertão',
      Observações: 'Monitoramento ativo com visitas trimestrais.',
      DataCadastro: new Date('2024-03-28'),
      DataUltimaAtualizacao: new Date('2024-06-02'),
      ResponsavelContato: 'Maria Silva',
      ContatoInstalacao: 'João Pereira',
      TelefoneContato: '(88) 99999-0000',
      StatusContato: 'Monitoramento',
      ProximaAcao: 'Agendar oficina de gestão comunitária',
      UltimoContato: new Date('2024-06-02'),
      ImpactoNoStatus: 'Comunidade abastecida diariamente e com governança ativa.',
      TipoPoco: 'Artesiano',
      SituacaoHidrica: 'Produzindo normalmente',
      AcoesPosInstalacao: 'Horta comunitária e capacitação de operadores.',
      UsoAguaComunitario: 'Consumo humano e irrigação comunitária'
    },
    {
      ID: idPoco2,
      Estado: 'PI',
      'Município': 'Picos',
      Comunidade: 'Assentamento Paz',
      'Região': 'Semiárido piauiense',
      Latitude: '-7.067',
      Longitude: '-41.467',
      'Beneficiários': 180,
      Investimento: 168000,
      'Vazão (L/H)': 2800,
      'Profundidade (m)': 120,
      Status: 'Instalação',
      ResumoStatus: 'Equipe técnica em campo executando montagem do sistema.',
      Solicitante: 'Carlos Nunes',
      ContatoSolicitante: 'carlos@missao.org | (86) 98888-1234',
      DataSolicitacao: new Date('2024-04-12'),
      DataOrcamentoPrevisto: new Date('2024-04-24'),
      DataInstalacao: new Date('2024-05-21'),
      OrcamentoPrevisto: 155000,
      OrcamentoAprovado: 168000,
      OrcamentoExecutado: 68000,
      'Valor Previsto Perfuração': 155000,
      'Valor Previsto Instalação': 0,
      'Valor Realizado': 68000,
      TermoAutorizacaoURL: 'https://drive.google.com/example-poco2-termo',
      NotaFiscalURL: '',
      ContatosJSON: contatosPoco2,
      EvidenciasJSON: evidenciasPoco2,
      LinhaDoTempoJSON: JSON.stringify([]),
      Doadores: `${doador2}`,
      'Empresa Responsável': 'Poços Brasil',
      Observações: 'Necessário validar cronograma revisado com fornecedor local.',
      DataCadastro: new Date('2024-04-12'),
      DataUltimaAtualizacao: new Date('2024-05-28'),
      ResponsavelContato: 'Carlos Nunes',
      ContatoInstalacao: 'Eng. Ana Ramos',
      TelefoneContato: '(86) 98888-1234',
      StatusContato: 'Em negociação',
      ProximaAcao: 'Concluir alinhamento logístico com fornecedor',
      UltimoContato: new Date('2024-05-28'),
      ImpactoNoStatus: 'Instalação depende da confirmação da logística revisada.',
      TipoPoco: 'Semiartesiano',
      SituacaoHidrica: 'Vazão reduzida',
      AcoesPosInstalacao: 'Capacitação para proteção da infraestrutura e uso produtivo.',
      UsoAguaComunitario: 'Consumo doméstico e apoio à escola local'
    },
    {
      ID: idPoco3,
      Estado: 'BA',
      'Município': 'Juazeiro',
      Comunidade: 'Vila Esperança',
      'Região': 'Vale do São Francisco',
      Latitude: '-9.430',
      Longitude: '-40.507',
      'Beneficiários': 320,
      Investimento: 210000,
      'Vazão (L/H)': 3500,
      'Profundidade (m)': 155,
      Status: 'Orçamento previsto',
      ResumoStatus: 'Estudo hidrogeológico concluído e aguardando liberação de recursos.',
      Solicitante: 'Luciana Prado',
      ContatoSolicitante: 'luciana@institutoaguaviva.org | (74) 97777-4567',
      DataSolicitacao: new Date('2024-04-28'),
      DataOrcamentoPrevisto: new Date('2024-05-10'),
      OrcamentoPrevisto: 198000,
      OrcamentoAprovado: 0,
      OrcamentoExecutado: 25000,
      'Valor Previsto Perfuração': 198000,
      'Valor Previsto Instalação': 0,
      'Valor Realizado': 25000,
      TermoAutorizacaoURL: '',
      NotaFiscalURL: '',
      ContatosJSON: contatosPoco3,
      EvidenciasJSON: evidenciasPoco3,
      LinhaDoTempoJSON: JSON.stringify([]),
      Doadores: `${doador3}`,
      'Empresa Responsável': 'Fonte Limpa Engenharia',
      Observações: 'Mobilização comunitária concluída, aguardando liberação financeira.',
      DataCadastro: new Date('2024-04-28'),
      DataUltimaAtualizacao: new Date('2024-05-20'),
      ResponsavelContato: 'Luciana Prado',
      ContatoInstalacao: 'Carlos Menezes',
      TelefoneContato: '(74) 97777-4567',
      StatusContato: 'Aguardando liberação',
      ProximaAcao: 'Confirmar disponibilidade de perfuratriz',
      UltimoContato: new Date('2024-05-20'),
      ImpactoNoStatus: 'Janela curta para perfuração exige decisão rápida.',
      TipoPoco: 'Artesiano',
      SituacaoHidrica: 'Em análise técnica',
      AcoesPosInstalacao: 'Treinamentos de governança da água e uso produtivo.',
      UsoAguaComunitario: 'Produção agrícola familiar e consumo humano'
    },
    {
      ID: idPoco4,
      Estado: 'MA',
      'Município': 'Codó',
      Comunidade: 'Serra Azul',
      'Região': 'Médio Itapecuru',
      Latitude: '-4.455',
      Longitude: '-43.890',
      'Beneficiários': 140,
      Investimento: 132000,
      'Vazão (L/H)': 2500,
      'Profundidade (m)': 110,
      Status: 'Solicitado',
      ResumoStatus: 'Licenciamento em análise e documentação complementar pendente.',
      Solicitante: 'Roberto Lima',
      ContatoSolicitante: 'roberto@cooperativasertao.org | (98) 95555-9988',
      DataSolicitacao: new Date('2024-05-25'),
      OrcamentoPrevisto: 130000,
      OrcamentoAprovado: 0,
      OrcamentoExecutado: 0,
      'Valor Previsto Perfuração': 130000,
      'Valor Previsto Instalação': 0,
      'Valor Realizado': 0,
      TermoAutorizacaoURL: '',
      NotaFiscalURL: '',
      ContatosJSON: contatosPoco4,
      EvidenciasJSON: JSON.stringify([]),
      LinhaDoTempoJSON: JSON.stringify([]),
      Doadores: `${doador4}`,
      'Empresa Responsável': 'Nordeste Perfurações',
      Observações: 'Em fase de licenciamento ambiental.',
      DataCadastro: new Date('2024-05-25'),
      DataUltimaAtualizacao: new Date('2024-05-27'),
      ResponsavelContato: 'Roberto Lima',
      ContatoInstalacao: 'Eng. Paula Duarte',
      TelefoneContato: '(98) 95555-9988',
      StatusContato: 'Documentação',
      ProximaAcao: 'Enviar certidões atualizadas ao órgão ambiental',
      UltimoContato: new Date('2024-05-27'),
      ImpactoNoStatus: 'Dependente de liberação ambiental para iniciar perfuração.',
      TipoPoco: 'Raso',
      SituacaoHidrica: 'Em análise técnica',
      AcoesPosInstalacao: 'Educação para uso responsável da água após licenciamento.',
      UsoAguaComunitario: 'Consumo humano e apoio a pequenos criadores'
    }
  ];

  pocosExemplo.forEach(poco => {
    const linha = COLUNAS_POCOS.map(coluna => poco[coluna] !== undefined ? poco[coluna] : '');
    shPocos.appendRow(linha);
  });

  shDoadores.appendRow([doador1,'Fundação Esperança','contato@fundesperanca.org','(11) 3000-0000',250000,new Date('2024-04-10'),idPoco1]);
  shDoadores.appendRow([doador2,'Igreja Luz Viva','doacoes@igrejaluzi.org','(31) 3555-5555',180000,new Date('2024-04-22'),`${idPoco1},${idPoco2}`]);
  shDoadores.appendRow([doador3,'Instituto Água Viva','parcerias@aguaviva.org','(21) 3666-6677',150000,new Date('2024-05-12'),idPoco3]);
  shDoadores.appendRow([doador4,'Cooperativa Sementes do Bem','relacionamento@sementesdobem.coop','(62) 3777-8899',95000,new Date('2024-05-18'),`${idPoco1},${idPoco4}`]);

  shPrest.appendRow([idPoco1,new Date('2024-04-25'),'Topografia e mobilização inicial',22000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-01'),'Compra de tubos e bombas',35000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-05'),'Serviço de instalação completa',42000,'','Instalação','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-18'),'Treinamento da comunidade',12000,'','Instalação','João Pereira']);
  shPrest.appendRow([idPoco1,new Date('2024-05-25'),'Pagamento fornecedor',85000,'','Pagamento','Maria Silva']);
  shPrest.appendRow([idPoco2,new Date('2024-05-14'),'Adiantamento fornecedor perfuração',35000,'','Perfuração','Carlos Nunes']);
  shPrest.appendRow([idPoco2,new Date('2024-05-20'),'Logística de equipamentos',18000,'','Logística','Carlos Nunes']);
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
