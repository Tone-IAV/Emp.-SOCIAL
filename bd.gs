function criarBaseDeDados() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abas = {
    "Poços": COLUNAS_POCOS,
    "Doadores": COLUNAS_DOADORES,
    "Depósitos": COLUNAS_DEPOSITOS,
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
  const shDepositos = ss.getSheetByName('Depósitos');

  const resetSheet = (sheet, headers) => {
    if (!sheet) return;
    sheet.clearContents();
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  };

  const normalizarTelefone = valor => (valor ? String(valor).replace(/\D+/g, '') : '');
  const criarContato = (nome, telefone, email, papel) => ({
    nome,
    telefone,
    telefoneNormalizado: normalizarTelefone(telefone),
    email,
    papel
  });

  resetSheet(shPocos, COLUNAS_POCOS);
  resetSheet(shDoadores, COLUNAS_DOADORES);
  resetSheet(shDepositos, COLUNAS_DEPOSITOS);
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
    criarContato('Maria Silva', '(88) 99999-0000', 'maria@pocos.org', 'Articulação local'),
    criarContato('João Pereira', '(88) 99888-2222', 'joao@comunidadevida.org', 'Liderança comunitária')
  ]);
  const evidenciasPoco1 = JSON.stringify([
    { tipo: 'imagem', titulo: 'Entrega do poço', url: 'https://drive.google.com/example-poco1-foto1', data: '2024-05-18' },
    { tipo: 'documento', titulo: 'Relatório técnico final', url: 'https://drive.google.com/example-poco1-relatorio', data: '2024-05-20' }
  ]);
  const timelinePoco1 = JSON.stringify([
    { id: 'monitoramento', titulo: 'Visita de monitoramento', descricao: 'Equipe validou funcionamento do poço.', status: 'concluido', data: '2024-06-10' }
  ]);

  const contatosPoco2 = JSON.stringify([
    criarContato('Carlos Nunes', '(86) 98888-1234', 'carlos@missao.org', 'Gestor de campo'),
    criarContato('Eng. Ana Ramos', '(86) 97777-5432', 'ana@pocosbr.com', 'Engenharia')
  ]);
  const evidenciasPoco2 = JSON.stringify([
    { tipo: 'imagem', titulo: 'Equipe em instalação', url: 'https://drive.google.com/example-poco2-foto1', data: '2024-05-22' }
  ]);
  const timelinePoco2 = JSON.stringify([
    { id: 'logistica', titulo: 'Logística revisada', descricao: 'Planejamento de entrega dos equipamentos atualizado.', status: 'andamento', data: '2024-05-26' }
  ]);

  const contatosPoco3 = JSON.stringify([
    criarContato('Luciana Prado', '(74) 97777-4567', 'luciana@institutoaguaviva.org', 'Mobilização regional')
  ]);
  const evidenciasPoco3 = JSON.stringify([
    { tipo: 'documento', titulo: 'Estudo hidrogeológico', url: 'https://drive.google.com/example-poco3-estudo', data: '2024-05-02' }
  ]);
  const timelinePoco3 = JSON.stringify([
    { id: 'mobilizacao', titulo: 'Mobilização comunitária', descricao: 'Reunião com lideranças locais para apresentação do projeto.', status: 'concluido', data: '2024-05-15' }
  ]);

  const contatosPoco4 = JSON.stringify([
    criarContato('Roberto Lima', '(98) 95555-9988', 'roberto@cooperativasertao.org', 'Articulação local')
  ]);
  const evidenciasPoco4 = JSON.stringify([]);
  const timelinePoco4 = JSON.stringify([
    { id: 'licenciamento', titulo: 'Protocolo de licenciamento', descricao: 'Documentação enviada ao órgão ambiental.', status: 'andamento', data: '2024-05-26' }
  ]);

  const pocosExemplo = [
    {
      ID: idPoco1,
      Estado: 'CE',
      'Município': 'Crateús',
      Comunidade: 'Comunidade Vida',
      'Região': 'Sertão dos Inhamuns',
      Latitude: -5.167,
      Longitude: -40.656,
      'Beneficiários': 250,
      Investimento: 185000,
      'Vazão (L/H)': 3200,
      'Profundidade (m)': 140,
      Status: 'Pago',
      ResumoStatus: 'Poço entregue, monitorado e com pagamento finalizado.',
      Solicitante: 'Maria Silva',
      ContatoSolicitante: 'maria@pocos.org | (88) 99999-0000',
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
      LinhaDoTempoJSON: timelinePoco1,
      Doadores: `${doador1},${doador4}`,
      'Empresa Responsável': 'Águas do Sertão',
      Observações: 'Monitoramento ativo com visitas trimestrais.',
      DataCadastro: new Date('2024-03-28'),
      DataUltimaAtualizacao: new Date('2024-06-02'),
      ResponsavelContato: 'Maria Silva',
      ContatoInstalacao: 'João Pereira',
      TelefoneContato: '(88) 99999-0000',
      TelefoneContatoNormalizado: normalizarTelefone('(88) 99999-0000'),
      StatusContato: 'Monitoramento',
      ProximaAcao: 'Agendar oficina de gestão comunitária',
      UltimoContato: new Date('2024-06-02'),
      ImpactoNoStatus: 'Comunidade abastecida diariamente e com governança ativa.',
      TipoPoco: 'Artesiano',
      SituacaoHidrica: 'Produzindo normalmente',
      AcoesPosInstalacao: 'Horta comunitária e capacitação de operadores.',
      UsoAguaComunitario: 'Consumo humano e irrigação comunitária',
      GeocodificacaoFonte: 'Cadastro manual',
      GeocodificacaoPrecisao: 'Equipe de campo',
      GeocodificacaoStatus: 'OK',
      GeocodificacaoTimestamp: new Date('2024-06-02')
    },
    {
      ID: idPoco2,
      Estado: 'PI',
      'Município': 'Picos',
      Comunidade: 'Assentamento Paz',
      'Região': 'Semiárido piauiense',
      Latitude: -7.067,
      Longitude: -41.467,
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
      LinhaDoTempoJSON: timelinePoco2,
      Doadores: `${doador2}`,
      'Empresa Responsável': 'Poços Brasil',
      Observações: 'Necessário validar cronograma revisado com fornecedor local.',
      DataCadastro: new Date('2024-04-12'),
      DataUltimaAtualizacao: new Date('2024-05-28'),
      ResponsavelContato: 'Carlos Nunes',
      ContatoInstalacao: 'Eng. Ana Ramos',
      TelefoneContato: '(86) 98888-1234',
      TelefoneContatoNormalizado: normalizarTelefone('(86) 98888-1234'),
      StatusContato: 'Em negociação',
      ProximaAcao: 'Concluir alinhamento logístico com fornecedor',
      UltimoContato: new Date('2024-05-28'),
      ImpactoNoStatus: 'Instalação depende da confirmação da logística revisada.',
      TipoPoco: 'Semiartesiano',
      SituacaoHidrica: 'Vazão reduzida',
      AcoesPosInstalacao: 'Capacitação para proteção da infraestrutura e uso produtivo.',
      UsoAguaComunitario: 'Consumo doméstico e apoio à escola local',
      GeocodificacaoFonte: 'Cadastro manual',
      GeocodificacaoPrecisao: 'Equipe de campo',
      GeocodificacaoStatus: 'OK',
      GeocodificacaoTimestamp: new Date('2024-05-28')
    },
    {
      ID: idPoco3,
      Estado: 'BA',
      'Município': 'Juazeiro',
      Comunidade: 'Vila Esperança',
      'Região': 'Vale do São Francisco',
      Latitude: -9.43,
      Longitude: -40.507,
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
      LinhaDoTempoJSON: timelinePoco3,
      Doadores: `${doador3}`,
      'Empresa Responsável': 'Fonte Limpa Engenharia',
      Observações: 'Mobilização comunitária concluída, aguardando liberação financeira.',
      DataCadastro: new Date('2024-04-28'),
      DataUltimaAtualizacao: new Date('2024-05-20'),
      ResponsavelContato: 'Luciana Prado',
      ContatoInstalacao: 'Carlos Menezes',
      TelefoneContato: '(74) 97777-4567',
      TelefoneContatoNormalizado: normalizarTelefone('(74) 97777-4567'),
      StatusContato: 'Aguardando liberação',
      ProximaAcao: 'Confirmar disponibilidade de perfuratriz',
      UltimoContato: new Date('2024-05-20'),
      ImpactoNoStatus: 'Janela curta para perfuração exige decisão rápida.',
      TipoPoco: 'Artesiano',
      SituacaoHidrica: 'Em análise técnica',
      AcoesPosInstalacao: 'Treinamentos de governança da água e uso produtivo.',
      UsoAguaComunitario: 'Produção agrícola familiar e consumo humano',
      GeocodificacaoFonte: 'Cadastro manual',
      GeocodificacaoPrecisao: 'Equipe de campo',
      GeocodificacaoStatus: 'OK',
      GeocodificacaoTimestamp: new Date('2024-05-20')
    },
    {
      ID: idPoco4,
      Estado: 'MA',
      'Município': 'Codó',
      Comunidade: 'Serra Azul',
      'Região': 'Médio Itapecuru',
      Latitude: -4.455,
      Longitude: -43.89,
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
      EvidenciasJSON: evidenciasPoco4,
      LinhaDoTempoJSON: timelinePoco4,
      Doadores: `${doador4}`,
      'Empresa Responsável': 'Nordeste Perfurações',
      Observações: 'Em fase de licenciamento ambiental.',
      DataCadastro: new Date('2024-05-25'),
      DataUltimaAtualizacao: new Date('2024-05-27'),
      ResponsavelContato: 'Roberto Lima',
      ContatoInstalacao: 'Eng. Paula Duarte',
      TelefoneContato: '(98) 95555-9988',
      TelefoneContatoNormalizado: normalizarTelefone('(98) 95555-9988'),
      StatusContato: 'Documentação',
      ProximaAcao: 'Enviar certidões atualizadas ao órgão ambiental',
      UltimoContato: new Date('2024-05-27'),
      ImpactoNoStatus: 'Dependente de liberação ambiental para iniciar perfuração.',
      TipoPoco: 'Raso',
      SituacaoHidrica: 'Em análise técnica',
      AcoesPosInstalacao: 'Educação para uso responsável da água após licenciamento.',
      UsoAguaComunitario: 'Consumo humano e apoio a pequenos criadores',
      GeocodificacaoFonte: 'Cadastro manual',
      GeocodificacaoPrecisao: 'Equipe de campo',
      GeocodificacaoStatus: 'Pendente',
      GeocodificacaoTimestamp: ''
    }
  ];

  pocosExemplo.forEach(poco => {
    const linha = COLUNAS_POCOS.map(coluna => poco[coluna] !== undefined ? poco[coluna] : '');
    shPocos.appendRow(linha);
  });

  const doadoresExemplo = [
    {
      ID: doador1,
      Nome: 'Fundação Esperança',
      Email: 'contato@fundesperanca.org',
      Telefone: '(11) 3000-0000',
      TelefoneNormalizado: normalizarTelefone('(11) 3000-0000'),
      Observacoes: 'Programa Água para o Sertão com foco em comunidades rurais.',
      CriadoEm: new Date('2024-03-10'),
      AtualizadoEm: new Date('2024-06-01'),
      'PoçosVinculados': `${idPoco1}`
    },
    {
      ID: doador2,
      Nome: 'Igreja Luz Viva',
      Email: 'doacoes@igrejaluzi.org',
      Telefone: '(31) 3555-5555',
      TelefoneNormalizado: normalizarTelefone('(31) 3555-5555'),
      Observacoes: 'Campanha anual de missões e apoio a perfurações.',
      CriadoEm: new Date('2024-03-22'),
      AtualizadoEm: new Date('2024-05-28'),
      'PoçosVinculados': `${idPoco1},${idPoco2}`
    },
    {
      ID: doador3,
      Nome: 'Instituto Água Viva',
      Email: 'parcerias@aguaviva.org',
      Telefone: '(21) 3666-6677',
      TelefoneNormalizado: normalizarTelefone('(21) 3666-6677'),
      Observacoes: 'Investimento social privado em inovação hídrica.',
      CriadoEm: new Date('2024-04-05'),
      AtualizadoEm: new Date('2024-05-20'),
      'PoçosVinculados': `${idPoco3}`
    },
    {
      ID: doador4,
      Nome: 'Cooperativa Sementes do Bem',
      Email: 'relacionamento@sementesdobem.coop',
      Telefone: '(62) 3777-8899',
      TelefoneNormalizado: normalizarTelefone('(62) 3777-8899'),
      Observacoes: 'Produtores solidários financiando água para o semiárido.',
      CriadoEm: new Date('2024-04-18'),
      AtualizadoEm: new Date('2024-05-30'),
      'PoçosVinculados': `${idPoco1},${idPoco4}`
    }
  ];

  doadoresExemplo.forEach(doador => {
    const linha = COLUNAS_DOADORES.map(coluna => doador[coluna] !== undefined ? doador[coluna] : '');
    shDoadores.appendRow(linha);
  });

  const depositosExemplo = [
    {
      ID: Utilities.getUuid(),
      DoadorID: doador1,
      Valor: 100000,
      DataDeposito: new Date('2024-03-20'),
      Metodo: 'PIX',
      Observacoes: 'Entrada do projeto Comunidade Vida',
      RegistradoEm: new Date('2024-03-20')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador1,
      Valor: 150000,
      DataDeposito: new Date('2024-04-10'),
      Metodo: 'TED',
      Observacoes: 'Complemento para instalação',
      RegistradoEm: new Date('2024-04-10')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador2,
      Valor: 90000,
      DataDeposito: new Date('2024-04-22'),
      Metodo: 'Transferência',
      Observacoes: 'Campanha Luz Viva 2024',
      RegistradoEm: new Date('2024-04-22')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador2,
      Valor: 90000,
      DataDeposito: new Date('2024-05-18'),
      Metodo: 'Transferência',
      Observacoes: 'Complemento para Assentamento Paz',
      RegistradoEm: new Date('2024-05-18')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador3,
      Valor: 150000,
      DataDeposito: new Date('2024-05-12'),
      Metodo: 'PIX corporativo',
      Observacoes: 'Projeto Vila Esperança',
      RegistradoEm: new Date('2024-05-12')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador4,
      Valor: 55000,
      DataDeposito: new Date('2024-05-18'),
      Metodo: 'PIX',
      Observacoes: 'Campanha cooperada - primeira parcela',
      RegistradoEm: new Date('2024-05-18')
    },
    {
      ID: Utilities.getUuid(),
      DoadorID: doador4,
      Valor: 40000,
      DataDeposito: new Date('2024-05-30'),
      Metodo: 'Boleto',
      Observacoes: 'Complemento após assembleia',
      RegistradoEm: new Date('2024-05-30')
    }
  ];

  depositosExemplo.forEach(deposito => {
    const linha = COLUNAS_DEPOSITOS.map(coluna => deposito[coluna] !== undefined ? deposito[coluna] : '');
    shDepositos.appendRow(linha);
  });

  shPrest.appendRow([idPoco1,new Date('2024-04-25'),'Topografia e mobilização inicial',22000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-01'),'Compra de tubos e bombas',35000,'','Perfuração','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-05'),'Serviço de instalação completa',42000,'','Instalação','Maria Silva']);
  shPrest.appendRow([idPoco1,new Date('2024-05-18'),'Treinamento da comunidade',12000,'','Instalação','João Pereira']);
  shPrest.appendRow([idPoco1,new Date('2024-05-25'),'Pagamento fornecedor',85000,'','Pagamento','Maria Silva']);
  shPrest.appendRow([idPoco2,new Date('2024-05-14'),'Adiantamento fornecedor perfuração',35000,'','Perfuração','Carlos Nunes']);
  shPrest.appendRow([idPoco2,new Date('2024-05-20'),'Logística de equipamentos',18000,'','Logística','Carlos Nunes']);
  shPrest.appendRow([idPoco3,new Date('2024-05-22'),'Consultoria hidrogeológica',25000,'','Perfuração','Luciana Prado']);
  shPrest.appendRow([idPoco3,new Date('2024-05-28'),'Reserva de materiais elétricos',8000,'','Instalação','Carlos Menezes']);

  shContatos.appendRow([Utilities.getUuid(), idPoco1, 'Maria Silva', 'João Pereira', 'Comunidade Vida', new Date('2024-06-02'),'Monitoramento pós-instalação', 'Organizar treinamento da comunidade', 'Monitoramento', 'Indicadores positivos mantidos', 'Maria Silva']);
  shContatos.appendRow([Utilities.getUuid(), idPoco2, 'Carlos Nunes', 'Eng. Ana Ramos', 'Poços Brasil', new Date('2024-05-28'),'Reunião para alinhar logística', 'Validar plano de instalação revisado', 'Em negociação', 'Prazo de instalação depende do novo cronograma', 'Carlos Nunes']);
  shContatos.appendRow([Utilities.getUuid(), idPoco3, 'Luciana Prado', 'Carlos Menezes', 'Fonte Limpa Engenharia', new Date('2024-05-20'), 'Análise de disponibilidade do equipamento', 'Confirmar disponibilidade do perfuratriz', 'Aguardando liberação', 'Atraso pode impactar início do projeto', 'Luciana Prado']);
  shContatos.appendRow([Utilities.getUuid(), idPoco4, 'Roberto Lima', 'Eng. Paula Duarte', 'Nordeste Perfurações', new Date('2024-05-27'), 'Licenciamento em andamento', 'Enviar documentação complementar', 'Documentação', 'Dependente de licenças ambientais', 'Roberto Lima']);

  if (shEmpresas) {
    shEmpresas.appendRow([Utilities.getUuid(),'Águas do Sertão','12.345.678/0001-90','Perfuração','(11) 3000-0000','Responsável pelo poço Comunidade Vida']);
    shEmpresas.appendRow([Utilities.getUuid(),'Poços Brasil','98.765.432/0001-10','Instalação','(86) 98888-1234','Equipe destacada para Assentamento Paz']);
    shEmpresas.appendRow([Utilities.getUuid(),'Fonte Limpa Engenharia','54.321.678/0001-55','Consultoria','(74) 97777-4567','Parceria técnica para Vila Esperança']);
    shEmpresas.appendRow([Utilities.getUuid(),'Nordeste Perfurações','43.210.987/0001-32','Perfuração','(98) 95555-9988','Responsável pela perfuração em Serra Azul']);
  }

  return '✅ Dados de exemplo carregados com sucesso. Ajuste conforme necessário.';
}

