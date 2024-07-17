function FormRegistroDocumentos() {
// 12:25
  var Planilha = SpreadsheetApp.getActiveSpreadsheet();
  // Referenciando a aba da Planilha
  var GuiaDocumento = Planilha.getSheetByName("Documentos");
  var ultimaLinha = GuiaDocumento.getLastRow();
  var GuiaAcervo = Planilha.getSheetByName("Listas_Suspensas");
  var GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
  var GuiaTipoPesquisa = Planilha.getSheetByName("Listas_Suspensas");
  var GuiaColaborador = Planilha.getSheetByName("Funcionarios");
  var GuiaAcervo = Planilha.getSheetByName("Listas_Suspensas");
  var GuiaStatus = Planilha.getSheetByName("Listas_Suspensas");


  var DadosDocumentos = GuiaDocumento.getRange(2, 2, ultimaLinha, 1).getValues();

  // Pegar última linha
  var ultimaLinhaPesquisador = GuiaPesquisador.getLastRow() - 1;
  var ultimaLinha = GuiaColaborador.getLastRow() - 1;


  // Percorrer todos os dados do array DadosDocumentos
  var acervo = {}
  for (var i = 0; i < DadosDocumentos.length; i++) {
    acervo[DadosDocumentos[i][0]] = DadosDocumentos[i][0];
  }
  var listaUnicaAcervo = [];
  for (var key in acervo) {
    listaUnicaAcervo.push([key]);
  }

  var list = GuiaPesquisador.getRange(2, 3, GuiaPesquisador.getRange("C2").getDataRegion().getLastRow(), 1).getValues();
  list.sort();
  var listaAcervo = GuiaAcervo.getRange(2, 15, GuiaAcervo.getRange("O2").getDataRegion().getLastRow(), 1).getValues();
  var listaTipoPesquisa = GuiaTipoPesquisa.getRange(2, 17, GuiaTipoPesquisa.getRange("Q2").getDataRegion().getLastRow(), 1).getValues();
  var listaColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 1).getValues();
  var listaStatus = GuiaStatus.getRange(2, 21, GuiaStatus.getRange("U2").getDataRegion().getLastRow(), 1).getValues();
  

  // Transferindo os dados para outra variável
  DadosDocumentos.length = 0;
  var pesquisaListaAcervo = listaUnicaAcervo;
  pesquisaListaAcervo.sort();
  listaTipoPesquisa.sort();
  listaColaborador.sort();

   // Criar a visualização do html
  var Form = HtmlService.createTemplateFromFile("FormRegistroPesquisa");
  Form.listaAcervo = listaAcervo.map(function(r) {return r[0];});
  Form.pesquisaListaAcervo = pesquisaListaAcervo.map(function(r) {return r[0];});
  Form.list = list.map(function (r) { return r[0]; });
  Form.listaTipoPesquisa = listaTipoPesquisa.map(function (r) { return r[0]; });
  Form.listaColaborador = listaColaborador.map(function(r) {return r[0];});
  Form.listaStatus = listaStatus.map(function(r) {return r[0];});

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  MostrarForm.setTitle("Registro de consulta").setHeight(530).setWidth(1500);
  // Exibindo formulário
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Registro de consulta");
}

    // Função para carregar lista suspensa dependente

    function ListaEspecies(Acervo) {
      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaAcervo = Planilha.getSheetByName("Documentos");
      var ultimaLinha = GuiaAcervo.getLastRow();
      var DadosEspecie = GuiaAcervo.getRange(2, 2, ultimaLinha, 2).getValues();
      var especies = [];

      // Capturar as espécies apenas da linha selecionada
      for (var i = 0; i < DadosEspecie.length; i++) {
        if (DadosEspecie[i][0] == Acervo) {
          especies.push([DadosEspecie[i][1]]);
        }
      }

      // Remover duplicatas da lista
      var conjuntoUnico = new Set(especies.map(JSON.stringify));
      var especiesUnicas = Array.from(conjuntoUnico).map(JSON.parse);

      DadosEspecie.length = 0;

      // Retornando a lista de identificadores referentes à linha selecionada
      return especiesUnicas.sort();
    }

    // Função para carregar lista suspensa dependente

  function ListaIdentificador1(Acervo, Especie) {
    var Planilha = SpreadsheetApp.getActiveSpreadsheet();
    var GuiaAcervo = Planilha.getSheetByName("Documentos");
    var ultimaLinha = GuiaAcervo.getLastRow();
    var DadosIdentificador1 = GuiaAcervo.getRange(2, 2, ultimaLinha, 3).getValues();
    var identificador1 = [];

    // Capturar os identificadores apenas da espécie documental selecionada
    for (var i = 0; i < DadosIdentificador1.length; i++) {
      if (DadosIdentificador1[i][0] == Acervo && DadosIdentificador1[i][1] == Especie) {
        identificador1.push([DadosIdentificador1[i][2]]);
      }
    }

    // Remover duplicatas da lista
    var conjuntoUnico = new Set(identificador1.map(JSON.stringify));
    var identificadoresUnicos = Array.from(conjuntoUnico).map(JSON.parse);

    DadosIdentificador1.length = 0;
    
    // Retornando a lista de identificadores referente à espécie documental selecionada
    return identificadoresUnicos.sort();
}


 // Função para carregar lista suspensa dependente

function ListaIdentificador2(Acervo, Especie, Identificador1) {
  var Planilha = SpreadsheetApp.getActiveSpreadsheet();
  var GuiaAcervo = Planilha.getSheetByName("Documentos");
  var ultimaLinha = GuiaAcervo.getLastRow();
  var DadosIdentificador2 = GuiaAcervo.getRange(2, 2, ultimaLinha, 4).getValues();
  var identificador2 = [];

  // Capturar os identificadores apenas do Identificador1 selecionado
  for (var i = 0; i < DadosIdentificador2.length; i++) {
    if (DadosIdentificador2[i][0] == Acervo && DadosIdentificador2[i][1] == Especie && DadosIdentificador2[i][2] == Identificador1) {
      identificador2.push([DadosIdentificador2[i][3]]);
    }
  }

  // Remover duplicatas da lista
  var conjuntoUnico = new Set(identificador2.map(JSON.stringify));
  var identificadoresUnicos = Array.from(conjuntoUnico).map(JSON.parse);

  DadosIdentificador2.length = 0;

  // Retornando a lista de identificadores referente ao Identificador1 selecionado
  return identificadoresUnicos.sort();
}


//  // Função para carregar lista suspensa dependente
 function ListaDocumento(Acervo, Especie, Identificador1, Identificador2) {
      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaAcervo = Planilha.getSheetByName("Documentos");
      var ultimaLinha = GuiaAcervo.getLastRow();
      var DadosDocumento = GuiaAcervo.getRange(2, 2, ultimaLinha, 5).getValues();
      var documento = [];

      // Capturar os identificadores apenas da espécie documental selecionada
      for (var i = 0; i < DadosDocumento.length; i++) {
          if (DadosDocumento[i][0] == Acervo && DadosDocumento[i][1] == Especie && DadosDocumento[i][2] == Identificador1 && DadosDocumento[i][3] == Identificador2) {
              documento.push([DadosDocumento[i][4]]);
          }
      }
      DadosDocumento.length = 0;
    // Retornando a lista de identificadores referente à espécie documental selecionada
    return documento.sort();
}

// Função para salvar os registros na tabela

function SalvarRegistroDocumento(Dados) {
  const user = LockService.getScriptLock();
  // Tempo para nova tentativa de realizar o script
  user.tryLock(10000);

  // Se usuário estiver liberado, seguirá adiante, para salvar os dados na Planilha
  if (user.hasLock) {
    var Planilha = SpreadsheetApp.getActiveSpreadsheet();
    var GuiaRegistroPesquisa = Planilha.getSheetByName("Registros_de_Pesquisas");

    let ultimaLinha = GuiaRegistroPesquisa.getLastRow();

    // Verificar se existem registros anteriores para comparar
    if (ultimaLinha > 1) {
      var ultimaData = GuiaRegistroPesquisa.getRange(ultimaLinha, 6).getValue();
      var ultimoPesquisador = GuiaRegistroPesquisa.getRange(ultimaLinha, 4).getValue();
      var ultimoId = GuiaRegistroPesquisa.getRange(ultimaLinha, 2).getValue();
      
      // Verificar se os valores de data e pesquisador são iguais aos do último registro
      if (ultimaData == Dados.DataPesquisa && ultimoPesquisador == Dados.Pesquisador) {
        // Atribuir o mesmo ID do registro anterior
        GuiaRegistroPesquisa.getRange(ultimaLinha + 1, 2).setValue(ultimoId);
      } else {
        // Atribuir um novo ID sequencial
        GuiaRegistroPesquisa.getRange(ultimaLinha + 1, 2).setValue(ultimoId+1);
      }
    } else {
      // Se for o primeiro registro, atribuir ID sequencial 1
      GuiaRegistroPesquisa.getRange(2, 2).setValue(1);
    }
    

    // Continuar salvando os outros dados
    let linha = ultimaLinha + 1;
    GuiaRegistroPesquisa.getRange(linha, 1).setValue(linha - 1); // Adiciona ID sequencial
    GuiaRegistroPesquisa.getRange(linha, 3).setValue(Dados.Pesquisador_Id);
    GuiaRegistroPesquisa.getRange(linha, 4).setValue(Dados.Pesquisador);
    GuiaRegistroPesquisa.getRange(linha, 5).setValue(Dados.Assunto);
    GuiaRegistroPesquisa.getRange(linha, 6).setValue(Dados.DataPesquisa);
    GuiaRegistroPesquisa.getRange(linha, 7).setValue(Dados.TipoPesquisa);
    GuiaRegistroPesquisa.getRange(linha, 8).setValue(Dados.IdColaborador);
    GuiaRegistroPesquisa.getRange(linha, 9).setValue(Dados.ListaColaborador);
    GuiaRegistroPesquisa.getRange(linha, 10).setValue(Dados.Status);
    GuiaRegistroPesquisa.getRange(linha, 11).setValue(Dados.Acervo);
    GuiaRegistroPesquisa.getRange(linha, 12).setValue(Dados.EspecieDocumental);
    GuiaRegistroPesquisa.getRange(linha, 13).setValue(Dados.Identificador1);
    GuiaRegistroPesquisa.getRange(linha, 14).setValue(Dados.Identificador2);
    GuiaRegistroPesquisa.getRange(linha, 15).setValue(Dados.Documento);
    GuiaRegistroPesquisa.getRange(linha, 16).setValue(Dados.IdDocumento);
    GuiaRegistroPesquisa.getRange(linha, 17).setValue(Dados.Periodo);
    GuiaRegistroPesquisa.getRange(linha, 18).setValue(Dados.Volume);
    GuiaRegistroPesquisa.getRange(linha, 19).setValue(Dados.Observacao);

    // Remover o formato de data
    GuiaRegistroPesquisa.getRange("F:F").setNumberFormat("@");
    GuiaRegistroPesquisa.getRange("Q:Q").setNumberFormat("@");

    return "Registro de consulta registrada com sucesso!";
  }
}

// Função para salvar todos os dados do formulário e apagar todos os campos posteriormente para adicionar novo pesquisador
function SalvareRegistrarNovoPesquisador(Dados) {
  const user = LockService.getScriptLock();
  // Tempo para nova tentativa de realizar o script
  user.tryLock(10000);

  // Se usuário estiver liberado, seguirá adiante, para salvar os dados na Planilha
  if (user.hasLock) {
    var Planilha = SpreadsheetApp.getActiveSpreadsheet();
    var GuiaRegistroPesquisa = Planilha.getSheetByName("Registros_de_Pesquisas");

    let ultimaLinha = GuiaRegistroPesquisa.getLastRow();

    // Verificar se existem registros anteriores para comparar
    if (ultimaLinha > 1) {
      var ultimaData = GuiaRegistroPesquisa.getRange(ultimaLinha, 6).getValue();
      var ultimoPesquisador = GuiaRegistroPesquisa.getRange(ultimaLinha, 4).getValue();
      var ultimoId = GuiaRegistroPesquisa.getRange(ultimaLinha, 2).getValue();
      
      // Verificar se os valores de data e pesquisador são iguais aos do último registro
      if (ultimaData == Dados.DataPesquisa && ultimoPesquisador == Dados.Pesquisador) {
        // Atribuir o mesmo ID do registro anterior
        GuiaRegistroPesquisa.getRange(ultimaLinha + 1, 2).setValue(ultimoId);
      } else {
        // Atribuir um novo ID sequencial
        GuiaRegistroPesquisa.getRange(ultimaLinha + 1, 2).setValue(ultimoId+1);
      }
    } else {
      // Se for o primeiro registro, atribuir ID sequencial 1
      GuiaRegistroPesquisa.getRange(2, 2).setValue(1);
    }
    

    // Continuar salvando os outros dados
    let linha = ultimaLinha + 1;
    GuiaRegistroPesquisa.getRange(linha, 1).setValue(linha - 1); // Adiciona ID sequencial
    GuiaRegistroPesquisa.getRange(linha, 3).setValue(Dados.Pesquisador_Id);
    GuiaRegistroPesquisa.getRange(linha, 4).setValue(Dados.Pesquisador);
    GuiaRegistroPesquisa.getRange(linha, 5).setValue(Dados.Assunto);
    GuiaRegistroPesquisa.getRange(linha, 6).setValue(Dados.DataPesquisa);
    GuiaRegistroPesquisa.getRange(linha, 7).setValue(Dados.TipoPesquisa);
    GuiaRegistroPesquisa.getRange(linha, 8).setValue(Dados.IdColaborador);
    GuiaRegistroPesquisa.getRange(linha, 9).setValue(Dados.ListaColaborador);
    GuiaRegistroPesquisa.getRange(linha, 10).setValue(Dados.Status);
    GuiaRegistroPesquisa.getRange(linha, 11).setValue(Dados.Acervo);
    GuiaRegistroPesquisa.getRange(linha, 12).setValue(Dados.EspecieDocumental);
    GuiaRegistroPesquisa.getRange(linha, 13).setValue(Dados.Identificador1);
    GuiaRegistroPesquisa.getRange(linha, 14).setValue(Dados.Identificador2);
    GuiaRegistroPesquisa.getRange(linha, 15).setValue(Dados.Documento);
    GuiaRegistroPesquisa.getRange(linha, 16).setValue(Dados.IdDocumento);
    GuiaRegistroPesquisa.getRange(linha, 17).setValue(Dados.Periodo);
    GuiaRegistroPesquisa.getRange(linha, 18).setValue(Dados.Volume);
    GuiaRegistroPesquisa.getRange(linha, 19).setValue(Dados.Observacao);

    // Remover o formato de data
    GuiaRegistroPesquisa.getRange("F:F").setNumberFormat("@");
    GuiaRegistroPesquisa.getRange("Q:Q").setNumberFormat("@");

    return "Registro de consulta registrada com sucesso!";
  }
}


  // Função para atualizar as linhas
    function AtualizarListaAcervos(){

    var Planilha = SpreadsheetApp.getActiveSpreadsheet();
    var GuiaAcervo = Planilha.getSheetByName("Documentos");

    var ultimaLinha = GuiaAcervo.getLastRow();
   // Capturando os dados da coluna A

    var DadosAcervo = GuiaAcervo.getRange(2,2,ultimaLinha,1).getValues();
    var b = {};

    for(var i = 0; i < DadosAcervo.length; i++){
      b[DadosAcervo[i][0]] = DadosAcervo[i][0];
    }
    // Vai montar uma lista única das linhas dos acervos
    var listaUnica = [];

    for(var key in b){
      listaUnica.push([key]);
    }

    DadosAcervo.length = 0;
    return listaUnica.sort();
  }


// Função para carregar o campo ID do pesquisador 
function Pesquisador_Id(Pesquisador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaPesquisador.getLastRow();
  // let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 19).getValues();
  let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 3).getValues();

  // // Pega a informação do número do identificador do pesquisador
  let IdPesquisador = GuiaPesquisador.getRange(2, 1, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosPesquisador.length; i++) {
    if(DadosPesquisador[i][0] == Pesquisador){
      let Id_Pesquisador = IdPesquisador[i];
      DadosPesquisador.length = 0;
       return ([Id_Pesquisador]);
    }
  };
}

// Função para pesquisar assunto de pesquisador
function Pesquisa_assunto(Pesquisador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaPesquisador.getLastRow();
  let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 3).getValues();

  // Pega a informação do assunto do pesquisador
  var Assunto = GuiaPesquisador.getRange(2, 20, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosPesquisador.length; i++) {
    if(DadosPesquisador[i][0] == Pesquisador){
      var Assunto = Assunto[i];
      DadosPesquisador.length = 0;
       return ([Assunto]);
    }
  };
}

// Função para carregar o campo ID do pesquisador 
function Colaborador_Id(Colaborador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaColaborador = Planilha.getSheetByName("Funcionarios");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaColaborador.getLastRow();
  let DadosColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 3).getValues();

  // // Pega a informação do número do identificador do pesquisador
  let IdColaborador = GuiaColaborador.getRange(2, 1, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosColaborador.length; i++) {
    if(DadosColaborador[i][0] == Colaborador){
      let Id_Colaborador = IdColaborador[i];
      DadosColaborador.length = 0;
       return ([Id_Colaborador]);
    }
  };
}

function Documento_Id(Documento){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaDocumento = Planilha.getSheetByName("Documentos");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaDocumento.getLastRow();
  let DadosDocumento = GuiaDocumento.getRange(2, 6, ultimaLinha, 6).getValues();

  // // Pega a informação do número do identificador do pesquisador
  let IdDocumento = GuiaDocumento.getRange(2, 1, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosDocumento.length; i++) {
    if(DadosDocumento[i][0] == Documento){
      let Id_Documento = IdDocumento[i];
      DadosDocumento.length = 0;
       return ([Id_Documento]);
    }
  };
}

