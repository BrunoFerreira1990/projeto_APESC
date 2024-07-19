function FormDocumentos() {
// 12:25
  var Planilha = SpreadsheetApp.getActiveSpreadsheet();
  // Referenciando a aba da Planilha
  var GuiaDocumento = Planilha.getSheetByName("Documentos");
  var ultimaLinha = GuiaDocumento.getLastRow();
  var GuiaAcervo = Planilha.getSheetByName("Listas_Suspensas");
  var DadosDocumentos = GuiaDocumento.getRange(2, 2, ultimaLinha, 1).getValues();

  // Percorrer todos os dados do array DadosDocumentos
  var acervo = {}
  for (var i = 0; i < DadosDocumentos.length; i++) {
    acervo[DadosDocumentos[i][0]] = DadosDocumentos[i][0];
  }
  var listaUnicaAcervo = [];
  for (var key in acervo) {
    listaUnicaAcervo.push([key]);
  }

  var listaAcervo = GuiaAcervo.getRange(2, 15, GuiaAcervo.getRange("O2").getDataRegion().getLastRow(), 1).getValues();

  // Transferindo os dados para outra variável
  DadosDocumentos.length = 0;
  var pesquisaListaAcervo = listaUnicaAcervo;
  pesquisaListaAcervo.sort();

   // Criar a visualização do html
  var Form = HtmlService.createTemplateFromFile("FormDocumentos");
  Form.listaAcervo = listaAcervo.map(function(r) {return r[0];});
  Form.pesquisaListaAcervo = pesquisaListaAcervo.map(function(r) {return r[0];});

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  MostrarForm.setTitle("Topográfico").setHeight(450).setWidth(1500);
  // Exibindo formulário
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Topográfico");
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

  // Função para salvar os especies
    function SalvarDocumento(Dados) {
      const user = LockService.getScriptLock();
    // Tempo para nova tentativa de realizar o script
    user.tryLock(10000);

    // Se usuário estiver liberado, seguirá adiante, para salvar os dados na Planilha
    if (user.hasLock) {
      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      var Identificador2 = Dados.Identificador2;
      var Documento = Dados.Documento;
      var Periodo = Dados.Periodo;
      var Volume = Dados.Volume;
      var Local = Dados.Local;
      var Observacao = Dados.Observacao;
     
      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaDocumento = Planilha.getSheetByName("Documentos");
      
      var ultimaLinha = GuiaDocumento.getLastRow();

      // Verificar se o documento já existe na Planilha
      let DadosDocumento = GuiaDocumento.getRange(2, 2, ultimaLinha, 2).getValues();

      for (var i = 0; i < DadosDocumento.length; i++) {
        if (DadosDocumento[i][0] == Acervo && DadosDocumento[i][1] == EspecieDocumental && DadosDocumento[i][2] == Identificador1 && DadosDocumento[i][3] == Identificador2 && DadosDocumento[i][4] == Documento) {
          return "Documento/Códice já cadastrado!";
        }
      }
      DadosDocumento.length = 0;
      // Se caso percorreu a lista e não tem o documento cadastrado, será um documento novo
      let linha = GuiaDocumento.getLastRow() + 1;

      GuiaDocumento.getRange(linha, 1).setValue(linha - 1); // Adiciona ID sequencial
      GuiaDocumento.getRange(linha, 2).setValue(Dados.Acervo);
      GuiaDocumento.getRange(linha, 3).setValue(Dados.EspecieDocumental);
      GuiaDocumento.getRange(linha, 4).setValue(Dados.Identificador1);
      GuiaDocumento.getRange(linha, 5).setValue(Dados.Identificador2);
      GuiaDocumento.getRange(linha, 6).setValue(Dados.Documento);
      GuiaDocumento.getRange(linha, 7).setValue(Dados.Periodo);
      GuiaDocumento.getRange(linha, 8).setValue(Dados.Volume);
      GuiaDocumento.getRange(linha, 9).setValue(Dados.Local);
      GuiaDocumento.getRange(linha, 10).setValue(Dados.Observacao);

      // Classificando a lista de documentos deixando em ordem alfabética
      // GuiaDocumento.getRange("B2:J").sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 3, ascending: true}, {column: 4, ascending: true}, {column: 5, ascending: true}]);
    
      return "Documento/Códice registrado com sucesso!";
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

  // Função de pesquisar espécie

  
  function PesquisarEspecie(Dados){

      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaEspecie = Planilha.getSheetByName("Documentos");
      var ultimaLinha = GuiaEspecie.getLastRow();
      var DadosEspecie = GuiaEspecie.getRange(2,2,ultimaLinha,3).getValues();

      for(var i = 0; i <DadosEspecie.length; i++){

        if(DadosEspecie[i][0] == Acervo && DadosEspecie[i][1] == EspecieDocumental && DadosEspecie[i][2] == Identificador1){ 
          var Identificador1 = DadosEspecie[i][2];
          DadosEspecie.length = 0;
          return ([Acervo, EspecieDocumental, Identificador1]);
        }
        }
          DadosEspecie.length = 0;
       
        // return "NÃO ENCONTRADO!";
    }

    // Função de pesquisar id1

  function PesquisarIdentificador1(Dados){

      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      // var Identificador2 = Dados.Identificador2;
      // var Documento = Dados.Documento;

      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaIdentificador1 = Planilha.getSheetByName("Documentos");
      var ultimaLinha = GuiaIdentificador1.getLastRow();
      
      var DadosIdentificador1 = GuiaIdentificador1.getRange(2,2,ultimaLinha,4).getValues();

      for(var i = 0; i <DadosIdentificador1.length; i++){

        if(DadosIdentificador1[i][0] == Acervo && DadosIdentificador1[i][1] == EspecieDocumental && DadosIdentificador1[i][2] == Identificador1){ 
          var Identificador2 = DadosIdentificador1[i][3];

          // var Preco = DadosIdentificador1[i][2].toLocaleString({style: 'decimal',decimal: 'pt-BR'});
          // var Preco = Preco.replace(/\./g,"");
          DadosIdentificador1.length = 0;

          return ([Acervo, EspecieDocumental, Identificador1, Identificador2]);

        }
        }
        DadosIdentificador1.length = 0;
        // return "NÃO ENCONTRADO!";
    }

       // Função de pesquisar id2

  function PesquisarIdentificador2(Dados){

      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      var Identificador2 = Dados.Identificador2;
      // var Documento = Dados.Documento;

      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaIdentificador2 = Planilha.getSheetByName("Documentos");

      var ultimaLinha = GuiaIdentificador2.getLastRow();

      var DadosIdentificador2 = GuiaIdentificador2.getRange(2,2,ultimaLinha,5).getValues();

      for(var i = 0; i <DadosIdentificador2.length; i++){

        if(DadosIdentificador2[i][0] == Acervo && DadosIdentificador2[i][1] == EspecieDocumental && DadosIdentificador2[i][2] == Identificador1  && DadosIdentificador2[i][3] == Identificador2 ){ 


          var Documento = DadosIdentificador2[i][4];
          
          DadosIdentificador2.length = 0;

          return ([Acervo, EspecieDocumental, Identificador1, Identificador2, Documento]); //, Documento, Periodo, Volume, Local, Observacao

        }
        }

        DadosIdentificador2.length = 0;
        // return "NÃO ENCONTRADO!";
    }

          // Função de pesquisar documento

  function PesquisarDocumento(Dados){

      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      var Identificador2 = Dados.Identificador2;
      var Documento = Dados.Documento;

      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaDocumento = Planilha.getSheetByName("Documentos");

      var ultimaLinha = GuiaDocumento.getLastRow();

      var DadosDocumento = GuiaDocumento.getRange(2,2,ultimaLinha,10).getValues();

      for(var i = 0; i <DadosDocumento.length; i++){

        if(DadosDocumento[i][0] == Acervo && DadosDocumento[i][1] == EspecieDocumental && DadosDocumento[i][2] == Identificador1  && DadosDocumento[i][3] == Identificador2 &&  DadosDocumento[i][4] == Documento){ 

          var Periodo = DadosDocumento[i][5];
          var Volume = DadosDocumento[i][6];
          var Local = DadosDocumento[i][7];
          let Observacao = DadosDocumento[i][8];

          
          DadosDocumento.length = 0;

          return ([Acervo, EspecieDocumental, Identificador1, Identificador2, Documento, Periodo, Volume, Local, Observacao]); //, Documento, Periodo, Volume, Local, Observacao
        }
        }
        DadosDocumento.length = 0;
        // return "NÃO ENCONTRADO!";
    }



    // Função de edição
    function EditarDocumento(Dados){

    const user = LockService.getScriptLock();
    user.tryLock(10000);

    if(user.hasLock()){

      var PesquisaAcervo = Dados.PesquisaAcervo;
      var PesquisaEspecie = Dados.PesquisaEspecie;
      var PesquisaId1 = Dados.PesquisaId1;
      var PesquisaId2 = Dados.PesquisaId2;
      var PesquisaDocumento = Dados.PesquisaDocumento;

      var Acervo = Dados.Acervo;
      var EspecieDocumental = Dados.EspecieDocumental;
      var Identificador1 = Dados.Identificador1;
      var Identificador2 = Dados.Identificador2;
      var Documento = Dados.Documento;
      var Volume = Dados.Volume;
      var Periodo = Dados.Periodo;
      var Local = Dados.Local;
      var Observacao = Dados.Observacao; 

      var Planilha = SpreadsheetApp.getActiveSpreadsheet();
      var GuiaAcervo = Planilha.getSheetByName("Documentos");
      // var guiaPedido = Planilha.getSheetByName("Pedidos");

      var ultimaLinha = GuiaAcervo.getLastRow();

      var DadosAcervo = GuiaAcervo.getRange(2,2,ultimaLinha,9).getValues();

      // var ultimaLinha = guiaPedido.getLastRow();

      // var dadosPedidos = guiaPedido.getRange(2,6,ultimaLinha,2).getValues();

      // var ver = dadosPedidos.filter(function(value,i,arr){
      //   return LinhaLista == arr[i][0] && ProdutoLista == arr[i][1];
      // })

      for(var i = 0; i <DadosAcervo.length; i++){

        if(DadosAcervo[i][0] == PesquisaAcervo && DadosAcervo[i][1] == PesquisaEspecie  && DadosAcervo[i][2] == PesquisaId1 && DadosAcervo[i][3] == PesquisaId2 && DadosAcervo[i][4] == PesquisaDocumento){

            var linha = i + 2;

            // if(ver.length < 1){

              GuiaAcervo.getRange(linha,2).setValue(Acervo);
              GuiaAcervo.getRange(linha,3).setValue(EspecieDocumental);
              GuiaAcervo.getRange(linha,4).setValue(Identificador1);
              GuiaAcervo.getRange(linha,5).setValue(Identificador2);
              GuiaAcervo.getRange(linha,6).setValue(Documento);
              GuiaAcervo.getRange(linha,7).setValue(Periodo);
              GuiaAcervo.getRange(linha,8).setValue(Volume);
              GuiaAcervo.getRange(linha,9).setValue(Local);
              GuiaAcervo.getRange(linha,10).setValue(Observacao);
              
              DadosAcervo.length = 0;
              // dadosPedidos.length = 0;

              return "Editado com sucesso!";

            }

            // if(ver.length > 0){

            //   GuiaAcervo.getRange(linha,3).setValue(Preco);

            //   DadosAcervo.length = 0;
            //   // dadosPedidos.length = 0;
            //   // ver.length = 0;

            //   return "EDITADO APENAS O PREÇO. PRODUTO JÁ POSSUI LANÇAMENTO DE PEDIDO.";

            // }

        }

      }

    DadosAcervo.length = 0;
    // dadosPedidos.length = 0;
    // ver.length = 0;

    return "Documento não encontrado!";

}


