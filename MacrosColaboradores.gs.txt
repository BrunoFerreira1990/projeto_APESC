function FormColaboradores() {
// Referenciar a planilha ativa
  var Planilha = SpreadsheetApp.getActiveSpreadsheet();
  // Referenciando a aba da planilha
  var GuiaColaborador = Planilha.getSheetByName("Funcionarios");
  var GuiaCargo = Planilha.getSheetByName("Listas_Suspensas");
  var GuiaLotacao = Planilha.getSheetByName("Listas_Suspensas");


// Carregando a lista dos colaboradores
  var ultimaLinha = GuiaColaborador.getLastRow() - 1;

  if (ultimaLinha == 0) {
    ultimaLinha = 1
  }

  let listaColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 1).getValues();
  listaColaborador.sort();
  let listaCargo = GuiaCargo.getRange(2, 13, GuiaCargo.getRange("M2").getDataRegion().getLastRow(), 1).getValues();
  let listaLotacao = GuiaLotacao.getRange(2, 11, GuiaLotacao.getRange("K2").getDataRegion().getLastRow(), 1).getValues();

  // Criar a visualização do html
  var Form = HtmlService.createTemplateFromFile("FormColaboradores");
  Form.listaColaborador = listaColaborador.map(function(r) {return r[0];});
  Form.listaCargo = listaCargo.map(function(r2) {return r2[0];});
  Form.listaLotacao = listaLotacao.map(function(r3) {return r3[0];});

  

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  MostrarForm.setTitle("Cadastro de Agentes Públicos").setHeight(600).setWidth(700);
  // Exibindo formulário
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Cadastro de Agentes Públicos");
}

function SalvarColaborador(Dados) {
  const user = LockService.getScriptLock();
  // Tempo para nova tentativa de realizar o script
  user.tryLock(10000);

  // Se usuário estiver liberado, seguirá adiante, para salvar os dados na planilha
  if (user.hasLock) {
    let Planilha = SpreadsheetApp.getActiveSpreadsheet();
    let GuiaColaborador = Planilha.getSheetByName("Funcionarios");

    let ultimaLinha = GuiaColaborador.getLastRow();

    // Verificar o pesquisador já existe na planilha
    let DadosColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 1).getValues();

    for (var i = 0; i < DadosColaborador.length; i++) {
      if (DadosColaborador[i][0] == Dados.Colaborador) {
        return "Colaborador já cadastrado!";
      }
    }

    DadosColaborador.length = 0;

    // Se caso percorreu a lista e não tem o nome da pessoa, será um colaborador novo
    let Linha = ultimaLinha + 1;

    var Data = new Date();

    GuiaColaborador.getRange(Linha, 1).setValue(Linha - 1); // Adiciona ID sequencial
    GuiaColaborador.getRange(Linha, 2).setValue(Data);
    GuiaColaborador.getRange(Linha, 3).setValue(Dados.Colaborador);
    GuiaColaborador.getRange(Linha, 4).setValue(Dados.Matricula);
    GuiaColaborador.getRange(Linha, 5).setValue(Dados.ListaCargo);
    GuiaColaborador.getRange(Linha, 6).setValue(Dados.ListaLotacao);
    GuiaColaborador.getRange(Linha, 7).setValue(Dados.Telefone);
    GuiaColaborador.getRange(Linha, 8).setValue(Dados.Email);

    // Remover o formato de data
    GuiaColaborador.getRange("B:B").setNumberFormat("@");
    
    // Obter os valores da coluna B como matriz
    var valores = GuiaColaborador.getRange("B2:B" + ultimaLinha).getValues();

    // Converter os valores para strings
    for (var i = 0; i < valores.length; i++) {
      valores[i][0] = valores[i][0].toString();
    }

    // Definir os novos valores na coluna B
    GuiaColaborador.getRange("B2:B" + ultimaLinha).setValues(valores);


    return "Funcionário registrado com sucesso!";
  }
}

// Função para preencher os campos com os dados do colaborador selecionado pela lista
function PesquisarColaborador(NomeColaborador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaColaborador = Planilha.getSheetByName("Funcionarios");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaColaborador.getLastRow();
  let DadosColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 6).getValues();

  for(let i = 0; i < DadosColaborador.length; i++) {
    if(DadosColaborador[i][0] == NomeColaborador){

      // let IdNumero = IdPesquisador[i];
      // let RegistroData = DataRegistro[i];
      let Colaborador = DadosColaborador[i][0];
      let Matricula = DadosColaborador[i][1];
      let ListaCargo = DadosColaborador[i][2];
      let ListaLotacao = DadosColaborador[i][3];
      let Telefone = DadosColaborador[i][4];
      let Email = DadosColaborador[i][5];
      

      DadosColaborador.length = 0;

      return ([Colaborador, Matricula, ListaCargo, ListaLotacao, Telefone, Email]);
    }
  };
  // Caso o pesquisador não for encontrado
  DadosColaborador.length = 0;
  return "Agente Público não encontrado!"
}

// Editar colaboradors
function EditarColaborador(Dados) {

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if (user.hasLock()) {
    let Planilha = SpreadsheetApp.getActiveSpreadsheet();
    let GuiaColaborador = Planilha.getSheetByName("Funcionarios");

    let ultimaLinha = GuiaColaborador.getLastRow();
    let DadosColaborador = GuiaColaborador.getRange(2, 3, ultimaLinha, 1).getValues();


    for(var i = 0; i < DadosColaborador.length; i++){
      if(DadosColaborador[i][0] == Dados.ListaColaborador){
        let Linha = i + 2;
        GuiaColaborador.getRange(Linha, 3).setValue(Dados.Colaborador);
        GuiaColaborador.getRange(Linha, 4).setValue(Dados.Matricula);
        GuiaColaborador.getRange(Linha, 5).setValue(Dados.ListaCargo);
        GuiaColaborador.getRange(Linha, 6).setValue(Dados.ListaLotacao);
        GuiaColaborador.getRange(Linha, 7).setValue(Dados.Telefone);
        GuiaColaborador.getRange(Linha, 8).setValue(Dados.Email);

        GuiaColaborador.getRange("A:A").setNumberFormat("@");
        GuiaColaborador.getRange("D:D").setNumberFormat("@");


      
        return "Agente Público editado com sucesso!";
      }
    }
    DadosColaborador.length = 0;
    // DadosRegistroPesquisa.length = 0;

    return "Pesquisador não encontrado!";
  }
}


// Chama a função na abertura do arquivo html para puxar os demais scripts

function Chamar(Arquivo) {
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}


// Atualizar a lista dos colaboradores
function AtualizarListaColaboradores() {

  // Referenciando a planilha
  var Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaColaborador = Planilha.getSheetByName("Funcionarios");

  let ultimaLinha = GuiaColaborador.getLastRow() - 1;
  let list = GuiaColaborador.getRange(2, 3, ultimaLinha, 1).getValues();

  return list.sort();
}





