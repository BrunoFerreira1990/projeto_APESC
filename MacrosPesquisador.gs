function FormPesquisador() {
  // Está pegando a tabela dos pesquisadores
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
  let GuiaEstado = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaFormacao = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaNacionalidade = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaPeriodoEstudo = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaAreaTrabalho = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaSexo = Planilha.getSheetByName("Listas_Suspensas");

  // Captura a última linha com dados
  let ultimaLinhaPesquisador = GuiaPesquisador.getLastRow() - 1;

  // Captura a lista de nome informando o número de início da linha e a coluna, e quantas colunas deseja capturar
  let list = GuiaPesquisador.getRange(2, 3, GuiaPesquisador.getRange("C2").getDataRegion().getLastRow(), 1).getValues();
  let list2 = GuiaNacionalidade.getRange(2, 5, GuiaNacionalidade.getRange("E2").getDataRegion().getLastRow(), 1).getValues();
  let list3 = GuiaEstado.getRange(2, 1, GuiaEstado.getRange("A2").getDataRegion().getLastRow(), 1).getValues();
  let list4 = GuiaFormacao.getRange(2, 3, GuiaFormacao.getRange("C2").getDataRegion().getLastRow(), 1).getValues();
  let list5 = GuiaPeriodoEstudo.getRange(2, 7, GuiaPeriodoEstudo.getRange("G2").getDataRegion().getLastRow(), 1).getValues();
  let list6 = GuiaAreaTrabalho.getRange(2, 9, GuiaAreaTrabalho.getRange("I2").getDataRegion().getLastRow(), 1).getValues();
  let list7 = GuiaSexo.getRange(2, 19, GuiaSexo.getRange("S2").getDataRegion().getLastRow(), 1).getValues();

  // Coloca a lista em ordem alfabética
  list.sort();

  // Abre o template html
  let Form = HtmlService.createTemplateFromFile("FormPesquisadores");

  // Pega cada nome da tabela de pesquisadores e transfere para o formulário no campo de seleção
  Form.list = list.map(function (r) { return r[0]; });
  Form.list2 = list2.map(function (r2) { return r2[0]; });
  Form.list3 = list3.map(function (r3) { return r3[0]; });
  Form.list4 = list4.map(function (r4) { return r4[0]; });
  Form.list5 = list5.map(function (r5) { return r5[0]; });
  Form.list6 = list6.map(function (r6) { return r6[0]; });
  Form.list7 = list7.map(function (r7) { return r7[0]; });

  let MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  // Insere um título no formulário
  MostrarForm.setTitle("Cadastro de Pesquisadores").setHeight(1000).setWidth(800);

  // Exibir o formulário
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Cadastro de Pesquisadores");
}

// Chama a função na abertura do arquivo html para puxar os demais scripts
function Chamar(Arquivo) {
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}

// Função trazida do FormPesquisadores-js
function SalvarPesquisador(Dados) {
  const user = LockService.getScriptLock();
  // Tempo para nova tentativa de realizar o script
  user.tryLock(10000);

  // Se usuário estiver liberado, seguirá adiante, para salvar os dados na planilha
  if (user.hasLock) {
    let Planilha = SpreadsheetApp.getActiveSpreadsheet();
    let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");

    let ultimaLinha = GuiaPesquisador.getLastRow();

    // Verificar o pesquisador já existe na planilha
    let DadosPesquisador = GuiaPesquisador.getRange(2, 6, ultimaLinha, 1).getValues();

    for (var i = 0; i < DadosPesquisador.length; i++) {
      if (DadosPesquisador[i][0] == Dados.Cpf) {
        return "Pesquisador já cadastrado!";
      }
    }
    DadosPesquisador.length = 0;

    // Se caso percorreu a lista e não tem o nome da pessoa, será um pesquisador novo
    let Linha = ultimaLinha + 1;

    var Data = new Date();

    GuiaPesquisador.getRange(Linha, 1).setValue(Linha - 1); // Adiciona ID sequencial
    GuiaPesquisador.getRange(Linha, 2).setValue(Data);
    GuiaPesquisador.getRange(Linha, 3).setValue(Dados.Pesquisador);
    GuiaPesquisador.getRange(Linha, 4).setValue(Dados.Nascimento);
    GuiaPesquisador.getRange(Linha, 5).setValue(Dados.Sexo);
    GuiaPesquisador.getRange(Linha, 6).setValue(Dados.Cpf);
    GuiaPesquisador.getRange(Linha, 7).setValue(Dados.Nacionalidade);
    GuiaPesquisador.getRange(Linha, 8).setValue(Dados.Telefone);
    GuiaPesquisador.getRange(Linha, 9).setValue(Dados.Email);
    GuiaPesquisador.getRange(Linha, 10).setValue(Dados.Endereco);
    GuiaPesquisador.getRange(Linha, 11).setValue(Dados.NumCasa);
    GuiaPesquisador.getRange(Linha, 12).setValue(Dados.Complemento);
    GuiaPesquisador.getRange(Linha, 13).setValue(Dados.Bairro);
    GuiaPesquisador.getRange(Linha, 14).setValue(Dados.Cidade);
    GuiaPesquisador.getRange(Linha, 15).setValue(Dados.Estado);
    GuiaPesquisador.getRange(Linha, 16).setValue(Dados.Formacao);
    GuiaPesquisador.getRange(Linha, 17).setValue(Dados.Instituicao);
    GuiaPesquisador.getRange(Linha, 18).setValue(Dados.Curso);
    GuiaPesquisador.getRange(Linha, 19).setValue(Dados.Profissao);
    GuiaPesquisador.getRange(Linha, 20).setValue(Dados.Assunto);
    GuiaPesquisador.getRange(Linha, 21).setValue(Dados.Finalidade);
    GuiaPesquisador.getRange(Linha, 22).setValue(Dados.PeriodoEstudo);
    GuiaPesquisador.getRange(Linha, 23).setValue(Dados.AreaTrabalho);
    GuiaPesquisador.getRange(Linha, 24).setValue(Dados.Observacoes);

    // Remover o formato de data e número
    GuiaPesquisador.getRange("B:B").setNumberFormat("@");
    GuiaPesquisador.getRange("D:D").setNumberFormat("@");

    
    return "Pesquisador registrado com sucesso!";
  }
}


// Função para preencher os campos com os dados do pesquisador selecionado pela lista
function PesquisarPesquisador(NomePesquisador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaPesquisador.getLastRow();
  let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 24).getValues();
      // Pega a informação do número do identificador do pesquisador
  let IdPesquisador = GuiaPesquisador.getRange(2, 1, ultimaLinha, 1).getValues();
  // Pega a informação da data de registro do pesquisador
  let DataRegistro = GuiaPesquisador.getRange(2, 2, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosPesquisador.length; i++) {
    if(DadosPesquisador[i][0] == NomePesquisador){

      let IdNumero = IdPesquisador[i];
      let RegistroData = DataRegistro[i];
      let Pesquisador = DadosPesquisador[i][0];
      let Nascimento = DadosPesquisador[i][1];
      let Sexo = DadosPesquisador[i][2];
      let Cpf = DadosPesquisador[i][3];
      let Nacionalidade = DadosPesquisador[i][4];
      let Telefone = DadosPesquisador[i][5];
      let Email = DadosPesquisador[i][6];
      let Endereco = DadosPesquisador[i][7];
      let NumCasa = DadosPesquisador[i][8];
      let Complemento = DadosPesquisador[i][9];
      let Bairro = DadosPesquisador[i][10];
      let Cidade = DadosPesquisador[i][11];
      let Estado = DadosPesquisador[i][12];
      let Formacao = DadosPesquisador[i][13];
      let Instituicao = DadosPesquisador[i][14];
      let Curso = DadosPesquisador[i][15];
      let Profissao = DadosPesquisador[i][16];
      let Assunto = DadosPesquisador[i][17];
      let Finalidade = DadosPesquisador[i][18];
      let PeriodoEstudo = DadosPesquisador[i][19];
      let AreaTrabalho = DadosPesquisador[i][20];
      let Observacoes = DadosPesquisador[i][21];

      DadosPesquisador.length = 0;

      return ([IdNumero, RegistroData, Pesquisador, Nascimento, Sexo, Cpf, Nacionalidade, Telefone, Email, Endereco, NumCasa, Complemento, Bairro, Cidade, Estado, Formacao, Instituicao, Curso, Profissao, Assunto, Finalidade, PeriodoEstudo, AreaTrabalho, Observacoes]);
    }
  };
  // Caso o pesquisador não for encontrado
  DadosPesquisador.length = 0;
  return "Pesquisador não encontrado!"
}


// Função para preencher os campos com os dados do pesquisador selecionado pelo cpf
function PesquisarPesquisadorCpf(CpfPesquisador){
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
// Se adicionar mais campos pra preencher, é necessário alterar o número de colunas
  let ultimaLinha = GuiaPesquisador.getLastRow();
  let DadosPesquisador = GuiaPesquisador.getRange(2, 6, ultimaLinha, 24).getValues();
    // Pega a informação do nome do pesquisador na planilha
  let NomePesquisa = GuiaPesquisador.getRange(2, 3, ultimaLinha, 1).getValues();
  // Pega a informação da data de nascimento na planilha
  let NascimentoPesquisa = GuiaPesquisador.getRange(2, 4, ultimaLinha, 1).getValues();
  // Pega a informação do número do identificador do pesquisador
  let IdPesquisador = GuiaPesquisador.getRange(2, 1, ultimaLinha, 1).getValues();
  // Pega a informação da data de registro do pesquisador
  let DataRegistro = GuiaPesquisador.getRange(2, 2, ultimaLinha, 1).getValues();
  // Pega a informação do sexo
  let SexoPesquisador = GuiaPesquisador.getRange(2, 5, ultimaLinha, 1).getValues();

  for(let i = 0; i < DadosPesquisador.length; i++) {
    if(DadosPesquisador[i][0] == CpfPesquisador){

      let IdNumero = IdPesquisador[i];
      let RegistroData = DataRegistro[i];
      let Pesquisador = NomePesquisa[i];
      let Nascimento = NascimentoPesquisa[i];
      let Sexo = SexoPesquisador[i];
      let Cpf = DadosPesquisador[i][0];
      let Nacionalidade = DadosPesquisador[i][1];
      let Telefone = DadosPesquisador[i][2];
      let Email = DadosPesquisador[i][3];
      let Endereco = DadosPesquisador[i][4];
      let NumCasa = DadosPesquisador[i][5];
      let Complemento = DadosPesquisador[i][6];
      let Bairro = DadosPesquisador[i][7];
      let Cidade = DadosPesquisador[i][8];
      let Estado = DadosPesquisador[i][9];
      let Formacao = DadosPesquisador[i][10];
      let Instituicao = DadosPesquisador[i][11];
      let Curso = DadosPesquisador[i][12];
      let Profissao = DadosPesquisador[i][13];
      let Assunto = DadosPesquisador[i][14];
      let Finalidade = DadosPesquisador[i][15];
      let PeriodoEstudo = DadosPesquisador[i][16];
      let AreaTrabalho = DadosPesquisador[i][17];
      let Observacoes = DadosPesquisador[i][18];

      DadosPesquisador.length = 0;

       return ([IdNumero, RegistroData, Pesquisador, Nascimento, Sexo, Cpf, Nacionalidade, Telefone, Email, Endereco, NumCasa, Complemento, Bairro, Cidade, Estado, Formacao, Instituicao, Curso, Profissao, Assunto, Finalidade, PeriodoEstudo, AreaTrabalho, Observacoes]);
    }
  };
  // Caso o pesquisador não for encontrado
  DadosPesquisador.length = 0;
  return "Pesquisador não encontrado!"
}



  // Função para editar pesquisador

function EditarPesquisador(Dados) {

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if (user.hasLock()) {
    let Planilha = SpreadsheetApp.getActiveSpreadsheet();
    let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");

    let ultimaLinha = GuiaPesquisador.getLastRow();
    let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 1).getValues();
    let DadosPesquisadorCpf = GuiaPesquisador.getRange(2, 6, ultimaLinha, 1).getValues();


    for(var i = 0; i < DadosPesquisador.length; i++){
      if(DadosPesquisador[i][0] == Dados.ListaPesquisador || DadosPesquisador[i][2] == Dados.PesquisaCpf){
        let Linha = i + 2;
        GuiaPesquisador.getRange(Linha, 3).setValue(Dados.NomePesquisador);
        GuiaPesquisador.getRange(Linha, 4).setValue(Dados.Nascimento);
        GuiaPesquisador.getRange(Linha, 5).setValue(Dados.Sexo);
        GuiaPesquisador.getRange(Linha, 6).setValue(Dados.Cpf);
        GuiaPesquisador.getRange(Linha, 7).setValue(Dados.Nacionalidade);
        GuiaPesquisador.getRange(Linha, 8).setValue(Dados.Telefone);
        GuiaPesquisador.getRange(Linha, 9).setValue(Dados.Email);
        GuiaPesquisador.getRange(Linha, 10).setValue(Dados.Endereco);
        GuiaPesquisador.getRange(Linha, 11).setValue(Dados.NumCasa);
        GuiaPesquisador.getRange(Linha, 12).setValue(Dados.Complemento);
        GuiaPesquisador.getRange(Linha, 13).setValue(Dados.Bairro);
        GuiaPesquisador.getRange(Linha, 14).setValue(Dados.Cidade);
        GuiaPesquisador.getRange(Linha, 15).setValue(Dados.Estado);
        GuiaPesquisador.getRange(Linha, 16).setValue(Dados.Formacao);
        GuiaPesquisador.getRange(Linha, 17).setValue(Dados.Instituicao);
        GuiaPesquisador.getRange(Linha, 18).setValue(Dados.Curso);
        GuiaPesquisador.getRange(Linha, 19).setValue(Dados.Profissao);
        GuiaPesquisador.getRange(Linha, 20).setValue(Dados.Assunto);
        GuiaPesquisador.getRange(Linha, 21).setValue(Dados.Finalidade);
        GuiaPesquisador.getRange(Linha, 22).setValue(Dados.PeriodoEstudo);
        GuiaPesquisador.getRange(Linha, 23).setValue(Dados.AreaTrabalho);
        GuiaPesquisador.getRange(Linha, 24).setValue(Dados.Observacoes);

        GuiaPesquisador.getRange("B:B").setNumberFormat("@");
        GuiaPesquisador.getRange("D:D").setNumberFormat("@");

        DadosPesquisador.length = 0;

        return "Pesquisador editado com sucesso!";
      }
    }
    DadosPesquisador.length = 0;

    return "Pesquisador não encontrado!";
  }
}

// Função para o botão excluir
function ExcluirPesquisador(ListaPesquisador){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock){
    let Planilha = SpreadsheetApp.getActiveSpreadsheet();
    let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
    let GuiaRegistroPesquisa = Planilha.getSheetByName("Registro_Pesquisa");

    let ultimaLinha = GuiaPesquisador.getLastRow();

    // Modificar quando o script de registro for criado
    let DadosPesquisador = GuiaPesquisador.getRange(2, 3, ultimaLinha, 1).getValues();
    let ultimaLinhaRegistroPesquisa = GuiaRegistroPesquisa.getLastRow();

    let DadosRegistroPesquisa = GuiaRegistroPesquisa.getRange(2, 2, ultimaLinhaRegistroPesquisa, 1).getValues();

    let Ver = DadosRegistroPesquisa.filter(function(value, i, arr){
      return ListaPesquisador == arr[i][0];
    });
    
    if(Ver.length > 0) {
      DadosPesquisador.length = 0;
      DadosRegistroPesquisa.length = 0;
      Ver.length = 0;
      return "Pesquisador não pode ser excluído, já tem lançamento no registro de pesquisa.";
    }

    for(var i = 0; i < DadosPesquisador.length; i++) {
      if (DadosPesquisador[i][0] == ListaPesquisador) {
        let Linha = i + 2;
        GuiaPesquisador.deleteRow(Linha);
        // fechando os arrays
        DadosPesquisador.length = 0;
        DadosRegistroPesquisa.length = 0;
        Ver.length = 0;
        return "Excluído com sucesso!";
      }
    }
    DadosPesquisador.length = 0;
    DadosRegistroPesquisa.length = 0;
    Ver.length = 0;
    return "Pesquisador não localizado!";
  }
}

// Função para puxar a tela de consulta e edição dos pesquisadores
function ConsultaPesquisador() {
  // Está pegando a tabela dos pesquisadores
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");
  let GuiaEstado = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaFormacao = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaNacionalidade = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaPeriodoEstudo = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaAreaTrabalho = Planilha.getSheetByName("Listas_Suspensas");
  let GuiaSexo = Planilha.getSheetByName("Listas_Suspensas");

  // Captura a última linha com dados
  let ultimaLinhaPesquisador = GuiaPesquisador.getLastRow() - 1;

  // Captura a lista de nome informando o número de início da linha e a coluna, e quantas colunas deseja capturar
  let list = GuiaPesquisador.getRange(2, 3, GuiaPesquisador.getRange("C2").getDataRegion().getLastRow(), 1).getValues();
  let list2 = GuiaNacionalidade.getRange(2, 5, GuiaNacionalidade.getRange("E2").getDataRegion().getLastRow(), 1).getValues();
  let list3 = GuiaEstado.getRange(2, 1, GuiaEstado.getRange("A2").getDataRegion().getLastRow(), 1).getValues();
  let list4 = GuiaFormacao.getRange(2, 3, GuiaFormacao.getRange("C2").getDataRegion().getLastRow(), 1).getValues();
  let list5 = GuiaPeriodoEstudo.getRange(2, 7, GuiaPeriodoEstudo.getRange("G2").getDataRegion().getLastRow(), 1).getValues();
  let list6 = GuiaAreaTrabalho.getRange(2, 9, GuiaAreaTrabalho.getRange("I2").getDataRegion().getLastRow(), 1).getValues();
  let list7 = GuiaSexo.getRange(2, 19, GuiaSexo.getRange("S2").getDataRegion().getLastRow(), 1).getValues();

  // Coloca a lista em ordem alfabética
  list.sort();

  // Abre o template html
  let Form = HtmlService.createTemplateFromFile("ConsultaPesquisadores");

  // Pega cada nome da tabela de pesquisadores e transfere para o formulário no campo de seleção
  Form.list = list.map(function (r) { return r[0]; });
  Form.list2 = list2.map(function (r2) { return r2[0]; });
  Form.list3 = list3.map(function (r3) { return r3[0]; });
  Form.list4 = list4.map(function (r4) { return r4[0]; });
  Form.list5 = list5.map(function (r5) { return r5[0]; });
  Form.list6 = list6.map(function (r6) { return r6[0]; });
  Form.list7 = list7.map(function (r7) { return r7[0]; });


  let MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);

  // Insere um título no formulário
  MostrarForm.setTitle("Cadastro de Pesquisadores").setHeight(1000).setWidth(800);

  // Exibir o formulário
  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Consulta sobre Pesquisadores");

}

// Atualizar a lista dos colaboradores
function AtualizarListaPesquisadores() {

  // Referenciando a planilha
  let Planilha = SpreadsheetApp.getActiveSpreadsheet();
  let GuiaPesquisador = Planilha.getSheetByName("Pesquisadores");

  let ultimaLinha = GuiaPesquisador.getLastRow() - 1;
  let list = GuiaPesquisador.getRange(2, 3, ultimaLinha, 1).getValues();

  return list.sort();
}

// Chama a função na abertura do arquivo html para puxar os demais scripts
function Chamar(Arquivo) {
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}
