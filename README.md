Sistema desenvolvido com scripts no Google Sheets para a criação de cadastro dos pesquisadores, de funcionários, cadastro e localização dos documentos e dos registros de consulta dos materiais utilizados nas pesquisas pelos pesquisadores com objetivo de não utilizar mais as fichas em papel.

Para desenvolver o sistema no Google Sheets, aprendi como programar observando o Canal SGP no youtube. Devido a isso, mantive algumas características do estilo de programação do instrutor no vídeo, como, por exemplo, deixar o nome das variáveis com a primeira letra em maiúsculo e o uso do Framework Materialize, que poderá ser trocado pelo Bootstrap posteriormente.

Cada formulário de cadastro adiciona as informações inseridas em abas diferentes da planilha, no qual os dados integramm-se no formulário de Registro de Consulta.

-------------

CADASTRO DO PESQUISADOR

São inseridas as informações dos pesquisadores que eram utilizadas na ficha em papel, como nome, endereço, e-mail, telefone, nacionalidade, formação, assunto e finalidade da pesquisa, data de nascimento, profissão, escolaridade, campos de assunto selecionáveis, período de estudo (colonial, imperial e republicano). Também é criado um ID contínuo e individual para cada pesquisador que respeita a ordem de registro.

-------------

CADASTRO DO FUNCIONÁRIO

Formulário com objetivo de inserir informações básicas de quem utilizará o sistema, como nome, matrícula, setor e se é agente público efetivo ou terceirizado. Também é criado número de ID individual.

-------------

TOPOGRÁFICO - Registro dos documentos

Formulário de cadastro e pesquisa dos documentos cadastros, sendo informado o tipo de acervo (documental, bibliográfico, cartográfico, iconográfico etc), o tipo de documento (ofício, ata etc), os códigos que contém nas caixas ou códices, o nome completo do documento, periodo da documentação, quantidade de volume, local de armazenamento e observações. É criado ID individual para cada documento inserido na planilha.

-------------

REGISTRO DE CONSULTA DO PESQUISADOR

O registro de consulta é o formulário onde é inserido o nome do pesquisador que realizou a pesquisa, data da pesquisa, se presencial ou online, quem atendeu (agente público), os documentos que foram consultados, o período da documentação consultada e a quantidade de volume. Também há um campo de observação para anotação de informações complementares a depender do documento registrado. Aqui é onde há a integração das informações dos formulários acima, sendo informado o número de ID do pesquisador, do agente público e do documento. Também são criados dois IDs a mais: um sequencial, para cada registro, e outro para cada pesquisa. Como cada pesquisa pode ter mais de um documento consultado, então caso a pesquisa da linha anterior o nome do pesquisador e a data serem os mesmos, o número do ID de pesquisa será o mesmo. Caso alguma destas informações seja diferente, então o ID terá o número seguinte. Desta forma, é possível contar de forma correta quantas pesquisas foram realizadas.

--------------

OBSERVAÇÕES

A implementação dos registros nas planilhas seguem conceito de banco de dados relacional, entretanto, por ser uma planilha e não uma banco de dados, não foi possível implementar as devidas normalizações. No último formulário acima há também a duplicação de informações: é colocado na planilha o ID dos pesquisadores, dos agentes públicos e dos documentos ao mesmo tempo que é informado o nome de cada um. A duplicação não é correta seguindo os conceitos de banco de dados relacional, o ideal seria o uso das chaves secundárias, no entanto como o levantamento dos dados para os relatórios estão sendo feitos por meio de tabelas dinâmmicas, então foi adotado desta forma. Assim que for possível trabalhar apenas com os IDs, as demais colunas serão excluídas.