
/************************************************************
 SISTEMA DE SORTEIO – NATAL E ANO NOVO
 PLANILHA: SORTEIO FIM DE ANO

 REGRAS:
 1. Um apartamento = uma inscrição válida.
 2. A última resposta substitui a anterior.
 3. Cada ambiente possui seu próprio sorteio.
 4. Desistência exclui o ganhador somente daquele ambiente.
 5. Histórico dos sorteios é preservado.
 6. Sorteios protegidos contra execuções simultâneas.
************************************************************/


const CONFIG = {

  NOME_FORMULARIO:
    "Inscrição – Sorteio dos Salões – Natal e Ano Novo",

  ABA_RESPOSTAS:
    "Respostas ao formulário 2",

  AMBIENTES: [
    "Salão Gourmet – Natal (25 pessoas)",
    "Salão de Festas – Natal (50 pessoas)",
    "Salão Gourmet – Ano Novo (25 pessoas)",
    "Salão de Festas – Ano Novo (50 pessoas)"
  ],

  ABAS: [
    "Gourmet – Natal",
    "Festas – Natal",
    "Gourmet – Ano Novo",
    "Festas – Ano Novo",
    "Resultados dos Sorteios",
    "Histórico dos Sorteios"
  ]

};


const PROPRIEDADES =
  PropertiesService.getScriptProperties();


/************************************************************
 1. MENU PRINCIPAL
************************************************************/

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("SORTEIO DOS SALÕES")

    .addItem(
      "1 – Criar formulário",
      "criarFormulario"
    )

    .addItem(
      "2 – Atualizar listas",
      "atualizarListas"
    )

    .addSeparator()

    .addItem(
      "Sortear todos os ambientes",
      "sortearTodos"
    )

    .addItem(
      "Sortear Gourmet – Natal",
      "sortearGourmetNatal"
    )

    .addItem(
      "Sortear Festas – Natal",
      "sortearFestasNatal"
    )

    .addItem(
      "Sortear Gourmet – Ano Novo",
      "sortearGourmetAnoNovo"
    )

    .addItem(
      "Sortear Festas – Ano Novo",
      "sortearFestasAnoNovo"
    )

    .addSeparator()

    .addItem(
      "Registrar desistência e novo sorteio",
      "registrarDesistencia"
    )

    .addToUi();

}


/************************************************************
 2. CRIAR AS ABAS DO SISTEMA
************************************************************/

function criarAbas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  CONFIG.ABAS.forEach(nome => {

    let aba = ss.getSheetByName(nome);

    if (!aba) {

      aba = ss.insertSheet(nome);

    }

    if (aba.getLastRow() === 0) {

      aba.appendRow([
        "Nome completo",
        "Apartamento",
        "Data da inscrição"
      ]);

    }

  });


  const resultados = ss.getSheetByName(
    "Resultados dos Sorteios"
  );

  if (resultados.getLastRow() === 0) {

    resultados.appendRow([
      "Ambiente",
      "Nome do ganhador",
      "Apartamento",
      "Data do sorteio",
      "Situação"
    ]);

  }


  const historico = ss.getSheetByName(
    "Histórico dos Sorteios"
  );

  if (historico.getLastRow() === 0) {

    historico.appendRow([
      "Data e hora",
      "Ambiente",
      "Tipo de registro",
      "Nome completo",
      "Apartamento",
      "Situação",
      "Observação"
    ]);

  }

}


/************************************************************
 3. CRIAR FORMULÁRIO
************************************************************/

function criarFormulario() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();

  let form = null;

  const idExistente =
    PROPRIEDADES.getProperty("ID_FORMULARIO");


  if (idExistente) {

    try {

      form = FormApp.openById(idExistente);

    } catch (erro) {

      form = null;

    }

  }


  if (!form) {

    form = FormApp.create(
      CONFIG.NOME_FORMULARIO
    );

    form.setDescription(

      "INSCRIÇÃO PARA OS SORTEIOS DE NATAL E ANO NOVO.\n\n" +

      "Você poderá participar de um, dois, três ou dos quatro sorteios.\n\n" +

      "ATENÇÃO: Cada apartamento poderá ter apenas uma inscrição válida. " +

      "Caso preencha novamente, será considerada somente a última resposta enviada, " +

      "que substituirá integralmente a anterior."

    );


    form.addTextItem()

      .setTitle("Nome completo")

      .setRequired(true);


    form.addTextItem()

      .setTitle("Apartamento")

      .setRequired(true);


    form.addCheckboxItem()

      .setTitle(
        "Em quais sorteios deseja participar?"
      )

      .setHelpText(
        "Você pode participar de um, dois, três ou dos quatro sorteios."
      )

      .setChoiceValues(CONFIG.AMBIENTES)

      .setRequired(true);


    form.setConfirmationMessage(
      "Sua inscrição foi recebida com sucesso!"
    );


    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      ss.getId()
    );


    PROPRIEDADES.setProperty(
      "ID_FORMULARIO",
      form.getId()
    );

  }


  // Criar gatilho somente se ainda não existir.

  const gatilhos = ScriptApp.getProjectTriggers();

  const existeGatilho = gatilhos.some(

    t =>
      t.getHandlerFunction() ===
      "aoEnviarFormulario"

  );


  if (!existeGatilho) {

    ScriptApp.newTrigger("aoEnviarFormulario")

      .forSpreadsheet(ss)

      .onFormSubmit()

      .create();

  }


  SpreadsheetApp.getUi().alert(

    "FORMULÁRIO PRONTO!\n\n" +

    "LINK PARA OS MORADORES:\n\n" +

    form.getPublishedUrl() +

    "\n\nLINK PARA EDITAR:\n\n" +

    form.getEditUrl()

  );

}


/************************************************************
 4. RECEBER RESPOSTAS DO FORMULÁRIO
************************************************************/

function aoEnviarFormulario(e) {

  const aba = e.range.getSheet();

  const linha = e.range.getRow();

  const ultimaColuna = aba.getLastColumn();


  if (ultimaColuna >= 2) {

    const intervalo = aba.getRange(

      linha,

      2,

      1,

      ultimaColuna - 1

    );


    const valores = intervalo.getValues();


    const valoresMaiusculos = valores.map(linha =>

      linha.map(valor =>

        typeof valor === "string"

          ? valor.toLocaleUpperCase("pt-BR")

          : valor

      )

    );


    intervalo.setValues(valoresMaiusculos);

  }


  atualizarListas_(false);

}


/************************************************************
 5. NORMALIZAR TEXTO
************************************************************/

function normalizar_(valor) {

  return String(valor || "")

    .normalize("NFD")

    .replace(/[\u0300-\u036f]/g, "")

    .toLowerCase()

    .trim();

}


/************************************************************
 6. LOCALIZAR A ABA OFICIAL DE RESPOSTAS
************************************************************/

function localizarAbaRespostas_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const aba = ss.getSheetByName(
    CONFIG.ABA_RESPOSTAS
  );


  if (!aba) {

    return null;

  }


  if (aba.getLastRow() < 1) {

    return null;

  }


  const cabecalhos = aba.getRange(

    1,

    1,

    1,

    aba.getLastColumn()

  ).getDisplayValues()[0].map(normalizar_);


  const nome = cabecalhos.some(

    v => v.includes("nome completo")

  );


  const apartamento = cabecalhos.some(

    v => v.includes("apartamento")

  );


  const escolhas = cabecalhos.some(

    v => v.includes("em quais sorteios")

  );


  return nome && apartamento && escolhas

    ? aba

    : null;

}


/************************************************************
 7. ATUALIZAR LISTAS
************************************************************/

function atualizarListas() {

  const resultado = atualizarListas_(true);


  if (resultado) {

    SpreadsheetApp.getUi().alert(

      "As quatro listas foram atualizadas corretamente!\n\n" +

      "Foi considerada somente a última resposta de cada apartamento."

    );

  }

}


/************************************************************
 8. PROCESSAR AS RESPOSTAS
************************************************************/

function atualizarListas_(mostrarErro) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();

  const origem = localizarAbaRespostas_();


  if (!origem) {

    if (mostrarErro) {

      SpreadsheetApp.getUi().alert(

        "Não encontrei a aba oficial de respostas.\n\n" +

        "Verifique se existe a aba:\n" +

        CONFIG.ABA_RESPOSTAS

      );

    }

    return false;

  }


  const dados = origem.getDataRange().getValues();

  const exibidos = origem.getDataRange().getDisplayValues();


  if (dados.length < 2) {

    return false;

  }


  const cabecalhos =
    exibidos[0].map(normalizar_);


  const colData = cabecalhos.findIndex(

    v =>
      v.includes("carimbo") ||
      v.includes("timestamp")

  );


  const colNome = cabecalhos.findIndex(

    v => v.includes("nome completo")

  );


  const colApartamento = cabecalhos.findIndex(

    v => v.includes("apartamento")

  );


  const colEscolhas = cabecalhos.findIndex(

    v => v.includes("em quais sorteios")

  );


  if (

    colData < 0 ||
    colNome < 0 ||
    colApartamento < 0 ||
    colEscolhas < 0

  ) {

    if (mostrarErro) {

      SpreadsheetApp.getUi().alert(

        "Não foi possível identificar todas as colunas do formulário."

      );

    }

    return false;

  }


  /********************************************************
   GUARDA A ÚLTIMA RESPOSTA DE CADA APARTAMENTO
  ********************************************************/

  const respostasMaisRecentes = new Map();


  for (let i = 1; i < dados.length; i++) {

    const nome = String(

      exibidos[i][colNome] || ""

    ).trim();


    const apartamento = String(

      exibidos[i][colApartamento] || ""

    ).trim();


    const escolhas = String(

      exibidos[i][colEscolhas] || ""

    ).trim();


    const data = dados[i][colData];


    if (

      !nome ||
      !apartamento ||
      !escolhas ||
      !(data instanceof Date) ||
      isNaN(data.getTime())

    ) {

      continue;

    }


    const chave = normalizar_(apartamento);

    const anterior =
      respostasMaisRecentes.get(chave);


    // Se houver outra resposta mais recente,
    // ela substitui integralmente a anterior.

    if (

      !anterior ||
      data.getTime() >= anterior.data.getTime()

    ) {

      respostasMaisRecentes.set(chave, {

        nome: nome,

        apartamento: apartamento,

        escolhas: escolhas,

        data: data

      });

    }

  }


  /********************************************************
   PRAZO DE INSCRIÇÃO: ATÉ 10 DE DEZEMBRO
  ********************************************************/

  const anoAtual = new Date().getFullYear();

  const limite = new Date(

    anoAtual,

    11,

    10,

    23,

    59,

    59

  );


  const listas = [

    [],
    [],
    [],
    []

  ];


  /********************************************************
   DISTRIBUI SOMENTE A ÚLTIMA RESPOSTA
  ********************************************************/

  respostasMaisRecentes.forEach(resposta => {

    if (

      resposta.data.getFullYear() !== anoAtual ||
      resposta.data > limite

    ) {

      return;

    }


    CONFIG.AMBIENTES.forEach(

      (ambiente, indice) => {

        if (

          resposta.escolhas.includes(ambiente)

        ) {

          listas[indice].push([

            resposta.nome,

            resposta.apartamento,

            resposta.data

          ]);

        }

      }

    );

  });


  /********************************************************
   ESCREVER AS QUATRO LISTAS
  ********************************************************/

  const nomesAbas = [

    "Gourmet – Natal",

    "Festas – Natal",

    "Gourmet – Ano Novo",

    "Festas – Ano Novo"

  ];


  nomesAbas.forEach((nomeAba, indice) => {

    const aba = ss.getSheetByName(nomeAba);

    aba.clearContents();


    aba.getRange(1, 1, 1, 3).setValues([[

      "Nome completo",

      "Apartamento",

      "Data da inscrição"

    ]]);


    if (listas[indice].length > 0) {

      aba.getRange(

        2,

        1,

        listas[indice].length,

        3

      ).setValues(listas[indice]);

    }


    aba.setFrozenRows(1);

    aba.autoResizeColumns(1, 3);

  });


  return true;

}


/************************************************************
 9. SORTEIOS INDIVIDUAIS
************************************************************/

function sortearGourmetNatal() {

  const resultado = sortearAmbiente_(
    "Gourmet – Natal"
  );

  SpreadsheetApp.getUi().alert(resultado);

}


function sortearFestasNatal() {

  const resultado = sortearAmbiente_(
    "Festas – Natal"
  );

  SpreadsheetApp.getUi().alert(resultado);

}


function sortearGourmetAnoNovo() {

  const resultado = sortearAmbiente_(
    "Gourmet – Ano Novo"
  );

  SpreadsheetApp.getUi().alert(resultado);

}


function sortearFestasAnoNovo() {

  const resultado = sortearAmbiente_(
    "Festas – Ano Novo"
  );

  SpreadsheetApp.getUi().alert(resultado);

}


/************************************************************
 10. SORTEAR TODOS OS AMBIENTES
************************************************************/

function sortearTodos() {

  if (!atualizarListas_(false)) {

    SpreadsheetApp.getUi().alert(

      "Não foi possível atualizar as listas."

    );

    return;

  }


  const ambientes = [

    "Gourmet – Natal",

    "Festas – Natal",

    "Gourmet – Ano Novo",

    "Festas – Ano Novo"

  ];


  const mensagens = [];


  ambientes.forEach(ambiente => {

    mensagens.push(

      sortearAmbiente_(ambiente)

    );

  });


  atualizarResultados_();


  SpreadsheetApp.getUi().alert(

    "RESULTADOS DOS SORTEIOS\n\n" +

    mensagens.join("\n\n")

  );

}


/************************************************************
 11. EXECUTAR SORTEIO DE UM AMBIENTE
 COM PROTEÇÃO CONTRA EXECUÇÕES SIMULTÂNEAS
************************************************************/

function sortearAmbiente_(ambiente) {

  const bloqueio = LockService.getScriptLock();


  // Aguarda até 10 segundos para obter o bloqueio.

  if (!bloqueio.tryLock(10000)) {

    return "Não foi possível realizar o sorteio de " +

      ambiente +

      ". Outra operação está em andamento. Tente novamente.";

  }


  try {

    return sortearAmbienteInterno_(ambiente);

  } catch (erro) {

    console.error(erro);

    return "Erro ao realizar o sorteio de " +

      ambiente + ": " + erro.message;

  } finally {

    // Libera o bloqueio mesmo se ocorrer algum erro.

    if (bloqueio.hasLock()) {

      bloqueio.releaseLock();

    }

  }

}


/************************************************************
 11.1. EXECUÇÃO INTERNA DO SORTEIO
 NÃO ADQUIRE OUTRO BLOQUEIO
************************************************************/

function sortearAmbienteInterno_(ambiente) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();


  // Atualiza as listas antes de cada sorteio.

  if (!atualizarListas_(false)) {

    return "Não foi possível atualizar as listas.";

  }


  const abaLista = ss.getSheetByName(ambiente);

  const historico = ss.getSheetByName(

    "Histórico dos Sorteios"

  );


  const dadosHistorico =
    historico.getDataRange().getValues();


  // Impede sorteio duplicado enquanto houver
  // um ganhador ativo naquele ambiente.

  const ganhadorAtivo = dadosHistorico

    .slice(1)

    .some(linha =>

      linha[1] === ambiente &&
      linha[5] === "Ganhador"

    );


  if (ganhadorAtivo) {

    return ambiente +

      ": já possui um ganhador ativo.";

  }


  const dadosLista =
    abaLista.getDataRange().getValues();


  if (dadosLista.length < 2) {

    return ambiente +

      ": não há inscritos.";

  }


  // Recupera os apartamentos que desistiram
  // somente deste ambiente.

  const desistentes = new Set();


  dadosHistorico.slice(1).forEach(linha => {

    if (

      linha[1] === ambiente &&
      linha[5] === "Desistiu"

    ) {

      desistentes.add(

        normalizar_(linha[4])

      );

    }

  });


  const candidatos = dadosLista

    .slice(1)

    .filter(linha => {

      const nome = String(

        linha[0] || ""

      ).trim();


      const apartamento = String(

        linha[1] || ""

      ).trim();


      return (

        nome &&
        apartamento &&

        !desistentes.has(

          normalizar_(apartamento)

        )

      );

    });


  if (candidatos.length === 0) {

    return ambiente +

      ": não há candidatos disponíveis.";

  }


  // Sorteio aleatório.

  const indice = Math.floor(

    Math.random() * candidatos.length

  );


  const ganhador = candidatos[indice];


  // Registrar no histórico.

  historico.appendRow([

    new Date(),

    ambiente,

    "Sorteio",

    ganhador[0],

    ganhador[1],

    "Ganhador",

    ""

  ]);


  atualizarResultados_();


  return (

    ambiente +

    "\nGanhador: " + ganhador[0] +

    "\nApartamento: " + ganhador[1]

  );

}


/************************************************************
 12. REGISTRAR DESISTÊNCIA
 COM PROTEÇÃO CONTRA EXECUÇÕES SIMULTÂNEAS
************************************************************/

function registrarDesistencia() {

  const ui = SpreadsheetApp.getUi();


  const ambientes = [

    "Gourmet – Natal",

    "Festas – Natal",

    "Gourmet – Ano Novo",

    "Festas – Ano Novo"

  ];


  const respostaAmbiente = ui.prompt(

    "REGISTRAR DESISTÊNCIA",

    "Digite exatamente o nome do ambiente:\n\n" +

    ambientes.join("\n"),

    ui.ButtonSet.OK_CANCEL

  );


  if (

    respostaAmbiente.getSelectedButton() !==

    ui.Button.OK

  ) {

    return;

  }


  const ambiente = ambientes.find(

    nome =>

      normalizar_(nome) ===

      normalizar_(

        respostaAmbiente.getResponseText()

      )

  );


  if (!ambiente) {

    ui.alert("Nome do ambiente inválido.");

    return;

  }


  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const historico = ss.getSheetByName(

    "Histórico dos Sorteios"

  );


  const dados = historico.getDataRange().getValues();


  let linhaGanhador = -1;


  for (

    let i = dados.length - 1;

    i >= 1;

    i--

  ) {

    if (

      dados[i][1] === ambiente &&

      dados[i][5] === "Ganhador"

    ) {

      linhaGanhador = i + 1;

      break;

    }

  }


  if (linhaGanhador < 0) {

    ui.alert(

      "Não existe ganhador ativo nesse ambiente."

    );

    return;

  }


  const ganhador = dados[linhaGanhador - 1];


  const confirmacao = ui.alert(

    "CONFIRMAR DESISTÊNCIA",

    "Ganhador: " + ganhador[3] +

    "\nApartamento: " + ganhador[4] +

    "\n\nO apartamento será excluído somente deste ambiente." +

    "\n\nDeseja confirmar e realizar novo sorteio?",

    ui.ButtonSet.YES_NO

  );


  if (confirmacao !== ui.Button.YES) {

    return;

  }


  // Bloqueia outras operações durante a desistência
  // e o novo sorteio.

  const bloqueio = LockService.getScriptLock();


  if (!bloqueio.tryLock(10000)) {

    ui.alert(

      "Outra operação está em andamento. " +

      "Tente registrar a desistência novamente."

    );

    return;

  }


  try {

    // Confere novamente o ganhador após obter o bloqueio.

    const dadosAtualizados =
      historico.getDataRange().getValues();


    let linhaAtual = -1;


    for (

      let i = dadosAtualizados.length - 1;

      i >= 1;

      i--

    ) {

      if (

        dadosAtualizados[i][1] === ambiente &&

        dadosAtualizados[i][5] === "Ganhador"

      ) {

        linhaAtual = i + 1;

        break;

      }

    }


    if (linhaAtual < 0) {

      ui.alert(

        "O ganhador ativo foi alterado por outra operação. " +

        "Verifique o histórico antes de continuar."

      );

      return;

    }


    const ganhadorAtual =
      dadosAtualizados[linhaAtual - 1];


    // Marcar o antigo ganhador como desistente.

    historico.getRange(

      linhaAtual,

      6

    ).setValue("Desistiu");


    historico.getRange(

      linhaAtual,

      7

    ).setValue(

      "Desistência registrada em " +

      Utilities.formatDate(

        new Date(),

        Session.getScriptTimeZone(),

        "dd/MM/yyyy HH:mm"

      )

    );


    // Realizar novo sorteio sem adquirir
    // novamente o bloqueio já obtido.

    const resultado = sortearAmbienteInterno_(

      ambiente

    );


    atualizarResultados_();


    ui.alert(

      "NOVO SORTEIO REALIZADO\n\n" +

      "Desistente: " + ganhadorAtual[3] +

      "\nApartamento: " + ganhadorAtual[4] +

      "\n\n" + resultado

    );

  } catch (erro) {

    console.error(erro);

    ui.alert(

      "Ocorreu um erro ao registrar a desistência:\n\n" +

      erro.message

    );

  } finally {

    // Libera o bloqueio mesmo se ocorrer algum erro.

    if (bloqueio.hasLock()) {

      bloqueio.releaseLock();

    }

  }

}


/************************************************************
 13. ATUALIZAR RESULTADOS
************************************************************/

function atualizarResultados_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();


  const aba = ss.getSheetByName(

    "Resultados dos Sorteios"

  );


  const historico = ss.getSheetByName(

    "Histórico dos Sorteios"

  );


  const dados = historico.getDataRange().getValues();


  aba.clearContents();


  aba.appendRow([

    "Ambiente",

    "Nome do ganhador",

    "Apartamento",

    "Data do sorteio",

    "Situação"

  ]);


  const ambientes = [

    "Gourmet – Natal",

    "Festas – Natal",

    "Gourmet – Ano Novo",

    "Festas – Ano Novo"

  ];


  ambientes.forEach(ambiente => {

    let ganhador = null;


    for (

      let i = dados.length - 1;

      i >= 1;

      i--

    ) {

      if (

        dados[i][1] === ambiente &&

        dados[i][5] === "Ganhador"

      ) {

        ganhador = dados[i];

        break;

      }

    }


    if (ganhador) {

      aba.appendRow([

        ambiente,

        ganhador[3],

        ganhador[4],

        ganhador[0],

        "Ganhador"

      ]);

    } else {

      aba.appendRow([

        ambiente,

        "Aguardando sorteio",

        "",

        "",

        "Pendente"

      ]);

    }

  });


  aba.setFrozenRows(1);

  aba.autoResizeColumns(1, 5);

}


/************************************************************
 14. APLICATIVO ADMINISTRATIVO – INTEGRAÇÃO
************************************************************/

const EMAILS_AUTORIZADOS = [
  "rseneda@gmail.com"
];

const AMBIENTES_PAINEL = [
  "Gourmet – Natal",
  "Festas – Natal",
  "Gourmet – Ano Novo",
  "Festas – Ano Novo"
];


/************************************************************
 VERIFICAR ACESSO DO ADMINISTRADOR
************************************************************/

function verificarAcessoPainel_() {

  const email = Session.getActiveUser()
    .getEmail()
    .toLowerCase()
    .trim();

  if (
    !email ||
    !EMAILS_AUTORIZADOS.includes(email)
  ) {
    throw new Error(
      "Acesso não autorizado. Entre com uma conta Google autorizada."
    );
  }

  return email;
}


/************************************************************
 15. ABRIR O APLICATIVO
************************************************************/

function doGet() {

  try {

    verificarAcessoPainel_();

    return HtmlService
      .createHtmlOutputFromFile("Index")
      .setTitle("Painel Administrativo – Sorteio dos Salões")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1"
      );

  } catch (erro) {

    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;text-align:center;padding:60px">' +
      '<h2>Acesso restrito</h2>' +
      '<p>Esta área é exclusiva para administradores autorizados.</p>' +
      '<p>Entre com a conta Google cadastrada para acessar o sistema.</p>' +
      '</div>'
    ).setTitle("Acesso restrito");

  }

}


/************************************************************
 16. LOCALIZAR GANHADOR ATIVO
************************************************************/

function localizarGanhadorPainel_(ambiente) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const historico = ss.getSheetByName(
    "Histórico dos Sorteios"
  );

  if (!historico) return null;

  const dados = historico.getDataRange().getValues();

  for (let i = dados.length - 1; i >= 1; i--) {

    if (
      dados[i][1] === ambiente &&
      dados[i][5] === "Ganhador"
    ) {

      return {
        nome: String(dados[i][3]),
        apartamento: String(dados[i][4]),
        data: dados[i][0]
      };

    }

  }

  return null;

}


/************************************************************
 17. DADOS GERAIS DO PAINEL
************************************************************/

function obterDadosPainel() {

  verificarAcessoPainel_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();

  atualizarListas_(false);

  const nomesAbas = [
    "Gourmet – Natal",
    "Festas – Natal",
    "Gourmet – Ano Novo",
    "Festas – Ano Novo"
  ];

  const quantidades = [];

  const ganhadores = [];

  nomesAbas.forEach(ambiente => {

    const aba = ss.getSheetByName(ambiente);

    quantidades.push(
      Math.max(0, aba.getLastRow() - 1)
    );

    ganhadores.push(
      localizarGanhadorPainel_(ambiente)
    );

  });

  return {
    quantidades: quantidades,
    ganhadores: ganhadores
  };

}


/************************************************************
 18. EXECUTAR SORTEIO PELO PAINEL
************************************************************/

function executarSorteioPainel(indice) {

  verificarAcessoPainel_();

  indice = Number(indice);

  if (
    !Number.isInteger(indice) ||
    indice < 0 ||
    indice >= AMBIENTES_PAINEL.length
  ) {
    throw new Error("Ambiente inválido.");
  }

  const ambiente = AMBIENTES_PAINEL[indice];

  const resultado = sortearAmbiente_(ambiente);

  atualizarResultados_();

  return resultado;

}


/************************************************************
 19. CONSULTAR RESULTADOS
************************************************************/

function obterResultadosPainel() {

  verificarAcessoPainel_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();

  atualizarResultados_();

  const aba = ss.getSheetByName(
    "Resultados dos Sorteios"
  );

  const dados = aba.getDataRange().getValues();

  return dados.slice(1).map(linha => ({

    ambiente: String(linha[0] || ""),

    nome: String(linha[1] || ""),

    apartamento: String(linha[2] || ""),

    data: linha[3] instanceof Date
      ? Utilities.formatDate(
          linha[3],
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm"
        )
      : "",

    situacao: String(linha[4] || "")

  }));

}


/************************************************************
 20. CONSULTAR HISTÓRICO
************************************************************/

function obterHistoricoPainel() {

  verificarAcessoPainel_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbas();

  const aba = ss.getSheetByName(
    "Histórico dos Sorteios"
  );

  const dados = aba.getDataRange().getValues();

  return dados.slice(1).map(linha => ({

    data: linha[0] instanceof Date
      ? Utilities.formatDate(
          linha[0],
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm:ss"
        )
      : "",

    ambiente: String(linha[1] || ""),

    tipo: String(linha[2] || ""),

    nome: String(linha[3] || ""),

    apartamento: String(linha[4] || ""),

    situacao: String(linha[5] || ""),

    observacao: String(linha[6] || "")

  }));

}


/************************************************************
 21. CONSULTAR GANHADOR PARA DESISTÊNCIA
************************************************************/

function obterGanhadorPainel(indice) {

  verificarAcessoPainel_();

  indice = Number(indice);

  if (
    !Number.isInteger(indice) ||
    indice < 0 ||
    indice >= AMBIENTES_PAINEL.length
  ) {
    throw new Error("Ambiente inválido.");
  }

  return localizarGanhadorPainel_(
    AMBIENTES_PAINEL[indice]
  );

}


/************************************************************
 22. REGISTRAR DESISTÊNCIA PELO PAINEL
************************************************************/

function registrarDesistenciaPainel(indice) {

  verificarAcessoPainel_();

  indice = Number(indice);

  if (
    !Number.isInteger(indice) ||
    indice < 0 ||
    indice >= AMBIENTES_PAINEL.length
  ) {
    throw new Error("Ambiente inválido.");
  }

  const ambiente = AMBIENTES_PAINEL[indice];

  const bloqueio = LockService.getScriptLock();

  if (!bloqueio.tryLock(10000)) {

    throw new Error(
      "Outra operação está em andamento. Tente novamente."
    );

  }

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const historico = ss.getSheetByName(
      "Histórico dos Sorteios"
    );

    const dados = historico.getDataRange().getValues();

    let linhaGanhador = -1;

    for (let i = dados.length - 1; i >= 1; i--) {

      if (
        dados[i][1] === ambiente &&
        dados[i][5] === "Ganhador"
      ) {

        linhaGanhador = i + 1;

        break;

      }

    }

    if (linhaGanhador < 0) {

      throw new Error(
        "Não existe ganhador ativo neste ambiente."
      );

    }

    const ganhador = dados[linhaGanhador - 1];

    // Registrar a desistência no histórico.

    historico.getRange(linhaGanhador, 6)
      .setValue("Desistiu");

    historico.getRange(linhaGanhador, 7)
      .setValue(
        "Desistência registrada em " +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm"
        )
      );

    // Realizar novo sorteio sem adquirir outro bloqueio.

    const resultado = sortearAmbienteInterno_(ambiente);

    atualizarResultados_();

    return (
      "DESISTÊNCIA REGISTRADA\n\n" +
      "Ganhador anterior: " + ganhador[3] +
      "\nApartamento: " + ganhador[4] +
      "\n\n" + resultado
    );

  } finally {

    if (bloqueio.hasLock()) {
      bloqueio.releaseLock();
    }

  }

}

