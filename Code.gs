
/************************************************************
 SISTEMA DE SORTEIO – NATAL E ANO NOVO
 PLANILHA: SORTEIO FIM DE ANO
 ************************************************************/

const CONFIG = {
  NOME_FORMULARIO: "SORTEIO FIM DE ANO",
  ABA_RESPOSTAS: "Respostas ao formulário 2",

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
  ],

  PRAZO_INSCRICAO: {
    MES: 12,
    DIA: 10
  }
};

const EMAILS_AUTORIZADOS = ["rseneda@gmail.com"];
const AMBIENTES_PAINEL = CONFIG.AMBIENTES;
const PROPRIEDADES = PropertiesService.getScriptProperties();


/************************************************************
 MENU
 ************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("SORTEIO DOS SALÕES")
    .addItem("Criar / Corrigir abas", "criarAbas")
    .addItem("Atualizar listas", "atualizarListas")
    .addSeparator()
    .addItem("Sortear Gourmet – Natal", "sortearGourmetNatal")
    .addItem("Sortear Festas – Natal", "sortearFestasNatal")
    .addItem("Sortear Gourmet – Ano Novo", "sortearGourmetAnoNovo")
    .addItem("Sortear Festas – Ano Novo", "sortearFestasAnoNovo")
    .addSeparator()
    .addItem("Registrar desistência – Gourmet Natal", "registrarDesistenciaGourmetNatal")
    .addItem("Registrar desistência – Festas Natal", "registrarDesistenciaFestasNatal")
    .addItem("Registrar desistência – Gourmet Ano Novo", "registrarDesistenciaGourmetAnoNovo")
    .addItem("Registrar desistência – Festas Ano Novo", "registrarDesistenciaFestasAnoNovo")
    .addSeparator()
    .addItem("Atualizar resultados", "atualizarResultados")
    .addToUi();
}


/************************************************************
 NORMALIZAÇÃO E VALIDAÇÃO
 ************************************************************/

function normalizar_(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function normalizarComparacao_(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function validarIndice_(indice) {
  const numero = Number(indice);

  if (!Number.isInteger(numero) || numero < 0 || numero >= 4) {
    throw new Error("Ambiente inválido. Verifique o índice do sorteio.");
  }

  return numero;
}


/************************************************************
 CRIAÇÃO DAS ABAS
 ************************************************************/

function criarAbas() {
  criarAbasSemMensagem_();
  SpreadsheetApp.getUi().alert("Abas criadas e corrigidas com sucesso.");
}

function criarAbasSemMensagem_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cabecalhoAmbiente = [
    "Nome completo",
    "Apartamento",
    "Data da inscrição"
  ];

  for (let i = 0; i < 4; i++) {
    let aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (!aba) aba = ss.insertSheet(CONFIG.ABAS[i]);

    aba.getRange(1, 1, 1, 3).setValues([cabecalhoAmbiente]);
    aba.setFrozenRows(1);
  }

  let resultados = ss.getSheetByName("Resultados dos Sorteios");

  if (!resultados) resultados = ss.insertSheet("Resultados dos Sorteios");

  resultados.getRange(1, 1, 1, 5).setValues([[
    "Ambiente",
    "Ganhador",
    "Apartamento",
    "Data do sorteio",
    "Status"
  ]]);

  resultados.setFrozenRows(1);

  let historico = ss.getSheetByName("Histórico dos Sorteios");

  if (!historico) historico = ss.insertSheet("Histórico dos Sorteios");

  historico.getRange(1, 1, 1, 5).setValues([[
    "Data/Hora",
    "Ambiente",
    "Nome",
    "Apartamento",
    "Status"
  ]]);

  historico.setFrozenRows(1);
}


/************************************************************
 LOCALIZAÇÃO DA ABA DE RESPOSTAS
 ************************************************************/

function localizarAbaRespostas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let aba = ss.getSheetByName(CONFIG.ABA_RESPOSTAS);

  if (aba) return aba;

  for (const sheet of ss.getSheets()) {
    if (CONFIG.ABAS.includes(sheet.getName())) continue;
    if (sheet.getLastColumn() === 0) continue;

    const cabecalhos = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(normalizar_);

    const temNome = cabecalhos.some(x => x.includes("nome completo"));
    const temApartamento = cabecalhos.some(x => x.includes("apartamento"));
    const temSorteios = cabecalhos.some(x => x.includes("sorteios"));

    if (temNome && temApartamento && temSorteios) return sheet;
  }

  throw new Error("Não foi encontrada a aba de respostas do formulário.");
}


/************************************************************
 IDENTIFICAÇÃO DAS COLUNAS
 ************************************************************/

function localizarColunasRespostas_(aba) {
  const cabecalhos = aba
    .getRange(1, 1, 1, aba.getLastColumn())
    .getValues()[0];

  let timestamp = -1;
  let nome = -1;
  let apartamento = -1;
  let sorteios = -1;

  cabecalhos.forEach((valor, i) => {
    const c = normalizar_(valor);

    if (timestamp === -1 &&
        (c.includes("carimbo") || c.includes("timestamp") || c.includes("data"))) {
      timestamp = i;
    }

    if (nome === -1 && c.includes("nome completo")) nome = i;
    if (apartamento === -1 && c.includes("apartamento")) apartamento = i;
    if (sorteios === -1 && c.includes("sorteios")) sorteios = i;
  });

  if (timestamp === -1) timestamp = 0;
  if (nome === -1) throw new Error("Coluna Nome completo não encontrada.");
  if (apartamento === -1) throw new Error("Coluna Apartamento não encontrada.");
  if (sorteios === -1) throw new Error("Coluna de escolha dos sorteios não encontrada.");

  return { timestamp, nome, apartamento, sorteios };
}

function obterDataResposta_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return valor;

  const data = new Date(valor);

  return isNaN(data.getTime()) ? null : data;
}


/************************************************************
 ATUALIZAÇÃO DAS LISTAS DE PARTICIPANTES
 ************************************************************/

function atualizarListas() {
  atualizarListas_(true);
}

function atualizarListas_(mostrarMensagem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbasSemMensagem_();

  const respostas = localizarAbaRespostas_();
  const colunas = localizarColunasRespostas_(respostas);

  const ultimaLinha = respostas.getLastRow();
  const ultimaColuna = respostas.getLastColumn();

  for (let i = 0; i < 4; i++) {
    const aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (aba.getLastRow() > 1) {
      aba.getRange(2, 1, aba.getLastRow() - 1, 3).clearContent();
    }
  }

  if (ultimaLinha < 2) {
    if (mostrarMensagem) {
      SpreadsheetApp.getUi().alert("Ainda não existem inscrições.");
    }
    return;
  }

  const dados = respostas
    .getRange(2, 1, ultimaLinha - 1, ultimaColuna)
    .getValues();

  const ultimasRespostas = {};

  dados.forEach(linha => {
    const apartamento = String(linha[colunas.apartamento] || "").trim();
    const chave = normalizarComparacao_(apartamento);

    if (!chave) return;

    const data = obterDataResposta_(linha[colunas.timestamp]);

    if (!data) return;

    const resposta = {
      data: data,
      nome: String(linha[colunas.nome] || "").trim(),
      apartamento: apartamento,
      escolhas: String(linha[colunas.sorteios] || "").trim()
    };

    if (
      !ultimasRespostas[chave] ||
      data.getTime() >= ultimasRespostas[chave].data.getTime()
    ) {
      ultimasRespostas[chave] = resposta;
    }
  });

  const anoAtual = new Date().getFullYear();

  const prazo = new Date(
    anoAtual,
    CONFIG.PRAZO_INSCRICAO.MES - 1,
    CONFIG.PRAZO_INSCRICAO.DIA,
    23, 59, 59
  );

  const listas = [[], [], [], []];

  Object.values(ultimasRespostas).forEach(resposta => {
    if (resposta.data.getFullYear() !== anoAtual) return;
    if (resposta.data.getTime() > prazo.getTime()) return;
    if (!resposta.nome || !resposta.apartamento) return;

    const escolhas = normalizarComparacao_(resposta.escolhas);

    CONFIG.AMBIENTES.forEach((ambiente, indice) => {
      const ambienteNormalizado = normalizarComparacao_(ambiente);

      if (escolhas.includes(ambienteNormalizado)) {
        listas[indice].push([
          resposta.nome,
          resposta.apartamento,
          resposta.data
        ]);
      }
    });
  });

  for (let i = 0; i < 4; i++) {
    const aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (listas[i].length > 0) {
      aba.getRange(2, 1, listas[i].length, 3).setValues(listas[i]);

      aba.getRange(2, 3, listas[i].length, 1)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }

  if (mostrarMensagem) {
    SpreadsheetApp.getUi().alert("Listas atualizadas com sucesso.");
  }
}


/************************************************************
 RECEBIMENTO DO FORMULÁRIO
 ************************************************************/

function aoEnviarFormulario(e) {
  if (!e || !e.range) return;

  const aba = e.range.getSheet();

  if (aba.getName() !== CONFIG.ABA_RESPOSTAS) return;

  const linha = e.range.getRow();
  const ultimaColuna = aba.getLastColumn();

  if (ultimaColuna >= 2) {
    const intervalo = aba.getRange(linha, 2, 1, ultimaColuna - 1);
    const valores = intervalo.getValues();

    valores[0] = valores[0].map(valor => {
      if (valor instanceof Date) return valor;
      return String(valor || "").toUpperCase().trim();
    });

    intervalo.setValues(valores);
  }

  atualizarListas_(false);
}


/************************************************************
 CONFIGURAÇÃO DO FORMULÁRIO
 ************************************************************/

function criarFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let form = null;
  const formId = PROPRIEDADES.getProperty("FORM_ID");

  if (formId) {
    try {
      form = FormApp.openById(formId);
    } catch (e) {
      form = null;
    }
  }

  if (!form) {
    const arquivos = DriveApp.getFilesByName(CONFIG.NOME_FORMULARIO);

    while (arquivos.hasNext()) {
      const arquivo = arquivos.next();

      if (arquivo.getMimeType() === MimeType.GOOGLE_FORMS) {
        form = FormApp.openById(arquivo.getId());
        break;
      }
    }
  }

  if (!form) {
    form = FormApp.create(CONFIG.NOME_FORMULARIO);
    PROPRIEDADES.setProperty("FORM_ID", form.getId());
  }

  if (form.getItems().length === 0) {
    form.addTextItem()
      .setTitle("Nome completo")
      .setRequired(true);

    form.addTextItem()
      .setTitle("Apartamento")
      .setRequired(true);

    form.addCheckboxItem()
      .setTitle("Em quais sorteios deseja participar?")
      .setChoiceValues(CONFIG.AMBIENTES)
      .setRequired(true);
  }

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "aoEnviarFormulario") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("aoEnviarFormulario")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert(
    "Formulário configurado.\n\nID: " + form.getId()
  );
}


/************************************************************
 FUNÇÕES DOS QUATRO SORTEIOS
 ************************************************************/

function sortearGourmetNatal() {
  return sortearAmbiente_(0);
}

function sortearFestasNatal() {
  return sortearAmbiente_(1);
}

function sortearGourmetAnoNovo() {
  return sortearAmbiente_(2);
}

function sortearFestasAnoNovo() {
  return sortearAmbiente_(3);
}

function sortearAmbiente_(indice) {
  const numero = validarIndice_(indice);
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    return sortearAmbienteInterno_(numero);
  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 SORTEIO INTERNO
 ************************************************************/

function sortearAmbienteInterno_(indice) {
  const numero = validarIndice_(indice);

  atualizarListas_(false);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ambiente = CONFIG.AMBIENTES[numero];
  const aba = ss.getSheetByName(CONFIG.ABAS[numero]);
  const historico = ss.getSheetByName("Histórico dos Sorteios");

  if (historico.getLastRow() > 1) {
    const dadosHistorico = historico
      .getRange(2, 1, historico.getLastRow() - 1, 5)
      .getDisplayValues();

    for (let i = dadosHistorico.length - 1; i >= 0; i--) {
      if (
        normalizarComparacao_(dadosHistorico[i][1]) ===
        normalizarComparacao_(ambiente)
      ) {
        if (normalizarComparacao_(dadosHistorico[i][4]) === "ganhador") {
          throw new Error("Este ambiente já possui um ganhador ativo.");
        }
        break;
      }
    }
  }

  if (aba.getLastRow() < 2) {
    throw new Error("Não existem inscrições válidas para este ambiente.");
  }

  const candidatos = aba
    .getRange(2, 1, aba.getLastRow() - 1, 3)
    .getValues();

  const desistentes = obterDesistentesDoAmbiente_(ambiente);

  const validos = candidatos.filter(candidato => {
    const apartamento = normalizarComparacao_(candidato[1]);
    return apartamento && !desistentes.has(apartamento);
  });

  if (validos.length === 0) {
    throw new Error(
      "Não existem candidatos disponíveis para " + ambiente +
      ". Os participantes registrados já desistiram ou não há outras inscrições."
    );
  }

  const ganhador = validos[Math.floor(Math.random() * validos.length)];
  const agora = new Date();

  historico.appendRow([
    agora,
    ambiente,
    ganhador[0],
    ganhador[1],
    "Ganhador"
  ]);

  atualizarResultados_();

  ss.toast(
    "Ganhador: " + ganhador[0] + " – Apto " + ganhador[1],
    "SORTEIO",
    8
  );

  return {
    ambiente: ambiente,
    nome: ganhador[0],
    apartamento: ganhador[1],
    data: agora
  };
}


/************************************************************
 LISTA DE APARTAMENTOS QUE DESISTIRAM
 ************************************************************/

function obterDesistentesDoAmbiente_(ambiente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historico = ss.getSheetByName("Histórico dos Sorteios");
  const desistentes = new Set();

  if (!historico || historico.getLastRow() < 2) return desistentes;

  const dados = historico
    .getRange(2, 1, historico.getLastRow() - 1, 5)
    .getDisplayValues();

  dados.forEach(registro => {
    if (
      normalizarComparacao_(registro[1]) === normalizarComparacao_(ambiente) &&
      normalizarComparacao_(registro[4]) === "desistiu"
    ) {
      desistentes.add(normalizarComparacao_(registro[3]));
    }
  });

  return desistentes;
}


/************************************************************
 REGISTRO DE DESISTÊNCIA – MENU
 ************************************************************/

function registrarDesistenciaGourmetNatal() {
  return registrarDesistencia(0);
}

function registrarDesistenciaFestasNatal() {
  return registrarDesistencia(1);
}

function registrarDesistenciaGourmetAnoNovo() {
  return registrarDesistencia(2);
}

function registrarDesistenciaFestasAnoNovo() {
  return registrarDesistencia(3);
}

function registrarDesistencia(indice) {
  const numero = validarIndice_(indice);
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const resultado = registrarDesistenciaInterna_(numero);

    if (resultado.semNovoGanhador) {
      SpreadsheetApp.getUi().alert(
        "Desistência registrada.\n\n" +
        "Ganhador anterior: " + resultado.nomeAnterior +
        "\nApartamento: " + resultado.apartamentoAnterior +
        "\n\nNão existe outro participante disponível neste ambiente." +
        "\nO apartamento que desistiu foi excluído deste sorteio." +
        "\n\nQuando houver uma nova inscrição válida, será possível realizar outro sorteio."
      );
    } else {
      SpreadsheetApp.getUi().alert(
        "Desistência registrada com sucesso!\n\n" +
        "Ganhador anterior: " + resultado.nomeAnterior +
        "\n\nNovo ganhador: " + resultado.nome +
        "\nApartamento: " + resultado.apartamento
      );
    }

    return resultado;

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 PROCESSAMENTO INTERNO DA DESISTÊNCIA
 ************************************************************/

function registrarDesistenciaInterna_(indice) {
  const numero = validarIndice_(indice);

  atualizarListas_(false);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ambiente = CONFIG.AMBIENTES[numero];
  const aba = ss.getSheetByName(CONFIG.ABAS[numero]);
  const historico = ss.getSheetByName("Histórico dos Sorteios");

  if (!historico || historico.getLastRow() < 2) {
    throw new Error("Não existe ganhador registrado.");
  }

  const ultimaLinha = historico.getLastRow();

  const dados = historico
    .getRange(2, 1, ultimaLinha - 1, 5)
    .getDisplayValues();

  let linhaGanhador = 0;
  let nomeAnterior = "";
  let apartamentoAnterior = "";

  for (let i = dados.length - 1; i >= 0; i--) {
    const ambienteRegistro = normalizarComparacao_(dados[i][1]);

    if (ambienteRegistro === normalizarComparacao_(ambiente)) {
      if (normalizarComparacao_(dados[i][4]) === "ganhador") {
        linhaGanhador = i + 2;
        nomeAnterior = dados[i][2];
        apartamentoAnterior = dados[i][3];
      }
      break;
    }
  }

  if (!linhaGanhador) {
    throw new Error(
      "Não foi encontrado ganhador ativo para " + ambiente +
      ". Confira o histórico antes de tentar novamente."
    );
  }

  // Identifica os desistentes anteriores.
  const desistentes = obterDesistentesDoAmbiente_(ambiente);

  // O ganhador atual também será excluído do próximo sorteio.
  const apartamentoAtual = normalizarComparacao_(apartamentoAnterior);

  const candidatos = aba.getLastRow() > 1
    ? aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues()
    : [];

  const outrosCandidatos = candidatos.filter(candidato => {
    const apartamento = normalizarComparacao_(candidato[1]);

    return apartamento &&
      apartamento !== apartamentoAtual &&
      !desistentes.has(apartamento);
  });

  // Registra a desistência mesmo se não houver substituto.
  historico.getRange(linhaGanhador, 5).setValue("Desistiu");

  atualizarResultados_();

  // Se não houver outro candidato, encerra sem tentar sortear.
  if (outrosCandidatos.length === 0) {
    return {
      semNovoGanhador: true,
      ambiente: ambiente,
      nomeAnterior: nomeAnterior,
      apartamentoAnterior: apartamentoAnterior
    };
  }

  // Havendo outros participantes, realiza novo sorteio.
  const novoGanhador = sortearAmbienteInterno_(numero);

  atualizarResultados_();

  return {
    semNovoGanhador: false,
    ambiente: ambiente,
    nomeAnterior: nomeAnterior,
    apartamentoAnterior: apartamentoAnterior,
    nome: novoGanhador.nome,
    apartamento: novoGanhador.apartamento,
    data: novoGanhador.data
  };
}


/************************************************************
 RESULTADOS DOS SORTEIOS
 ************************************************************/

function atualizarResultados() {
  atualizarResultados_();
  SpreadsheetApp.getUi().alert("Resultados atualizados.");
}

function atualizarResultados_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbasSemMensagem_();

  const historico = ss.getSheetByName("Histórico dos Sorteios");
  const resultados = ss.getSheetByName("Resultados dos Sorteios");

  if (resultados.getLastRow() > 1) {
    resultados.getRange(
      2, 1, resultados.getLastRow() - 1, 5
    ).clearContent();
  }

  if (!historico || historico.getLastRow() < 2) return;

  const dados = historico
    .getRange(2, 1, historico.getLastRow() - 1, 5)
    .getValues();

  const ativos = {};

  dados.forEach(registro => {
    const chave = normalizarComparacao_(registro[1]);
    const status = normalizarComparacao_(registro[4]);

    if (!chave) return;

    if (status === "ganhador") {
      ativos[chave] = registro;
    } else if (status === "desistiu") {
      delete ativos[chave];
    }
  });

  const linhas = [];

  CONFIG.AMBIENTES.forEach(ambiente => {
    const registro = ativos[normalizarComparacao_(ambiente)];

    if (registro) {
      linhas.push([
        registro[1],
        registro[2],
        registro[3],
        registro[0],
        registro[4]
      ]);
    }
  });

  if (linhas.length > 0) {
    resultados.getRange(2, 1, linhas.length, 5).setValues(linhas);

    resultados.getRange(2, 4, linhas.length, 1)
      .setNumberFormat("dd/MM/yyyy HH:mm:ss");
  }
}


/************************************************************
 PAINEL ADMINISTRATIVO
 ************************************************************/

function verificarAcessoPainel_() {
  const email = String(
    Session.getActiveUser().getEmail() || ""
  ).trim().toLowerCase();

  if (!email) {
    throw new Error("Não foi possível identificar a conta Google.");
  }

  const autorizados = EMAILS_AUTORIZADOS.map(x => x.toLowerCase());

  if (!autorizados.includes(email)) {
    throw new Error("Acesso não autorizado.");
  }

  return true;
}

function doGet() {
  try {
    verificarAcessoPainel_();

    return HtmlService
      .createHtmlOutputFromFile("Index")
      .setTitle("Painel – Sorteio dos Salões")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (erro) {
    return HtmlService.createHtmlOutput(
      "<h2>Acesso restrito</h2><p>" + erro.message + "</p>"
    );
  }
}

function localizarGanhadorPainel_(indice) {
  const numero = validarIndice_(indice);
  const ambiente = AMBIENTES_PAINEL[numero];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historico = ss.getSheetByName("Histórico dos Sorteios");

  if (!historico || historico.getLastRow() < 2) return null;

  const dados = historico
    .getRange(2, 1, historico.getLastRow() - 1, 5)
    .getValues();

  for (let i = dados.length - 1; i >= 0; i--) {
    if (
      normalizarComparacao_(dados[i][1]) ===
      normalizarComparacao_(ambiente)
    ) {
      if (normalizarComparacao_(dados[i][4]) === "ganhador") {
        return {
          ambiente: dados[i][1],
          nome: dados[i][2],
          apartamento: dados[i][3],
          data: dados[i][0],
          status: dados[i][4]
        };
      }

      return null;
    }
  }

  return null;
}

function obterDadosPainel() {
  verificarAcessoPainel_();

  return AMBIENTES_PAINEL.map((ambiente, indice) => ({
    indice: indice,
    ambiente: ambiente,
    ganhador: localizarGanhadorPainel_(indice)
  }));
}

function executarSorteioPainel(indice) {
  verificarAcessoPainel_();
  return sortearAmbiente_(validarIndice_(indice));
}

function obterResultadosPainel() {
  verificarAcessoPainel_();

  atualizarResultados_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Resultados dos Sorteios");

  if (!aba || aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 5)
    .getValues()
    .map(linha => ({
      ambiente: linha[0],
      nome: linha[1],
      apartamento: linha[2],
      data: linha[3],
      status: linha[4]
    }));
}

function obterHistoricoPainel() {
  verificarAcessoPainel_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Histórico dos Sorteios");

  if (!aba || aba.getLastRow() < 2) return [];

  return aba.getRange(2, 1, aba.getLastRow() - 1, 5)
    .getValues()
    .reverse()
    .map(linha => ({
      data: linha[0],
      ambiente: linha[1],
      nome: linha[2],
      apartamento: linha[3],
      status: linha[4]
    }));
}

function obterGanhadorPainel(indice) {
  verificarAcessoPainel_();
  return localizarGanhadorPainel_(validarIndice_(indice));
}

function registrarDesistenciaPainel(indice) {
  verificarAcessoPainel_();

  const numero = validarIndice_(indice);
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const resultado = registrarDesistenciaInterna_(numero);

    if (resultado.semNovoGanhador) {
      throw new Error(
        "Desistência registrada, mas não há outro participante disponível para " +
        resultado.ambiente + ". O apartamento " +
        resultado.apartamentoAnterior +
        " foi excluído deste ambiente."
      );
    }

    return {
      ambiente: resultado.ambiente,
      nome: resultado.nome,
      apartamento: resultado.apartamento,
      data: resultado.data
    };

  } finally {
    lock.releaseLock();
  }
}
