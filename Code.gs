/************************************************************
 SISTEMA DE SORTEIO – NATAL E ANO NOVO
 PLANILHA: SORTEIO FIM DE ANO

 REGRAS:
 1. Um apartamento = uma inscrição válida.
 2. A última resposta substitui a anterior.
 3. Cada ambiente possui seu próprio sorteio.
 4. Desistência exclui o ganhador somente daquele ambiente.
 5. Histórico dos sorteios é preservado.
 6. Sorteios protegidos contra execução simultânea.
 7. A aba de respostas mantém todas as respostas do formulário.
 8. As abas dos ambientes mostram somente a inscrição válida
    mais recente de cada apartamento.
 ************************************************************/


/************************************************************
 CONFIGURAÇÃO
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


/************************************************************
 EMAILS AUTORIZADOS AO PAINEL ADMINISTRATIVO
 ************************************************************/

const EMAILS_AUTORIZADOS = [
  "rseneda@gmail.com"
];


/************************************************************
 AMBIENTES DO PAINEL
 ************************************************************/

const AMBIENTES_PAINEL = [
  "Salão Gourmet – Natal (25 pessoas)",
  "Salão de Festas – Natal (50 pessoas)",
  "Salão Gourmet – Ano Novo (25 pessoas)",
  "Salão de Festas – Ano Novo (50 pessoas)"
];


/************************************************************
 PROPRIEDADES
 ************************************************************/

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
 NORMALIZAÇÃO
 ************************************************************/

/*
 * Normalização simples:
 * usada para nomes, apartamentos etc.
 */
function normalizar_(valor) {

  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}


/*
 * Normalização extremamente robusta para comparar
 * respostas do formulário com os nomes dos ambientes.
 *
 * Remove:
 * - acentos
 * - maiúsculas/minúsculas
 * - espaços
 * - pontuação
 * - travessões diferentes
 * - caracteres especiais
 */
function normalizarComparacao_(valor) {

  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}


/************************************************************
 CRIAÇÃO / CORREÇÃO DAS ABAS
 ************************************************************/

function criarAbas() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /*
   * Abas dos quatro ambientes
   */
  const cabecalhoAmbiente = [
    "Nome completo",
    "Apartamento",
    "Data da inscrição"
  ];

  for (let i = 0; i < 4; i++) {

    let aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (!aba) {
      aba = ss.insertSheet(CONFIG.ABAS[i]);
    }

    /*
     * Se estiver completamente vazia,
     * cria o cabeçalho.
     */
    if (aba.getLastRow() === 0) {

      aba.getRange(1, 1, 1, 3)
        .setValues([cabecalhoAmbiente]);

      aba.setFrozenRows(1);

    } else {

      /*
       * Garante que o cabeçalho esteja correto.
       */
      aba.getRange(1, 1, 1, 3)
        .setValues([cabecalhoAmbiente]);

      aba.setFrozenRows(1);
    }
  }


  /**********************************************************
   ABA RESULTADOS
   **********************************************************/

  let resultados = ss.getSheetByName("Resultados dos Sorteios");

  if (!resultados) {
    resultados = ss.insertSheet("Resultados dos Sorteios");
  }

  const cabecalhoResultados = [
    "Ambiente",
    "Ganhador",
    "Apartamento",
    "Data do sorteio",
    "Status"
  ];

  /*
   * Corrige somente a primeira linha.
   * Não apaga resultados existentes.
   */
  resultados.getRange(
    1,
    1,
    1,
    cabecalhoResultados.length
  ).setValues([cabecalhoResultados]);

  resultados.setFrozenRows(1);


  /**********************************************************
   ABA HISTÓRICO
   **********************************************************/

  let historico = ss.getSheetByName("Histórico dos Sorteios");

  if (!historico) {
    historico = ss.insertSheet("Histórico dos Sorteios");
  }

  const cabecalhoHistorico = [
    "Data/Hora",
    "Ambiente",
    "Nome",
    "Apartamento",
    "Status"
  ];

  /*
   * Corrige somente a primeira linha.
   * O histórico existente é preservado.
   */
  historico.getRange(
    1,
    1,
    1,
    cabecalhoHistorico.length
  ).setValues([cabecalhoHistorico]);

  historico.setFrozenRows(1);


  SpreadsheetApp.getUi().alert(
    "Abas criadas/corrigidas com sucesso."
  );
}


/************************************************************
 LOCALIZA ABA DE RESPOSTAS
 ************************************************************/

function localizarAbaRespostas_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let aba = ss.getSheetByName(CONFIG.ABA_RESPOSTAS);

  if (aba) {
    return aba;
  }

  /*
   * Caso o nome tenha sido alterado,
   * procura pela aba que possui os cabeçalhos esperados.
   */

  const abas = ss.getSheets();

  for (let i = 0; i < abas.length; i++) {

    const sheet = abas[i];

    if (CONFIG.ABAS.indexOf(sheet.getName()) !== -1) {
      continue;
    }

    if (sheet.getLastColumn() === 0) {
      continue;
    }

    const cabecalhos = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(function(valor) {
        return normalizar_(valor);
      });

    const temNome = cabecalhos.some(function(c) {
      return c.indexOf("nome completo") !== -1;
    });

    const temApartamento = cabecalhos.some(function(c) {
      return c.indexOf("apartamento") !== -1;
    });

    const temSorteios = cabecalhos.some(function(c) {
      return c.indexOf("sorteios") !== -1;
    });

    if (temNome && temApartamento && temSorteios) {
      return sheet;
    }
  }

  throw new Error(
    "Não foi encontrada a aba de respostas do formulário."
  );
}


/************************************************************
 LOCALIZA COLUNAS
 ************************************************************/

function localizarColunasRespostas_(aba) {

  const ultimaColuna = aba.getLastColumn();

  if (ultimaColuna === 0) {
    throw new Error("A aba de respostas está vazia.");
  }

  const cabecalhos = aba
    .getRange(1, 1, 1, ultimaColuna)
    .getValues()[0];

  let colunaTimestamp = -1;
  let colunaNome = -1;
  let colunaApartamento = -1;
  let colunaSorteios = -1;

  cabecalhos.forEach(function(cabecalho, index) {

    const normalizado = normalizar_(cabecalho);

    if (
      colunaTimestamp === -1 &&
      (
        normalizado.indexOf("carimbo") !== -1 ||
        normalizado.indexOf("timestamp") !== -1 ||
        normalizado.indexOf("data") !== -1
      )
    ) {
      colunaTimestamp = index;
    }

    if (
      colunaNome === -1 &&
      normalizado.indexOf("nome completo") !== -1
    ) {
      colunaNome = index;
    }

    if (
      colunaApartamento === -1 &&
      normalizado.indexOf("apartamento") !== -1
    ) {
      colunaApartamento = index;
    }

    if (
      colunaSorteios === -1 &&
      normalizado.indexOf("sorteios") !== -1
    ) {
      colunaSorteios = index;
    }
  });


  if (colunaTimestamp === -1) {
    colunaTimestamp = 0;
  }

  if (colunaNome === -1) {
    throw new Error(
      "Não encontrei a coluna 'Nome completo'."
    );
  }

  if (colunaApartamento === -1) {
    throw new Error(
      "Não encontrei a coluna 'Apartamento'."
    );
  }

  if (colunaSorteios === -1) {
    throw new Error(
      "Não encontrei a coluna de escolha dos sorteios."
    );
  }


  return {
    timestamp: colunaTimestamp,
    nome: colunaNome,
    apartamento: colunaApartamento,
    sorteios: colunaSorteios
  };
}


/************************************************************
 FORMATA DATA
 ************************************************************/

function obterDataResposta_(valor) {

  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  const data = new Date(valor);

  if (!isNaN(data.getTime())) {
    return data;
  }

  return new Date();
}


/************************************************************
 ATUALIZAR LISTAS
 ************************************************************/

function atualizarListas() {

  atualizarListas_(true);
}


function atualizarListas_(mostrarMensagem) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  criarAbasSemMensagem_();

  const abaRespostas = localizarAbaRespostas_();

  const colunas = localizarColunasRespostas_(abaRespostas);

  const ultimaLinha = abaRespostas.getLastRow();
  const ultimaColuna = abaRespostas.getLastColumn();

  /*
   * Limpa as quatro abas operacionais.
   */
  for (let i = 0; i < 4; i++) {

    const aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (aba.getLastRow() > 1) {

      aba.getRange(
        2,
        1,
        aba.getLastRow() - 1,
        3
      ).clearContent();
    }
  }


  if (ultimaLinha < 2) {

    if (mostrarMensagem) {
      SpreadsheetApp.getUi().alert(
        "A aba de respostas ainda não possui inscrições."
      );
    }

    return;
  }


  const dados = abaRespostas
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      ultimaColuna
    )
    .getValues();


  /**********************************************************
   * ÚLTIMA RESPOSTA DE CADA APARTAMENTO
   **********************************************************/

  const ultimaRespostaPorApartamento = {};


  dados.forEach(function(linha) {

    const apartamento = String(
      linha[colunas.apartamento] || ""
    ).trim();

    if (!apartamento) {
      return;
    }

    const chaveApartamento =
      normalizarComparacao_(apartamento);

    if (!chaveApartamento) {
      return;
    }

    const timestamp =
      obterDataResposta_(linha[colunas.timestamp]);

    const resposta = {
      timestamp: timestamp,
      nome: String(
        linha[colunas.nome] || ""
      ).trim(),
      apartamento: apartamento,
      escolhas: String(
        linha[colunas.sorteios] || ""
      ).trim()
    };


    /*
     * A última resposta substitui a anterior.
     */
    if (
      !ultimaRespostaPorApartamento[chaveApartamento] ||
      timestamp.getTime() >=
        ultimaRespostaPorApartamento[
          chaveApartamento
        ].timestamp.getTime()
    ) {

      ultimaRespostaPorApartamento[
        chaveApartamento
      ] = resposta;
    }
  });


  /**********************************************************
   * DATA ATUAL / PRAZO
   **********************************************************/

  const hoje = new Date();

  const anoAtual = hoje.getFullYear();

  const prazo = new Date(
    anoAtual,
    CONFIG.PRAZO_INSCRICAO.MES - 1,
    CONFIG.PRAZO_INSCRICAO.DIA,
    23,
    59,
    59
  );


  /**********************************************************
   * LISTAS DOS AMBIENTES
   **********************************************************/

  const listas = [
    [],
    [],
    [],
    []
  ];


  Object.keys(
    ultimaRespostaPorApartamento
  ).forEach(function(chave) {

    const resposta =
      ultimaRespostaPorApartamento[chave];

    /*
     * Ignora respostas de anos anteriores.
     */
    if (
      resposta.timestamp.getFullYear() !== anoAtual
    ) {
      return;
    }

    /*
     * Ignora inscrições depois do prazo.
     */
    if (
      resposta.timestamp.getTime() > prazo.getTime()
    ) {
      return;
    }

    if (!resposta.nome || !resposta.apartamento) {
      return;
    }


    /*
     * NORMALIZAÇÃO ROBUSTA
     *
     * Isso resolve o problema de:
     *
     * SALÃO GOURMET – NATAL
     *
     * versus
     *
     * Salão Gourmet - Natal
     *
     * ou diferenças de acentos, espaços etc.
     */
    const escolhasNormalizadas =
      normalizarComparacao_(
        resposta.escolhas
      );


    CONFIG.AMBIENTES.forEach(function(
      ambiente,
      indice
    ) {

      const ambienteNormalizado =
        normalizarComparacao_(ambiente);


      /*
       * Verifica se o ambiente escolhido
       * está realmente na resposta.
       */
      if (
        escolhasNormalizadas.indexOf(
          ambienteNormalizado
        ) !== -1
      ) {

        listas[indice].push([
          resposta.nome,
          resposta.apartamento,
          resposta.timestamp
        ]);
      }
    });
  });


  /**********************************************************
   * ESCREVE AS LISTAS
   **********************************************************/

  for (let i = 0; i < 4; i++) {

    const aba = ss.getSheetByName(
      CONFIG.ABAS[i]
    );

    if (listas[i].length > 0) {

      aba.getRange(
        2,
        1,
        listas[i].length,
        3
      ).setValues(listas[i]);

      aba.getRange(
        2,
        3,
        listas[i].length,
        1
      ).setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );
    }
  }


  if (mostrarMensagem) {

    SpreadsheetApp.getUi().alert(
      "Listas atualizadas com sucesso."
    );
  }
}


/************************************************************
 CRIA ABAS SEM MOSTRAR ALERTA
 ************************************************************/

function criarAbasSemMensagem_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cabecalhoAmbiente = [
    "Nome completo",
    "Apartamento",
    "Data da inscrição"
  ];

  for (let i = 0; i < 4; i++) {

    let aba = ss.getSheetByName(CONFIG.ABAS[i]);

    if (!aba) {
      aba = ss.insertSheet(CONFIG.ABAS[i]);
    }

    aba.getRange(
      1,
      1,
      1,
      3
    ).setValues([
      cabecalhoAmbiente
    ]);

    aba.setFrozenRows(1);
  }


  let resultados =
    ss.getSheetByName(
      "Resultados dos Sorteios"
    );

  if (!resultados) {
    resultados =
      ss.insertSheet(
        "Resultados dos Sorteios"
      );
  }

  resultados.getRange(
    1,
    1,
    1,
    5
  ).setValues([[
    "Ambiente",
    "Ganhador",
    "Apartamento",
    "Data do sorteio",
    "Status"
  ]]);

  resultados.setFrozenRows(1);


  let historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );

  if (!historico) {
    historico =
      ss.insertSheet(
        "Histórico dos Sorteios"
      );
  }

  historico.getRange(
    1,
    1,
    1,
    5
  ).setValues([[
    "Data/Hora",
    "Ambiente",
    "Nome",
    "Apartamento",
    "Status"
  ]]);

  historico.setFrozenRows(1);
}


/************************************************************
 EVENTO DO FORMULÁRIO
 ************************************************************/

function aoEnviarFormulario(e) {

  if (!e || !e.range) {
    return;
  }

  const aba = e.range.getSheet();

  /*
   * Só processa a aba correta.
   */
  if (
    aba.getName() !== CONFIG.ABA_RESPOSTAS
  ) {
    return;
  }


  /*
   * Converte Nome, Apartamento e demais
   * respostas para maiúsculas.
   *
   * Mantém o timestamp intacto.
   */
  const linha = e.range.getRow();

  const ultimaColuna = aba.getLastColumn();

  if (ultimaColuna >= 2) {

    const intervalo =
      aba.getRange(
        linha,
        2,
        1,
        ultimaColuna - 1
      );

    const valores =
      intervalo.getValues();

    valores[0] = valores[0].map(
      function(valor) {

        if (
          valor instanceof Date
        ) {
          return valor;
        }

        return String(valor || "")
          .toUpperCase()
          .trim();
      }
    );

    intervalo.setValues(valores);
  }


  /*
   * Atualiza imediatamente as listas.
   */
  atualizarListas_(false);
}


/************************************************************
 FUNÇÃO DE TESTE MANUAL
 ************************************************************/

function testarAtualizacao() {

  atualizarListas_(true);
}


/************************************************************
 SORTEIOS
 ************************************************************/

function sortearGourmetNatal() {

  return sortearAmbiente_(
    0
  );
}


function sortearFestasNatal() {

  return sortearAmbiente_(
    1
  );
}


function sortearGourmetAnoNovo() {

  return sortearAmbiente_(
    2
  );
}


function sortearFestasAnoNovo() {

  return sortearAmbiente_(
    3
  );
}


/************************************************************
 SORTEIO COM LOCK
 ************************************************************/

function sortearAmbiente_(indice) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    return sortearAmbienteInterno_(
      indice
    );

  } finally {

    lock.releaseLock();
  }
}


/************************************************************
 SORTEIO INTERNO
 ************************************************************/

function sortearAmbienteInterno_(indice) {

  if (
    indice < 0 ||
    indice >= CONFIG.AMBIENTES.length
  ) {
    throw new Error(
      "Ambiente inválido."
    );
  }


  /*
   * Sempre atualiza as listas antes do sorteio.
   */
  atualizarListas_(false);


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const ambiente =
    CONFIG.AMBIENTES[indice];

  const aba =
    ss.getSheetByName(
      CONFIG.ABAS[indice]
    );


  if (!aba) {
    throw new Error(
      "A aba do ambiente não existe."
    );
  }


  /**********************************************************
   * VERIFICA SE JÁ EXISTE GANHADOR ATIVO
   **********************************************************/

  const historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );

  const ultimaLinha =
    historico.getLastRow();


  if (ultimaLinha > 1) {

    const dadosHistorico =
      historico.getRange(
        2,
        1,
        ultimaLinha - 1,
        5
      ).getValues();


    /*
     * Procura o último registro desse ambiente.
     */
    for (
      let i = dadosHistorico.length - 1;
      i >= 0;
      i--
    ) {

      const registro =
        dadosHistorico[i];

      const ambienteRegistro =
        normalizarComparacao_(
          registro[1]
        );

      const ambienteAtual =
        normalizarComparacao_(
          ambiente
        );


      if (
        ambienteRegistro ===
        ambienteAtual
      ) {

        const status =
          normalizarComparacao_(
            registro[4]
          );

        /*
         * Se já existe ganhador ativo,
         * não faz outro sorteio.
         */
        if (
          status ===
          normalizarComparacao_(
            "Ganhador"
          )
        ) {

          throw new Error(
            "Este ambiente já possui um ganhador ativo."
          );
        }

        /*
         * Se o último registro for desistência,
         * o novo sorteio pode continuar.
         */
        break;
      }
    }
  }


  /**********************************************************
   * LÊ CANDIDATOS
   **********************************************************/

  const ultimaLinhaAba =
    aba.getLastRow();

  if (ultimaLinhaAba < 2) {

    throw new Error(
      "Não existem inscrições válidas para este ambiente."
    );
  }


  const candidatos =
    aba.getRange(
      2,
      1,
      ultimaLinhaAba - 1,
      3
    ).getValues();


  /*
   * Apartamentos que já desistiram desse ambiente.
   */
  const desistentes =
    obterDesistentesDoAmbiente_(
      ambiente
    );


  const candidatosValidos =
    candidatos.filter(
      function(candidato) {

        const apartamento =
          normalizarComparacao_(
            candidato[1]
          );

        return (
          apartamento &&
          !desistentes.has(
            apartamento
          )
        );
      }
    );


  if (
    candidatosValidos.length === 0
  ) {

    throw new Error(
      "Não existem candidatos válidos para este sorteio."
    );
  }


  /**********************************************************
   * SORTEIO
   **********************************************************/

  const indiceSorteado =
    Math.floor(
      Math.random() *
      candidatosValidos.length
    );


  const ganhador =
    candidatosValidos[
      indiceSorteado
    ];


  const agora =
    new Date();


  /**********************************************************
   * REGISTRA NO HISTÓRICO
   **********************************************************/

  historico.appendRow([
    agora,
    ambiente,
    ganhador[0],
    ganhador[1],
    "Ganhador"
  ]);


  /*
   * Atualiza resultados.
   */
  atualizarResultados_();


  SpreadsheetApp.getActiveSpreadsheet()
    .toast(
      "Ganhador: " +
      ganhador[0] +
      " – Apto " +
      ganhador[1],
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
 DESISTENTES DE UM AMBIENTE
 ************************************************************/

function obterDesistentesDoAmbiente_(ambiente) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );

  const desistentes =
    new Set();


  if (
    !historico ||
    historico.getLastRow() < 2
  ) {
    return desistentes;
  }


  const dados =
    historico.getRange(
      2,
      1,
      historico.getLastRow() - 1,
      5
    ).getValues();


  const ambienteNormalizado =
    normalizarComparacao_(
      ambiente
    );


  dados.forEach(
    function(registro) {

      const ambienteRegistro =
        normalizarComparacao_(
          registro[1]
        );

      const status =
        normalizarComparacao_(
          registro[4]
        );


      if (
        ambienteRegistro ===
        ambienteNormalizado &&
        status ===
        normalizarComparacao_(
          "Desistiu"
        )
      ) {

        const apartamento =
          normalizarComparacao_(
            registro[3]
          );

        if (apartamento) {
          desistentes.add(
            apartamento
          );
        }
      }
    }
  );


  return desistentes;
}


/************************************************************
 DESISTÊNCIAS
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


/************************************************************
 REGISTRAR DESISTÊNCIA
 ************************************************************/

function registrarDesistencia(indice) {

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {

    if (
      indice < 0 ||
      indice >= CONFIG.AMBIENTES.length
    ) {

      throw new Error(
        "Ambiente inválido."
      );
    }


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const ambiente =
      CONFIG.AMBIENTES[indice];


    const historico =
      ss.getSheetByName(
        "Histórico dos Sorteios"
      );


    if (
      !historico ||
      historico.getLastRow() < 2
    ) {

      throw new Error(
        "Não existe ganhador registrado."
      );
    }


    const dados =
      historico.getRange(
        2,
        1,
        historico.getLastRow() - 1,
        5
      ).getValues();


    const ambienteNormalizado =
      normalizarComparacao_(
        ambiente
      );


    let linhaGanhador = -1;


    /*
     * Procura de trás para frente.
     * Isso garante que encontramos o estado
     * mais recente do ambiente.
     */
    for (
      let i = dados.length - 1;
      i >= 0;
      i--
    ) {

      const ambienteRegistro =
        normalizarComparacao_(
          dados[i][1]
        );

      const status =
        normalizarComparacao_(
          dados[i][4]
        );


      if (
        ambienteRegistro ===
        ambienteNormalizado
      ) {

        if (
          status ===
          normalizarComparacao_(
            "Ganhador"
          )
        ) {

          linhaGanhador =
            i + 2;

        } else {

          /*
           * O último registro não é ganhador ativo.
           */
          throw new Error(
            "Este ambiente não possui ganhador ativo."
          );
        }

        break;
      }
    }


    if (
      linhaGanhador === -1
    ) {

      throw new Error(
        "Não foi encontrado ganhador ativo."
      );
    }


    /*
     * Marca o ganhador como desistente.
     */
    historico
      .getRange(
        linhaGanhador,
        5
      )
      .setValue(
        "Desistiu"
      );


    /*
     * Faz novo sorteio usando a função INTERNA,
     * porque o LockService já está ativo.
     */
    const novoGanhador =
      sortearAmbienteInterno_(
        indice
      );


    atualizarResultados_();


    SpreadsheetApp.getUi().alert(
      "Desistência registrada.\n\n" +
      "Novo ganhador:\n" +
      novoGanhador.nome +
      "\nApartamento: " +
      novoGanhador.apartamento
    );


    return novoGanhador;

  } finally {

    lock.releaseLock();
  }
}


/************************************************************
 RESULTADOS
 ************************************************************/

function atualizarResultados() {

  atualizarResultados_();

  SpreadsheetApp.getUi().alert(
    "Resultados atualizados."
  );
}


function atualizarResultados_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  criarAbasSemMensagem_();


  const historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );


  const resultados =
    ss.getSheetByName(
      "Resultados dos Sorteios"
    );


  /*
   * Limpa somente os resultados antigos,
   * preservando o cabeçalho.
   */
  if (
    resultados.getLastRow() > 1
  ) {

    resultados
      .getRange(
        2,
        1,
        resultados.getLastRow() - 1,
        5
      )
      .clearContent();
  }


  if (
    !historico ||
    historico.getLastRow() < 2
  ) {

    return;
  }


  const dados =
    historico.getRange(
      2,
      1,
      historico.getLastRow() - 1,
      5
    ).getValues();


  const ganhadores = {};


  /*
   * Percorre em ordem cronológica.
   * O último status de cada ambiente prevalece.
   */
  dados.forEach(
    function(registro) {

      const ambiente =
        String(
          registro[1] || ""
        ).trim();

      if (!ambiente) {
        return;
      }

      const chave =
        normalizarComparacao_(
          ambiente
        );


      const status =
        String(
          registro[4] || ""
        ).trim();


      if (
        normalizarComparacao_(
          status
        ) ===
        normalizarComparacao_(
          "Ganhador"
        )
      ) {

        ganhadores[chave] = {
          ambiente: ambiente,
          nome: registro[2],
          apartamento: registro[3],
          data: registro[0],
          status: "Ganhador"
        };

      } else if (
        normalizarComparacao_(
          status
        ) ===
        normalizarComparacao_(
          "Desistiu"
        )
      ) {

        delete ganhadores[chave];
      }
    }
  );


  const linhas = [];


  CONFIG.AMBIENTES.forEach(
    function(ambiente) {

      const chave =
        normalizarComparacao_(
          ambiente
        );

      const ganhador =
        ganhadores[chave];


      if (ganhador) {

        linhas.push([
          ganhador.ambiente,
          ganhador.nome,
          ganhador.apartamento,
          ganhador.data,
          ganhador.status
        ]);
      }
    }
  );


  if (linhas.length > 0) {

    resultados
      .getRange(
        2,
        1,
        linhas.length,
        5
      )
      .setValues(linhas);


    resultados
      .getRange(
        2,
        4,
        linhas.length,
        1
      )
      .setNumberFormat(
        "dd/MM/yyyy HH:mm:ss"
      );
  }
}


/************************************************************
 CRIAR FORMULÁRIO
 ************************************************************/

function criarFormulario() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  let formId =
    PROPRIEDADES.getProperty(
      "FORM_ID"
    );


  let form;


  if (formId) {

    try {

      form =
        FormApp.openById(
          formId
        );

    } catch (erro) {

      form = null;
    }
  }


  /*
   * Procura formulário pelo nome se o ID
   * não estiver disponível.
   */
  if (!form) {

    const arquivos =
      DriveApp.getFilesByName(
        CONFIG.NOME_FORMULARIO
      );


    while (
      arquivos.hasNext()
    ) {

      const arquivo =
        arquivos.next();


      if (
        arquivo.getMimeType() ===
        MimeType.GOOGLE_FORMS
      ) {

        form =
          FormApp.openById(
            arquivo.getId()
          );

        break;
      }
    }
  }


  /*
   * Cria o formulário se ainda não existir.
   */
  if (!form) {

    form =
      FormApp.create(
        CONFIG.NOME_FORMULARIO
      );

    formId =
      form.getId();


    PROPRIEDADES.setProperty(
      "FORM_ID",
      formId
    );
  }


  /*
   * Verifica se já existem perguntas.
   */
  const itens =
    form.getItems();


  if (itens.length === 0) {

    form.addTextItem()
      .setTitle(
        "Nome completo"
      )
      .setRequired(true);


    form.addTextItem()
      .setTitle(
        "Apartamento"
      )
      .setRequired(true);


    form.addCheckboxItem()
      .setTitle(
        "Em quais sorteios deseja participar?"
      )
      .setChoiceValues(
        CONFIG.AMBIENTES
      )
      .setRequired(true);
  }


  /*
   * Vincula à planilha.
   */
  form.setDestination(
    FormApp.DestinationType.SPREADSHEET,
    ss.getId()
  );


  /*
   * Remove triggers antigos da função
   * para evitar duplicidade.
   */
  const triggers =
    ScriptApp.getProjectTriggers();


  triggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        "aoEnviarFormulario"
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );
      }
    }
  );


  /*
   * Cria novo trigger.
   */
  ScriptApp.newTrigger(
    "aoEnviarFormulario"
  )
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();


  SpreadsheetApp.getUi().alert(
    "Formulário configurado.\n\n" +
    "ID do formulário:\n" +
    form.getId()
  );
}


/************************************************************
 PAINEL ADMINISTRATIVO
 ************************************************************/

function verificarAcessoPainel_() {

  const email =
    String(
      Session.getActiveUser()
        .getEmail() || ""
    )
    .trim()
    .toLowerCase();


  if (!email) {

    throw new Error(
      "Não foi possível identificar a conta Google que acessou o painel."
    );
  }


  const autorizado =
    EMAILS_AUTORIZADOS
      .map(function(item) {
        return item.toLowerCase();
      })
      .indexOf(email) !== -1;


  if (!autorizado) {

    throw new Error(
      "Acesso não autorizado."
    );
  }


  return true;
}


/************************************************************
 DO GET
 ************************************************************/

function doGet() {

  try {

    verificarAcessoPainel_();


    return HtmlService
      .createHtmlOutputFromFile(
        "Index"
      )
      .setTitle(
        "Painel – Sorteio dos Salões"
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  } catch (erro) {

    return HtmlService
      .createHtmlOutput(
        "<h2>Acesso restrito</h2>" +
        "<p>" +
        erro.message +
        "</p>"
      );
  }
}


/************************************************************
 DADOS DO PAINEL
 ************************************************************/

function obterDadosPainel() {

  verificarAcessoPainel_();


  const dados =
    AMBIENTES_PAINEL.map(
      function(ambiente, indice) {

        return {
          indice: indice,
          ambiente: ambiente,
          ganhador:
            localizarGanhadorPainel_(
              indice
            )
        };
      }
    );


  return dados;
}


/************************************************************
 LOCALIZAR GANHADOR
 ************************************************************/

function localizarGanhadorPainel_(
  indice
) {

  const ambiente =
    AMBIENTES_PAINEL[indice];


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );


  if (
    !historico ||
    historico.getLastRow() < 2
  ) {

    return null;
  }


  const dados =
    historico.getRange(
      2,
      1,
      historico.getLastRow() - 1,
      5
    ).getValues();


  const ambienteNormalizado =
    normalizarComparacao_(
      ambiente
    );


  for (
    let i = dados.length - 1;
    i >= 0;
    i--
  ) {

    const ambienteRegistro =
      normalizarComparacao_(
        dados[i][1]
      );


    if (
      ambienteRegistro ===
      ambienteNormalizado
    ) {

      const status =
        normalizarComparacao_(
          dados[i][4]
        );


      if (
        status ===
        normalizarComparacao_(
          "Ganhador"
        )
      ) {

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


/************************************************************
 EXECUTAR SORTEIO PELO PAINEL
 ************************************************************/

function executarSorteioPainel(
  indice
) {

  verificarAcessoPainel_();


  const numero =
    Number(indice);


  if (
    isNaN(numero) ||
    numero < 0 ||
    numero >= AMBIENTES_PAINEL.length
  ) {

    throw new Error(
      "Ambiente inválido."
    );
  }


  return sortearAmbiente_(
    numero
  );
}


/************************************************************
 RESULTADOS DO PAINEL
 ************************************************************/

function obterResultadosPainel() {

  verificarAcessoPainel_();


  atualizarResultados_();


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const resultados =
    ss.getSheetByName(
      "Resultados dos Sorteios"
    );


  if (
    !resultados ||
    resultados.getLastRow() < 2
  ) {

    return [];
  }


  return resultados
    .getRange(
      2,
      1,
      resultados.getLastRow() - 1,
      5
    )
    .getValues()
    .map(
      function(linha) {

        return {
          ambiente: linha[0],
          nome: linha[1],
          apartamento: linha[2],
          data: linha[3],
          status: linha[4]
        };
      }
    );
}


/************************************************************
 HISTÓRICO DO PAINEL
 ************************************************************/

function obterHistoricoPainel() {

  verificarAcessoPainel_();


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const historico =
    ss.getSheetByName(
      "Histórico dos Sorteios"
    );


  if (
    !historico ||
    historico.getLastRow() < 2
  ) {

    return [];
  }


  return historico
    .getRange(
      2,
      1,
      historico.getLastRow() - 1,
      5
    )
    .getValues()
    .reverse()
    .map(
      function(linha) {

        return {
          data: linha[0],
          ambiente: linha[1],
          nome: linha[2],
          apartamento: linha[3],
          status: linha[4]
        };
      }
    );
}


/************************************************************
 GANHADOR DO PAINEL
 ************************************************************/

function obterGanhadorPainel(
  indice
) {

  verificarAcessoPainel_();


  const numero =
    Number(indice);


  if (
    isNaN(numero) ||
    numero < 0 ||
    numero >= AMBIENTES_PAINEL.length
  ) {

    throw new Error(
      "Ambiente inválido."
    );
  }


  return localizarGanhadorPainel_(
    numero
  );
}


/************************************************************
 DESISTÊNCIA PELO PAINEL
 ************************************************************/

function registrarDesistenciaPainel(
  indice
) {

  verificarAcessoPainel_();


  const numero =
    Number(indice);


  if (
    isNaN(numero) ||
    numero < 0 ||
    numero >= AMBIENTES_PAINEL.length
  ) {

    throw new Error(
      "Ambiente inválido."
    );
  }


  const lock =
    LockService.getScriptLock();


  lock.waitLock(30000);


  try {

    const ambiente =
      AMBIENTES_PAINEL[
        numero
      ];


    const ss =
      SpreadsheetApp.getActiveSpreadsheet();


    const historico =
      ss.getSheetByName(
        "Histórico dos Sorteios"
      );


    if (
      !historico ||
      historico.getLastRow() < 2
    ) {

      throw new Error(
        "Não existe histórico para este ambiente."
      );
    }


    const dados =
      historico.getRange(
        2,
        1,
        historico.getLastRow() - 1,
        5
      ).getValues();


    const ambienteNormalizado =
      normalizarComparacao_(
        ambiente
      );


    let linhaGanhador = -1;


    /*
     * Procura o ganhador ativo atual.
     */
    for (
      let i = dados.length - 1;
      i >= 0;
      i--
    ) {

      const ambienteRegistro =
        normalizarComparacao_(
          dados[i][1]
        );


      if (
        ambienteRegistro !==
        ambienteNormalizado
      ) {
        continue;
      }


      const status =
        normalizarComparacao_(
          dados[i][4]
        );


      if (
        status ===
        normalizarComparacao_(
          "Ganhador"
        )
      ) {

        linhaGanhador =
          i + 2;

      } else {

        throw new Error(
          "Este ambiente não possui ganhador ativo."
        );
      }


      break;
    }


    if (
      linhaGanhador === -1
    ) {

      throw new Error(
        "Não foi encontrado ganhador ativo."
      );
    }


    /*
     * Registra desistência.
     */
    historico
      .getRange(
        linhaGanhador,
        5
      )
      .setValue(
        "Desistiu"
      );


    /*
     * Novo sorteio sem adquirir outro lock.
     */
    const novoGanhador =
      sortearAmbienteInterno_(
        numero
      );


    atualizarResultados_();


    return novoGanhador;

  } finally {

    lock.releaseLock();
  }
}
