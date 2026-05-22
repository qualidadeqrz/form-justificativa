const ABA_GESTORES = "Gestores";
const COL_NOME  = 1; // A = Nome
const COL_CPF   = 2; // B = CPF
const COL_CARGO = 3; // C = Cargo
const COL_LOJA  = 4; // D = Loja

// Script Property que guarda o ID da pasta do Drive com os Excel de referência.
// Configure via menu "Justificativas → Configurar pasta de referência".
const PROP_PASTA_ID = "PASTA_REFERENCIA_ID";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action;

  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Content-Type": "application/json"
  };

  try {
    let result;

    if (action === "validarCPF") {
      result = validarCPF(params.cpf);
    } else if (action === "salvarResposta") {
      result = salvarResposta(params);
    } else if (action === "buscarRegistros") {
      result = buscarRegistros(params.cpf, params.data);
    } else {
      result = { ok: false, erro: "Ação inválida." };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function validarCPF(cpfRaw) {
  if (!cpfRaw) return { ok: false, erro: "CPF não informado." };

  const cpf = cpfRaw.replace(/\D/g, "");
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_GESTORES);

  if (!aba) return { ok: false, erro: `Aba "${ABA_GESTORES}" não encontrada.` };

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const cpfPlanilha = String(dados[i][COL_CPF - 1]).replace(/\D/g, "");
    if (cpfPlanilha === cpf) {
      return {
        ok:    true,
        nome:  dados[i][COL_NOME  - 1],
        cpf:   cpfPlanilha,
        cargo: dados[i][COL_CARGO - 1],
        loja:  dados[i][COL_LOJA  - 1]
      };
    }
  }

  return { ok: false, erro: "CPF não autorizado. Verifique o número digitado." };
}

function salvarResposta(dados) {
  const { cpf, nome, cargo, loja, data_referente, data_resposta, horario, registro, setor, justificativa } = dados;

  if (!cpf || !data_referente || !registro || !setor || !justificativa) {
    return { ok: false, erro: "Dados incompletos para salvar." };
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "Respostas_" + data_referente;
  let aba       = ss.getSheetByName(nomeAba);

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    aba.appendRow([
      "Loja", "Registro", "Setor", "Data Referente", "Data da Resposta", "Horário da Resposta", "Nome", "Cargo", "Justificativa"
    ]);
    const header = aba.getRange(1, 1, 1, 9);
    header.setFontWeight("bold");
    header.setBackground("#1e3a5f");
    header.setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  const [a, m, d] = data_referente.split("-");
  const dataRefFormatada = `${d}/${m}/${a}`;

  aba.appendRow([loja, registro, setor, dataRefFormatada, data_resposta, horario, nome, cargo || "", justificativa]);

  return { ok: true, mensagem: "Resposta salva com sucesso." };
}

function buscarRegistros(cpf, data) {
  if (!cpf || !data) return { ok: true, registros: [] };

  // Identifica gestor pelo CPF para obter nome e loja (não estão salvos nas Respostas)
  const gestor = validarCPF(cpf);
  if (!gestor.ok) return { ok: true, registros: [] };

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const aba     = ss.getSheetByName("Respostas_" + data);
  if (!aba)     return { ok: true, registros: [] };

  // Colunas: Loja(0) | Registro(1) | Setor(2) | Data Ref(3) | Data Resp(4) | Horário(5) | Nome(6) | Cargo(7) | Justificativa(8)
  const registros = aba.getDataRange().getValues().slice(1)
    .filter(row => String(row[0]).trim() === gestor.loja && String(row[6]).trim() === gestor.nome)
    .map(row => ({
      registro:      String(row[1]),
      setor:         String(row[2]),
      justificativa: String(row[8])
    }));

  return { ok: true, registros };
}

// ------------------------------------------------------------
// Lê o Excel de referência do Drive e retorna pares { loja, setor }
// Espera exatamente um arquivo na pasta configurada em PASTA_REFERENCIA_ID.
// Estrutura esperada: REGIONAL(A) | LOJA(B) | SETOR(C) | FONTE(D) | DATA(E)
// ------------------------------------------------------------
function lerReferenciaDrive() {
  const pastaId = PropertiesService.getScriptProperties().getProperty(PROP_PASTA_ID);
  if (!pastaId) {
    throw new Error(
      'Script Property "PASTA_REFERENCIA_ID" não configurada.\n' +
      'Use o menu Justificativas → Configurar pasta de referência.'
    );
  }

  const pasta    = DriveApp.getFolderById(pastaId);
  const arquivos = pasta.getFiles();

  if (!arquivos.hasNext()) {
    throw new Error("Nenhum arquivo encontrado na pasta do Drive configurada.\nVerifique se o Excel foi enviado para a pasta correta.");
  }

  let alvo = arquivos.next();
  while (arquivos.hasNext()) {
    const proximo = arquivos.next();
    if (proximo.getLastUpdated() > alvo.getLastUpdated()) alvo = proximo;
  }

  const ss    = SpreadsheetApp.openById(alvo.getId());
  const aba   = ss.getSheets()[0];
  const dados = aba.getDataRange().getValues().slice(1); // pula cabeçalho

  const pares  = [];
  const vistos = new Set();

  let dataReferencia = "";
  dados.forEach(row => {
    const loja  = String(row[1]).trim(); // coluna B
    const setor = String(row[2]).trim(); // coluna C
    const fonte = String(row[3]).trim(); // coluna D
    const data  = row[4];               // coluna E — DATA
    if (!loja || !setor || !fonte) return;
    const chave = `${loja}||${setor}||${fonte}`;
    if (!vistos.has(chave)) {
      vistos.add(chave);
      pares.push({ loja, setor, fonte });
    }
    if (!dataReferencia && data) {
      dataReferencia = data instanceof Date
        ? Utilities.formatDate(data, Session.getScriptTimeZone(), "dd-MM-yyyy")
        : String(data).replace(/\//g, "-");
    }
  });

  return { pares, nomeArquivo: alvo.getName(), dataReferencia };
}

// ------------------------------------------------------------
// CONSOLIDAÇÃO
// Chame pelo menu "Justificativas → Consolidar respostas"
// Lê o Excel de referência do Drive para saber quais pares
// LOJA+SETOR devem aparecer no consolidado.
// A aba Respostas_DATA é mantida para backup — apague manualmente.
// ------------------------------------------------------------
function consolidar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Busca automaticamente a aba Respostas_ mais recente
  const abas    = ss.getSheets().map(a => a.getName());
  const recente = abas.filter(n => n.startsWith("Respostas_")).sort().reverse()[0];

  if (!recente) {
    ui.alert("❌ Nenhuma aba de Respostas encontrada.");
    return;
  }

  const data    = recente.replace("Respostas_", "");
  const abaResp = ss.getSheetByName(recente);

  // Lê os pares LOJA+SETOR+FONTE e a data de referência do Excel no Drive
  let referencia, nomeArquivo, dataConsolidado;
  try {
    ({ pares: referencia, nomeArquivo, dataReferencia: dataConsolidado } = lerReferenciaDrive());
  } catch (err) {
    ui.alert("❌ Erro ao ler referência do Drive:\n\n" + err.message);
    return;
  }
  if (!dataConsolidado) dataConsolidado = data;

  if (referencia.length === 0) {
    ui.alert("❌ O arquivo de referência não contém nenhum par LOJA+SETOR válido.");
    return;
  }

  // Lê respostas — pula cabeçalho
  // Colunas: Loja | Registro | Setor | Data Referente | Data da Resposta | Horário | Nome | Cargo | Justificativa
  const respostaDados = abaResp.getDataRange().getValues().slice(1);

  // Cria (ou recria) aba consolidada
  const nomeConsolidado = "Consolidado_" + dataConsolidado;
  const abaExistente    = ss.getSheetByName(nomeConsolidado);
  if (abaExistente) ss.deleteSheet(abaExistente);

  const abaConsolidado = ss.insertSheet(nomeConsolidado);

  // Cabeçalho
  abaConsolidado.appendRow([
    "Loja", "Registro", "Setor", "Data Referente", "Data da Resposta", "Horário da Resposta", "Nome", "Cargo", "Justificativa"
  ]);
  const cabecalho = abaConsolidado.getRange(1, 1, 1, 9);
  cabecalho.setFontWeight("bold");
  cabecalho.setBackground("#1e3a5f");
  cabecalho.setFontColor("#ffffff");
  abaConsolidado.setFrozenRows(1);

  // Para cada par LOJA+SETOR da referência: insere respostas ou linha de ausência
  let totalRespondidos = 0;
  let totalAusentes    = 0;

  referencia.forEach(({ loja, setor, fonte }) => {
    const linhas = respostaDados.filter(r =>
      String(r[0]).trim() === loja &&
      String(r[2]).trim().toLowerCase() === setor.toLowerCase() &&
      String(r[1]).trim().toLowerCase() === fonte.toLowerCase()
    );

    if (linhas.length > 0) {
      linhas.forEach(linha => abaConsolidado.appendRow(linha));
      totalRespondidos += linhas.length;
    } else {
      abaConsolidado.appendRow([loja, fonte, setor, dataConsolidado.replace(/-/g, "/"), "", "", "", "", "AUSENTE - sem justificativa"]);
      totalAusentes++;
    }
  });

  // Formata linhas ausentes em vermelho claro
  const todosValores = abaConsolidado.getDataRange().getValues();
  todosValores.forEach((row, i) => {
    if (i === 0) return;
    if (row[8] === "AUSENTE - sem justificativa") {
      abaConsolidado.getRange(i + 1, 1, 1, 9)
        .setBackground("#fde8e8")
        .setFontColor("#9b1c1c")
        .setFontWeight("bold");
    }
  });

  abaConsolidado.autoResizeColumns(1, 9);

  ui.alert(
    `✅ Consolidado_${dataConsolidado} criado!\n\n` +
    `• Referência: ${nomeArquivo}\n` +
    `• ${referencia.length} par(es) LOJA+SETOR+FONTE na referência\n` +
    `• ${totalRespondidos} justificativa(s) registrada(s)\n` +
    `• ${totalAusentes} par(es) sem resposta (marcados em vermelho)\n\n` +
    `A aba Respostas_${data} foi mantida para backup.`
  );
}

// ------------------------------------------------------------
// Configura o ID da pasta do Drive com os Excel de referência.
// Execute uma única vez pelo menu ou manualmente no editor.
// ------------------------------------------------------------
function configurarPastaReferencia() {
  const ui       = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    "Configurar pasta de referência",
    "Cole o ID da pasta do Google Drive que contém os Excel diários:\n" +
    "(ex: 1ABCxyz_exemplo_de_id_de_pasta)",
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) return;

  const id = resposta.getResponseText().trim();
  if (!id) {
    ui.alert("❌ ID não informado.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty(PROP_PASTA_ID, id);
  ui.alert(`✅ Pasta configurada com sucesso!\n\nID salvo: ${id}`);
}

// ------------------------------------------------------------
// Menu no Google Sheets — aparece automaticamente ao abrir
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎰 Justificativas")
    .addItem("Consolidar respostas do dia", "consolidar")
    .addSeparator()
    .addItem("Configurar pasta de referência (Drive)", "configurarPastaReferencia")
    .addToUi();
}

// ------------------------------------------------------------
// Teste manual
// ------------------------------------------------------------
function testeManual() {
  const resultado = salvarResposta({
    cpf:            "11122233344",
    nome:           "TESTE",
    cargo:          "GERENTE DE LOJA",
    loja:           "HQ MOSSORÓ",
    data_referente: "2026-04-30",
    data_resposta:  "05/05/2026",
    horario:        "10:00:00",
    registro:       "Avarias",
    setor:          "Padaria",
    justificativa:  "Teste manual direto pelo script"
  });
  console.log(JSON.stringify(resultado));
}
