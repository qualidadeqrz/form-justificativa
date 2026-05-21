const ABA_GESTORES = "Gestores";
const COL_NOME  = 1; // A = Nome
const COL_CPF   = 2; // B = CPF
const COL_CARGO = 3; // C = Cargo
const COL_LOJA  = 4; // D = Loja

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
// CONSOLIDAÇÃO
// Chame pelo menu "📋 Justificativas → Consolidar respostas"
// Pede a data no formato AAAA-MM-DD, monta o Consolidado_DATA
// garantindo que todas as lojas apareçam pelo menos uma vez.
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

  if (!abaResp) {
    ui.alert(`❌ Aba "${nomeAbaResp}" não encontrada.\n\nVerifique se a data está correta e se há respostas registradas.`);
    return;
  }

  // Lê respostas — pula cabeçalho
  // Colunas: Loja | Registro | Setor | Data Referente | Data da Resposta | Horário | Nome | Cargo | Justificativa
  const respostaDados = abaResp.getDataRange().getValues().slice(1);

  // Lê todas as lojas cadastradas (sem duplicatas, mantendo ordem da planilha)
  const abaGestores   = ss.getSheetByName(ABA_GESTORES);
  const gestoresDados = abaGestores.getDataRange().getValues().slice(1);

  const lojasVistas = new Set();
  const lojaOrdem   = [];
  gestoresDados.forEach(row => {
    const loja = String(row[COL_LOJA - 1]).trim();
    if (loja && !lojasVistas.has(loja)) {
      lojasVistas.add(loja);
      lojaOrdem.push(loja);
    }
  });

  // Lojas que responderam (para identificar ausentes)
  const lojasQueResponderam = new Set(respostaDados.map(r => String(r[0]).trim()));

  // Cria (ou recria) aba consolidada
  const nomeConsolidado = "Consolidado_" + data;
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

  // Para cada loja: insere respostas ou linha de ausência
  lojaOrdem.forEach(loja => {
    const linhasDaLoja = respostaDados.filter(r => String(r[0]).trim() === loja);

    if (linhasDaLoja.length > 0) {
      linhasDaLoja.forEach(linha => abaConsolidado.appendRow(linha));
    } else {
      // Linha de ausência — só a loja identificada, restante vazio
      abaConsolidado.appendRow([loja, "", "", "", "", "", "", "", "AUSENTE - sem justificativa"]);
    }
  });

  // Formata linhas ausentes em vermelho claro
  const todosValores = abaConsolidado.getDataRange().getValues();
  todosValores.forEach((row, i) => {
    if (i === 0) return;
    if (row[8] === "AUSENTE – sem justificativa") {
      abaConsolidado.getRange(i + 1, 1, 1, 9)
        .setBackground("#fde8e8")
        .setFontColor("#9b1c1c")
        .setFontWeight("bold");
    }
  });

  abaConsolidado.autoResizeColumns(1, 9);

  // Resumo final
  const total    = respostaDados.length;
  const ausentes = lojaOrdem.length - lojasQueResponderam.size;

  ui.alert(
    `✅ Consolidado_${data} criado!\n\n` +
    `• ${total} justificativa(s) registrada(s)\n` +
    `• ${ausentes} loja(s) sem resposta (marcadas em vermelho)\n\n` +
    `A aba Respostas_${data} foi mantida para backup.`
  );
}

// ------------------------------------------------------------
// Menu no Google Sheets — aparece automaticamente ao abrir
// ------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Justificativas")
    .addItem("Consolidar respostas do dia", "consolidar")
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
