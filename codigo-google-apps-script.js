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

  const cpf = cpfRaw.replace(/\D/g, ""); // remove pontos e traço
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
  // Aba nomeada pela data referente (ex: Respostas_2026-04-30)
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
