const ABA_GESTORES = "Gestores";
const COL_NOME = 1;
const COL_CPF  = 2;
const COL_LOJA = 3;

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
      const body = e.postData ? JSON.parse(e.postData.contents) : params;
      result = salvarResposta(body);
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
        ok:   true,
        nome: dados[i][COL_NOME - 1],
        cpf:  cpfPlanilha,
        loja: dados[i][COL_LOJA - 1]
      };
    }
  }

  return { ok: false, erro: "CPF não autorizado. Verifique o número digitado." };
}

function salvarResposta(dados) {
  const { cpf, nome, loja, data, registro, setor, justificativa, horario } = dados;

  if (!cpf || !data || !registro || !setor || !justificativa) {
    return { ok: false, erro: "Dados incompletos para salvar." };
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "Respostas_" + data; // ex: Respostas_2026-04-29
  let aba       = ss.getSheetByName(nomeAba);

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    
    aba.appendRow([
      "Data", "Horário", "CPF", "Nome", "Loja", "Registro", "Setor", "Justificativa"
    ]);
    const header = aba.getRange(1, 1, 1, 8);
    header.setFontWeight("bold");
    header.setBackground("#1e3a5f");
    header.setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  aba.appendRow([data, horario, cpf, nome, loja, registro, setor, justificativa]);

  return { ok: true, mensagem: "Resposta salva com sucesso." };
}
