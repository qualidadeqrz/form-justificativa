const ABA_GESTORES = "Gestores";
const COL_NOME  = 1;
const COL_CPF   = 2;
const COL_CARGO = 3;
const COL_LOJA  = 4;

const PROP_PASTA_ID = "PASTA_REFERENCIA_ID";

// IDs fixos — devem ser idênticos aos do planilha_ausencias.py
// Usados apenas como fallback; o cruzamento principal usa os IDs lidos do Excel de referência.
const FONTE_ID = { 'Avarias': 1, 'Transferência de Produção': 2, 'Relatórios - Moki': 3 };
const SETOR_ID = { 'Açougue': 1, 'Frios': 2, 'Hortifruti': 3, 'Padaria': 4, 'Refeição': 5, 'Gestores': 6 };

// Busca case-insensitive e sem acento para evitar falhas de codificação
function _normStr(s) {
  return String(s).normalize("NFC").trim().toLowerCase();
}
function _lookupId(dict, key) {
  const k = _normStr(key);
  for (const [dk, dv] of Object.entries(dict)) {
    if (_normStr(dk) === k) return dv;
  }
  return 0;
}

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action;
  const headers = { "Access-Control-Allow-Origin": "*", "Content-Type": "application/json" };

  try {
    let result;
    if      (action === "validarCPF")      result = validarCPF(params.cpf);
    else if (action === "salvarResposta")  result = salvarResposta(params);
    else if (action === "buscarRegistros") result = buscarRegistros(params.cpf, params.data);
    else                                   result = { ok: false, erro: "Ação inválida." };

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: err.message })).setMimeType(ContentService.MimeType.JSON);
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

  // Extrai ID numérico da loja — ex: "09 - HQ ALTO DA CONCEIÇÃO" → 9
  const lojaIdMatch = String(loja).match(/^(\d+)/);
  const id_loja  = lojaIdMatch ? parseInt(lojaIdMatch[1]) : 0;
  const id_fonte = _lookupId(FONTE_ID, registro);
  const id_setor = _lookupId(SETOR_ID, setor);

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "Respostas_" + data_referente;
  let aba       = ss.getSheetByName(nomeAba);

  if (!aba) {
    aba = ss.insertSheet(nomeAba);
    aba.appendRow(["Loja", "Registro", "Setor", "Data Referente", "Data da Resposta", "Horário da Resposta", "Nome", "Cargo", "Justificativa", "ID_LOJA", "ID_FONTE", "ID_SETOR"]);
    const header = aba.getRange(1, 1, 1, 12);
    header.setFontWeight("bold");
    header.setBackground("#1e3a5f");
    header.setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  const [a, m, d] = data_referente.split("-");
  const dataRefFormatada = `${d}/${m}/${a}`;

  aba.appendRow([loja, registro, setor, dataRefFormatada, data_resposta, horario, nome, cargo || "", justificativa, id_loja, id_fonte, id_setor]);

  return { ok: true, mensagem: "Resposta salva com sucesso." };
}

function buscarRegistros(cpf, data) {
  if (!cpf || !data) return { ok: true, registros: [] };

  const gestor = validarCPF(cpf);
  if (!gestor.ok) return { ok: true, registros: [] };

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Respostas_" + data);
  if (!aba) return { ok: true, registros: [] };

  // Colunas: Loja(0) | Registro(1) | Setor(2) | ... | Nome(6) | ... | ID_LOJA(9) | ID_FONTE(10) | ID_SETOR(11)
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
// Lê o Excel de referência do Drive
// Estrutura: REGIONAL(A) | LOJA(B) | SETOR(C) | FONTE(D) | DATA(E) | ID_LOJA(F) | ID_SETOR(G) | ID_FONTE(H)
// ------------------------------------------------------------
function lerReferenciaDrive() {
  const pastaId = PropertiesService.getScriptProperties().getProperty(PROP_PASTA_ID);
  if (!pastaId) throw new Error('Script Property "PASTA_REFERENCIA_ID" não configurada.');

  const pasta    = DriveApp.getFolderById(pastaId);
  const arquivos = pasta.getFiles();
  if (!arquivos.hasNext()) throw new Error("Nenhum arquivo encontrado na pasta do Drive configurada.");

  let alvo = arquivos.next();
  while (arquivos.hasNext()) {
    const proximo = arquivos.next();
    if (proximo.getLastUpdated() > alvo.getLastUpdated()) alvo = proximo;
  }

  const mimeXlsx   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const mimeSheets = "application/vnd.google-apps.spreadsheet";
  let   ssId   = alvo.getId();
  let   tempId = null;

  if (alvo.getMimeType() === mimeXlsx) {
    const copia = Drive.Files.copy({ title: "_temp_ausencias", mimeType: mimeSheets }, ssId);
    tempId = copia.id;
    ssId   = tempId;
  }

  const ss    = SpreadsheetApp.openById(ssId);
  const aba   = ss.getSheets()[0];
  const dados = aba.getDataRange().getValues().slice(1);

  const pares  = [];
  const vistos = new Set();
  let dataReferencia = "";

  dados.forEach(row => {
    const loja     = String(row[1]).trim(); // B
    const setor    = String(row[2]).trim(); // C
    const fonte    = String(row[3]).trim(); // D
    const data     = row[4];                // E
    const id_loja  = Number(row[5]);        // F
    const id_setor = Number(row[6]);        // G
    const id_fonte = Number(row[7]);        // H

    if (!loja || !setor || !fonte) return;

    const chave = `${id_loja}||${id_setor}||${id_fonte}`;
    if (!vistos.has(chave)) {
      vistos.add(chave);
      pares.push({ loja, setor, fonte, id_loja, id_setor, id_fonte });
    }
    if (!dataReferencia && data) {
      dataReferencia = data instanceof Date
        ? Utilities.formatDate(data, Session.getScriptTimeZone(), "dd-MM-yyyy")
        : String(data).replace(/\//g, "-");
    }
  });

  if (tempId) Drive.Files.remove(tempId);

  return { pares, nomeArquivo: alvo.getName(), dataReferencia };
}

// ------------------------------------------------------------
// Consolidação — compara por ID, não por texto
// ------------------------------------------------------------
function consolidar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abas    = ss.getSheets().map(a => a.getName());
  const recente = abas.filter(n => n.startsWith("Respostas_")).sort().reverse()[0];
  if (!recente) { ui.alert("❌ Nenhuma aba de Respostas encontrada."); return; }

  const data    = recente.replace("Respostas_", "");
  const abaResp = ss.getSheetByName(recente);

  let referencia, nomeArquivo, dataConsolidado;
  try {
    ({ pares: referencia, nomeArquivo, dataReferencia: dataConsolidado } = lerReferenciaDrive());
  } catch (err) {
    ui.alert("❌ Erro ao ler referência do Drive:\n\n" + err.message);
    return;
  }
  if (!dataConsolidado) dataConsolidado = data;
  if (referencia.length === 0) { ui.alert("❌ O arquivo de referência não contém nenhum par válido."); return; }

  // Colunas respostas: Loja(0) Registro(1) Setor(2) DataRef(3) DataResp(4) Horario(5) Nome(6) Cargo(7) Just(8) ID_LOJA(9) ID_FONTE(10) ID_SETOR(11)
  const respostaDados = abaResp.getDataRange().getValues().slice(1);

  const nomeConsolidado = "Consolidado_" + dataConsolidado;
  const abaExistente    = ss.getSheetByName(nomeConsolidado);
  if (abaExistente) ss.deleteSheet(abaExistente);

  const abaConsolidado = ss.insertSheet(nomeConsolidado);
  abaConsolidado.appendRow(["Loja", "Registro", "Setor", "Data Referente", "Data da Resposta", "Horário da Resposta", "Nome", "Cargo", "Justificativa"]);
  const cabecalho = abaConsolidado.getRange(1, 1, 1, 9);
  cabecalho.setFontWeight("bold");
  cabecalho.setBackground("#1e3a5f");
  cabecalho.setFontColor("#ffffff");
  abaConsolidado.setFrozenRows(1);

  let totalRespondidos = 0;
  let totalAusentes    = 0;

  referencia.forEach(({ loja, setor, fonte, id_loja, id_setor, id_fonte }) => {
    const linhas = respostaDados.filter(r => {
      const rId9  = r[9];
      // Linha nova (tem IDs): compara por ID — imune a acentos e capitalização
      if (rId9 !== "" && rId9 != null && !isNaN(Number(rId9))) {
        return Number(r[9])  === id_loja  &&
               Number(r[10]) === id_fonte &&
               Number(r[11]) === id_setor;
      }
      // Linha antiga (sem IDs): fallback por string normalizada
      return _normStr(r[0]) === _normStr(loja)  &&
             _normStr(r[2]) === _normStr(setor) &&
             _normStr(r[1]) === _normStr(fonte);
    });

    if (linhas.length > 0) {
      linhas.forEach(linha => abaConsolidado.appendRow(linha.slice(0, 9)));
      totalRespondidos += linhas.length;
    } else {
      abaConsolidado.appendRow([loja, fonte, setor, dataConsolidado.replace(/-/g, "/"), "", "", "", "", "AUSENTE - sem justificativa"]);
      totalAusentes++;
    }
  });

  const todosValores = abaConsolidado.getDataRange().getValues();
  todosValores.forEach((row, i) => {
    if (i === 0) return;
    if (row[8] === "AUSENTE - sem justificativa") {
      abaConsolidado.getRange(i + 1, 1, 1, 9).setBackground("#fde8e8").setFontColor("#9b1c1c").setFontWeight("bold");
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

function configurarPastaReferencia() {
  const ui       = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    "Configurar pasta de referência",
    "Cole o ID da pasta do Google Drive que contém os Excel diários:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resposta.getSelectedButton() !== ui.Button.OK) return;
  const id = resposta.getResponseText().trim();
  if (!id) { ui.alert("❌ ID não informado."); return; }
  PropertiesService.getScriptProperties().setProperty(PROP_PASTA_ID, id);
  ui.alert(`✅ Pasta configurada com sucesso!\n\nID salvo: ${id}`);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔰 Justificativas")
    .addItem("Consolidar respostas do dia", "consolidar")
    .addSeparator()
    .addItem("Configurar pasta de referência (Drive)", "configurarPastaReferencia")
    .addToUi();
}
