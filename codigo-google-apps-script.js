const ABA_GESTORES = "Gestores";
const COL_NOME    = 1;
const COL_CPF     = 2;
const COL_CARGO   = 3;
const COL_LOJA    = 4;
const COL_ID_LOJA = 5;

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
// Converte "dd/MM/yyyy" ou "dd-MM-yyyy" → "yyyy-MM-dd" (ISO, nome da aba)
function _dataParaISO(s) {
  const str = String(s).trim();
  const m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  return "";
}
// getValues() pode retornar Date objects quando a célula contém uma data
function _cellToISO(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return _dataParaISO(String(val));
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
    else if (action === "buscarRegistros") result = buscarRegistros(params.cpf, params.data, params.periodo);
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
        ok:      true,
        nome:    dados[i][COL_NOME    - 1],
        cpf:     cpfPlanilha,
        cargo:   dados[i][COL_CARGO   - 1],
        loja:    dados[i][COL_LOJA    - 1],
        id_loja: Number(dados[i][COL_ID_LOJA - 1]) || 0
      };
    }
  }
  return { ok: false, erro: "CPF não autorizado. Verifique o número digitado." };
}

function salvarResposta(dados) {
  const { cpf, nome, cargo, loja, data_referente, data_resposta, horario, registro, setor, justificativa, periodo } = dados;

  if (!cpf || !data_referente || !registro || !setor || !justificativa) {
    return { ok: false, erro: "Dados incompletos para salvar." };
  }

  // Extrai ID numérico da loja — ex: "09 - HQ ALTO DA CONCEIÇÃO" → 9
  const lojaIdMatch = String(loja).match(/^(\d+)/);
  const id_loja  = lojaIdMatch ? parseInt(lojaIdMatch[1]) : 0;
  const id_fonte = _lookupId(FONTE_ID, registro);
  const id_setor = _lookupId(SETOR_ID, setor);

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  // Quando há período (ex: "2026-05-22_a_2026-05-24"), todas as datas vão para a mesma aba
  const nomeAba = "Respostas_" + (periodo && periodo !== data_referente ? periodo : data_referente);
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

function buscarRegistros(cpf, data, periodo) {
  if (!cpf || !data) return { ok: true, registros: [] };

  const gestor = validarCPF(cpf);
  if (!gestor.ok) return { ok: true, registros: [] };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Busca na aba do período; se não existir, tenta a aba da data específica
  let aba = periodo ? ss.getSheetByName("Respostas_" + periodo) : null;
  if (!aba) aba = ss.getSheetByName("Respostas_" + data);
  if (!aba) return { ok: true, registros: [] };

  // Colunas: Loja(0) | Registro(1) | Setor(2) | DataRef(3) | ... | ID_LOJA(9) | ID_FONTE(10) | ID_SETOR(11)
  const registros = aba.getDataRange().getValues().slice(1)
    .filter(row => Number(row[9]) === gestor.id_loja)
    .map(row => ({
      registro:        String(row[1]),
      setor:           String(row[2]),
      data_referente:  row[3] instanceof Date
        ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "dd/MM/yyyy")
        : String(row[3])
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

  const pares    = [];
  const vistos   = new Set();
  const datasISO = new Set(); // datas únicas em formato yyyy-MM-dd (nomes das abas)
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

    // Converte DATA (col E) para os dois formatos necessários
    let dataRefFmt = ""; // dd-MM-yyyy (exibição e nome do consolidado)
    let dataRefISO = ""; // yyyy-MM-dd (nome da aba Respostas_)
    if (data) {
      dataRefFmt = data instanceof Date
        ? Utilities.formatDate(data, Session.getScriptTimeZone(), "dd-MM-yyyy")
        : String(data).replace(/\//g, "-");
      dataRefISO = _dataParaISO(dataRefFmt);
      if (dataRefISO) datasISO.add(dataRefISO);
      if (!dataReferencia) dataReferencia = dataRefFmt;
    }

    // Chave inclui a data para suportar Excel com múltiplas datas
    const chave = `${id_loja}||${id_setor}||${id_fonte}||${dataRefISO}`;
    if (!vistos.has(chave)) {
      vistos.add(chave);
      pares.push({ loja, setor, fonte, id_loja, id_setor, id_fonte, dataRefISO, dataRefFmt });
    }
  });

  if (tempId) Drive.Files.remove(tempId);

  return { pares, nomeArquivo: alvo.getName(), dataReferencia, datasISO: Array.from(datasISO) };
}

// ------------------------------------------------------------
// Consolidação — compara por ID, não por texto
// ------------------------------------------------------------
function consolidar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let referencia, nomeArquivo, dataConsolidado, datasISO;
  try {
    ({ pares: referencia, nomeArquivo, dataReferencia: dataConsolidado, datasISO } = lerReferenciaDrive());
  } catch (err) {
    ui.alert("❌ Erro ao ler referência do Drive:\n\n" + err.message);
    return;
  }
  if (referencia.length === 0) { ui.alert("❌ O arquivo de referência não contém nenhum par válido."); return; }

  // Coleta respostas de TODAS as abas Respostas_ presentes na planilha
  const respostaDados   = [];
  const abasEncontradas = [];

  const abasResp = ss.getSheets().map(a => a.getName()).filter(n => n.startsWith("Respostas_"));
  if (abasResp.length === 0) {
    ui.alert("❌ Nenhuma aba de Respostas encontrada.");
    return;
  }
  abasResp.forEach(nomeAba => {
    ss.getSheetByName(nomeAba).getDataRange().getValues().slice(1).forEach(r => respostaDados.push(r));
    abasEncontradas.push(nomeAba);
  });

  if (!dataConsolidado) dataConsolidado = datasISO[0]
    ? (() => { const [a,m,d] = datasISO[0].split("-"); return `${d}-${m}-${a}`; })()
    : "sem-data";

  // Quando o período tem múltiplas datas, inclui "inicio_a_fim" no título — igual ao nome da aba de Respostas
  let nomeConsolidado;
  if (datasISO.length > 1) {
    const sorted = [...datasISO].sort();
    const [a1, m1, d1] = sorted[0].split("-");
    const [a2, m2, d2] = sorted[sorted.length - 1].split("-");
    nomeConsolidado = `Consolidado_${d1}-${m1}-${a1}_a_${d2}-${m2}-${a2}`;
  } else {
    nomeConsolidado = "Consolidado_" + dataConsolidado;
  }
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

  referencia.forEach(({ loja, setor, fonte, id_loja, id_setor, id_fonte, dataRefISO, dataRefFmt }) => {
    const linhas = respostaDados.filter(r => {
      // Filtra pela data de referência (coluna D da resposta = "dd/MM/yyyy")
      if (dataRefISO) {
        if (_cellToISO(r[3]) !== dataRefISO) return false;
      }
      // Linha nova (tem IDs): compara por ID
      const rId9 = r[9];
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

    const dataDisplay = dataRefFmt ? dataRefFmt.replace(/-/g, "/") : dataConsolidado.replace(/-/g, "/");

    if (linhas.length > 0) {
      linhas.forEach(linha => abaConsolidado.appendRow(linha.slice(0, 9)));
      totalRespondidos += linhas.length;
    } else {
      abaConsolidado.appendRow([loja, fonte, setor, dataDisplay, "", "", "", "", "AUSENTE - sem justificativa"]);
      totalAusentes++;
    }
  });

  abaConsolidado.autoResizeColumns(1, 9);

  const infoAbas = `• Abas lidas: ${abasEncontradas.join(", ")}\n`;

  ui.alert(
    `✅ ${nomeConsolidado} criado!\n\n` +
    `• Referência: ${nomeArquivo}\n` +
    infoAbas +
    `• ${referencia.length} par(es) LOJA+SETOR+FONTE+DATA na referência\n` +
    `• ${totalRespondidos} justificativa(s) registrada(s)\n` +
    `• ${totalAusentes} par(es) sem resposta (marcados em vermelho)\n\n` +
    `As abas de Respostas foram mantidas para backup.`
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
