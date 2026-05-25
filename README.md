# Formulário de Justificativas — Guia de Configuração

## Arquivos do projeto

```
form-justificativa/
├── index.html                    ← Formulário web (publicado no GitHub Pages)
└── codigo-google-apps-script.js  ← Backend — cola no Apps Script da planilha
```

O arquivo de ausências (`.xlsx`) é gerado por um script Python externo e enviado para uma pasta do Google Drive. O Apps Script lê esse arquivo diretamente durante a consolidação.

---

## PASSO 1 — Preparar a planilha Google Sheets

1. Crie ou abra a planilha que vai receber as respostas
2. Renomeie a aba principal para **`Gestores`**
3. Certifique-se de que as colunas estejam nessa ordem:

| Coluna | Campo  |
|--------|--------|
| A      | Nome   |
| B      | CPF    |
| C      | Cargo  |
| D      | Loja   |

A linha 1 deve ser o cabeçalho. O CPF pode ter máscara ou só números.

---

## PASSO 2 — Instalar o Apps Script

1. Na planilha, clique em **Extensões → Apps Script**
2. Apague o código padrão e cole o conteúdo de `codigo-google-apps-script.js`
3. Salve (`Ctrl+S`)

### Ativar o serviço Drive API

Necessário para leitura de arquivos `.xlsx` do Drive:

1. No menu lateral esquerdo do editor, clique em **"+"** ao lado de **Serviços**
2. Selecione **Google Drive API** e clique em **Adicionar**
3. Abra o arquivo `appsscript.json` (ative em ⚙️ Configurações do projeto → "Mostrar arquivo de manifesto") e confirme que está assim:

```json
{
  "timeZone": "America/Sao_Paulo",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v2",
        "serviceId": "drive"
      }
    ]
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui"
  ],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  }
}
```

### Forçar autorização com Drive

Para garantir que os escopos do Drive sejam autorizados, crie temporariamente essa função, execute pelo editor (▶ Executar) e aceite todas as permissões:

```javascript
function forcarAutorizacao() {
  DriveApp.getRootFolder();
  Drive.Files.list({ maxResults: 1 });
}
```

Após autorizar, apague essa função.

---

## PASSO 3 — Implantar como App da Web

1. Clique em **Implantar → Nova implantação**
2. Tipo: **App da Web**
3. Configure:
   - Executar como: **Eu mesmo**
   - Quem tem acesso: **Qualquer pessoa**
4. Clique em **Implantar** e copie a URL gerada:
   ```
   https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXX/exec
   ```

> Sempre que alterar o script, crie uma **nova implantação** — não basta salvar.

---

## PASSO 4 — Configurar o formulário HTML

Abra `index.html` e localize a linha com `API_URL`:

```js
const API_URL = "SUA_URL_DO_APPS_SCRIPT_AQUI";
```

Substitua pela URL copiada no PASSO 3:

```js
const API_URL = "https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXX/exec";
```

---

## PASSO 5 — Publicar no GitHub Pages

1. Faça commit e push para o repositório
2. Nas configurações do repositório: **Settings → Pages**
3. Source: branch `main`, pasta `/ (root)` → **Save**
4. Compartilhe a URL gerada com os gestores

---

## PASSO 6 — Configurar a pasta de referência no Drive

O script lê o arquivo `.xlsx` de ausências de uma pasta do Google Drive.

1. Crie uma pasta no Google Drive para receber os arquivos gerados pelo Python
2. Copie o **ID da pasta** da URL:
   ```
   https://drive.google.com/drive/folders/  ID_DA_PASTA_AQUI
   ```
3. No Google Sheets, clique em **Justificativas → Configurar pasta de referência (Drive)**
4. Cole o ID e confirme — fica salvo permanentemente

> Essa configuração é feita uma única vez.

---

## Fluxo diário de uso

```
1. Script Python gera o .xlsx de ausências → salva na pasta do Drive
2. Gestores acessam o formulário e preencheram as justificativas
3. No Sheets: menu Justificativas → Consolidar respostas do dia
4. O script lê o .xlsx mais recente da pasta e gera a aba Consolidado_
```

### Estrutura esperada do arquivo `.xlsx`

| REGIONAL | LOJA | SETOR | FONTE | DATA | ID_LOJA | ID_SETOR | ID_FONTE |
|----------|------|-------|-------|------|---------|----------|----------|

### Consolidado gerado

- Um dia: `Consolidado_22-05-2026`
- Múltiplos dias (ex: fim de semana): `Consolidado_22-05-2026_a_24-05-2026`
- Linhas com resposta: exibidas normalmente
- Linhas sem resposta: marcadas em vermelho com `AUSENTE - sem justificativa`

---

## Estrutura das abas no Google Sheets

| Aba | Descrição |
|-----|-----------|
| `Gestores` | Cadastro de gestores autorizados (Nome, CPF, Cargo, Loja) |
| `Respostas_AAAA-MM-DD` | Respostas brutas por data — mantidas como backup |
| `Consolidado_DD-MM-AAAA` | Resultado final da consolidação |
