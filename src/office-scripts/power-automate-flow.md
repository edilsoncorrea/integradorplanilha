Power Automate — exemplo de fluxo para integrar Office Scripts e API

Resumo do fluxo (alto nível)
1) Executar script `Autenticacao` (action: `buildAuthPayload`) no workbook -> retorna { url, method, payload }
2) Fazer HTTP POST para url retornada (token endpoint). Extrair `access_token` do JSON de resposta.
3) Executar script `Documento` (action: `buildDocumentoPayloads`) -> retorna lista `payloads` com { sheetRow, payload }
4) Para cada payload na lista: fazer HTTP POST para `HOST + /api/documentos` com header Authorization: Bearer <token>. Coletar resultado (identificador ou erro).
5) Montar array `results` com objetos { sheetRow, notaCriada: 'Sim'|'Não', retorno: string }
6) Executar script `Documento` (action: `applyDocumentoResults`) com inputs.results = results

Detalhes das actions

A) Run script: Autenticacao
- Script: `Autenticacao`
- Inputs (JSON):
  {
    "action": "buildAuthPayload",
    "host": "https://seu-host",
    "username": "bimerapi",
    "senha": "SENHA_AQUI",
    "nonce": "123456789"
  }
- Saída esperada:
  { "url": "https://.../oauth/token", "method": "POST", "payload": { client_id:..., username:..., password:..., grant_type:..., nonce:... } }

B) HTTP (token)
- Method: POST
- URL: use a saída `url` do passo anterior
- Headers: Content-Type: application/json (ou application/x-www-form-urlencoded se a API exigir)
- Body: use `payload` do passo anterior (se precisar form-urlencoded, transforme o JSON)
- Parsear resposta JSON e armazenar `access_token` (ex: body('HTTP')['access_token'])

C) Run script: Documento (build)
- Script: `Documento`
- Inputs:
  { "action": "buildDocumentoPayloads" }
- Saída: { "payloads": [ { "sheetRow": 3, "payload": {...} }, ... ] }

D) Apply to each (payloads)
- Para cada item em payloads:
  1) HTTP POST para: HOST + '/api/documentos'
     - Headers:
       - Authorization: Bearer <access_token>
       - Content-Type: application/json
     - Body: item.payload (JSON)
  2) Parsear a resposta e gerar um objeto de resultado para a planilha, ex:
     {
       "sheetRow": item.sheetRow,
       "notaCriada": (res.ListaObjetos && res.ListaObjetos.length>0) ? 'Sim' : 'Não',
       "retorno": (res.ListaObjetos && res.ListaObjetos.length>0) ? res.ListaObjetos[0].Identificador : (res.Message || (res.Erros && res.Erros[0] && res.Erros[0].ErrorMessage) || 'Erro desconhecido')
     }
  3) Append to array variable `results`

E) Run script: Documento (apply)
- Script: `Documento`
- Inputs:
  { "action": "applyDocumentoResults", "results": <the results array> }
- Saída: { ok: true, updated: N }

Notas de segurança e operações
- Armazenamento de credenciais: preferir usar Azure Key Vault ou os conectores de autenticação do Power Automate para não colocar senhas nos scripts.
- Retentativas/erros: adicione controle de erros no Flow para repetir chamadas falhas ou enviar notificações (Teams/email).
- Execução: salve os Office Scripts no workbook (Excel Online) e aponte o Run script action para o arquivo correto.

Exemplo simplificado de transformação para x-www-form-urlencoded (caso a API não aceite JSON para token):
- Se payload = { username: 'bimer', password: 'abc' }
- Body (raw): username=bimer&password=abc&client_id=IntegracaoBimer.js&grant_type=password&nonce=...

Esta documentação serve como guia para implementar o Flow que usará os Office Scripts gerados nesta pasta. Se quiser, eu posso gerar um export JSON do Flow de exemplo ou um passo-a-passo com capturas de tela (manual) — diga a opção desejada.
