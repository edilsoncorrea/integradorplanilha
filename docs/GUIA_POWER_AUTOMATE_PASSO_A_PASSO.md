# 🚀 Guia Completo: Configurar Power Automate com Office Scripts

## 📋 Pré-requisitos

- ✅ Excel Online (Microsoft 365) com a planilha salva no OneDrive ou SharePoint
- ✅ Office Scripts já publicado no Excel (o arquivo `IntegradorCompleto.ts`)
- ✅ Conta Power Automate (incluída no Microsoft 365)
- ✅ Acesso à API Bimer (credenciais)

---

## 🎯 Visão Geral do Fluxo

```
┌─────────────────────────────────────────────────────────────┐
│  FLUXO POWER AUTOMATE - CRIAR DOCUMENTOS BIMER             │
└─────────────────────────────────────────────────────────────┘

1️⃣ [Botão Manual] → Iniciar fluxo

2️⃣ [Office Script] → action='buildAuthPayload'
   └─> Retorna: { url, payload }

3️⃣ [HTTP POST] → Autenticar na API Bimer
   └─> Retorna: { access_token }

4️⃣ [Office Script] → action='buildValidationQueries' (opcional)
   └─> Retorna: { queries[] }

5️⃣ [Para cada query] → HTTP GET buscar identificadores
   └─> Preencher IDs faltantes

6️⃣ [Office Script] → action='applyValidationResults'

7️⃣ [Office Script] → action='buildDocumentos'
   └─> Retorna: { payloads[] }

8️⃣ [Para cada payload] → HTTP POST criar documento
   └─> Coletar respostas

9️⃣ [Office Script] → action='applyResults'
   └─> Escrever resultados na planilha
```

---

## 📝 PASSO 1: Publicar o Office Script no Excel

### 1.1. Abrir Excel Online
1. Acesse **OneDrive** ou **SharePoint**
2. Abra sua planilha Excel Online
3. Vá na aba **Automatizar** → **Office Scripts**

### 1.2. Criar novo Script
1. Clique em **Novo Script**
2. Delete o código de exemplo
3. Cole todo o conteúdo do arquivo `IntegradorCompleto.ts`
4. Clique em **Salvar Script**
5. Nomeie como: **"Integrador Completo"**

✅ **Pronto!** O script está disponível para o Power Automate usar.

---

## 🔧 PASSO 2: Criar o Flow no Power Automate

### 2.1. Acessar Power Automate
1. Acesse: https://make.powerautomate.com
2. Entre com sua conta Microsoft 365
3. No menu lateral, clique em **Criar**
4. Escolha **Fluxo de nuvem instantâneo**

### 2.2. Configurar Trigger
1. Nome do fluxo: **"Criar Documentos Bimer - Excel"**
2. Escolha o gatilho: **"Acionar um fluxo manualmente"**
3. Clique em **Criar**

---

## 🎬 PASSO 3: Adicionar Ações ao Flow

### ⚙️ Ação 1: Inicializar Variável (Array de Resultados)

**Clique em "Nova etapa"** → Pesquise: **"Inicializar variável"**

```
Nome: ResultadosArray
Tipo: Matriz
Valor: [] (deixe vazio)
```

---

### 📊 Ação 2: Executar Script - Build Auth Payload

**Nova etapa** → Pesquise: **"Excel Online (Business)"** → **"Executar script"**

**Configuração:**
```
Local: OneDrive (ou SharePoint)
Biblioteca de Documentos: OneDrive
Arquivo: [Navegue até sua planilha Excel]
Script: Integrador Completo
```

**Clique em "Mostrar opções avançadas"** e adicione os inputs:

```json
{
  "action": "buildAuthPayload",
  "host": "https://homologacaowisepcp.alterdata.com.br/BimerApi",
  "username": "supervisor",
  "senha": "Senhas123",
  "nonce": "123456789"
}
```

> 💡 **Dica:** Use variáveis para senha e username para maior segurança!

---

### 🔐 Ação 3: HTTP POST - Obter Token

**Nova etapa** → Pesquise: **"HTTP"**

**Configuração:**
```
Método: POST
URI: [Conteúdo dinâmico] → result → url
```

**Cabeçalhos:**
```
Content-Type: application/x-www-form-urlencoded
```

**Corpo:**

Clique em "Mostrar opções avançadas" e adicione expressão:

```
concat(
  'client_id=', outputs('Executar_script')?['body']?['result']?['payload']?['client_id'],
  '&username=', outputs('Executar_script')?['body']?['result']?['payload']?['username'],
  '&password=', outputs('Executar_script')?['body']?['result']?['payload']?['password'],
  '&grant_type=', outputs('Executar_script')?['body']?['result']?['payload']?['grant_type'],
  '&nonce=', outputs('Executar_script')?['body']?['result']?['payload']?['nonce']
)
```

**Configurações Avançadas:**
```
Autenticação: Nenhum
```

---

### 🔍 Ação 4: Analisar JSON - Token

**Nova etapa** → Pesquise: **"Analisar JSON"**

**Configuração:**
```
Conteúdo: [Conteúdo dinâmico] → Corpo (da ação HTTP)
```

**Esquema:**
```json
{
  "type": "object",
  "properties": {
    "access_token": {
      "type": "string"
    },
    "token_type": {
      "type": "string"
    },
    "expires_in": {
      "type": "integer"
    }
  }
}
```

---

### 📋 Ação 5: Executar Script - Build Documentos

**Nova etapa** → Pesquise: **"Excel Online (Business)"** → **"Executar script"**

**Configuração:**
```
Local: OneDrive (ou SharePoint)
Biblioteca de Documentos: OneDrive
Arquivo: [Mesma planilha]
Script: Integrador Completo
```

**Inputs:**
```json
{
  "action": "buildDocumentos"
}
```

---

### 🔁 Ação 6: Aplicar a cada - Payloads

**Nova etapa** → Pesquise: **"Aplicar a cada"**

**Configuração:**
```
Selecione uma saída das etapas anteriores:
[Conteúdo dinâmico] → result → payloads
```

#### Dentro do "Aplicar a cada":

##### 6.1. HTTP POST - Criar Documento

**Adicionar uma ação** → **"HTTP"**

**Configuração:**
```
Método: POST
URI: https://homologacaowisepcp.alterdata.com.br/BimerApi/api/documentos
```

**Cabeçalhos:**
```
Authorization: Bearer [Conteúdo dinâmico → access_token do Parse JSON]
Content-Type: application/json
```

**Corpo:**
```
[Conteúdo dinâmico] → Current item → payload
```

##### 6.2. Analisar JSON - Resposta API

**Adicionar uma ação** → **"Analisar JSON"**

**Conteúdo:**
```
[Corpo da resposta HTTP anterior]
```

**Esquema:**
```json
{
  "type": "object",
  "properties": {
    "ListaObjetos": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "Identificador": {
            "type": "string"
          }
        }
      }
    },
    "Message": {
      "type": "string"
    },
    "Erros": {
      "type": "array"
    }
  }
}
```

##### 6.3. Condição - Verificar Sucesso

**Adicionar uma ação** → **"Condição"**

**Expressão:**
```
Se:
  length(body('Analisar_JSON_Resposta')?['ListaObjetos'])
  é maior que
  0
```

**Se sim (sucesso):**

Adicione ação **"Acrescentar à variável de matriz"**
```
Nome: ResultadosArray
Valor: 
{
  "sheetRow": [Conteúdo dinâmico] → Current item → sheetRow,
  "notaCriada": "Sim",
  "retorno": [Expressão] first(body('Analisar_JSON_Resposta')?['ListaObjetos'])?['Identificador']
}
```

**Se não (erro):**

Adicione ação **"Acrescentar à variável de matriz"**
```
Nome: ResultadosArray
Valor: 
{
  "sheetRow": [Conteúdo dinâmico] → Current item → sheetRow,
  "notaCriada": "Não",
  "retorno": [Expressão] coalesce(body('Analisar_JSON_Resposta')?['Message'], 'Erro desconhecido')
}
```

---

### ✍️ Ação 7: Executar Script - Aplicar Resultados

**Nova etapa** (FORA do "Aplicar a cada") → **"Excel Online (Business)"** → **"Executar script"**

**Configuração:**
```
Local: OneDrive (ou SharePoint)
Biblioteca de Documentos: OneDrive
Arquivo: [Mesma planilha]
Script: Integrador Completo
```

**Inputs:**
```json
{
  "action": "applyResults",
  "results": [Variável ResultadosArray]
}
```

---

## ✅ PASSO 4: Salvar e Testar

### 4.1. Salvar o Flow
1. Clique em **Salvar** no topo
2. Aguarde confirmação

### 4.2. Testar
1. Clique em **Testar** no topo
2. Escolha **"Manualmente"**
3. Clique em **"Testar"**
4. Clique em **"Executar fluxo"**
5. Acompanhe a execução

### 4.3. Verificar Resultados
1. Abra sua planilha no Excel Online
2. Verifique a coluna **"Nota Criada"** (coluna W - índice 21)
3. Verifique a coluna **"Retorno API"** (coluna X - índice 22)

---

## 🐛 Troubleshooting - Problemas Comuns

### ❌ Erro: "Script não encontrado"
**Solução:** Verifique se o script está salvo no workbook correto e com nome "Integrador Completo"

### ❌ Erro: "401 Unauthorized" na API
**Solução:** 
- Verifique credenciais (username/senha)
- Confirme que o token foi obtido corretamente
- Verifique o formato do header Authorization

### ❌ Erro: "Cannot read property 'result'"
**Solução:** 
- O script não está retornando no formato esperado
- Adicione ação "Compor" para visualizar a saída do script

### ❌ Payloads vazios
**Solução:**
- Verifique se há linhas na planilha sem "Nota Criada"
- Confirme que todos os identificadores obrigatórios estão preenchidos
- Execute action='buildValidationQueries' primeiro

---

## 📊 Estrutura da Planilha

Certifique-se que sua planilha tem estas colunas (índices 0-based):

| Índice | Coluna | Nome                           |
|--------|--------|--------------------------------|
| 0      | A      | Código da Empresa              |
| 1      | B      | Código Cliente                 |
| 2      | C      | Nome do Cliente                |
| 3      | D      | **Identificador Cliente**      |
| 4      | E      | Código da Operação             |
| 5      | F      | **Identificador Operação**     |
| 6      | G      | CFOP                           |
| 7      | H      | Código do Serviço              |
| 8      | I      | **Identificador Serviço**      |
| 9      | J      | Nome do Serviço                |
| 10     | K      | Quantidade                     |
| 11     | L      | Valor                          |
| 12     | M      | Discriminação 1                |
| 13     | N      | Discriminação 2                |
| 14     | O      | Código Prazo                   |
| 15     | P      | Identificador Prazo            |
| 16     | Q      | Forma Pagamento Entrada        |
| 17     | R      | Código Forma de Pagamento      |
| 18     | S      | **Identificador Forma Pag.**   |
| 19     | T      | Data Emissão                   |
| 20     | U      | Vencimento Fatura              |
| 21     | V      | **Nota Criada** ✅             |
| 22     | W      | **Retorno API** ✅             |

---

## 🔒 Segurança - Boas Práticas

### Use Variáveis de Ambiente

No Power Automate, crie variáveis para dados sensíveis:

1. Vá em **Soluções** → **Nova Solução**
2. Adicione **Variáveis de Ambiente**
3. Configure:
   - `BIMERApiHost`
   - `BIMERUsername` 
   - `BIMERPassword`

### Azure Key Vault (Recomendado)

Para ambientes de produção, use Azure Key Vault:

1. Crie um Key Vault no Azure
2. Armazene credenciais como secrets
3. No Power Automate, use conector "Azure Key Vault"
4. Recupere secrets durante execução

---

## 📱 Notificações

### Adicionar notificação de sucesso/erro:

Após a última ação, adicione:

**Nova etapa** → **"Enviar um email (V2)"** ou **"Postar mensagem no Teams"**

**Configuração:**
```
Para: seu@email.com
Assunto: Documentos Bimer - Processamento Concluído
Corpo: 
Total de documentos processados: [length(variables('ResultadosArray'))]
Sucesso: [Usar expressão para contar "Sim"]
Erro: [Usar expressão para contar "Não"]
```

---

## 📚 Recursos Adicionais

- [Documentação Power Automate](https://learn.microsoft.com/power-automate/)
- [Office Scripts Reference](https://learn.microsoft.com/office/dev/scripts/)
- [Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Community/ct-p/MPACommunity)

---

## ✨ Próximos Passos

1. ✅ Configure o fluxo básico
2. ✅ Teste com 1-2 linhas primeiro
3. ✅ Adicione tratamento de erros robusto
4. ✅ Configure notificações
5. ✅ Documente o processo para sua equipe
6. ✅ Configure execução agendada (se necessário)

---

**🎉 Pronto! Seu fluxo Power Automate está configurado!**

Qualquer dúvida durante a implementação, me consulte! 🚀
