# 🚀 Guia Rápido - Office Scripts com HTTP

## ✅ O Script Agora Funciona!

Obrigado pela correção! **Office Scripts SUPORTA `fetch()` sim!** 🎉

O script foi atualizado para fazer **chamadas HTTP diretas** à API Bimer.

---

## 📖 Como Usar (Super Simples)

### 1️⃣ Preparar a Planilha

**Estrutura necessária:**

| Col | Nome                        | Obrigatório? | Preenchido por |
|-----|-----------------------------|--------------|----------------|
| A   | Código da Empresa           | ✅           | Você           |
| B   | Código Cliente              | ✅           | Você           |
| C   | Nome do Cliente             | ❌           | Script         |
| D   | **Identificador Cliente**   | ⚠️           | Script/Você    |
| E   | Código da Operação          | ✅           | Você           |
| F   | **Identificador Operação**  | ⚠️           | Script/Você    |
| G   | CFOP                        | ✅           | Você           |
| H   | Código do Serviço           | ✅           | Você           |
| I   | **Identificador Serviço**   | ⚠️           | Script/Você    |
| J   | Nome do Serviço             | ✅           | Você           |
| K   | Quantidade                  | ✅           | Você           |
| L   | Valor                       | ✅           | Você           |
| M   | Discriminação 1             | ❌           | Você           |
| N   | Discriminação 2             | ❌           | Você           |
| O   | Código Prazo                | ❌           | Você           |
| P   | Identificador Prazo         | ❌           | Você           |
| Q   | Forma Pagamento Entrada     | ❌           | Você           |
| R   | Código Forma de Pagamento   | ✅           | Você           |
| S   | **Identificador Forma Pag.**| ⚠️           | Script/Você    |
| T   | Data Emissão                | ✅           | Você           |
| U   | Vencimento Fatura           | ✅           | Você           |
| V   | **Nota Criada**             | 🤖           | **Script**     |
| W   | **Retorno API**             | 🤖           | **Script**     |

⚠️ **Identificadores:** Se você não souber, deixe vazio - o script busca automaticamente!

---

### 2️⃣ Criar o Script no Excel Online

1. Abra sua planilha no **Excel Online** (navegador)
2. Vá em **Automatizar** → **Office Scripts**
3. Clique em **Novo Script**
4. Delete o código de exemplo
5. Cole o conteúdo completo de `IntegradorCompleto.ts`
6. **Salve** com o nome: **"Integrador Completo"**

---

### 3️⃣ Executar o Script

#### Método 1: Botão "Executar" (Mais Simples)

1. No painel Office Scripts, selecione **"Integrador Completo"**
2. Clique em **"Executar"**
3. Aguarde o processamento
4. Verifique as colunas V e W atualizadas!

#### Método 2: Com Parâmetros Personalizados

1. Clique em **"Executar com entrada"**
2. Cole este JSON:

```json
{
  "action": "executarCompleto",
  "host": "https://homologacaowisepcp.alterdata.com.br/BimerApi",
  "username": "supervisor",
  "senha": "Senhas123"
}
```

3. Clique em **"Executar"**

> 💡 **Dica:** Se as credenciais padrão funcionam para você, nem precisa passar parâmetros!

---

## 📊 O Que o Script Faz Automaticamente

### Passo 1: Autenticação 🔐
```
Conecta na API Bimer usando MD5
Obtém token de acesso
```

### Passo 2: Validação 🔍
```
Para cada linha sem "Nota Criada":
  - Se falta Identificador Cliente → Busca na API
  - Se falta Identificador Operação → Busca na API
  - Se falta Identificador Serviço → Busca na API
  - Se falta Identificador Forma Pag. → Busca na API
  Preenche automaticamente na planilha
```

### Passo 3: Criação de Documentos 📄
```
Para cada linha validada:
  - Monta payload do documento
  - Envia POST para /api/documentos
  - Recebe identificador do documento criado
  - Atualiza coluna "Nota Criada" = "Sim"
  - Atualiza coluna "Retorno API" = identificador
```

### Passo 4: Tratamento de Erros ⚠️
```
Se algo der errado:
  - Atualiza "Nota Criada" = "Não"
  - Atualiza "Retorno API" = mensagem de erro
  - Continua processando próximas linhas
```

---

## ✅ Verificar Resultados

Após executar, verifique:

### ✔️ Sucesso
```
Coluna V (Nota Criada): "Sim"
Coluna W (Retorno API): "ABC123DEF" (identificador)
```

### ❌ Erro
```
Coluna V (Nota Criada): "Não"
Coluna W (Retorno API): "Erro: 401 Unauthorized" (mensagem)
```

---

## 🔧 Configurações Avançadas

### Alterar Credenciais

Edite o script diretamente ou passe nos inputs:

```json
{
  "action": "executarCompleto",
  "host": "https://sua-api.exemplo.com.br/BimerApi",
  "username": "seu-usuario",
  "senha": "sua-senha-segura"
}
```

### Processar Apenas Linhas Específicas

O script **pula automaticamente** linhas que já têm:
- ✅ "Nota Criada" preenchida

Para **reprocessar** uma linha:
1. Limpe o conteúdo das colunas V e W
2. Execute o script novamente

---

## 🐛 Troubleshooting

### ❌ Erro: "Planilha 'Documento' não encontrada"
**Solução:** Renomeie a aba para exatamente **"Documento"**

### ❌ Erro: "401 Unauthorized"
**Solução:** Verifique username e senha nas credenciais

### ❌ Erro: "Campos obrigatórios faltando"
**Solução:** 
- Verifique se Código Cliente, Operação e Serviço estão preenchidos
- Execute novamente - o script tenta buscar identificadores automaticamente

### ❌ Nada acontece
**Solução:**
- Verifique se há linhas SEM "Nota Criada"
- Adicione `console.log()` no script para debug

### ❌ Erro de CORS
**Solução:** 
- A API precisa permitir requisições do domínio do Excel Online
- Entre em contato com o administrador da API

---

## 📈 Exemplo de Resultado

**Antes:**
| Cliente | Identificador Cliente | Nota Criada | Retorno API |
|---------|----------------------|-------------|-------------|
| CLI001  | (vazio)              | (vazio)     | (vazio)     |

**Depois:**
| Cliente | Identificador Cliente        | Nota Criada | Retorno API           |
|---------|------------------------------|-------------|-----------------------|
| CLI001  | 12345678-abcd-1234-efgh-... | Sim         | DOC-2024-00123       |

---

## 🎯 Fluxo Completo Visual

```
┌─────────────────────────────────────────────────┐
│ VOCÊ: Preenche dados básicos na planilha       │
│      (códigos, valores, quantidades)            │
└─────────────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────┐
│ SCRIPT: Executa automaticamente                 │
│   1️⃣ Autentica na API                           │
│   2️⃣ Busca identificadores faltantes            │
│   3️⃣ Cria documentos                            │
│   4️⃣ Atualiza resultados na planilha            │
└─────────────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────┐
│ RESULTADO: Planilha atualizada com             │
│   ✅ "Nota Criada" = Sim/Não                    │
│   ✅ "Retorno API" = ID ou erro                 │
└─────────────────────────────────────────────────┘
```

---

## 🔒 Segurança

### ⚠️ IMPORTANTE: Credenciais

**NÃO recomendado para produção:**
```typescript
// Credenciais hardcoded no script ❌
const username = 'supervisor';
const senha = 'Senhas123';
```

**Recomendado:**
- Use Azure Key Vault para armazenar credenciais
- Passe credenciais via inputs (não salve no script)
- Configure políticas de acesso adequadas

---

## 📚 Ações Disponíveis

| Action                      | Descrição                                    |
|-----------------------------|----------------------------------------------|
| `executarCompleto`          | ⭐ **Executa tudo automaticamente**         |
| `buildAuthPayload`          | Gera payload de autenticação                 |
| `buildValidationQueries`    | Lista o que precisa validar                  |
| `applyValidationResults`    | Aplica IDs validados na planilha             |
| `buildDocumentos`           | Gera payloads de documentos                  |
| `applyResults`              | Aplica resultados de APIs na planilha        |
| `help`                      | Mostra ajuda                                 |

---

## 🎉 Pronto!

Agora você pode:
✅ Processar planilhas com 1 clique  
✅ Criar documentos na API Bimer automaticamente  
✅ Validar identificadores automaticamente  
✅ Rastrear sucessos e erros  

**Qualquer dúvida, consulte a documentação completa no cabeçalho do script!** 📖
