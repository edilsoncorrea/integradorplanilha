# 📘 Guia de Uso: Integrador Completo Office Scripts

## 🎯 Visão Geral

O arquivo `IntegradorCompleto.ts` é um script único e consolidado que centraliza todas as funcionalidades do projeto para uso no **Office Scripts do Excel Online (Automatizar)**.

## 🔑 Funcionalidades Integradas

### 1. **Autenticação na API Bimer**
- Gera payload MD5 para autenticação
- Cria requisição para obter access_token

### 2. **Validação de Dados da Planilha**
- Identifica campos faltantes (Cliente, Operação, Serviço, Prazo)
- Gera lista de consultas GET necessárias
- Aplica resultados na planilha

### 3. **Criação de Pedidos de Venda**
- Lê dados da planilha "Documento"
- Gera payloads completos para API de Pedidos
- Inclui itens, prazos e formas de pagamento

### 4. **Criação de Documentos Fiscais**
- Gera payloads para documentos fiscais
- Inclui itens e pagamentos
- Formata observações e valores

### 5. **Aplicação de Resultados**
- Escreve retornos da API na planilha
- Atualiza campos NotaCriada e RetornoAPI

## 📋 Estrutura da Planilha

A planilha deve ter uma aba chamada **"Documento"** com as seguintes colunas (0-based):

| Índice | Coluna | Descrição |
|--------|--------|-----------|
| 0 | CodigoDaEmpresa | Código da empresa |
| 1 | CodigoCliente | Código do cliente |
| 2 | NomeDoCliente | Nome do cliente |
| 3 | IdentificadorCliente | ID interno do cliente |
| 4 | CodigoDaOperacao | Código da operação |
| 5 | IdentificadorOperacao | ID interno da operação |
| 6 | CFOP | Código Fiscal |
| 7 | CodigoDoServico | Código do serviço/produto |
| 8 | IdentificadorServico | ID interno do serviço |
| 9 | NomeDoServico | Nome do serviço |
| 10 | Quantidade | Quantidade |
| 11 | Valor | Valor unitário |
| 12 | Descriminacao1 | Descrição 1 |
| 13 | Descriminacao2 | Descrição 2 |
| 14 | Codigoprazo | Código do prazo |
| 15 | IdentificadorPrazo | ID interno do prazo |
| 16 | FormaPagamentoEntrada | SIM/NÃO |
| 17 | CodigoDaFormaDePagamento | Código forma de pagamento |
| 18 | IdentificadorFormaPagamento | ID interno forma de pagamento |
| 19 | DataEmissao | Data de emissão |
| 20 | VencimentoFatura | Data de vencimento |
| 21 | NotaCriada | Número da nota criada (preenchido pela API) |
| 22 | RetornoAPI | Retorno da API (preenchido automaticamente) |

## 🚀 Como Usar no Office Scripts

### Passo 1: Adicionar o Script ao Excel Online

1. Abra sua planilha no **Excel Online** (Office 365)
2. Vá em **Automatizar** > **Novo Script**
3. Copie todo o conteúdo de `IntegradorCompleto.ts`
4. Cole no editor do Office Scripts
5. Salve com o nome "Integrador Completo"

### Passo 2: Testar o Script Manualmente

Para testar ações individuais:

```typescript
// No Power Automate ou Office Scripts, chame:
main(workbook, { action: 'help' })
```

Isso retorna todas as ações disponíveis.

## 🔄 Fluxo Completo com Power Automate

### 📌 Fluxo Recomendado

```
┌─────────────────────────────────────────────────────────────┐
│ 1. AUTENTICAÇÃO                                              │
├─────────────────────────────────────────────────────────────┤
│ Script: { action: 'buildAuthPayload' }                      │
│ ↓                                                            │
│ HTTP POST: /oauth/token                                     │
│ ↓                                                            │
│ Salvar: access_token                                        │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. VALIDAÇÃO DA PLANILHA                                    │
├─────────────────────────────────────────────────────────────┤
│ Script: { action: 'buildValidationQueries' }                │
│ ↓                                                            │
│ Loop: Para cada query                                       │
│   HTTP GET: query.endpoint                                  │
│   Processar resposta                                        │
│ ↓                                                            │
│ Script: { action: 'applyValidationResults',                 │
│          results: [...] }                                   │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. CRIAR PEDIDOS OU DOCUMENTOS                              │
├─────────────────────────────────────────────────────────────┤
│ Script: { action: 'buildPedidos' }                          │
│         OU                                                  │
│         { action: 'buildDocumentos' }                       │
│ ↓                                                            │
│ Loop: Para cada payload                                     │
│   HTTP POST: /api/pedidosVenda ou /api/documentos          │
│   Coletar: notaCriada e retorno                            │
│ ↓                                                            │
│ Script: { action: 'applyResults',                           │
│          results: [{sheetRow, notaCriada, retorno}] }      │
└─────────────────────────────────────────────────────────────┘
```

## 📝 Exemplos de Uso

### Exemplo 1: Autenticação

```typescript
// Input
{
  action: 'buildAuthPayload',
  host: 'https://homologacaowisepcp.alterdata.com.br/BimerApi',
  username: 'supervisor',
  senha: 'Senhas123',
  nonce: '123456789'
}

// Output
{
  url: 'https://homologacaowisepcp.alterdata.com.br/BimerApi/oauth/token',
  method: 'POST',
  payload: {
    client_id: 'IntegracaoBimer.js',
    username: 'supervisor',
    password: '9a3f2c...', // MD5 hash
    grant_type: 'password',
    nonce: '123456789'
  },
  note: 'Use este payload no Power Automate para fazer POST e obter access_token'
}
```

### Exemplo 2: Validação

```typescript
// Input
{ action: 'buildValidationQueries' }

// Output
{
  queries: [
    {
      sheetRow: 3,
      method: 'GET',
      endpoint: '/api/pessoas/codigo/CLI001',
      field: 'IdentificadorCliente',
      codigo: 'CLI001'
    },
    {
      sheetRow: 3,
      method: 'GET',
      endpoint: '/api/formasPagamento',
      field: 'IdentificadorFormaPagamento',
      codigo: 'FP01'
    }
    // ... mais queries
  ],
  total: 5,
  note: 'Execute cada query no Power Automate e chame applyValidationResults com os resultados'
}
```

### Exemplo 3: Criar Pedidos

```typescript
// Input
{ action: 'buildPedidos' }

// Output
{
  payloads: [
    {
      sheetRow: 3,
      tipo: 'pedido',
      payload: {
        CodigoEmpresa: 1,
        DataEmissao: '2024-12-24',
        IdentificadorOperacao: '00A000000R',
        IdentificadorCliente: '00A0000063',
        Observacao: 'Serviço de consultoria (10 X R$ 150,00) - 150...',
        Itens: [
          {
            CFOP: '5101',
            IdentificadorProduto: '00A00000SQ',
            QuantidadePedida: 10,
            Valor: 1500,
            ValorUnitario: 150
          }
        ],
        Status: 'A',
        TipoFrete: 'E'
      }
    }
  ],
  total: 1,
  note: 'POST cada payload para /api/pedidosVenda e chame applyResults com as respostas'
}
```

### Exemplo 4: Aplicar Resultados

```typescript
// Input
{
  action: 'applyResults',
  results: [
    {
      sheetRow: 3,
      notaCriada: 'PV-12345',
      retorno: 'Pedido criado com sucesso'
    },
    {
      sheetRow: 4,
      notaCriada: 'PV-12346',
      retorno: 'Pedido criado com sucesso'
    }
  ]
}

// Output
{
  ok: true,
  updated: 4,
  message: '4 resultado(s) aplicado(s) na planilha'
}
```

## ⚙️ Configuração do Power Automate

### Template de Flow Básico

1. **Gatilho**: "Quando um botão de fluxo é clicado" ou "Recorrência"

2. **Executar script do Office (Autenticação)**
   - Script: Integrador Completo
   - Parâmetros: `{ "action": "buildAuthPayload" }`

3. **HTTP (POST Token)**
   - Método: POST
   - URI: `@{outputs('Script_Auth').url}`
   - Cabeçalhos: Content-Type: application/x-www-form-urlencoded
   - Corpo: `@{outputs('Script_Auth').payload}`

4. **Analisar JSON (Token)**
   - Conteúdo: `@{body('HTTP_Token')}`

5. **Executar script do Office (Build Pedidos)**
   - Script: Integrador Completo
   - Parâmetros: `{ "action": "buildPedidos" }`

6. **Apply to each (Payloads)**
   - Para cada: `@{outputs('Script_Pedidos').payloads}`
   
7. **HTTP (POST Pedido)**
   - Método: POST
   - URI: `https://HOST/api/pedidosVenda`
   - Cabeçalhos: 
     - Authorization: `Bearer @{body('Parse_Token').access_token}`
     - Content-Type: application/json
   - Corpo: `@{items('Apply_to_each').payload}`

8. **Acrescentar à variável de matriz (Resultados)**
   ```json
   {
     "sheetRow": @{items('Apply_to_each').sheetRow},
     "notaCriada": "@{body('HTTP_Pedido').Identificador}",
     "retorno": "@{body('HTTP_Pedido')}"
   }
   ```

9. **Executar script do Office (Aplicar Resultados)**
   - Script: Integrador Completo
   - Parâmetros: 
   ```json
   {
     "action": "applyResults",
     "results": @{variables('ResultadosArray')}
   }
   ```

## 🔍 Troubleshooting

### Erro: "Planilha 'Documento' não encontrada"
- Certifique-se que sua planilha tem uma aba chamada exatamente "Documento"

### Erro: "Campos obrigatórios faltando"
- Execute primeiro a validação (`buildValidationQueries` + `applyValidationResults`)
- Verifique se os campos IdentificadorCliente, IdentificadorOperacao e IdentificadorServico estão preenchidos

### Linhas não são processadas
- Verifique se a coluna NotaCriada está vazia (o script pula linhas com nota já criada)
- Certifique-se que as linhas começam na linha 3 (índice 2)

### Valores não formatam corretamente
- O script aceita valores em formato brasileiro (R$ 1.500,00) ou decimal (1500.00)
- Verifique se a coluna Valor tem dados válidos

## 📊 Monitoramento e Logs

O script retorna objetos estruturados que facilitam o monitoramento:

```typescript
// Todas as respostas incluem:
{
  ok: true/false,          // Status da operação
  error: "mensagem",       // Em caso de erro
  total: número,           // Quantidade de itens processados
  note: "orientação"       // Dicas de próximos passos
}
```

## 🔐 Segurança

- **Credenciais**: Configure no Power Automate usando variáveis de ambiente
- **Token**: Armazene o access_token em variável do Flow, nunca na planilha
- **HTTPS**: Sempre use conexões seguras (https://)
- **Logs**: Evite logar dados sensíveis (senhas, tokens) no Power Automate

## 📚 Referências

- [Documentação Office Scripts](https://learn.microsoft.com/office/dev/scripts/)
- [Power Automate](https://make.powerautomate.com/)
- [API Bimer](https://homologacaowisepcp.alterdata.com.br/BimerApi)

## 💡 Dicas de Otimização

1. **Processamento em Lote**: Processe múltiplas linhas em um único flow
2. **Tratamento de Erros**: Use blocos try-catch no Power Automate
3. **Cache de Token**: Reutilize o access_token enquanto válido (não autentique a cada requisição)
4. **Validação Incremental**: Execute validação apenas quando necessário
5. **Logs Estruturados**: Use o histórico do Power Automate para debug

## ✅ Checklist de Implementação

- [ ] Script copiado para Office Scripts
- [ ] Planilha "Documento" criada com estrutura correta
- [ ] Flow do Power Automate configurado
- [ ] Autenticação testada
- [ ] Validação testada
- [ ] Criação de pedido/documento testada
- [ ] Aplicação de resultados testada
- [ ] Tratamento de erros implementado
- [ ] Monitoramento configurado

---

**Versão**: 1.0  
**Data**: 24/12/2025  
**Autor**: Sistema Integrador PWC
