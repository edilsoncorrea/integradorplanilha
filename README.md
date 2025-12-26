# 📊 Integrador de Planilha - API Bimer

Sistema de integração entre planilhas Excel/Google Sheets e API Bimer para criação automatizada de pedidos de venda e documentos fiscais.

## ✅ Status do Projeto

**🎉 Versão 1.0.0 - TOTALMENTE FUNCIONAL**

Este projeto está **completo e validado** com todas as funcionalidades operacionais:
- ✅ Office Scripts funcionando no Excel Online
- ✅ Integração com Power Automate configurada
- ✅ Criação de documentos validada na API Bimer
- ✅ Validação automática de identificadores
- ✅ Documentação completa disponível

## 🚀 Funcionalidades

- ✅ **Autenticação MD5** nativa na API Bimer (sem dependências)
- ✅ **Validação automática** de identificadores (Cliente, Operação, Serviço, Forma de Pagamento)
- ✅ **Criação de Pedidos de Venda** via API
- ✅ **Criação de Documentos Fiscais** via API
- ✅ **Integração com Power Automate** para automação completa
- ✅ **Sistema modular de ações** para controle de fluxo
- ✅ **Atualização automática** de resultados na planilha
- ✅ **Debug completo** com breakpoints e logs detalhados
- ✅ **Compatibilidade** com Office Scripts e Google Apps Script

## 📁 Estrutura do Projeto

```
src/
├── office-scripts/              # Office Scripts para Excel Online ⭐
│   ├── IntegradorCompleto.ts   # Script principal modular
│   ├── Autenticacao.ts          # Lógica de autenticação
│   ├── DocumentoScript.ts       # Criação de documentos
│   └── PedidoDeVenda.ts        # Criação de pedidos
├── api/                         # Cliente da API Bimer
├── mocks/                       # Dados simulados para testes
├── debug/                       # Sistema de debug
└── utils/                       # Utilitários diversos

docs/                           # 📚 Documentação completa
├── GUIA_RAPIDO_OFFICE_SCRIPTS.md           # Início rápido
├── GUIA_POWER_AUTOMATE_PASSO_A_PASSO.md   # Power Automate
├── ALTERNATIVAS_SEM_POWER_AUTOMATE.md      # Outras soluções
└── GUIA_INTEGRADOR_COMPLETO.md             # Guia técnico

test/                          # Testes e debugs
├── debug-auth-step-by-step.test.ts
├── debug-pedidos-csv.test.ts
├── debug-documentos-api-real.test.ts
└── debug-fluxo-completo-documentos.test.ts
```

## 🛠️ Instalação

```bash
# Clonar repositório
git clone https://git.alterdata.com.br/edilson.dsn.erp/integradorplanilha.git
cd integradorplanilha

# Instalar dependências
npm install

# Compilar TypeScript
npm run build
```

## 🔧 Configuração

### API Bimer
- **Host**: `https://homologacaowisepcp.alterdata.com.br/BimerApi`
- **Autenticação**: OAuth2 com hash MD5
- **Endpoints**: `/oauth/token`, `/api/documentos`, `/api/pedidosVenda`

### Variáveis de Ambiente (Opcional)
```bash
BIMER_HOST=https://homologacaowisepcp.alterdata.com.br/BimerApi
BIMER_USERNAME=supervisor
BIMER_PASSWORD=Senhas123
```

### Configuração do Office Script
As credenciais padrão estão em [IntegradorCompleto.ts](src/office-scripts/IntegradorCompleto.ts):
- Podem ser sobrescritas via inputs do Power Automate
- Recomendado: Use Azure Key Vault em produção

## 🚀 Como Usar

### 📖 Guias Disponíveis

1. **[Guia Rápido Office Scripts](docs/GUIA_RAPIDO_OFFICE_SCRIPTS.md)** - Para começar imediatamente
2. **[Guia Power Automate](docs/GUIA_POWER_AUTOMATE_PASSO_A_PASSO.md)** - Configuração completa
3. **[Alternativas](docs/ALTERNATIVAS_SEM_POWER_AUTOMATE.md)** - Outras soluções (Node.js, Python, Azure)

### 🎯 Início Rápido

#### Opção 1: Excel Online + Office Scripts (Recomendado)
```
1. Abra sua planilha no Excel Online
2. Vá em Automatizar → Office Scripts
3. Cole o código de IntegradorCompleto.ts
4. Execute com um clique!
```

#### Opção 2: Excel Online + Power Automate (Automação Total)
```
1. Configure o Office Script conforme Opção 1
2. Crie um Flow no Power Automate
3. Configure as ações HTTP para a API
4. Execute manualmente ou agende
```

#### Opção 3: Desenvolvimento Local
```bash
# Clonar repositório
git clone <url-do-repo>
cd PWC

# Instalar dependências
npm install

# Compilar TypeScript
npm run build

# Executar testes
npx ts-node test/debug-fluxo-completo-documentos.test.ts
```

## 📊 Exemplo de Uso - IntegradorCompleto.ts

### Estrutura Modular com Actions

```typescript
// No Excel Online, execute o script com:
{ "action": "executarCompleto" }

// OU configure via Power Automate:

// 1. Autenticar
{ "action": "buildAuthPayload" }
// Retorna: { url, method, payload }

// 2. Validar Identificadores
{ "action": "buildValidationQueries" }
// Retorna: { queries: [...] }

// 3. Aplicar Validações
{ 
  "action": "applyValidationResults",
  "results": [...]
}

// 4. Gerar Payloads de Documentos
{ "action": "buildDocumentos" }
// Retorna: { payloads: [...] }

// 5. Aplicar Resultados da API
{
  "action": "applyResults",
  "results": [
    {
      "sheetRow": 3,
      "notaCriada": "Sim",
      "retorno": "ABC123XYZ"
    }
  ]
}
```

### Colunas da Planilha (índices 0-based)

| Índice | Coluna | Nome                      | Obrigatório |
|--------|--------|---------------------------|-------------|
| 0      | A      | Código da Empresa         | ✅          |
| 1      | B      | Código Cliente            | ✅          |
| 3      | D      | Identificador Cliente     | ⚠️ *       |
| 5      | F      | Identificador Operação    | ⚠️ *       |
| 8      | I      | Identificador Serviço     | ⚠️ *       |
| 18     | S      | Identificador Forma Pag.  | ⚠️ *       |
| 21     | V      | **Nota Criada** (output)  | 🤖          |
| 22     | W      | **Retorno API** (output)  | 🤖          |

*⚠️ Se não preenchido, o script busca automaticamente na API*

## 🔍 Sistema de Debug

O projeto inclui um sistema completo de debug com:

- **Breakpoints automáticos** em pontos estratégicos
- **Logs detalhados** de requisições e respostas
- **Inspeção de variáveis** em tempo real
- **Trace de rede** com tempos de resposta

### Testes Disponíveis
```bash
# Teste de autenticação passo a passo
npx ts-node test/debug-auth-step-by-step.test.ts

# Teste de criação de documentos (fluxo completo)
npx ts-node test/debug-fluxo-completo-documentos.test.ts

# Teste de validação de planilha
npx ts-node test/debug-validar-planilha.test.ts

# Teste com dados reais da API
npx ts-node test/debug-documentos-api-real.test.ts
```

### Breakpoints Disponíveis
- `AUTH_CHECK`: Verificação de autenticação
- `BEFORE_API_CALL`: Antes de chamadas da API
- `AFTER_API_RESPONSE`: Após respostas da API
- `VALIDATION_START`: Início da validação
- `PAYLOAD_BUILD`: Construção de payloads

## 📝 Últimos Ajustes (v1.0.0)

### ✅ Correções Implementadas
1. **IntegradorCompleto.ts**: Corrigido problema de constantes não definidas
2. **Passagem de parâmetros**: Todas as funções agora recebem constantes explicitamente
3. **Remoção de variáveis globais**: Código mais limpo e sem side effects
4. **Documentação completa**: 4 guias detalhados adicionados
5. **Estrutura modular**: Sistema de ações totalmente funcional

### 🎯 Funcionalidades Validadas
- ✅ Autenticação MD5 funcionando
- ✅ Criação de documentos na API
- ✅ Validação automática de identificadores
- ✅ Atualização de resultados na planilha
- ✅ Integração com Power Automate

### 📚 Documentação Adicionada
- [CHANGELOG.md](CHANGELOG.md) - Histórico completo de mudanças
- [CONTRIBUTING.md](CONTRIBUTING.md) - Guia de contribuição
- [docs/GUIA_RAPIDO_OFFICE_SCRIPTS.md](docs/GUIA_RAPIDO_OFFICE_SCRIPTS.md)
- [docs/GUIA_POWER_AUTOMATE_PASSO_A_PASSO.md](docs/GUIA_POWER_AUTOMATE_PASSO_A_PASSO.md)
- [docs/ALTERNATIVAS_SEM_POWER_AUTOMATE.md](docs/ALTERNATIVAS_SEM_POWER_AUTOMATE.md)

## 📋 Estrutura dos Dados da API

### Request - Criar Documento
```json
{
  "StatusNotaFiscalEletronica": "A",
  "TipoDocumento": "S",
  "TipoPagamento": "0",
  "CodigoEmpresa": "000012",
  "DataEmissao": "2024-12-26",
  "DataReferencia": "2024-12-26",
  "IdentificadorOperacao": "00A000000X",
  "IdentificadorPessoa": "00A0000023",
  "Itens": [{
    "CFOP": "5933",
    "IdentificadorProduto": "00A00000DK",
    "Quantidade": 1,
    "ValorUnitario": 1500.00
  }],
  "Pagamentos": [{
    "Aliquota": 100,
    "DataVencimento": "2024-12-26",
    "IdentificadorFormaPagamento": "00A000000P",
    "Valor": 1500.00
  }]
}
```

### Response - Documento Criado
```json
{
  "Erros": [],
  "ListaObjetos": [{
    "Identificador": "00A000002B",
    "Codigo": "000043",
    "DataEmissao": "2024-12-26",
    "ValorTotal": 1500.00
  }]
}
```

## 🛡️ Arquitetura

O projeto segue os princípios de **Domain-Driven Design (DDD)** e **Clean Architecture**:

- **Separação de responsabilidades**
- **Código testável e modular**
- **Compatibilidade entre plataformas**
- **Debug e monitoramento integrados**

## 📝 Scripts Disponíveis

```bash
npm run build          # Compilar TypeScript
npm run test           # Executar testes
npm run debug:auth     # Debug de autenticação
npm run debug:pedidos  # Debug de pedidos
```

## 🤝 Contribuição

Leia o [CONTRIBUTING.md](CONTRIBUTING.md) para detalhes sobre:
- Padrões de código
- Processo de Pull Request
- Como reportar bugs
- Como sugerir melhorias

## 📝 Changelog

Veja [CHANGELOG.md](CHANGELOG.md) para histórico completo de versões.

## 📄 Licença

Este projeto é propriedade da Alterdata Software.

---

## 🎯 Próximos Passos

Consulte o [ROADMAP.md](ROADMAP.md) para ver funcionalidades planejadas.

---

**🎉 Status**: Versão 1.0.0 - Totalmente Funcional  
**📅 Última Atualização**: 26 de Dezembro de 2024  
**👨‍💻 Desenvolvido por**: Edilson DSN ERP  
**🏢 Empresa**: Alterdata Software  

---

## 📞 Suporte

Para dúvidas ou problemas:
1. Consulte a documentação em `docs/`
2. Verifique a seção de Troubleshooting nos guias
3. Abra uma issue no repositório
4. Entre em contato com o time de desenvolvimento

---

**⭐ Se este projeto te ajudou, considere dar uma estrela no repositório!**