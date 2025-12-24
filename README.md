# 📊 Integrador de Planilha - API Bimer

Sistema de integração entre planilhas Excel/Google Sheets e API Bimer para criação automatizada de pedidos de venda.

## 🚀 Funcionalidades

- ✅ **Autenticação automática** na API Bimer
- ✅ **Conversão de dados** de planilha para payloads da API
- ✅ **Criação de pedidos** de venda via API
- ✅ **Debug completo** com breakpoints e logs detalhados
- ✅ **Recuperação de identificadores** dos pedidos criados
- ✅ **Compatibilidade** com Office Scripts e Google Apps Script

## 📁 Estrutura do Projeto

```
src/
├── office-scripts/     # Conversões para Office Scripts (Excel Online)
├── src_appscripts/     # Scripts originais do Google Apps Script
├── api/               # Cliente da API Bimer
├── mocks/             # Dados simulados para testes
├── debug/             # Sistema de debug e breakpoints
└── utils/             # Utilitários diversos

test/                  # Testes e debugs
├── debug-auth-step-by-step.test.ts
├── debug-pedidos-csv.test.ts
├── debug-pedidos-interativo.test.ts
└── debug-pedidos-criados.test.ts
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

### Credenciais da API
As credenciais estão configuradas em `src/office-scripts/Autenticacao.ts`:
- **Host**: `https://087344bimerapi.alterdata.cloud`
- **Username**: `bimerapi`
- **Senha**: `123456`

### Dados de Teste
Os dados simulados estão em `src/mocks/excel-mock.ts` e incluem:
- Headers da planilha
- Dados de empresas, clientes e produtos
- Configurações de prazo e pagamento

## 🚀 Como Usar

### 1. Debug no VS Code (Recomendado)
```bash
# Pressionar Ctrl+Shift+D no VS Code
# Selecionar uma das opções:
# - 🔐 Debug Autenticação
# - 📦 Debug Pedidos CSV → API  
# - 🚀 Debug Pedidos Interativo
# - 🔍 Debug Pedidos Criados
```

### 2. Linha de Comando
```bash
# Testar autenticação
npx ts-node test/debug-auth-step-by-step.test.ts

# Testar criação de pedidos
npx ts-node test/debug-pedidos-csv.test.ts

# Verificar pedidos criados
npx ts-node test/debug-pedidos-criados.test.ts
```

## 📊 Exemplo de Uso

```typescript
import { MockWorkbook } from './src/mocks/excel-mock';
import { main as pedidoMain } from './src/office-scripts/PedidoDeVenda';
import { RealAPIClient } from './src/api/real-api-client';

// 1. Carregar dados da planilha
const workbook = new MockWorkbook().loadRealData();

// 2. Gerar payloads
const result = pedidoMain(workbook, { 
  action: 'buildPedidoVendaFromSheet' 
});

// 3. Autenticar e criar pedidos
const apiClient = new RealAPIClient();
const authResult = await apiClient.authenticate();

for (const pedido of result.payloads) {
  const pedidoResult = await apiClient.createPedido(pedido.payload);
  console.log(`Pedido criado: ${pedidoResult.data.ListaObjetos[0].Identificador}`);
}
```

## 🔍 Sistema de Debug

O projeto inclui um sistema completo de debug com:

- **Breakpoints automáticos** em pontos estratégicos
- **Logs detalhados** de requisições e respostas
- **Inspeção de variáveis** em tempo real
- **Trace de rede** com tempos de resposta

### Breakpoints Disponíveis
- `AUTH_CHECK`: Verificação de autenticação
- `BEFORE_API_CALL`: Antes de chamadas da API
- `AFTER_API_RESPONSE`: Após respostas da API

## 📋 Estrutura dos Dados

### Planilha de Entrada
| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| CodigoEmpresa | Código da empresa | 000012 |
| IdentificadorCliente | ID do cliente | 00A0000023 |
| IdentificadorOperacao | ID da operação | 00A000000X |
| IdentificadorServico | ID do serviço | 00A00000DK |
| Quantidade | Quantidade | 01 |
| Valor | Valor do item | R$ 1.500,00 |

### Resposta da API
```json
{
  "Erros": [],
  "ListaObjetos": [
    {
      "Identificador": "00A000002B",
      "Codigo": "000043"
    }
  ]
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

1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

## 📄 Licença

Este projeto é propriedade da Alterdata Software.

---

**Desenvolvido por**: Edilson DSN ERP  
**Empresa**: Alterdata Software  
**Repositório**: https://git.alterdata.com.br/edilson.dsn.erp/integradorplanilha