# 🎯 Integrador Completo - Script Único para Office Scripts

## 📌 Sobre

Este diretório contém o script consolidado **`IntegradorCompleto.ts`** que centraliza TODAS as funcionalidades do projeto em um único arquivo TypeScript pronto para ser usado no **Office Scripts do Excel Online**.

## 🎁 O que você encontra aqui

### ✅ Arquivo Principal
- **`IntegradorCompleto.ts`** - Script único e completo com todas as funcionalidades

### 📚 Documentação
- **`GUIA_INTEGRADOR_COMPLETO.md`** (em /docs/) - Manual completo de uso
- **`power-automate-flow-completo.json`** - Template de Flow do Power Automate

### 📁 Arquivos Individuais (Mantidos para Manutenção)
Todos os arquivos TypeScript individuais continuam disponíveis para facilitar a manutenção:
- `Autenticacao.ts`
- `ValidarPlanilha.ts`
- `PedidoDeVenda.ts`
- `DocumentoScript.ts`
- `Constantes.ts`
- etc.

## 🚀 Quick Start - 3 Passos Simples

### 1️⃣ Copiar o Script

1. Abra o arquivo [`IntegradorCompleto.ts`](./IntegradorCompleto.ts)
2. Copie TODO o conteúdo (Ctrl+A, Ctrl+C)
3. Abra sua planilha no Excel Online (Office 365)
4. Vá em **Automatizar** > **Novo Script**
5. Cole o código
6. Clique em **Salvar script** com o nome "Integrador Completo"

### 2️⃣ Preparar a Planilha

Certifique-se de que sua planilha tem:
- ✅ Uma aba chamada **"Documento"**
- ✅ As colunas corretas (veja estrutura abaixo)
- ✅ Dados preenchidos nas linhas (começando na linha 3)

### 3️⃣ Criar o Flow no Power Automate

Use o template [`power-automate-flow-completo.json`](./power-automate-flow-completo.json) como base:

1. Acesse [Power Automate](https://make.powerautomate.com/)
2. Crie novo Flow
3. Siga a estrutura do template JSON
4. Configure conexões (OneDrive, HTTP)
5. Teste com dados reais

## 📋 Estrutura da Planilha "Documento"

```
| A  | B       | C      | D          | E       | F          | G    | ... | V         | W          |
|----|---------|--------|------------|---------|------------|------|-----|-----------|------------|
| 0  | 1       | 2      | 3          | 4       | 5          | 6    | ... | 21        | 22         |
| Cód| Cód     | Nome   | ID         | Cód     | ID         | CFOP | ... | Nota      | Retorno    |
| Emp| Cliente | Cliente| Cliente    | Operação| Operação   |      | ... | Criada    | API        |
```

**Colunas obrigatórias para criar pedido/documento:**
- Coluna A (0): Código da Empresa
- Coluna D (3): Identificador Cliente*
- Coluna F (5): Identificador Operação*
- Coluna I (8): Identificador Serviço*
- Coluna K (10): Quantidade
- Coluna L (11): Valor
- Coluna T (19): Data Emissão

\* *Se não preenchidos, execute a validação primeiro*

## 🎮 Ações Disponíveis

| Ação | Descrição | Input | Output |
|------|-----------|-------|--------|
| `help` | Lista todas as ações | - | Objeto com documentação |
| `buildAuthPayload` | Gera payload de autenticação | host, username, senha, nonce | url, method, payload |
| `hash` | Calcula MD5 | value | md5 hash |
| `buildValidationQueries` | Lista campos faltantes | - | Array de queries GET |
| `applyValidationResults` | Aplica IDs na planilha | results[] | ok, updated |
| `buildPedidos` | Gera payloads de pedidos | - | Array de payloads |
| `buildDocumentos` | Gera payloads de documentos | - | Array de payloads |
| `applyResults` | Escreve resultados na planilha | results[] | ok, updated |

## 💡 Exemplos de Uso

### Teste Rápido no Office Scripts

```typescript
// Cole no editor do Office Scripts junto com o script
// e execute para testar

function testHelp() {
  const workbook = Excel.getCurrentWorkbook();
  const resultado = main(workbook, { action: 'help' });
  console.log(resultado);
}
```

### No Power Automate

```json
// Step: Executar Script do Office
{
  "script": "Integrador Completo",
  "inputs": {
    "action": "buildPedidos"
  }
}
```

## 🔄 Fluxo Completo Recomendado

```
1. AUTENTICAR
   ├─ Script: buildAuthPayload
   ├─ HTTP POST: /oauth/token
   └─ Salvar: access_token

2. VALIDAR (se necessário)
   ├─ Script: buildValidationQueries
   ├─ Loop HTTP GET: para cada query
   └─ Script: applyValidationResults

3. CRIAR PEDIDOS/DOCUMENTOS
   ├─ Script: buildPedidos (ou buildDocumentos)
   ├─ Loop HTTP POST: para cada payload
   └─ Coletar: respostas da API

4. APLICAR RESULTADOS
   └─ Script: applyResults com todas as respostas
```

## 🔍 Diferenças vs Arquivos Individuais

### ✅ Vantagens do Script Consolidado

- **Único arquivo**: Fácil de copiar e colar no Office Scripts
- **Sem dependências**: Tudo em um só lugar
- **Pronto para produção**: Testado e validado
- **Documentado**: Comentários detalhados no código
- **Manutenível**: Seções claramente separadas

### 📂 Por que manter arquivos individuais?

- **Desenvolvimento**: Mais fácil editar arquivos menores
- **Testes**: Testar funcionalidades isoladas
- **Histórico**: Rastrear mudanças específicas
- **Reutilização**: Copiar funções para outros projetos

## ⚠️ Limitações Importantes

O Office Scripts **NÃO PODE**:
- ❌ Fazer requisições HTTP diretamente à API externa
- ❌ Acessar recursos fora do Excel (arquivos, sistema)
- ❌ Executar código assíncrono complexo

**Solução**: Use Power Automate para:
- ✅ Fazer chamadas HTTP à API Bimer
- ✅ Orquestrar o fluxo completo
- ✅ Tratar erros e retentativas
- ✅ Enviar notificações

## 🛠️ Manutenção

### Atualizar o Script Consolidado

Se você editar arquivos individuais e quiser atualizar o consolidado:

1. Faça as alterações nos arquivos individuais
2. Teste cada funcionalidade
3. Atualize as seções correspondentes em `IntegradorCompleto.ts`
4. Teste o script consolidado completo
5. Atualize no Office Scripts

### Adicionar Nova Funcionalidade

1. Crie/edite o arquivo individual correspondente
2. Teste a funcionalidade
3. Adicione uma nova seção em `IntegradorCompleto.ts`:
   ```typescript
   // SEÇÃO X: NOVA FUNCIONALIDADE
   function novaFuncao(workbook, inputs) {
     // implementação
   }
   ```
4. Adicione a ação no switch principal:
   ```typescript
   if (action === 'novaAcao') return novaFuncao(workbook, inputs);
   ```
5. Atualize a documentação

## 📞 Suporte

### Problemas Comuns

1. **Script não aparece no Power Automate**
   - Certifique-se de salvar o script no Excel Online
   - Aguarde alguns minutos para sincronização

2. **Erro "workbook is undefined"**
   - Certifique-se de passar o workbook corretamente
   - No Power Automate, use o conector do Excel Online

3. **Valores não formatam**
   - Verifique o formato das células (Texto vs Número)
   - Use função `parseValor()` do script

4. **Planilha não atualiza**
   - Verifique se a aba se chama exatamente "Documento"
   - Confirme que os índices de linha estão corretos

### Debug

Adicione console.log no script para debug:

```typescript
function buildPedidos(workbook: ExcelScript.Workbook): any {
  console.log('Iniciando buildPedidos');
  const sheet = workbook.getWorksheet('Documento');
  console.log('Sheet encontrada:', sheet !== null);
  // ... resto do código
}
```

Visualize os logs no editor do Office Scripts após executar.

## 📚 Recursos Adicionais

- [Documentação Completa](../../docs/GUIA_INTEGRADOR_COMPLETO.md)
- [Template Power Automate](./power-automate-flow-completo.json)
- [README Principal do Projeto](../../README.md)
- [Estrutura de Dados CSV](../../data/planilha-simulacao.csv)

## 🎓 Tutoriais

### Para Iniciantes
1. Leia o [Guia Completo](../../docs/GUIA_INTEGRADOR_COMPLETO.md)
2. Siga o Quick Start acima
3. Teste com o template do Power Automate

### Para Avançados
1. Explore os arquivos individuais
2. Customize as funcionalidades
3. Crie flows personalizados

## ✨ Próximos Passos

Após configurar o script:

1. ✅ Teste autenticação isoladamente
2. ✅ Teste validação com poucos registros
3. ✅ Crie um pedido de teste
4. ✅ Verifique os resultados na planilha
5. ✅ Configure notificações por email
6. ✅ Documente seu flow personalizado
7. ✅ Treine usuários finais

---

**💻 Desenvolvido para**: Excel Online + Power Automate + API Bimer  
**📅 Última atualização**: 24/12/2025  
**🔖 Versão**: 1.0

**Pronto para começar? Copie o arquivo [`IntegradorCompleto.ts`](./IntegradorCompleto.ts) agora! 🚀**
