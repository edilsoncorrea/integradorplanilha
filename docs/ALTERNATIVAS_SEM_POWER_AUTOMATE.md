# 🚀 Alternativas para Integração com API Bimer (Sem Power Automate)

## ❌ Por que Office Scripts não funciona sozinho?

Office Scripts roda em **sandbox restrito** sem acesso a:
- ❌ `fetch()` ou `XMLHttpRequest`
- ❌ Bibliotecas externas (axios, node-fetch)
- ❌ Módulos Node.js
- ❌ Sistema de arquivos

## ✅ Soluções Viáveis

---

## 🎯 SOLUÇÃO 1: Script Node.js Local (RECOMENDADO)

### Vantagens
✅ Controle total da lógica  
✅ Usa bibliotecas modernas (axios)  
✅ Fácil de debugar  
✅ Pode rodar em servidor local ou agendado

### Como implementar

**Arquivo:** `scripts/criar-documentos-bimer.js`

```javascript
const axios = require('axios');
const ExcelJS = require('exceljs');

const HOST = 'https://homologacaowisepcp.alterdata.com.br/BimerApi';
const USERNAME = 'supervisor';
const SENHA = 'Senhas123';

async function main() {
  console.log('🚀 Iniciando integração com Bimer...\n');
  
  // 1. Carregar planilha
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('./planilha-documentos.xlsx');
  const sheet = workbook.getWorksheet('Documento');
  
  // 2. Autenticar
  const token = await autenticar();
  console.log('✅ Token obtido:', token.substring(0, 20) + '...\n');
  
  // 3. Processar linhas
  const resultados = [];
  
  for (let rowNum = 3; rowNum <= sheet.rowCount; rowNum++) {
    const row = sheet.getRow(rowNum);
    
    // Pular se já tem nota criada
    if (row.getCell(22).value) {
      console.log(`⏭️  Linha ${rowNum}: Já processada`);
      continue;
    }
    
    // Validar identificadores
    const identificadores = await validarIdentificadores(row, token);
    
    // Criar documento
    const payload = construirPayload(row);
    const resultado = await criarDocumento(payload, token);
    
    // Atualizar planilha
    row.getCell(22).value = resultado.sucesso ? 'Sim' : 'Não';
    row.getCell(23).value = resultado.mensagem;
    
    console.log(`${resultado.sucesso ? '✅' : '❌'} Linha ${rowNum}: ${resultado.mensagem}`);
    
    resultados.push(resultado);
  }
  
  // 4. Salvar planilha
  await workbook.xlsx.writeFile('./planilha-documentos.xlsx');
  
  console.log('\n📊 Resumo:');
  console.log(`Total: ${resultados.length}`);
  console.log(`Sucesso: ${resultados.filter(r => r.sucesso).length}`);
  console.log(`Erro: ${resultados.filter(r => !r.sucesso).length}`);
}

async function autenticar() {
  const crypto = require('crypto');
  const nonce = Date.now().toString();
  const password = crypto
    .createHash('md5')
    .update(USERNAME + nonce + SENHA)
    .digest('hex');
  
  const response = await axios.post(`${HOST}/oauth/token`, 
    new URLSearchParams({
      client_id: 'IntegracaoBimer.js',
      username: USERNAME,
      password: password,
      grant_type: 'password',
      nonce: nonce
    }),
    {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    }
  );
  
  return response.data.access_token;
}

async function validarIdentificadores(row, token) {
  // Buscar identificador do cliente se não existir
  if (!row.getCell(4).value) {
    const codigoCliente = row.getCell(2).value;
    const cliente = await axios.get(
      `${HOST}/api/pessoas/codigo/${codigoCliente}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    row.getCell(4).value = cliente.data.Identificador;
    row.getCell(3).value = cliente.data.Nome;
  }
  
  // Repetir para outros identificadores...
}

function construirPayload(row) {
  return {
    StatusNotaFiscalEletronica: 'A',
    TipoDocumento: 'S',
    TipoPagamento: '0',
    CodigoEmpresa: row.getCell(1).value,
    DataEmissao: row.getCell(20).value,
    DataReferencia: row.getCell(20).value,
    DataReferenciaPagamento: row.getCell(20).value,
    IdentificadorOperacao: row.getCell(6).value,
    IdentificadorPessoa: row.getCell(4).value,
    Itens: [{
      CFOP: row.getCell(7).value,
      IdentificadorProduto: row.getCell(9).value,
      Quantidade: row.getCell(11).value,
      ValorUnitario: parseFloat(row.getCell(12).value)
    }],
    Pagamentos: [{
      Aliquota: 100,
      DataVencimento: row.getCell(20).value,
      IdentificadorFormaPagamento: row.getCell(19).value,
      Valor: parseFloat(row.getCell(12).value)
    }]
  };
}

async function criarDocumento(payload, token) {
  try {
    const response = await axios.post(
      `${HOST}/api/documentos`,
      payload,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );
    
    const identificador = response.data.ListaObjetos?.[0]?.Identificador;
    return {
      sucesso: true,
      mensagem: identificador || 'Criado',
      dados: response.data
    };
  } catch (error) {
    return {
      sucesso: false,
      mensagem: error.response?.data?.Message || error.message,
      erro: error.response?.data
    };
  }
}

main().catch(console.error);
```

### Setup

```bash
# 1. Criar pasta do projeto
mkdir bimer-integration
cd bimer-integration

# 2. Inicializar Node.js
npm init -y

# 3. Instalar dependências
npm install axios exceljs dotenv

# 4. Criar arquivo .env
echo "BIMER_HOST=https://homologacaowisepcp.alterdata.com.br/BimerApi" > .env
echo "BIMER_USERNAME=supervisor" >> .env
echo "BIMER_SENHA=Senhas123" >> .env

# 5. Executar
node criar-documentos-bimer.js
```

---

## 🎯 SOLUÇÃO 2: Azure Function + Excel Add-in

### Arquitetura
```
[Excel] → [Azure Function (HTTP)] → [API Bimer] → [Retorna para Excel]
```

### Vantagens
✅ Sem instalar nada localmente  
✅ Excel chama via URL  
✅ Pode usar de qualquer lugar

### Azure Function (Node.js)

```javascript
// index.js
module.exports = async function (context, req) {
  const { action, data, token } = req.body;
  
  switch(action) {
    case 'autenticar':
      return await autenticar();
    
    case 'criarDocumento':
      return await criarDocumento(data, token);
    
    default:
      context.res = { status: 400, body: 'Action inválida' };
  }
};
```

### Chamada do Excel (Custom Function)

```typescript
// custom-functions.ts
/**
 * Criar documento no Bimer
 * @customfunction
 */
async function CRIAR_DOCUMENTO(
  codigoEmpresa: string,
  identificadorCliente: string,
  valor: number
): Promise<string> {
  const response = await fetch(
    'https://sua-function.azurewebsites.net/api/bimer',
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'criarDocumento',
        data: { codigoEmpresa, identificadorCliente, valor }
      })
    }
  );
  
  const result = await response.json();
  return result.identificador || result.erro;
}
```

---

## 🎯 SOLUÇÃO 3: Google Apps Script (Se puder usar Google Sheets)

### Vantagens
✅ Suporta UrlFetchApp (chamadas HTTP nativas)  
✅ Execução na nuvem  
✅ Triggers automáticos  
✅ Grátis

### Código

```javascript
function criarDocumentosBimer() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Autenticar
  const token = autenticarBimer();
  
  // Processar linhas
  for (let i = 2; i < data.length; i++) { // Pula cabeçalho
    const row = data[i];
    
    if (row[21]) continue; // Já processada
    
    const payload = {
      CodigoEmpresa: row[0],
      IdentificadorPessoa: row[3],
      // ... resto do payload
    };
    
    const resultado = criarDocumento(payload, token);
    
    // Atualizar célula
    sheet.getRange(i + 1, 22).setValue(resultado.sucesso ? 'Sim' : 'Não');
    sheet.getRange(i + 1, 23).setValue(resultado.mensagem);
  }
}

function autenticarBimer() {
  const nonce = new Date().getTime().toString();
  const password = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    'supervisor' + nonce + 'Senhas123'
  ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  
  const response = UrlFetchApp.fetch(
    'https://homologacaowisepcp.alterdata.com.br/BimerApi/oauth/token',
    {
      method: 'post',
      payload: {
        client_id: 'IntegracaoBimer.js',
        username: 'supervisor',
        password: password,
        grant_type: 'password',
        nonce: nonce
      }
    }
  );
  
  return JSON.parse(response.getContentText()).access_token;
}

function criarDocumento(payload, token) {
  try {
    const response = UrlFetchApp.fetch(
      'https://homologacaowisepcp.alterdata.com.br/BimerApi/api/documentos',
      {
        method: 'post',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload)
      }
    );
    
    const data = JSON.parse(response.getContentText());
    return {
      sucesso: true,
      mensagem: data.ListaObjetos[0].Identificador
    };
  } catch (error) {
    return {
      sucesso: false,
      mensagem: error.message
    };
  }
}
```

---

## 🎯 SOLUÇÃO 4: Python Script com openpyxl

### Vantagens
✅ Simples e direto  
✅ Ótimo para automação  
✅ Pode rodar em servidor

```python
import openpyxl
import requests
import hashlib

def criar_documentos():
    # Carregar planilha
    wb = openpyxl.load_workbook('planilha.xlsx')
    ws = wb['Documento']
    
    # Autenticar
    token = autenticar()
    
    # Processar linhas
    for row in ws.iter_rows(min_row=3, values_only=False):
        if row[21].value:  # Já processada
            continue
        
        payload = construir_payload(row)
        resultado = criar_documento(payload, token)
        
        row[21].value = 'Sim' if resultado['sucesso'] else 'Não'
        row[22].value = resultado['mensagem']
    
    wb.save('planilha.xlsx')

def autenticar():
    nonce = str(int(time.time()))
    password = hashlib.md5(
        f"supervisor{nonce}Senhas123".encode()
    ).hexdigest()
    
    response = requests.post(
        'https://homologacaowisepcp.alterdata.com.br/BimerApi/oauth/token',
        data={
            'client_id': 'IntegracaoBimer.js',
            'username': 'supervisor',
            'password': password,
            'grant_type': 'password',
            'nonce': nonce
        }
    )
    
    return response.json()['access_token']

def criar_documento(payload, token):
    try:
        response = requests.post(
            'https://homologacaowisepcp.alterdata.com.br/BimerApi/api/documentos',
            json=payload,
            headers={
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
        )
        
        data = response.json()
        return {
            'sucesso': True,
            'mensagem': data['ListaObjetos'][0]['Identificador']
        }
    except Exception as e:
        return {
            'sucesso': False,
            'mensagem': str(e)
        }

if __name__ == '__main__':
    criar_documentos()
```

---

## 📊 Comparação das Soluções

| Solução | Custo | Complexidade | Requer Servidor | Tempo Setup |
|---------|-------|--------------|-----------------|-------------|
| Node.js Local | Grátis | ⭐⭐ | ❌ | 10 min |
| Azure Function | ~$5/mês | ⭐⭐⭐ | ✅ | 30 min |
| Google Apps Script | Grátis | ⭐⭐ | ✅ (nuvem) | 15 min |
| Python Script | Grátis | ⭐ | ❌ | 10 min |

---

## 🎯 Minha Recomendação

**Para começar rápido: Script Node.js Local (Solução 1)**

Motivos:
1. ✅ Não precisa de servidor
2. ✅ Setup em 10 minutos
3. ✅ Fácil de debugar
4. ✅ Pode rodar manual ou agendado
5. ✅ Funciona offline

---

## 🚀 Próximos Passos

Escolha uma solução e eu te ajudo a implementar! Posso:

1. 📝 Criar o script completo e funcional
2. 🔧 Configurar o ambiente
3. 🐛 Ajudar com debugging
4. 📚 Documentar o processo

Qual solução você prefere?
