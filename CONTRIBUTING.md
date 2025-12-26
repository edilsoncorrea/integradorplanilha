````md
# CONTRIBUTING

Este projeto utiliza **Office Scripts (Excel + TypeScript)** conforme a especificação oficial da Microsoft.  
Office Scripts **não é TypeScript completo**, mas **possui APIs próprias de runtime**, incluindo **suporte oficial a `fetch`**.

Qualquer contribuição (humana ou assistida por IA) **deve seguir rigorosamente este documento**.

---

## 1. Papel do Autor do Código

Ao escrever ou sugerir código para este projeto, assuma **sempre** o papel de:

**Especialista em Office Scripts da Microsoft**

Isso implica:
- Conhecer as **APIs reais disponíveis** no ambiente Office Scripts
- Não negar funcionalidades **explicitamente suportadas pela Microsoft**
- Diferenciar corretamente:
  - APIs inexistentes  
  - APIs restritas  
  - APIs suportadas com limitações

---

## 2. Princípios Invioláveis

### 2.1 Código funcional não deve ser reinterpretado como erro de plataforma

Se um recurso:
- Já foi testado
- Já funciona no Office Scripts
- Está documentado ou validado empiricamente

👉 **É proibido afirmar que ele “não é suportado” sem prova concreta.**

Isso inclui, mas não se limita a:
- `fetch`
- Promises
- APIs HTTP do ambiente Office Scripts

---

### 2.2 Office Scripts ≠ TypeScript ≠ Node.js

Office Scripts:
- **Não é Node.js**
- **Não é navegador**
- **Não é TypeScript puro**

Mas **possui runtime próprio**, com APIs específicas fornecidas pela Microsoft.

É proibido:
- Assumir ausência de API apenas por não existir em Node ou Browser
- Tratar Office Scripts como ambiente “limitado por padrão”
- Negar APIs funcionais baseando-se apenas em suposições

---

## 3. Uso de `fetch` (Regra Específica e Explícita)

### ✅ `fetch` é SUPORTADO em Office Scripts

Este projeto **autoriza explicitamente** o uso de:

```ts
fetch(url, options)
````

Desde que:

* Seja usado apenas para HTTP/HTTPS
* Não dependa de objetos de navegador (window, document, DOM)
* Tenha tratamento de erro explícito
* Respeite tipagem rigorosa

❌ É proibido sugerir que `fetch` “não é suportado” ou “não existe” no Office Scripts.

❌ É proibido tentar “resolver” removendo `fetch` quando ele já funciona.

---

## 4. Tipagem é Obrigatória (Regra Absoluta)

Declarações explícitas de tipo são **inquebrantáveis**.

Sempre tipar:

* Variáveis
* Parâmetros
* Retornos
* Promises
* Respostas de `fetch`

Exemplo obrigatório:

```ts
let response: Response = await fetch(url);
let data: unknown = await response.json();
```

---

## 5. Funções

### 5.1 Assinaturas explícitas

Toda função deve declarar tipo de retorno, inclusive funções assíncronas.

```ts
async function carregarDados(url: string): Promise<unknown> {
}
```

---

### 5.2 Função `main`

A função `main` deve respeitar **exatamente** a assinatura do Office Scripts.

```ts
function main(workbook: ExcelScript.Workbook): void {
}
```

---

## 6. Uso das APIs ExcelScript

* Utilizar apenas APIs documentadas
* Não encadear chamadas de forma agressiva
* Priorizar clareza sobre concisão

---

## 7. Alterações em Código Existente

Antes de alterar código que:

* Usa `fetch`
* Usa Promises
* Usa APIs já testadas

É obrigatório:

1. Verificar se o código **já funciona**
2. Não reinterpretar funcionamento correto como “erro de plataforma”
3. Não “corrigir” algo válido com base em suposição

---

## 8. Comentários e Intenção

Quando usar APIs que:

* Parecem incomuns
* São frequentemente negadas por engano (ex: `fetch`)

Adicionar comentário explicando:

* Que a API é suportada no Office Scripts
* Que o uso é intencional

---

## 9. Compatibilidade

O código deve funcionar:

* No Excel Online
* No Excel Desktop
* No runtime oficial do Office Scripts

---

## 10. Regra Final

É proibido “resolver” problemas **negando capacidades reais da plataforma**.

Se o código funciona no Office Scripts:
👉 o problema **não é a existência da API**, mas sim a implementação.

Office Scripts é o ambiente.
TypeScript é apenas a linguagem de apoio.

```
```
