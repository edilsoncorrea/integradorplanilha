// Bairro.ts
// Conversão de Bairro.gs: constrói descritor para criar um bairro via API.

import { HOST } from './Constantes';

export function main(workbook: any, inputs?: any): any {
  // inputs: { action: 'criarBairro', payload: { Nome: '...' }, token: '...' }
  const action = inputs && inputs.action ? inputs.action : 'criarBairro';

  if (action === 'criarBairro') {
    const payload = inputs.payload || { Nome: 'TESTE OFFICE SCRIPTS' };
    const token = inputs.token || '';
    const url = (HOST && HOST.endsWith('/')) ? HOST.replace(/\/$/, '') + '/api/bairros' : HOST + '/api/bairros';
    const headers: any = { 'Content-Type': 'application/json' };
    if (token && token.length > 0) headers['Authorization'] = 'Bearer ' + token;

    return {
      method: 'POST',
      url: url,
      headers: headers,
      body: payload
    };
  }

  return { error: 'unknown action' };
}
