// ConsumirAPI.ts
// Script helper: constrói descritores de requisição HTTP para Power Automate.
// Não realiza fetch diretamente. O Power Automate (ou outro executor) deve usar esses descriptors.

import { HOST } from './Constantes';

export function main(workbook: any, inputs?: any): any {
  const action = inputs && inputs.action ? inputs.action : 'buildPost';

  if (action === 'buildPost') {
    const endpoint = inputs.endpoint || '/api/';
    const payload = inputs.payload || {};
    const token = inputs.token || '';
    const url = (HOST && HOST.endsWith('/')) ? HOST.replace(/\/$/, '') + endpoint : HOST + endpoint;
    const headers: any = { 'Content-Type': 'application/json' };
    if (token && token.length > 0) headers['Authorization'] = 'Bearer ' + token;

    return {
      method: 'POST',
      url: url,
      headers: headers,
      body: payload
    };
  }

  if (action === 'buildGet') {
    const endpoint = inputs.endpoint || '/api/';
    const token = inputs.token || '';
    const url = (HOST && HOST.endsWith('/')) ? HOST.replace(/\/$/, '') + endpoint : HOST + endpoint;
    const headers: any = { 'Content-Type': 'application/json' };
    if (token && token.length > 0) headers['Authorization'] = 'Bearer ' + token;

    return {
      method: 'GET',
      url: url,
      headers: headers
    };
  }

  return { error: 'unknown action' };
}
