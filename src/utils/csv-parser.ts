/**
 * Utilitário para converter CSV em dados de teste
 */

export function parseCSVToTestData(csvContent: string): any[][] {
  const lines = csvContent.split('\n');
  const data: any[][] = [];
  
  for (const line of lines) {
    if (line.trim()) {
      // Parse simples de CSV considerando aspas
      const row = parseCSVLine(line);
      data.push(row);
    }
  }
  
  return data;
}

function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"' && (i === 0 || line[i-1] === ',')) {
      inQuotes = true;
    } else if (char === '"' && inQuotes && (i === line.length - 1 || line[i+1] === ',')) {
      inQuotes = false;
    } else if (char === ',' && !inQuotes) {
      result.push(current.replace(/&quot;/g, '"').trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current.replace(/&quot;/g, '"').trim());
  return result;
}

// Dados CSV reais para facilitar testes
export const SAMPLE_CSV = `,,,,,,,,,,,,,,,,,,,,,,
Cd.empresa,Codigo Cliente,Nome do Cliente,Identificador do Cliente,Cd. da Operacao,Identificador da Operacao,CFOP,Cd. do serviço,Identificador do Serviço,Nome do serviço,Quantidade,Valor,Descriminação 1,Descriminação 2,Cd prazo,Identificador prazo,Forma Pagamento é entrada?,Código da F. pagamento,Identificador da  forma de Pagamento,Dt.Emissao,Vencimento,Nota Criada?,Retorno
000010,000001,ASPENTECH SOFTWARE BRASIL LTDA.,00A0000001,400061,00A000008J,6.102,000497,00A00000DX,TELECOMUNICAÇÃO,01,&quot;R$ 100,00&quot;,INVBRZ00022315,247035 JBS S.A.,,,,,,17/09/2025,17/09/2025,Sim,00A0000024
000012,000074,&quot;INTELIGENCIA DE NEGOCIOS, SISTEMAS E INFORMATICA L&quot;,00A0000023,200004,00A000000X,,000486,00A00000DK,LICENCIAMENTO OU CESSAO DE DIREITO DE USO DE PROGRAMAS DE COMPUTACAO,01,&quot;R$ 1.000,00&quot;,INVBRZ00022315,teste,000003,00A0000004,Não,000002,00A0000003,23/09/2025,23/10/2025,Sim,00A0000025
000012,000074,&quot;INTELIGENCIA DE NEGOCIOS, SISTEMAS E INFORMATICA L&quot;,00A0000023,200004,00A000000X,,000486,00A00000DK,LICENCIAMENTO OU CESSAO DE DIREITO DE USO DE PROGRAMAS DE COMPUTACAO,01,&quot;R$ 1.500,00&quot;,INVBRZ00022315,teste,000003,00A0000004,Não,000002,00A0000003,24/09/2025,24/10/2025,,`;