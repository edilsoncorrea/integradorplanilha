/**
 * Debug para testar a função ValidarPlanilha
 */

import { MockWorkbook } from '../src/mocks/excel-mock';
import { ValidarPlanilha, main } from '../src/office-scripts/ValidarPlanilhaCompleta';

async function debugValidarPlanilha() {
  console.log('🔍 DEBUG - VALIDAR PLANILHA\n');
  console.log('=' .repeat(60));
  
  // PASSO 1: Carregar dados simulados
  console.log('\n📊 PASSO 1: Carregando dados da planilha');
  const mockWorkbook = new MockWorkbook().loadRealData();
  const sheet = mockWorkbook.getWorksheet('Documento');
  
  if (!sheet) {
    console.log('❌ Planilha não encontrada');
    return;
  }
  
  const data = sheet.getUsedRange()?.getValues();
  console.log(`✅ Dados carregados: ${data?.length || 0} linhas`);
  
  // PASSO 2: Executar ValidarPlanilha
  console.log('\n🔍 PASSO 2: Executando ValidarPlanilha()');
  
  try {
    const result = ValidarPlanilha(mockWorkbook);
    
    console.log('\n📊 RESULTADO DA VALIDAÇÃO:');
    console.log(JSON.stringify(result, null, 2));
    
    // PASSO 3: Analisar resultados por tipo
    console.log('\n📋 PASSO 3: Análise detalhada por tipo');
    
    if (result.results) {
      Object.entries(result.results).forEach(([tipo, dados]: [string, any]) => {
        console.log(`\n${getIcon(tipo)} ${tipo.toUpperCase()}:`);
        console.log(`   Processados: ${dados.processed || 0}`);
        console.log(`   Consultas: ${dados.queries?.length || 0}`);
        
        if (dados.queries && dados.queries.length > 0) {
          console.log('   Exemplos:');
          dados.queries.slice(0, 2).forEach((query: any, index: number) => {
            console.log(`     ${index + 1}. Linha ${query.sheetRow}: ${query.endpoint}`);
          });
        }
      });
    }
    
    // PASSO 4: Testar via main() também
    console.log('\n🔧 PASSO 4: Testando via main()');
    const mainResult = main(mockWorkbook, { action: 'ValidarPlanilha' });
    
    console.log('Resultado via main():');
    console.log(`Sucesso: ${mainResult.success}`);
    console.log(`Mensagem: ${mainResult.message}`);
    
  } catch (error) {
    console.log(`\n❌ ERRO: ${error}`);
  }
  
  console.log('\n' + '=' .repeat(60));
  console.log('🎯 Debug ValidarPlanilha concluído!');
}

function getIcon(tipo: string): string {
  const icons: { [key: string]: string } = {
    'consultarCliente': '👤',
    'consultarFormaPagamento': '💳',
    'consultarOperacao': '⚙️',
    'consultarServico': '🛠️',
    'consultarPrazo': '📅'
  };
  return icons[tipo] || '📋';
}

debugValidarPlanilha().catch(console.error);