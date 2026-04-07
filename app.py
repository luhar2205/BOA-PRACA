import os
import time
from datetime import datetime
from core.processor import rodar_automacao_total
from core.exporter import salvar_excel_kildere

# =====================================================================
# 🛡️ BLINDAGEM DE DIRETÓRIO (Fim do erro de caminho não encontrado)
# O script descobre automaticamente a pasta onde ele próprio está salvo
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_XML = os.path.join(BASE_DIR, "XML")
# IMPORTANTE: Se você mudou para underline antes, troque o nome abaixo para "PRODUTOS_KILDERE.xlsx"
BASE_SANKHYA = os.path.join(BASE_DIR, "PRODUTOS_KILDERE.xlsx")
PASTA_SAIDA = BASE_DIR
# =====================================================================

def iniciar():
    os.system('cls' if os.name == 'nt' else 'clear')
    print("="*60)
    print(" 🚀 MOTOR DE INTELIGÊNCIA SANKHYA UA - VERSÃO 4.0 (PRO)")
    print("="*60)
    
    # 🚨 TRAVA DE SEGURANÇA: Verifica se a planilha realmente existe antes de rodar
    if not os.path.exists(BASE_SANKHYA):
        print(f"\n❌ ERRO FATAL: A planilha matriz não foi encontrada!")
        print(f"O sistema procurou exatamente neste local:\n{BASE_SANKHYA}")
        print("\nVerifique se o nome do arquivo tem espaços, underlines ou se a extensão é mesmo .xlsx")
        return

    timestamp = datetime.now().strftime("%d_%m_%Hh%M")
    arquivo_saida = os.path.join(PASTA_SAIDA, f"Analise_Cruzamento_{timestamp}.xlsx")

    try:
        start_time = time.time()
        print("\n⚙️  Iniciando varredura e cruzamento de XMLs...")
        
        df = rodar_automacao_total(PASTA_XML, BASE_SANKHYA)
        salvar_excel_kildere(df, arquivo_saida)
        
        tempo_total = round(time.time() - start_time, 2)
        
        print(f"\n✅ SUCESSO ABSOLUTO!")
        print(f"📦 Total de itens mapeados: {len(df)}")
        print(f"⏱️  Tempo de execução: {tempo_total} segundos")
        print(f"📄 Relatório gerado em: {os.path.basename(arquivo_saida)}")
        print("="*60)
        
        os.startfile(arquivo_saida)
        
    except Exception as e:
        print(f"\n❌ FALHA CRÍTICA NO MOTOR: {str(e)}")
        print("Verifique os caminhos das pastas e a integridade do arquivo matriz.")

if __name__ == "__main__":
    iniciar()