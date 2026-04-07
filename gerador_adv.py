import pandas as pd
import warnings
import re
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog

# Silenciando os avisos do terminal
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=UserWarning, module='xlsxwriter')

class RadarADVApp:
    def __init__(self):
        self.timestamp = datetime.now().strftime("%d_%m_%Hh%M")
        self.arquivo_saida = f"Dashboard_ADV_V5_{self.timestamp}.xlsx"
        
        self.dicionario_trocas = {
            "CORDEIRO": "Dos {qtd} faltantes, podemos completar com PERNIL OU PALETA?",
            "DOURADO": "Dos {qtd} faltantes, podemos completar com FILÉ DE PESCADA AMARELA?",
            "PESCADA AMARELA": "Dos {qtd} faltantes, podemos substituir por FILÉ DE PESCADA BRANCA?",
            "TILAPIA": "Faltaram {qtd}, podemos completar com o nosso FILÉ DE TILÁPIA?",
            "COUVE DE BRUXELAS": "CORTE. Não conseguimos. Substituir por BRÓCOLIS CONGELADO OU REPOLHO?",
            "MORTADELA": "Em ruptura. Podemos enviar MORTADELA DEFUMADA no lugar?"
        }

    def limpar_aba(self, nome):
        return re.sub(r'[\\/*?:\[\]]', '', str(nome))[:31]

    def gerar_texto(self, desc, qtd_faltante):
        desc_upper = str(desc).upper()
        qtd = f"{qtd_faltante:g}".replace(".", ",")
        
        for chave, texto in self.dicionario_trocas.items():
            if chave in desc_upper:
                if any(x in chave for x in ["CORDEIRO", "DOURADO", "PESCADA", "TILAPIA"]):
                    return f"{desc_upper} -> {texto.format(qtd=qtd+'kg')}"
                return f"{desc_upper} -> {texto.format(qtd=qtd)}"
                
        return f"{desc_upper} -> Faltam {qtd}. Informo o CORTE. Favor indicar substituto."

    def carregar_dados(self, caminho):
        print("📥 Extraindo e lapidando dados do Sankhya...")
        df = pd.read_excel(caminho, header=2)
        df.columns = [str(c).strip() for c in df.columns]
        
        for col in ['Navio', 'Descrição', 'Categoria', 'Qtd.Pedido', 'Estoque', 'Comprar', 'Comprador']:
            if col not in df.columns: df[col] = ""

        for col in ['Comprar', 'Qtd.Pedido', 'Estoque']:
            df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df_pendentes = df[df['Comprar'] > 0].copy()
        if df_pendentes.empty:
            return None
            
        df_pendentes = df_pendentes.sort_values(by=['Navio', 'Descrição'])
        df_pendentes['MENSAGEM ADV'] = df_pendentes.apply(lambda r: self.gerar_texto(r['Descrição'], r['Comprar']), axis=1)
        
        return df_pendentes

    def construir_excel(self, df):
        print("📊 Gerando Interface UX/UI (Modo Power BI)...")
        
        writer = pd.ExcelWriter(self.arquivo_saida, engine='xlsxwriter')
        workbook = writer.book

        # ================= UX/UI DESIGN SYSTEM =================
        FONTE = 'Segoe UI' # Fonte padrão de Dashboards e Windows
        COR_FUNDO = '#F3F4F6'
        
        # Estilos dos KPIs (Cartões do Topo)
        f_kpi_titulo = workbook.add_format({'font_name': FONTE, 'font_size': 10, 'font_color': '#6B7280', 'align': 'center', 'valign': 'vcenter', 'bg_color': 'white', 'top': 1, 'left': 1, 'right': 1, 'border_color': '#D1D5DB'})
        f_kpi_valor = workbook.add_format({'font_name': FONTE, 'font_size': 22, 'font_color': '#111827', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': 'white', 'bottom': 1, 'left': 1, 'right': 1, 'border_color': '#D1D5DB'})
        
        # Estilos da Tabela
        f_header = workbook.add_format({'font_name': FONTE, 'bg_color': '#1F2937', 'font_color': 'white', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#374151'})
        f_navio_group = workbook.add_format({'font_name': FONTE, 'bg_color': '#E5E7EB', 'font_color': '#111827', 'bold': True, 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#D1D5DB'})
        
        f_texto = workbook.add_format({'font_name': FONTE, 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#E5E7EB'})
        f_texto_sub = workbook.add_format({'font_name': FONTE, 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#E5E7EB', 'indent': 1, 'font_color': '#4B5563'})
        f_num = workbook.add_format({'font_name': FONTE, 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#E5E7EB'})
        f_alerta = workbook.add_format({'font_name': FONTE, 'bg_color': '#FEE2E2', 'font_color': '#991B1B', 'bold': True, 'align': 'center', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#E5E7EB'})
        
        f_msg = workbook.add_format({'font_name': FONTE, 'bg_color': '#FEF3C7', 'font_color': '#92400E', 'bold': True, 'align': 'left', 'valign': 'vcenter', 'bottom': 1, 'border_color': '#E5E7EB'})
        
        f_dropdown = workbook.add_format({'font_name': FONTE, 'bg_color': '#D1FAE5', 'font_color': '#065F46', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#A7F3D0', 'locked': False})
        f_fantasma = workbook.add_format({'font_name': FONTE, 'bg_color': '#F3F4F6', 'font_color': '#374151', 'align': 'center', 'valign': 'vcenter', 'bold': True, 'locked': True, 'hidden': True, 'bottom': 1, 'border_color': '#E5E7EB'})

        # ================= ABA 1: DASHBOARD POWER BI =================
        nome_aba_dash = "📊 PAINEL GERAL"
        ws_dash = workbook.add_worksheet(nome_aba_dash)
        ws_dash.hide_gridlines(2)
        ws_dash.set_tab_color('#1F2937') # Cor da aba
        
        # Background cinza clarinho tipo Power BI
        ws_dash.set_column('A:Z', None, workbook.add_format({'bg_color': COR_FUNDO}))

        # Calculando KPIs
        total_navios = df['Navio'].nunique()
        total_itens_falta = len(df)
        total_volume_kg = df['Comprar'].sum()

        # Desenhando Cartões KPI
        ws_dash.write('B2', 'NAVIOS COM RUPTURA', f_kpi_titulo)
        ws_dash.write('B3', f"{total_navios}", f_kpi_valor)
        
        ws_dash.write('D2', 'TOTAL DE ITENS PENDENTES', f_kpi_titulo)
        ws_dash.write('D3', f"{total_itens_falta}", f_kpi_valor)
        
        ws_dash.write('F2', 'VOLUME TOTAL EM FALTA', f_kpi_titulo)
        ws_dash.write('F3', f"{total_volume_kg:g}".replace('.', ','), f_kpi_valor)
        
        ws_dash.set_row(1, 20)
        ws_dash.set_row(2, 40) # Aumenta a altura da linha dos números

        # Cabeçalhos da Matriz de Dados (Drill-down)
        cabecalhos_dash = ['NAVIO / DESCRIÇÃO DO PRODUTO', 'ESTOQUE', 'PEDIDO', 'FALTA COMPRAR', 'MENSAGEM PARA O ADV (Copie e Cole)']
        linha = 5
        for col_num, value in enumerate(cabecalhos_dash):
            ws_dash.write(linha, col_num + 1, value, f_header)
        
        ws_dash.set_column('A:A', 3)  # Margem esquerda
        ws_dash.set_column('B:B', 50) # Produto
        ws_dash.set_column('C:E', 15) # Números
        ws_dash.set_column('F:F', 95) # Mensagem

        linha += 1
        
        # 🔥 O RETORNO DO DRILL-DOWN (Agrupamento no XlsxWriter)
        for navio, group in df.groupby('Navio'):
            # Linha Mestra (Navio)
            ws_dash.write(linha, 1, f"🚢 {navio}", f_navio_group)
            for c in range(2, 6): ws_dash.write_blank(linha, c, "", f_navio_group)
            
            # Formata a linha como PAI (nível 0)
            ws_dash.set_row(linha, 25, None, {'level': 0})
            linha += 1
            
            # Linhas Filhas (Produtos)
            for _, row in group.iterrows():
                ws_dash.write(linha, 1, f"↳ {row['Descrição']}", f_texto_sub)
                ws_dash.write(linha, 2, row['Estoque'], f_num)
                ws_dash.write(linha, 3, row['Qtd.Pedido'], f_num)
                ws_dash.write(linha, 4, row['Comprar'], f_alerta) # Destaque na falta
                ws_dash.write(linha, 5, row['MENSAGEM ADV'], f_msg)
                
                # 🔥 ESCONDE AS LINHAS FILHAS (Cria o '+' no Excel)
                ws_dash.set_row(linha, 20, None, {'level': 1, 'hidden': True})
                linha += 1

        # ================= ABAS DOS NAVIOS (Sem Visual) =================
        for navio in df['Navio'].unique():
            nome_aba = self.limpar_aba(navio)
            df_navio = df[df['Navio'] == navio].copy()
            ws = workbook.add_worksheet(nome_aba)
            ws.hide_gridlines(2)
            ws.freeze_panes(1, 0)

            # Cabeçalhos Limpos (Sem 'Visual')
            cabecalhos = ['PRODUTO', 'CATEGORIA', 'ESTOQUE', 'PEDIDO', 'FALTA', 'MENSAGEM PRO ADV (Copie)', 'RISCO (%)', 'STATUS (Selecione)']
            for col_num, value in enumerate(cabecalhos):
                ws.write(0, col_num, value, f_header)

            ws.set_column('A:A', 45)
            ws.set_column('B:B', 18)
            ws.set_column('C:E', 12)
            ws.set_column('F:F', 95)
            ws.set_column('G:G', 15)
            ws.set_column('H:H', 22) # Dropdown subiu de posição

            dados = df_navio[['Descrição', 'Categoria', 'Estoque', 'Qtd.Pedido', 'Comprar', 'MENSAGEM ADV']].values.tolist()

            for row_num, linha_dados in enumerate(dados):
                linha_excel = row_num + 1
                
                ws.write(linha_excel, 0, linha_dados[0], f_texto)
                ws.write(linha_excel, 1, linha_dados[1], f_texto)
                ws.write(linha_excel, 2, linha_dados[2], f_num)
                ws.write(linha_excel, 3, linha_dados[3], f_num)
                ws.write(linha_excel, 4, linha_dados[4], f_alerta) # Fundo salmão para a falta
                ws.write(linha_excel, 5, linha_dados[5], f_msg)

                # Fórmula Oculta (Nível Risco)
                formula_risco = f'=IF(E{linha_excel+1}>=D{linha_excel+1}, "ALERTA 100%", TEXT(E{linha_excel+1}/D{linha_excel+1}, "0%") & " CORTE")'
                ws.write_formula(linha_excel, 6, formula_risco, f_fantasma)

                # Célula Desbloqueada para o Dropdown (Agora na coluna H)
                ws.write_blank(linha_excel, 7, "", f_dropdown)

            # Data Validation (Dropdown) sem Emojis que corrompem
            ws.data_validation(1, 7, len(dados), 7, {
                'validate': 'list',
                'source': ['Aguardando ADV', 'Aprovado', 'Recusado', 'Resolvido no Estoque'],
                'input_title': 'Status',
                'input_message': 'Selecione o status atual.'
            })

            ws.autofilter(0, 0, len(dados), len(cabecalhos)-1)

            ws.protect('kildere123', {
                'select_locked_cells': True,
                'select_unlocked_cells': True,
                'format_columns': True,
                'autofilter': True,
                'sort': True
            })

        writer.close()
        print(f"✅ SUCESSO! Interface Power BI gerada: {self.arquivo_saida}")
        os.startfile(self.arquivo_saida)

    def executar(self):
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        print("🤖 Radar ADV V-5.0 (UX Premium) Iniciado...")
        caminho = filedialog.askopenfilename(title="Selecione a Análise de Navios", filetypes=[("Excel", "*.xlsx *.xls")])
        
        if caminho:
            df = self.carregar_dados(caminho)
            if df is not None:
                self.construir_excel(df)
            else:
                print("👍 Nenhum item pendente de compra encontrado.")

if __name__ == "__main__":
    app = RadarADVApp()
    app.executar()