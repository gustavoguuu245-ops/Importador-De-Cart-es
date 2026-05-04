import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# 1. Configurações e Mapeamento cnpjs teste 
MAPA_EMPRESAS = {
    '11.111.111/1111-11': {'cod': 601, 'banco': 20,  'nat rotativo': '1146', 'nat mensal': '1147'},
    '22.222.222/2222-22': {'cod': 602, 'banco': 20,  'nat rotativo': '1146', 'nat mensal': '1147'},
    '33.333.333/3333-33': {'cod': 301, 'banco': 19,  'nat rotativo': '1146'}, 
    '44.444.444/4444-44': {'cod': 304, 'banco': 19,  'nat rotativo': '221',  'nat mensal': '219'},
    '55.555.555/5555-55': {'cod': 15,  'banco': 701, 'nat rotativo': '1136', 'nat mensal': '1137'},
    '66.666.666/6666-66': {'cod': 400, 'banco': 17,  'nat rotativo': '1146'}
}
BANDEIRAS = {
    'DEBITO ELO': 9, 'DEBITO MASTER': 8, 'DEBITO VISA': 7,
    'CREDITO ELO': 14, 'CREDITO MASTER': 11, 'CREDITO VISA': 10
}

def processar_importacao(caminho_input, cnpj_selecionado):
    try:
        df = pd.read_excel(caminho_input)
        df.columns = df.columns.str.strip()

        # Filtra apenas a empresa selecionada
        df_filtrado = df[df['DOCUMENTO'] == cnpj_selecionado].copy()
        
        # Guardamos o total real para conferência
        total_real = df_filtrado['VALOR BRUTO'].sum()

        # Tratamento de Datas
        df_filtrado['EMISSAO_DT'] = pd.to_datetime(df_filtrado['DATA DA VENDA'], dayfirst=True).dt.date
        df_filtrado['VENC_DT'] = pd.to_datetime(df_filtrado['DATA DE VENCIMENTO ORIGINAL'], dayfirst=True).dt.date

        info = MAPA_EMPRESAS[cnpj_selecionado]
        
        # IDENTIFICAÇÃO DA BANDEIRA E PGTO PELO PRODUTO
        def identificar_pgto_completo(linha):
            band_original = str(linha['BANDEIRA']).upper()
            produto = str(linha['PRODUTO']).upper()
            
            # Normaliza Bandeira
            if 'MASTER' in band_original: band = 'MASTER'
            elif 'VISA' in band_original: band = 'VISA'
            elif 'ELO' in band_original: band = 'ELO'
            else: band = band_original
            
            # Define se é Crédito ou Débito pela coluna PRODUTO da Stone
            tipo = 'CREDITO' if 'CRÉDITO' in produto or 'CREDITO' in produto else 'DEBITO'
            
            chave = f"{tipo} {band}"
            cod = BANDEIRAS.get(chave, "")
            return str(cod).replace('.0', '') if cod != "" else ""

        df_filtrado['PGTO_FINAL'] = df_filtrado.apply(identificar_pgto_completo, axis=1)

        # Lógica de Natureza (Mensalista por valor ainda, mas agrupado por PGTO)
        df_filtrado['eh_mensalista'] = df_filtrado['VALOR BRUTO'] > 110
        
        def define_nat(linha):
            if linha['eh_mensalista']:
                return info.get('nat mensal', info['nat rotativo'])
            return info['nat rotativo']

        df_filtrado['NAT_FINAL'] = df_filtrado.apply(define_nat, axis=1)

        # Agrupamento Seguro: Incluindo PGTO_FINAL para separar Crédito de Débito no mesmo dia
        col_agrupar = ['EMISSAO_DT', 'VENC_DT', 'BANDEIRA', 'NAT_FINAL', 'PGTO_FINAL']
        resultado = df_filtrado.groupby(col_agrupar)['VALOR BRUTO'].sum().reset_index()

        #  MONTAGEM DO LAYOUT FINAL 
        df_final = pd.DataFrame()
        df_final['EMPRESA'] = [info['cod']] * len(resultado)
        df_final['DATA DE EMISSÃO'] = resultado['EMISSAO_DT']
        df_final['DATA DE VENCIMENTO'] = resultado['VENC_DT']
        df_final['TIPO'] = "Debito"
        df_final['BANDEIRA'] = resultado['BANDEIRA']
        df_final['STONE ID'] = ""
        df_final['N° CARTÃO'] = ""
        
        # H - VALOR BRUTO (Vírgula e sem .0)
        df_final['VALOR BRUTO'] = resultado['VALOR BRUTO'].apply(lambda x: '{0:g}'.format(float(round(x, 2))).replace('.', ','))
        
        df_final['NR TITULO'] = ""
        
        # J - PGTO (Sem .0)
        df_final['PGTO'] = resultado['PGTO_FINAL']
        
        df_final['BANCO'] = info['banco']
        df_final['NATUREZA'] = resultado['NAT_FINAL']
        df_final['PESSOA'] = 16
        df_final['NOME DO TITULO'] = "STONE PAGAMENTOS"
        df_final['OBSERVAÇÃO'] = "ROTATIVO"
        df_final['ALIQUOTA'] = ""

        user_profile = os.environ.get('USERPROFILE') or os.path.expanduser('~')
        desktop = os.path.join(user_profile, 'Desktop')
        nome_arquivo = f"IMPORT_EMPRESA_{info['cod']}.csv"
        caminho_final = os.path.join(desktop, nome_arquivo)

        df_final.to_csv(caminho_final, index=False, sep=';', date_format='%d/%m/%Y', encoding='latin-1')
        
        messagebox.showinfo("Sucesso!", f"Conferência Stone: R$ {total_real:.2f}\nArquivo salvo com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro: {e}")

def iniciar_sistema():
    root = tk.Tk()
    root.withdraw()
    caminho = filedialog.askopenfilename(title="Selecione a Planilha da Stone", filetypes=[("Excel", "*.xlsx *.xls")])
    if not caminho: return

    try:
        temp_df = pd.read_excel(caminho)
        temp_df.columns = temp_df.columns.str.strip()
        cnpjs = [c for c in temp_df['DOCUMENTO'].unique() if c in MAPA_EMPRESAS]
    except:
        messagebox.showerror("Erro", "Falha ao ler o arquivo.")
        return

    janela_sel = tk.Toplevel()
    janela_sel.title("Seletor de Empresa")
    janela_sel.geometry("350x200")
    
    tk.Label(janela_sel, text="Escolha a empresa:", pady=10).pack()
    seletor = ttk.Combobox(janela_sel, values=cnpjs, width=30, state="readonly")
    seletor.pack(pady=10)
    if cnpjs: seletor.current(0)

    tk.Button(janela_sel, text="GERAR IMPORTAÇÃO", bg="#2ecc71", fg="white", font=("Arial", 10, "bold"),
              command=lambda: [processar_importacao(caminho, seletor.get()), janela_sel.destroy()]).pack(pady=20)

    janela_sel.mainloop()

if __name__ == "__main__":
    iniciar_sistema()