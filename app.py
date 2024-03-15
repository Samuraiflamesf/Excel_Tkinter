import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Variáveis globais para armazenar os arquivos selecionados
arquivo1 = None
arquivo2 = None
lbl_arquivo1 = None
lbl_arquivo2 = None


def carregar_planilha(nome_arquivo):
    try:
        # Carregar a planilha usando pandas
        if nome_arquivo == arquivo1:
            planilha = pd.read_excel(nome_arquivo, skiprows=1)
            return planilha
        if nome_arquivo == arquivo2:
            planilha = pd.read_excel(nome_arquivo)
            return planilha
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo não encontrado.")
        return None


def acompGuia():
    global arquivo1
    arquivo1 = filedialog.askopenfilename(
        title="Selecione a planilha AcompGuia")
    lbl_arquivo1['text'] = f"{arquivo1}"
    return carregar_planilha(arquivo1)


def relacaoSaidaData():
    global arquivo2
    arquivo2 = filedialog.askopenfilename(
        title="Selecione a planilha relacaoSaidaData")
    lbl_arquivo2['text'] = f"{arquivo2}"
    return carregar_planilha(arquivo2)


def reset_arquivos():
    global arquivo1, arquivo2
    arquivo1 = None
    arquivo2 = None
    lbl_arquivo1['text'] = "AcompGuia: "
    lbl_arquivo2['text'] = "RelacaoSaidaData: "
    messagebox.showinfo("Informação", "Arquivos resetados com sucesso!")


def app():
    global arquivo1, arquivo2

    if arquivo1 is None or arquivo2 is None:
        messagebox.showerror('Erro', 'Selecione ambas as planilhas!')
        return
    try:
        planilha1 = carregar_planilha(arquivo1)
        planilha2 = carregar_planilha(arquivo2)
        messagebox.showinfo("Informação", "Planilhas carregadas com sucesso!")

        # Verificar se as colunas 'Guia' e 'Destino' existem em planilha1
        if 'Guia' not in planilha1.columns or 'Destino' not in planilha1.columns:
            messagebox.showerror(
                'Erro', 'Colunas necessárias não encontradas na planilha AcompGuia!')
            return

        # Exibir as primeiras linhas das planilhas
        planilha1 = planilha1[['Guia', 'Destino']]
        planilha1['Guia'] = planilha1['Guia'].str.replace(
            '.', '').str.replace('/', '')

        # Verificar se a coluna 'num_documento' existe em planilha2
        if 'num_documento' not in planilha2.columns:
            messagebox.showerror(
                'Erro', 'Coluna necessária não encontrada na planilha relacaoSaidaData!')
            return

        planilha2['num_documento'] = planilha2['num_documento'].str.replace(
            'Nº ', '')
        planilha2 = planilha2.rename(columns={'num_documento': 'Guia'})
        planilha2["Guia"] = planilha2["Guia"].str.strip()
        app_excel = pd.merge(planilha2, planilha1, on='Guia', how='left')
        app_excel = app_excel.drop_duplicates()
        app_excel = app_excel[['Guia', 'Destino', 'Item', 'Nome Item', 'unid_med_ent',
                              'data_lancamento', 'Desc_Mov', 'RMRS', 'qtde', 'Valor Unitário', 'Valor Total']]

        # Salvando o arquivo mesclado
        arquivo_saida = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if arquivo_saida:
            app_excel.to_excel(arquivo_saida, index=False)
            messagebox.showinfo("Informação", "Arquivo salvo com sucesso!")

    except FileNotFoundError:
        messagebox.showinfo('Erro', 'Selecione as duas planilhas!')


def fechar_janela(janela):
    if messagebox.askokcancel("Fechar", "Tem certeza que deseja fechar?"):
        janela.destroy()


def main():
    janela = tk.Tk()
    janela.title("Cruzamento Almox")

    # Definindo a margem
    margem = 5

    # Titulo
    text1 = tk.Label(janela, text="App desenvolvido para cruzar planilhas:")
    text1.grid(row=0, padx=margem, sticky='w')

    text2 = tk.Label(janela, text="AcompGuia e relacaoSaidaData")
    text2.grid(row=1, padx=margem, pady=(0, 20), sticky='w')

    # Rótulos para exibir o nome do arquivo importado
    global lbl_arquivo1, lbl_arquivo2
    lbl_arquivo1 = tk.Label(janela, text="AcompGuia: ")
    lbl_arquivo1.grid(row=3, columnspan=3, padx=margem,
                      pady=margem, sticky='w')

    lbl_arquivo2 = tk.Label(janela, text="RelacaoSaidaData: ")
    lbl_arquivo2.grid(row=6, columnspan=3, padx=margem,
                      pady=margem, sticky='w')

    # Botões e texto para carregar planilhas
    text3 = tk.Label(janela, text='Selecione planilha AcompGuia:')
    text3.grid(row=2, padx=margem, pady=margem, sticky='w')

    btn1 = tk.Button(janela, text="AcompGuias", command=acompGuia)
    btn1.grid(row=4, padx=margem, pady=margem, sticky='w')

    text4 = tk.Label(janela, text='Selecione planilha relacaoSaidaData:')
    text4.grid(row=5, padx=margem, pady=margem, sticky='w')

    btn2 = tk.Button(janela, text="RelaçãodeSaídaporData",
                     command=relacaoSaidaData)
    btn2.grid(row=7, padx=margem, pady=(0, 30), sticky='w')

    # Botão para executar o aplicativo e sair
    btn_app = tk.Button(janela, text="Executar", command=app)
    btn_app.grid(row=8, column=0, padx=margem, pady=margem, sticky='w')

    # Botão para resetar os arquivos
    btn_reset = tk.Button(janela, text="Resetar Arquivos",
                          command=reset_arquivos)
    btn_reset.grid(row=8, column=1, padx=margem, pady=margem, sticky='w')

    btn_fechar = tk.Button(janela, text="Sair",
                           command=lambda: fechar_janela(janela))
    btn_fechar.grid(row=8, column=2, padx=margem, pady=margem, sticky='w')

    janela.mainloop()


if __name__ == "__main__":
    main()
