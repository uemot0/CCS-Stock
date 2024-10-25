import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

# Define o caminho da pasta onde as planilhas serão armazenadas
planilha_dir = r'C:\Users\uemot\Desktop\ControleEstoque\planilhas'

# Função para inicializar as planilhas, se não existirem
def inicializar_planilhas():
    if not os.path.exists(planilha_dir):
        os.makedirs(planilha_dir)
        print(f"Pasta '{planilha_dir}' criada.")
    if not os.path.exists(os.path.join(planilha_dir, 'estoque.xlsx')):
        pd.DataFrame(columns=['Quantidade', 'Nome Produto', 'Código', 'Data de Validade', 'Loja']).to_excel(os.path.join(planilha_dir, 'estoque.xlsx'), index=False)
        print("Planilha 'estoque.xlsx' criada.")
    if not os.path.exists(os.path.join(planilha_dir, 'brooklin.xlsx')):
        pd.DataFrame(columns=['Quantidade', 'Nome Produto', 'Código', 'Data de Validade', 'Loja']).to_excel(os.path.join(planilha_dir, 'brooklin.xlsx'), index=False)
        print("Planilha 'brooklin.xlsx' criada.")
    if not os.path.exists(os.path.join(planilha_dir, 'eucaliptos.xlsx')):
        pd.DataFrame(columns=['Quantidade', 'Nome Produto', 'Código', 'Data de Validade', 'Loja']).to_excel(os.path.join(planilha_dir, 'eucaliptos.xlsx'), index=False)
        print("Planilha 'eucaliptos.xlsx' criada.")

# Função para resetar e atualizar o inventário
def fazer_inventario(dados):
    df = pd.DataFrame(dados, columns=['Quantidade', 'Nome Produto', 'Código', 'Data de Validade', 'Loja'])
    df.to_excel(os.path.join(planilha_dir, 'estoque.xlsx'), index=False)
    messagebox.showinfo("Sucesso", "Inventário atualizado com sucesso!")

# Função para adicionar nova carga
def nova_carga(dados):
    top = tk.Toplevel()
    top.title("Adicionar Nova Carga")

    labels = ['Quantidade', 'Nome Produto', 'Código', 'Data de Validade (AAAA-MM-DD)', 'Loja']
    entradas = {}

    for i, label in enumerate(labels):
        lbl = tk.Label(top, text=label)
        lbl.grid(row=i, column=0, padx=10, pady=5)
        entrada = tk.Entry(top)
        entrada.grid(row=i, column=1, padx=10, pady=5)
        entradas[label] = entrada

    def adicionar_ao_estoque():
        produto = [entradas[label].get() for label in labels]
        df = pd.read_excel(os.path.join(planilha_dir, 'estoque.xlsx'))

        # Verifica se o produto já existe no estoque
        produto_existente = df[df['Nome Produto'] == produto[1]]

        if not produto_existente.empty:
            # Adiciona a quantidade ao produto existente
            df.loc[df['Nome Produto'] == produto[1], 'Quantidade'] += int(produto[0])
        else:
            # Adiciona uma nova linha se o produto não existir
            df_novo = pd.DataFrame([produto], columns=['Quantidade', 'Nome Produto', 'Código', 'Data de Validade', 'Loja'])
            df = pd.concat([df, df_novo], ignore_index=True)

        df.to_excel(os.path.join(planilha_dir, 'estoque.xlsx'), index=False)
        messagebox.showinfo("Sucesso", "Nova carga adicionada ao estoque!")

    btn_add = tk.Button(top, text="Adicionar", command=adicionar_ao_estoque)
    btn_add.grid(row=len(labels), column=0, columnspan=2, pady=10)

    btn_voltar = tk.Button(top, text="Voltar", command=top.destroy)
    btn_voltar.grid(row=len(labels)+1, column=0, columnspan=2, pady=10)

# Função para mover itens do estoque para a loja Brooklin ou Eucaliptos
def carga_para_loja():
    top = tk.Toplevel()
    top.title("Mover Carga para Loja")

    label_loja = tk.Label(top, text="Escolha a loja:")
    label_loja.grid(row=0, column=0, padx=10, pady=5)

    btn_brooklin = tk.Button(top, text="Brooklin", command=lambda: mover_itens('brooklin'))
    btn_brooklin.grid(row=0, column=1, padx=10, pady=5)

    btn_eucaliptos = tk.Button(top, text="Eucaliptos", command=lambda: mover_itens('eucaliptos'))
    btn_eucaliptos.grid(row=0, column=2, padx=10, pady=5)

    btn_voltar = tk.Button(top, text="Voltar", command=top.destroy)
    btn_voltar.grid(row=1, column=0, columnspan=3, pady=10)

def mover_itens(loja):
    top = tk.Toplevel()
    top.title(f"Mover Carga para {loja.capitalize()}")

    df = pd.read_excel(os.path.join(planilha_dir, 'estoque.xlsx'))

    label_selecao = tk.Label(top, text="Escolha a linha do produto e a quantidade a mover:")
    label_selecao.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

    lista_produtos = ttk.Combobox(top, values=df['Nome Produto'].tolist())
    lista_produtos.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

    label_quantidade = tk.Label(top, text="Quantidade:")
    label_quantidade.grid(row=2, column=0, padx=10, pady=5)
    entrada_quantidade = tk.Entry(top)
    entrada_quantidade.grid(row=2, column=1, padx=10, pady=5)

    def confirmar_movimentacao():
        df = pd.read_excel(os.path.join(planilha_dir, 'estoque.xlsx'))
        produto_selecionado = lista_produtos.get()
        quantidade = int(entrada_quantidade.get())
        linha_selecionada = df[df['Nome Produto'] == produto_selecionado]

        if not linha_selecionada.empty and quantidade <= linha_selecionada.iloc[0]['Quantidade']:
            df.at[linha_selecionada.index[0], 'Quantidade'] -= quantidade

            if df.at[linha_selecionada.index[0], 'Quantidade'] == 0:
                df = df.drop(linha_selecionada.index[0])

            df_loja = pd.read_excel(os.path.join(planilha_dir, f'{loja}.xlsx'))
            linha_nova = linha_selecionada.copy()
            linha_nova.at[linha_selecionada.index[0], 'Quantidade'] = quantidade
            df_loja = pd.concat([df_loja, linha_nova], ignore_index=True)

            df.to_excel(os.path.join(planilha_dir, 'estoque.xlsx'), index=False)
            df_loja.to_excel(os.path.join(planilha_dir, f'{loja}.xlsx'), index=False)
            messagebox.showinfo("Sucesso", f"Carga movida para {loja.capitalize()}!")
        else:
            messagebox.showerror("Erro", "Quantidade inválida ou produto não encontrado.")

    btn_confirmar = tk.Button(top, text="Confirmar", command=confirmar_movimentacao)
    btn_confirmar.grid(row=3, column=0, columnspan=2, pady=10)

    btn_voltar = tk.Button(top, text="Voltar", command=top.destroy)
    btn_voltar.grid(row=4, column=0, columnspan=2, pady=10)

# Função para checar itens com validade próxima (próximos 30 dias)
def checar_validade():
    estoque_df = pd.read_excel(os.path.join(planilha_dir, 'estoque.xlsx'))
    brooklin_df = pd.read_excel(os.path.join(planilha_dir, 'brooklin.xlsx'))
    eucaliptos_df = pd.read_excel(os.path.join(planilha_dir, 'eucaliptos.xlsx'))

    validade_proxima_estoque = estoque_df[pd.to_datetime(estoque_df['Data de Validade']) <= pd.Timestamp.today() + pd.DateOffset(days=30)]
    validade_proxima_brooklin = brooklin_df[pd.to_datetime(brooklin_df['Data de Validade']) <= pd.Timestamp.today() + pd.DateOffset(days=30)]
    validade_proxima_eucaliptos = eucaliptos_df[pd.to_datetime(eucaliptos_df['Data de Validade']) <= pd.Timestamp.today() + pd.DateOffset(days=30)]

    with pd.ExcelWriter(os.path.join(planilha_dir, 'itens_vencendo.xlsx')) as writer:
        validade_proxima_estoque.to_excel(writer, sheet_name='Estoque', index=False)
        validade_proxima_brooklin.to_excel(writer, sheet_name='Brooklin', index=False)
        validade_proxima_eucaliptos.to_excel(writer, sheet_name='Eucaliptos', index=False)

    messagebox.showinfo("Sucesso", "Itens com validade próxima foram identificados e salvos.")

# Função para interface de adicionar produtos
def adicionar_produto(dados):
    top = tk.Toplevel()
    top.title("Adicionar Produto")

    labels = ['Quantidade', 'Nome Produto', 'Código', 'Data de Validade (AAAA-MM-DD)', 'Loja']
    entradas = {}

    for i, label in enumerate(labels):
        lbl = tk.Label(top, text=label)
        lbl.grid(row=i, column=0, padx=10, pady=5)
        entrada = tk.Entry(top)
        entrada.grid(row=i, column=1, padx=10, pady=5)
        entradas[label] = entrada

    def adicionar_ao_estoque():
        produto = [entradas[label].get() for label in labels]
        dados.append(produto)
        messagebox.showinfo("Sucesso", "Produto adicionado temporariamente!")

    def finalizar_inventario():
        fazer_inventario(dados)
        top.destroy()

    btn_add = tk.Button(top, text="Adicionar Outro Produto", command=adicionar_ao_estoque)
    btn_add.grid(row=len(labels), column=0, pady=10)

    btn_finalizar = tk.Button(top, text="Finalizar", command=finalizar_inventario)
    btn_finalizar.grid(row=len(labels), column=1, pady=10)

    btn_voltar = tk.Button(top, text="Voltar", command=top.destroy)
    btn_voltar.grid(row=len(labels)+1, column=0, columnspan=2, pady=10)

# Função principal da interface
def iniciar_interface():
    root = tk.Tk()
    root.title("Controle de Estoque")
    root.geometry("400x300")  # Define o tamanho da janela
    dados_inventario = []

    btn_inventario = tk.Button(root, text="Fazer Inventário", command=lambda: adicionar_produto(dados_inventario), width=20, height=2)
    btn_inventario.grid(row=0, column=0, padx=10, pady=10)

    btn_carga = tk.Button(root, text="Nova Carga", command=lambda: nova_carga(dados_inventario), width=20, height=2)
    btn_carga.grid(row=1, column=0, padx=10, pady=10)

    btn_carga_loja = tk.Button(root, text="Carga para Loja", command=carga_para_loja, width=20, height=2)
    btn_carga_loja.grid(row=2, column=0, padx=10, pady=10)

    btn_validade = tk.Button(root, text="Checar Validade", command=checar_validade, width=20, height=2)
    btn_validade.grid(row=3, column=0, padx=10, pady=10)

    btn_sair = tk.Button(root, text="Sair", command=root.quit, width=20, height=2)
    btn_sair.grid(row=4, column=0, padx=10, pady=10)

    root.mainloop()

# Inicializa as planilhas e a interface
if __name__ == '__main__':
    inicializar_planilhas()
    iniciar_interface()