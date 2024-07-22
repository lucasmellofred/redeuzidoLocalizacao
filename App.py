import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import fdb
import openpyxl
from collections import defaultdict

def select_database():
    file_path = filedialog.askopenfilename(filetypes=[("Database files", "*.fdb")])
    if file_path:
        db_entry.delete(0, tk.END)
        db_entry.insert(0, file_path)

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def show_confirmation(modifications):
    confirm_window = tk.Toplevel(root)
    confirm_window.title("Confirmar Modificações")

    tk.Label(confirm_window, text="Os seguintes itens serão modificados:").pack(pady=10)

    # Criar o frame para a Treeview e a barra de rolagem
    frame = tk.Frame(confirm_window)
    frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    # Criar a barra de rolagem vertical
    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)

    # Criar a Treeview
    columns = ('reduzido', 'descricao', 'localizacao_antigo', 'localizacao_novo')
    tree = ttk.Treeview(frame, columns=columns, show='headings', yscrollcommand=scrollbar.set)
    tree.heading('reduzido', text='ITEM_REDUZIDO')
    tree.heading('descricao', text='ITEM_DESCRICAO')
    tree.heading('localizacao_antigo', text='ITEM_LOCALIZACAO ANTIGO')
    tree.heading('localizacao_novo', text='ITEM_LOCALIZACAO NOVO')

    # Configurar a barra de rolagem
    scrollbar.config(command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    for mod in modifications:
        tree.insert('', tk.END, values=(mod['reduzido'], mod['descricao'], mod['localizacao_antigo'], mod['localizacao_novo']))

    def on_confirm():
        update_database(modifications)
        confirm_window.destroy()

    tk.Button(confirm_window, text="Confirmar", command=on_confirm).pack(side=tk.LEFT, padx=10, pady=10)
    tk.Button(confirm_window, text="Cancelar", command=confirm_window.destroy).pack(side=tk.RIGHT, padx=10, pady=10)

def load_data():
    db_path = db_entry.get()
    file_path = file_entry.get()

    if not db_path or not file_path:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo do banco de dados e o arquivo XLSX.")
        return

    try:
        # Conexão com o banco de dados Firebird
        con = fdb.connect(
            dsn=db_path,
            user='SYSDBA',
            password='masterkey',
            charset='UTF8'
        )
        cursor = con.cursor()

        # Carregar a planilha
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active

        # Encontrar índices das colunas relevantes
        col_names = [cell.value for cell in sheet[1]]
        rua_idx = col_names.index("Rua")
        posicao_idx = col_names.index("Posição")
        reduzido_idx = col_names.index("Reduzido")

        # Dicionário para acumular localizações por reduzido
        localizacao_dict = defaultdict(list)
        modificacoes = []

        # Iterar pelas linhas da planilha
        for row in sheet.iter_rows(min_row=2):  # Assumindo que a primeira linha é o cabeçalho
            reduzido = row[reduzido_idx].value
            rua = row[rua_idx].value
            posicao = row[posicao_idx].value

            # Garantir que reduzido, rua e posicao não são None
            if reduzido is None or rua is None or posicao is None:
                continue

            reduzido = str(reduzido).strip()
            localizacao = str(rua).strip() + str(posicao).strip()

            if len(localizacao) > 70:
                localizacao = localizacao[:70]

            # Acumular localizações no dicionário
            localizacao_dict[reduzido].append(localizacao)

        # Recuperar descrições dos itens e valores antigos
        for reduzido, localizacoes in localizacao_dict.items():
            localizacao_final = ' / '.join(localizacoes)

            cursor.execute("SELECT ITEM_DESCRICAO, ITEM_LOCALIZACAO FROM ITENS WHERE ITEM_REDUZIDO = ? AND EMPRESA_CODIGO = '0001'", (reduzido,))
            result = cursor.fetchone()
            if result:
                descricao = result[0]
                localizacao_antiga = result[1] if result[1] else ""
                modificacoes.append({
                    'reduzido': reduzido,
                    'descricao': descricao,
                    'localizacao_antigo': localizacao_antiga,
                    'localizacao_novo': localizacao_final
                })

        cursor.close()
        con.close()

        # Exibir a confirmação
        show_confirmation(modificacoes)

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def update_database(modifications):
    db_path = db_entry.get()

    try:
        # Conexão com o banco de dados Firebird
        con = fdb.connect(
            dsn=db_path,
            user='SYSDBA',
            password='masterkey',
            charset='UTF8'
        )
        cursor = con.cursor()

        # Atualizar o banco de dados
        for mod in modifications:
            print(f"Atualizando ITEM_REDUZIDO: {mod['reduzido']} com ITEM_LOCALIZACAO: {mod['localizacao_novo']}")
            cursor.execute(
                "UPDATE ITENS SET ITEM_LOCALIZACAO = ? WHERE ITEM_REDUZIDO = ? AND EMPRESA_CODIGO = '0001'",
                (mod['localizacao_novo'], mod['reduzido'])
            )

        con.commit()
        cursor.close()
        con.close()
        messagebox.showinfo("Sucesso", "Atualização concluída com sucesso.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Configuração da interface Tkinter
root = tk.Tk()
root.title("Atualizar Banco de Dados Firebird")

tk.Label(root, text="Arquivo do Banco de Dados:").grid(row=0, column=0, padx=10, pady=10)
db_entry = tk.Entry(root, width=50)
db_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar Banco de Dados", command=select_database).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Arquivo XLSX:").grid(row=1, column=0, padx=10, pady=10)
file_entry = tk.Entry(root, width=50)
file_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar Arquivo", command=select_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Carregar Dados", command=load_data).grid(row=2, column=0, columnspan=3, pady=20)

root.mainloop()
