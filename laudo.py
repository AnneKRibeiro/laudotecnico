import tkinter as tk
from tkinter import scrolledtext, messagebox
import sqlite3
import pyodbc

def limpar_campos():
    for entry in entries:
        entry.delete(0, tk.END)
    text_reclamacao_cliente.delete("1.0", tk.END)
    text_defeito_encontrado.delete("1.0", tk.END)
    text_solucao.delete("1.0", tk.END)

def fechar_janela_laudo():
    limpar_campos()
    if 'janela_laudo' in globals():
        janela_laudo.destroy()

def validar_entrada(char, campo):
    return char.isdigit() or char == ""

def gerar_laudo():
    dados = {}
    for entry, campo in zip(entries, campos):
        valor = entry.get()
        dados[campo] = valor

    numero_os = entry_numero_os.get()

    reclamacao_cliente = text_reclamacao_cliente.get("1.0", "end-1c")
    defeito_encontrado = text_defeito_encontrado.get("1.0", "end-1c")
    solucao = text_solucao.get("1.0", "end-1c")

    laudo = "---- LAUDO TÉCNICO ----\n"
    for campo, valor in dados.items():
        laudo += f"**{campo}: {valor}\n"

    laudo += f"\n**RECLAMAÇÃO DO CLIENTE:\n{reclamacao_cliente}\n\n"
    laudo += f"**DEFEITO ENCONTRADO:\n{defeito_encontrado}\n\n"
    laudo += f"**SOLUÇÃO:\n{solucao}"

    conexao_sqlite = sqlite3.connect('laudos.db')
    cursor_sqlite = conexao_sqlite.cursor()

    cursor_sqlite.execute('''
        CREATE TABLE IF NOT EXISTS laudos_tecnicos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            laudo TEXT
        )
    ''')

    cursor_sqlite.execute('INSERT INTO laudos_tecnicos (laudo) VALUES (?)', (laudo,))
    conexao_sqlite.commit()
    conexao_sqlite.close()

    conexao_sql_server = pyodbc.connect('DRIVER={SQL Server};SERVER=localhost;DATABASE=autocenter;UID=sa;PWD=macro01')
    cursor_sql_server = conexao_sql_server.cursor()

    query_verificar_os = "SELECT vedid FROM venda WHERE vedid = ?"
    cursor_sql_server.execute(query_verificar_os, (numero_os,))
    resultado = cursor_sql_server.fetchone()

    if resultado:
        global janela_laudo
        janela_laudo = tk.Toplevel(root)
        janela_laudo.title("Laudo Técnico")

        text_laudo = scrolledtext.ScrolledText(janela_laudo, wrap=tk.WORD, width=80, height=30)
        text_laudo.insert(tk.END, laudo)
        text_laudo.pack()

        button_fechar = tk.Button(janela_laudo, text="Fechar", command=fechar_janela_laudo)
        button_fechar.pack()
        
        janela_laudo.protocol("WM_DELETE_WINDOW", fechar_janela_laudo)
        
        query_verificar_laudo = "SELECT vedLaudoTec FROM venda WHERE vedid = ?"
        cursor_sql_server.execute(query_verificar_laudo, (numero_os,))
        laudo_atual = cursor_sql_server.fetchone()[0]

        if laudo_atual:
            resposta = messagebox.askyesno("Laudo Existente", "Já existe um laudo registrado para essa O.S. Deseja substituí-lo?")
            if not resposta:
                return

        query_inserir = "UPDATE venda SET vedLaudoTec = ? WHERE vedid = ?"
        cursor_sql_server.execute(query_inserir, (laudo.replace('\n', '\r\n'), numero_os))  # Manter as quebras de linha
        conexao_sql_server.commit()
        conexao_sql_server.close()

    else:
        messagebox.showerror("Erro", "Número de O.S inválido ou não encontrado na tabela venda.")

root = tk.Tk()
root.title("Laudo Técnico")

campos = [
    "Técnico", "HD/SSD", "Mem.RAM", "Aquecimento COOLER", "Sujeira Interna", "WI-FI", "USB", "Áudio",
    "Tela", "Webcam", "Teclado", "Touch Pad", "Carcaça", "Gravadora", "Bateria da BIOS", "Bateria Notebook",
    "Troca da Pasta Térmica", "Sistema Operacional", "Encaminhado para Sistema", "Fonte (Original ou Paralela)"
]

entries = []
for idx, campo in enumerate(campos):
    label = tk.Label(root, text=campo + ":")
    label.grid(row=idx, column=0, padx=5, pady=5, sticky="w")

    entry = tk.Entry(root)
    entry.grid(row=idx, column=1, padx=5, pady=5, sticky="we")
    entries.append(entry)

# Campo apenas para o número da O.S.
label_os = tk.Label(root, text="Número da O.S.:")
label_os.grid(row=len(campos), column=0, padx=5, pady=5, sticky="w")

validate_os = root.register(lambda P: validar_entrada(P, entry_numero_os))
entry_numero_os = tk.Entry(root, validate="key", validatecommand=(validate_os, "%P"))
entry_numero_os.grid(row=len(campos), column=1, padx=5, pady=5, sticky="we")

label_reclamacao = tk.Label(root, text="Reclamação do Cliente:")
label_reclamacao.grid(row=0, column=5, padx=5, pady=5, sticky="w")

label_defeito = tk.Label(root, text="Defeito Encontrado:")
label_defeito.grid(row=6, column=5, padx=5, pady=5, sticky="w")

label_solucao = tk.Label(root, text="Solução:")
label_solucao.grid(row=12, column=5, padx=5, pady=5, sticky="w")

text_reclamacao_cliente = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=50, height=5)
text_reclamacao_cliente.grid(row=0, column=6, rowspan=5, padx=5, pady=5, sticky="nswe")

text_defeito_encontrado = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=50, height=5)
text_defeito_encontrado.grid(row=6, column=6, rowspan=5, padx=5, pady=5, sticky="nswe")

text_solucao = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=50, height=5)
text_solucao.grid(row=12, column=6, rowspan=5, padx=5, pady=5, sticky="nswe")

button_gerar_laudo = tk.Button(root, text="Gerar Laudo", command=gerar_laudo)
button_gerar_laudo.grid(row=30, column=7, columnspan=3, padx=5, pady=5)

root.state('zoomed')  # Abre a janela maximizada

root.mainloop()