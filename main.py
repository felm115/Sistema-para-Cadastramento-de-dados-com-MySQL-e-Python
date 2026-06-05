import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import mysql.connector
from mysql.connector import Error
from openpyxl import Workbook
from datetime import datetime
import re


DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "SUA_SENHA_AQUI",
    "database": "sistema_alunos"
}


def conectar():
    return mysql.connector.connect(**DB_CONFIG)


def inicializar_banco():
    try:
        conexao_sem_db = mysql.connector.connect(
            host=DB_CONFIG["host"],
            user=DB_CONFIG["user"],
            password=DB_CONFIG["password"]
        )
        cursor = conexao_sem_db.cursor()
        cursor.execute("CREATE DATABASE IF NOT EXISTS sistema_alunos")
        conexao_sem_db.close()

        conexao = conectar()
        cursor = conexao.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS alunos (
                id_aluno VARCHAR(20) PRIMARY KEY,
                nome VARCHAR(120) NOT NULL,
                idade INT,
                genero VARCHAR(30),
                contato VARCHAR(30),
                email VARCHAR(120),
                criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conexao.commit()
        conexao.close()

    except Error as erro:
        messagebox.showerror(
            "Erro no Banco de Dados",
            f"Não foi possível conectar ao MySQL.\n\nErro:\n{erro}"
        )


def gerar_id_aluno():
    try:
        conexao = conectar()
        cursor = conexao.cursor()
        cursor.execute("""
            SELECT id_aluno
            FROM alunos
            ORDER BY CAST(SUBSTRING(id_aluno, 4) AS UNSIGNED) DESC
            LIMIT 1
        """)
        resultado = cursor.fetchone()
        conexao.close()

        if resultado is None:
            return "ALU101"

        numero = int(resultado[0][3:])
        return f"ALU{numero + 1}"

    except Exception:
        return "ALU101"


def validar_campos():
    if nome_var.get().strip() == "":
        messagebox.showerror("Erro", "O campo Nome é obrigatório.")
        return False

    if idade_var.get().strip():
        try:
            idade = int(idade_var.get())
            if idade <= 0 or idade > 120:
                messagebox.showerror("Erro", "Digite uma idade válida.")
                return False
        except ValueError:
            messagebox.showerror("Erro", "A idade deve conter apenas números.")
            return False

    if email_var.get().strip():
        padrao_email = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        if not re.match(padrao_email, email_var.get().strip()):
            messagebox.showerror("Erro", "Digite um e-mail válido.")
            return False

    return True


def adicionar_aluno():
    if not validar_campos():
        return

    try:
        conexao = conectar()
        cursor = conexao.cursor()

        sql = """
            INSERT INTO alunos (id_aluno, nome, idade, genero, contato, email)
            VALUES (%s, %s, %s, %s, %s, %s)
        """

        valores = (
            id_var.get(),
            nome_var.get().strip(),
            idade_var.get().strip() or None,
            genero_var.get(),
            contato_var.get().strip(),
            email_var.get().strip()
        )

        cursor.execute(sql, valores)
        conexao.commit()
        conexao.close()

        messagebox.showinfo("Sucesso", "Aluno cadastrado com sucesso.")
        mostrar_todos()
        limpar_campos()

    except Error as erro:
        messagebox.showerror("Erro no Banco de Dados", str(erro))


def excluir_aluno():
    if id_var.get().strip() == "":
        messagebox.showerror("Erro", "Selecione um aluno para excluir.")
        return

    confirmar = messagebox.askyesno(
        "Confirmar Exclusão",
        f"Deseja realmente excluir o aluno {id_var.get()}?"
    )

    if not confirmar:
        return

    try:
        conexao = conectar()
        cursor = conexao.cursor()
        cursor.execute("DELETE FROM alunos WHERE id_aluno=%s", (id_var.get(),))
        conexao.commit()
        conexao.close()

        messagebox.showinfo("Excluído", "Aluno excluído com sucesso.")
        mostrar_todos()
        limpar_campos()

    except Error as erro:
        messagebox.showerror("Erro no Banco de Dados", str(erro))


def limpar_campos():
    id_var.set(gerar_id_aluno())
    nome_var.set("")
    idade_var.set("")
    genero_var.set("")
    contato_var.set("")
    email_var.set("")


def preencher_campos(event=None):
    selecionado = tabela.focus()
    valores = tabela.item(selecionado, "values")

    if valores:
        id_var.set(valores[0])
        nome_var.set(valores[1])
        idade_var.set(valores[2])
        genero_var.set(valores[3])
        contato_var.set(valores[4])
        email_var.set(valores[5])


def atualizar_contador(total):
    contador_var.set(f"Total de registros: {total}")


def mostrar_todos():
    try:
        tabela.delete(*tabela.get_children())

        conexao = conectar()
        cursor = conexao.cursor()
        cursor.execute("""
            SELECT id_aluno, nome, idade, genero, contato, email
            FROM alunos
            ORDER BY CAST(SUBSTRING(id_aluno, 4) AS UNSIGNED) ASC
        """)

        registros = cursor.fetchall()
        conexao.close()

        for registro in registros:
            tabela.insert("", tk.END, values=registro)

        atualizar_contador(len(registros))

    except Error as erro:
        messagebox.showerror("Erro no Banco de Dados", str(erro))


def pesquisar_aluno():
    termo = pesquisa_var.get().strip()

    if termo == "":
        mostrar_todos()
        return

    mapa_colunas = {
        "ID do Aluno": "id_aluno",
        "Nome": "nome",
        "Idade": "idade",
        "Gênero": "genero",
        "Contato": "contato",
        "E-mail": "email"
    }

    coluna = mapa_colunas.get(pesquisar_por_var.get(), "nome")

    try:
        conexao = conectar()
        cursor = conexao.cursor()

        sql = f"""
            SELECT id_aluno, nome, idade, genero, contato, email
            FROM alunos
            WHERE CAST({coluna} AS CHAR) LIKE %s
            ORDER BY CAST(SUBSTRING(id_aluno, 4) AS UNSIGNED) ASC
        """

        cursor.execute(sql, (f"%{termo}%",))
        registros = cursor.fetchall()
        conexao.close()

        tabela.delete(*tabela.get_children())

        for registro in registros:
            tabela.insert("", tk.END, values=registro)

        atualizar_contador(len(registros))

    except Error as erro:
        messagebox.showerror("Erro no Banco de Dados", str(erro))


def exportar_excel():
    registros = [tabela.item(item, "values") for item in tabela.get_children()]

    if not registros:
        messagebox.showwarning("Sem dados", "Não há registros para exportar.")
        return

    caminho = filedialog.asksaveasfilename(
        title="Salvar arquivo Excel",
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
        initialfile=f"alunos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    if not caminho:
        return

    try:
        planilha = Workbook()
        aba = planilha.active
        aba.title = "Alunos"

        aba.append(["ID do Aluno", "Nome", "Idade", "Gênero", "Contato", "E-mail"])

        for registro in registros:
            aba.append(list(registro))

        for coluna in aba.columns:
            maior = 0
            letra = coluna[0].column_letter

            for celula in coluna:
                if celula.value:
                    maior = max(maior, len(str(celula.value)))

            aba.column_dimensions[letra].width = maior + 4

        planilha.save(caminho)
        messagebox.showinfo("Exportado", "Arquivo Excel exportado com sucesso.")

    except Exception as erro:
        messagebox.showerror("Erro", f"Erro ao exportar arquivo:\n{erro}")


janela = tk.Tk()
janela.title("Sistema de Cadastro de Alunos")
janela.geometry("1920x1080")
janela.minsize(1200, 700)

try:
    janela.state("zoomed")
except Exception:
    pass

janela.resizable(True, True)

COR_FUNDO = "#223548"
COR_HEADER = "#e84135"
COR_CARD = "#2f455c"
COR_BRANCO = "#ffffff"
COR_VERDE = "#2ecc71"
COR_VERMELHO = "#e74c3c"
COR_CINZA = "#95a5a6"
COR_AZUL = "#3498db"
COR_INPUT = "#ffffff"

janela.configure(bg=COR_FUNDO)

id_var = tk.StringVar()
nome_var = tk.StringVar()
idade_var = tk.StringVar()
genero_var = tk.StringVar()
contato_var = tk.StringVar()
email_var = tk.StringVar()
pesquisar_por_var = tk.StringVar(value="Nome")
pesquisa_var = tk.StringVar()
contador_var = tk.StringVar(value="Total de registros: 0")

style = ttk.Style()
style.theme_use("default")

style.configure(
    "Treeview",
    background="white",
    foreground="black",
    rowheight=38,
    fieldbackground="white",
    font=("Segoe UI", 12)
)

style.configure(
    "Treeview.Heading",
    font=("Segoe UI", 12, "bold"),
    background="#d5d8dc",
    foreground="black"
)

style.map("Treeview", background=[("selected", "#3498db")])

cabecalho = tk.Label(
    janela,
    text="Sistema de Cadastro de Alunos",
    bg=COR_HEADER,
    fg=COR_BRANCO,
    font=("Segoe UI", 30, "bold"),
    pady=18
)
cabecalho.pack(fill=tk.X)

container = tk.Frame(janela, bg=COR_FUNDO)
container.pack(fill=tk.BOTH, expand=True, padx=30, pady=25)

container.columnconfigure(0, weight=1)
container.columnconfigure(1, weight=4)
container.rowconfigure(0, weight=1)

painel_esquerdo = tk.Frame(container, bg=COR_CARD, bd=0)
painel_esquerdo.grid(row=0, column=0, sticky="nsew", padx=(0, 20))

# Fonte reduzida em aproximadamente 30%: 22 -> 15
titulo_form = tk.Label(
    painel_esquerdo,
    text="Gerenciar Alunos",
    bg=COR_CARD,
    fg=COR_BRANCO,
    font=("Segoe UI", 15, "bold")
)
titulo_form.pack(pady=(25, 20))


def criar_rotulo(parent, texto):
    return tk.Label(
        parent,
        text=texto,
        bg=COR_CARD,
        fg=COR_BRANCO,
        font=("Segoe UI", 13, "bold")
    )


def criar_entrada(parent, variavel, readonly=False):
    estado = "readonly" if readonly else "normal"
    return tk.Entry(
        parent,
        textvariable=variavel,
        font=("Segoe UI", 13),
        bg=COR_INPUT,
        fg="black",
        relief=tk.FLAT,
        state=estado
    )


form_frame = tk.Frame(painel_esquerdo, bg=COR_CARD)
form_frame.pack(fill=tk.X, padx=30)

# Campo ID
criar_rotulo(form_frame, "ID do Aluno").pack(anchor="w", pady=(12, 4))
criar_entrada(form_frame, id_var, readonly=True).pack(fill=tk.X, ipady=9)

# Campo Nome
criar_rotulo(form_frame, "Nome").pack(anchor="w", pady=(12, 4))
criar_entrada(form_frame, nome_var).pack(fill=tk.X, ipady=9)

# Campo Idade
criar_rotulo(form_frame, "Idade").pack(anchor="w", pady=(12, 4))
criar_entrada(form_frame, idade_var).pack(fill=tk.X, ipady=9)

# Campo Gênero corrigido: agora fica dentro do form_frame, abaixo de Idade e antes de Contato
criar_rotulo(form_frame, "Gênero").pack(anchor="w", pady=(12, 4))
combo_genero = ttk.Combobox(
    form_frame,
    textvariable=genero_var,
    values=["Masculino", "Feminino", "Outro"],
    state="readonly",
    font=("Segoe UI", 13)
)
combo_genero.pack(fill=tk.X, ipady=6)

# Campo Contato
criar_rotulo(form_frame, "Contato").pack(anchor="w", pady=(12, 4))
criar_entrada(form_frame, contato_var).pack(fill=tk.X, ipady=9)

# Campo E-mail
criar_rotulo(form_frame, "E-mail").pack(anchor="w", pady=(12, 4))
criar_entrada(form_frame, email_var).pack(fill=tk.X, ipady=9)

botoes = tk.Frame(painel_esquerdo, bg=COR_CARD)
botoes.pack(pady=35)


def botao(texto, cor, comando):
    return tk.Button(
        botoes,
        text=texto,
        bg=cor,
        fg=COR_BRANCO,
        activebackground=cor,
        activeforeground=COR_BRANCO,
        font=("Segoe UI", 13, "bold"),
        width=16,
        height=2,
        relief=tk.FLAT,
        cursor="hand2",
        command=comando
    )


# Botão Atualizar removido
botao("Adicionar", COR_VERDE, adicionar_aluno).grid(row=0, column=0, padx=8, pady=8)
botao("Excluir", COR_VERMELHO, excluir_aluno).grid(row=1, column=0, padx=8, pady=8)
botao("Limpar", COR_CINZA, limpar_campos).grid(row=1, column=1, padx=8, pady=8)

painel_direito = tk.Frame(container, bg=COR_FUNDO)
painel_direito.grid(row=0, column=1, sticky="nsew")
painel_direito.rowconfigure(1, weight=1)
painel_direito.columnconfigure(0, weight=1)

barra_pesquisa = tk.Frame(painel_direito, bg=COR_FUNDO)
barra_pesquisa.grid(row=0, column=0, sticky="ew", pady=(0, 18))
barra_pesquisa.columnconfigure(2, weight=1)

tk.Label(
    barra_pesquisa,
    text="Pesquisar por:",
    bg=COR_FUNDO,
    fg=COR_BRANCO,
    font=("Segoe UI", 13, "bold")
).grid(row=0, column=0, padx=(0, 10))

combo_pesquisa = ttk.Combobox(
    barra_pesquisa,
    textvariable=pesquisar_por_var,
    values=["ID do Aluno", "Nome", "Idade", "Gênero", "Contato", "E-mail"],
    state="readonly",
    font=("Segoe UI", 12),
    width=18
)
combo_pesquisa.grid(row=0, column=1, padx=8, ipady=5)

entrada_pesquisa = tk.Entry(
    barra_pesquisa,
    textvariable=pesquisa_var,
    font=("Segoe UI", 13),
    relief=tk.FLAT
)
entrada_pesquisa.grid(row=0, column=2, sticky="ew", padx=8, ipady=8)
entrada_pesquisa.bind("<KeyRelease>", lambda event: pesquisar_aluno())

tk.Button(
    barra_pesquisa,
    text="Pesquisar",
    bg=COR_AZUL,
    fg=COR_BRANCO,
    font=("Segoe UI", 12, "bold"),
    width=13,
    height=2,
    relief=tk.FLAT,
    cursor="hand2",
    command=pesquisar_aluno
).grid(row=0, column=3, padx=8)

tk.Button(
    barra_pesquisa,
    text="Mostrar Todos",
    bg=COR_CINZA,
    fg=COR_BRANCO,
    font=("Segoe UI", 12, "bold"),
    width=15,
    height=2,
    relief=tk.FLAT,
    cursor="hand2",
    command=mostrar_todos
).grid(row=0, column=4, padx=8)

tk.Button(
    barra_pesquisa,
    text="Exportar para Excel",
    bg=COR_VERDE,
    fg=COR_BRANCO,
    font=("Segoe UI", 12, "bold"),
    width=20,
    height=2,
    relief=tk.FLAT,
    cursor="hand2",
    command=exportar_excel
).grid(row=0, column=5, padx=(8, 0))

frame_tabela = tk.Frame(painel_direito, bg="white")
frame_tabela.grid(row=1, column=0, sticky="nsew")

scroll_y = tk.Scrollbar(frame_tabela, orient=tk.VERTICAL)
scroll_x = tk.Scrollbar(frame_tabela, orient=tk.HORIZONTAL)

colunas = ("ID do Aluno", "Nome", "Idade", "Gênero", "Contato", "E-mail")

tabela = ttk.Treeview(
    frame_tabela,
    columns=colunas,
    show="headings",
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

scroll_y.config(command=tabela.yview)
scroll_x.config(command=tabela.xview)

larguras = {
    "ID do Aluno": 170,
    "Nome": 330,
    "Idade": 120,
    "Gênero": 180,
    "Contato": 240,
    "E-mail": 390
}

for coluna in colunas:
    tabela.heading(coluna, text=coluna)
    tabela.column(coluna, width=larguras[coluna], anchor="center")

tabela.pack(fill=tk.BOTH, expand=True)
tabela.bind("<ButtonRelease-1>", preencher_campos)
tabela.bind("<Double-1>", preencher_campos)

rodape = tk.Frame(painel_direito, bg=COR_FUNDO)
rodape.grid(row=2, column=0, sticky="ew", pady=(12, 0))

tk.Label(
    rodape,
    textvariable=contador_var,
    bg=COR_FUNDO,
    fg=COR_BRANCO,
    font=("Segoe UI", 12, "bold")
).pack(side=tk.LEFT)

tk.Label(
    rodape,
    text="Dica: clique em um registro da tabela para excluir ou visualizar.",
    bg=COR_FUNDO,
    fg=COR_BRANCO,
    font=("Segoe UI", 11)
).pack(side=tk.RIGHT)

inicializar_banco()
limpar_campos()
mostrar_todos()

janela.mainloop()