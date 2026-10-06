import fitz  # PyMuPDF
import csv
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ============================================================
# EXTRATOR DE TEXTOS DE FOLDER PDF
# ============================================================
#
# Extrai produtos de um folder PDF e organiza em uma tabela:
#
# Selecionar | Qtd. | Página | Categoria | Código |
# Produto | Embalagem | Preço | Total
#
# Também permite:
#
# - Selecionar produtos
# - Informar quantidade
# - Calcular total por produto
# - Calcular total do pedido
# - Gerar pedido em CSV compatível com Excel
#
# ============================================================


# ============================================================
# EXPRESSÕES REGULARES
# ============================================================

# Exemplos:
# 12X330ML
# 24X270ML
# 6X2L
# 30X220G
# 12X100ML

PADRAO_EMBALAGEM = re.compile(
    r"^\d+\s*[Xx]\s*\d+(?:[.,]\d+)?\s*[A-Za-z]+$"
)


# Preço no formato brasileiro
# Exemplo:
# 4,39
# 12,90
# 199,99

PADRAO_PRECO = re.compile(
    r"^\d{1,3},\d{2}$"
)


# Código dos produtos
PADRAO_CODIGO = re.compile(
    r"^\d{3,6}$"
)


# ============================================================
# FUNÇÃO PARA LIMPAR TEXTO
# ============================================================

def limpar_texto(texto):
    """
    Remove espaços duplicados.
    """

    return re.sub(r"\s+", " ", texto).strip()


# ============================================================
# DESCOBRIR CATEGORIA
# ============================================================

def descobrir_categoria(page):
    """
    Procura no texto da página algo como:

    01- BEBIDAS
    02- ALIMENTOS
    """

    texto = page.get_text("text")

    for linha in texto.splitlines():

        linha = limpar_texto(linha)

        encontrado = re.match(
            r"^\d{2}\s*-\s*(.+)$",
            linha
        )

        if encontrado:
            return encontrado.group(1).strip()

    return ""


# ============================================================
# EXTRAIR PRODUTOS DO PDF
# ============================================================

def extrair_produtos(pdf_path):

    documento = fitz.open(pdf_path)

    produtos = []

    categoria_atual = ""

    for numero_pagina, page in enumerate(
        documento,
        start=1
    ):

        # ----------------------------------------------------
        # Descobre categoria
        # ----------------------------------------------------

        categoria = descobrir_categoria(page)

        if categoria:
            categoria_atual = categoria

        palavras = page.get_text("words")

        # ----------------------------------------------------
        # Localiza códigos
        # ----------------------------------------------------

        codigos = []

        for palavra in palavras:

            x0, y0, x1, y1, texto = palavra[:5]

            texto = texto.strip()

            if PADRAO_CODIGO.fullmatch(texto):

                codigos.append({
                    "codigo": texto,
                    "x": (x0 + x1) / 2,
                    "y": y0
                })

        if not codigos:
            continue

        # ----------------------------------------------------
        # Agrupa códigos pela linha vertical
        # ----------------------------------------------------

        linhas_y = []

        for item in sorted(
            codigos,
            key=lambda v: v["y"]
        ):

            colocado = False

            for y_existente in linhas_y:

                if abs(item["y"] - y_existente) <= 4:

                    colocado = True
                    break

            if not colocado:

                linhas_y.append(
                    item["y"]
                )

        linhas_y.sort()

        # ----------------------------------------------------
        # Divide página em 5 colunas
        # ----------------------------------------------------

        largura_coluna = (
            page.rect.width / 5
        )

        # ----------------------------------------------------
        # Processa cada produto
        # ----------------------------------------------------

        for item in codigos:

            codigo = item["codigo"]

            x_centro = item["x"]

            y_inicio = item["y"]

            # ------------------------------------------------
            # Descobre linha do produto
            # ------------------------------------------------

            linha_atual = min(
                linhas_y,
                key=lambda y:
                abs(y - y_inicio)
            )

            indice_linha = linhas_y.index(
                linha_atual
            )

            # ------------------------------------------------
            # Final da área da linha
            # ------------------------------------------------

            if (
                indice_linha + 1
                < len(linhas_y)
            ):

                y_final = (
                    linhas_y[indice_linha + 1]
                    - 1
                )

            else:

                y_final = page.rect.height

            # ------------------------------------------------
            # Descobre coluna
            # ------------------------------------------------

            coluna = int(
                x_centro / largura_coluna
            )

            if coluna > 4:
                coluna = 4

            x_inicio = (
                coluna * largura_coluna
            )

            x_final = (
                (coluna + 1)
                * largura_coluna
            )

            # ------------------------------------------------
            # Seleciona palavras da célula
            # ------------------------------------------------

            palavras_celula = []

            for palavra in palavras:

                x0, y0, x1, y1, texto = (
                    palavra[:5]
                )

                centro_x = (
                    x0 + x1
                ) / 2

                if (
                    x_inicio <= centro_x < x_final
                    and y_inicio - 2 <= y0 < y_final
                ):

                    palavras_celula.append(
                        palavra
                    )

            # ------------------------------------------------
            # Ordena palavras
            # ------------------------------------------------

            palavras_celula.sort(
                key=lambda w: (
                    round(w[1], 1),
                    w[0]
                )
            )

            # ------------------------------------------------
            # Reconstrói linhas
            # ------------------------------------------------

            linhas = []

            for palavra in palavras_celula:

                x0, y0, x1, y1, texto = (
                    palavra[:5]
                )

                texto = limpar_texto(texto)

                if not texto:
                    continue

                if (
                    not linhas
                    or abs(
                        y0 - linhas[-1]["y"]
                    ) > 3
                ):

                    linhas.append({
                        "y": y0,
                        "textos": [texto]
                    })

                else:

                    linhas[-1]["textos"].append(
                        texto
                    )

            textos = [
                limpar_texto(
                    " ".join(
                        linha["textos"]
                    )
                )
                for linha in linhas
            ]

            # Remove linhas vazias
            textos = [
                linha
                for linha in textos
                if linha
            ]

            # ------------------------------------------------
            # Localiza preço
            # ------------------------------------------------

            preco = ""

            for linha in textos:

                if "R$" in linha:

                    encontrado = re.search(
                        r"R\$\s*(\d{1,3},\d{2})",
                        linha
                    )

                    if encontrado:

                        preco = (
                            encontrado.group(1)
                        )

                        break

            # ------------------------------------------------
            # Localiza embalagem
            # ------------------------------------------------

            embalagem = ""

            indice_embalagem = -1

            for i, linha in enumerate(
                textos
            ):

                linha_limpa = (
                    linha.replace(" ", "")
                )

                if PADRAO_EMBALAGEM.fullmatch(
                    linha_limpa
                ):

                    embalagem = linha

                    indice_embalagem = i

                    break

            # ------------------------------------------------
            # Localiza código
            # ------------------------------------------------

            indice_codigo = -1

            for i, linha in enumerate(
                textos
            ):

                if linha == codigo:

                    indice_codigo = i

                    break

            # ------------------------------------------------
            # Monta nome do produto
            # ------------------------------------------------

            if indice_codigo >= 0:

                inicio_nome = (
                    indice_codigo + 1
                )

                if (
                    indice_embalagem
                    > inicio_nome
                ):

                    nome_partes = textos[
                        inicio_nome:
                        indice_embalagem
                    ]

                else:

                    nome_partes = textos[
                        inicio_nome:
                    ]

                # Remove preço
                nome_partes = [
                    p
                    for p in nome_partes
                    if not p.startswith("R$")
                ]

                nome = limpar_texto(
                    " ".join(nome_partes)
                )

            else:

                nome = ""

            # ------------------------------------------------
            # Evita registros inválidos
            # ------------------------------------------------

            if not nome:
                continue

            if nome.upper().startswith(
                "IMAGENS MERAMENTE"
            ):
                continue

            # ------------------------------------------------
            # Adiciona produto
            # ------------------------------------------------

            produtos.append({

                "pagina": numero_pagina,

                "categoria": categoria_atual,

                "codigo": codigo,

                "produto": nome,

                "embalagem": embalagem,

                "preco": preco,

                # NOVO
                "quantidade": 0

            })

    documento.close()

    return produtos


# ============================================================
# APLICAÇÃO
# ============================================================

class Aplicacao:

    # ========================================================
    # INICIALIZAÇÃO
    # ========================================================

    def __init__(self, janela):

        self.janela = janela

        self.janela.title(
            "Extrator de Textos - Folder PDF"
        )

        self.janela.geometry(
            "1250x700"
        )

        # ----------------------------------------------------
        # Variáveis
        # ----------------------------------------------------

        self.produtos = []

        self.selecionados = set()

        self.itens_exibidos = []

        # ----------------------------------------------------
        # BARRA SUPERIOR
        # ----------------------------------------------------

        topo = ttk.Frame(
            janela,
            padding=10
        )

        topo.pack(
            fill="x"
        )

        ttk.Button(
            topo,
            text="ABRIR PDF",
            command=self.abrir_pdf
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="SALVAR CSV",
            command=self.salvar_csv
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="SALVAR TXT",
            command=self.salvar_txt
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="SALVAR SELECIONADOS CSV",
            command=self.salvar_csv_selecionados
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="SALVAR SELECIONADOS TXT",
            command=self.salvar_txt_selecionados
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="GERAR PEDIDO CSV",
            command=self.gerar_pedido_csv
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="MARCAR TODOS",
            command=self.marcar_todos
        ).pack(
            side="left",
            padx=3
        )

        ttk.Button(
            topo,
            text="DESMARCAR TODOS",
            command=self.desmarcar_todos
        ).pack(
            side="left",
            padx=3
        )

        # ----------------------------------------------------
        # PESQUISA
        # ----------------------------------------------------

        ttk.Label(
            topo,
            text="Pesquisar:"
        ).pack(
            side="left",
            padx=(20, 5)
        )

        self.pesquisa = tk.StringVar()

        campo_pesquisa = ttk.Entry(
            topo,
            textvariable=self.pesquisa,
            width=30
        )

        campo_pesquisa.pack(
            side="left"
        )

        campo_pesquisa.bind(
            "<KeyRelease>",
            self.filtrar
        )

        # ----------------------------------------------------
        # INFORMAÇÃO
        # ----------------------------------------------------

        self.info = ttk.Label(
            janela,
            text=(
                "Nenhum PDF carregado. "
                "Clique na coluna ✓ para "
                "marcar produtos."
            )
        )

        self.info.pack(
            fill="x",
            padx=10,
            pady=(0, 5)
        )

        # ----------------------------------------------------
        # TOTAL DO PEDIDO
        # ----------------------------------------------------

        self.label_total = ttk.Label(
            janela,
            text=(
                "Produtos selecionados: 0 | "
                "Quantidade: 0 | "
                "TOTAL DO PEDIDO: R$ 0,00"
            ),
            font=("Arial", 12, "bold")
        )

        self.label_total.pack(
            fill="x",
            padx=10,
            pady=5
        )

        # ----------------------------------------------------
        # TABELA
        # ----------------------------------------------------

        frame_tabela = ttk.Frame(
            janela
        )

        frame_tabela.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        colunas = (

            "selecionar",

            "quantidade",

            "pagina",

            "categoria",

            "codigo",

            "produto",

            "embalagem",

            "preco",

            "total"

        )

        self.tabela = ttk.Treeview(

            frame_tabela,

            columns=colunas,

            show="headings"

        )

        # ----------------------------------------------------
        # CABEÇALHOS
        # ----------------------------------------------------

        self.tabela.heading(
            "selecionar",
            text="✓"
        )

        self.tabela.heading(
            "quantidade",
            text="Qtd."
        )

        self.tabela.heading(
            "pagina",
            text="Página"
        )

        self.tabela.heading(
            "categoria",
            text="Categoria"
        )

        self.tabela.heading(
            "codigo",
            text="Código"
        )

        self.tabela.heading(
            "produto",
            text="Produto"
        )

        self.tabela.heading(
            "embalagem",
            text="Embalagem"
        )

        self.tabela.heading(
            "preco",
            text="Preço"
        )

        self.tabela.heading(
            "total",
            text="Total"
        )

        # ----------------------------------------------------
        # LARGURAS
        # ----------------------------------------------------

        self.tabela.column(
            "selecionar",
            width=45,
            anchor="center",
            stretch=False
        )

        self.tabela.column(
            "quantidade",
            width=65,
            anchor="center",
            stretch=False
        )

        self.tabela.column(
            "pagina",
            width=60,
            anchor="center",
            stretch=False
        )

        self.tabela.column(
            "categoria",
            width=130,
            anchor="w"
        )

        self.tabela.column(
            "codigo",
            width=80,
            anchor="center"
        )

        self.tabela.column(
            "produto",
            width=450,
            anchor="w"
        )

        self.tabela.column(
            "embalagem",
            width=120,
            anchor="center"
        )

        self.tabela.column(
            "preco",
            width=90,
            anchor="center"
        )

        self.tabela.column(
            "total",
            width=100,
            anchor="center"
        )

        # ----------------------------------------------------
        # SCROLLBARS
        # ----------------------------------------------------

        barra_vertical = ttk.Scrollbar(

            frame_tabela,

            orient="vertical",

            command=self.tabela.yview

        )

        barra_horizontal = ttk.Scrollbar(

            frame_tabela,

            orient="horizontal",

            command=self.tabela.xview

        )

        self.tabela.configure(

            yscrollcommand=barra_vertical.set,

            xscrollcommand=barra_horizontal.set

        )

        self.tabela.grid(

            row=0,

            column=0,

            sticky="nsew"

        )

        barra_vertical.grid(

            row=0,

            column=1,

            sticky="ns"

        )

        barra_horizontal.grid(

            row=1,

            column=0,

            sticky="ew"

        )

        frame_tabela.rowconfigure(
            0,
            weight=1
        )

        frame_tabela.columnconfigure(
            0,
            weight=1
        )

        # ----------------------------------------------------
        # CLIQUE NA TABELA
        # ----------------------------------------------------

        self.tabela.bind(
            "<Button-1>",
            self.clique_tabela
        )

    # ========================================================
    # CONVERTER PREÇO
    # ========================================================

    def converter_preco(self, preco):

        try:

            return float(
                preco
                .replace(".", "")
                .replace(",", ".")
            )

        except:

            return 0.0

    # ========================================================
    # FORMATAR MOEDA
    # ========================================================

    def formatar_moeda(self, valor):

        return (
            f"R$ {valor:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    # ========================================================
    # CALCULAR TOTAL DE UM PRODUTO
    # ========================================================

    def calcular_total_produto(
        self,
        produto
    ):

        try:

            quantidade = int(
                produto.get(
                    "quantidade",
                    0
                )
            )

        except:

            quantidade = 0

        preco = self.converter_preco(
            produto.get(
                "preco",
                "0,00"
            )
        )

        return preco * quantidade

    # ========================================================
    # CALCULAR TOTAL DO PEDIDO
    # ========================================================

    def calcular_total_pedido(self):

        total = 0.0

        for produto in self.produtos:

            if id(produto) in self.selecionados:

                total += (
                    self.calcular_total_produto(
                        produto
                    )
                )

        return total

    # ========================================================
    # ATUALIZAR TOTAL GERAL
    # ========================================================

    def atualizar_total_geral(self):

        total = (
            self.calcular_total_pedido()
        )

        quantidade_total = 0

        produtos_selecionados = 0

        for produto in self.produtos:

            if id(produto) in self.selecionados:

                produtos_selecionados += 1

                try:

                    quantidade_total += int(
                        produto.get(
                            "quantidade",
                            0
                        )
                    )

                except:

                    pass

        self.label_total.config(

            text=(

                f"Produtos selecionados: "
                f"{produtos_selecionados} | "

                f"Quantidade: "
                f"{quantidade_total} | "

                f"TOTAL DO PEDIDO: "
                f"{self.formatar_moeda(total)}"

            )

        )

    # ========================================================
    # ABRIR PDF
    # ========================================================

    def abrir_pdf(self):

        caminho = filedialog.askopenfilename(

            title="Selecione o folder PDF",

            filetypes=[

                (
                    "Arquivos PDF",
                    "*.pdf"
                ),

                (
                    "Todos os arquivos",
                    "*.*"
                )

            ]

        )

        if not caminho:
            return

        try:

            self.info.config(
                text="Lendo PDF, aguarde..."
            )

            self.janela.update_idletasks()

            self.produtos = (
                extrair_produtos(
                    caminho
                )
            )

            self.selecionados.clear()

            self.mostrar_produtos(
                self.produtos
            )

            self.atualizar_total_geral()

            self.info.config(

                text=(

                    f"{len(self.produtos)} "
                    f"produtos encontrados | "
                    f"Arquivo: {caminho}"

                )

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    "ler o PDF.\n\n"
                    f"{erro}"
                )

            )

            self.info.config(
                text="Erro ao ler o PDF."
            )

    # ========================================================
    # MOSTRAR PRODUTOS
    # ========================================================

    def mostrar_produtos(
        self,
        produtos
    ):

        self.itens_exibidos = list(
            produtos
        )

        # Limpa tabela

        for item in self.tabela.get_children():

            self.tabela.delete(
                item
            )

        # Insere novamente

        for indice, produto in enumerate(
            self.itens_exibidos
        ):

            if (
                id(produto)
                in self.selecionados
            ):

                marcado = "☑"

            else:

                marcado = "☐"

            total_produto = (
                self.calcular_total_produto(
                    produto
                )
            )

            self.tabela.insert(

                "",

                "end",

                iid=str(indice),

                values=(

                    marcado,

                    produto.get(
                        "quantidade",
                        0
                    ),

                    produto.get(
                        "pagina",
                        ""
                    ),

                    produto.get(
                        "categoria",
                        ""
                    ),

                    produto.get(
                        "codigo",
                        ""
                    ),

                    produto.get(
                        "produto",
                        ""
                    ),

                    produto.get(
                        "embalagem",
                        ""
                    ),

                    (
                        self.formatar_moeda(
                            self.converter_preco(
                                produto.get(
                                    "preco",
                                    "0,00"
                                )
                            )
                        )
                        if produto.get(
                            "preco"
                        )
                        else ""
                    ),

                    self.formatar_moeda(
                        total_produto
                    )

                )

            )

    # ========================================================
    # CLIQUE NA TABELA
    # ========================================================

    def clique_tabela(
        self,
        evento
    ):

        regiao = self.tabela.identify(
            "region",
            evento.x,
            evento.y
        )

        if regiao != "cell":
            return

        coluna = self.tabela.identify_column(
            evento.x
        )

        linha = self.tabela.identify_row(
            evento.y
        )

        if not linha:
            return

        try:

            indice = int(linha)

            produto = (
                self.itens_exibidos[
                    indice
                ]
            )

        except (
            ValueError,
            IndexError
        ):

            return

        # ----------------------------------------------------
        # Coluna 1 = seleção
        # ----------------------------------------------------

        if coluna == "#1":

            chave = id(produto)

            if chave in self.selecionados:

                self.selecionados.remove(
                    chave
                )

            else:

                self.selecionados.add(
                    chave
                )

            self.mostrar_produtos(
                self.itens_exibidos
            )

            self.atualizar_total_geral()

        # ----------------------------------------------------
        # Coluna 2 = quantidade
        # ----------------------------------------------------

        elif coluna == "#2":

            self.editar_quantidade(
                linha,
                produto
            )

    # ========================================================
    # EDITAR QUANTIDADE
    # ========================================================

    def editar_quantidade(
        self,
        linha,
        produto
    ):

        bbox = self.tabela.bbox(
            linha,
            "#2"
        )

        if not bbox:
            return

        x, y, largura, altura = bbox

        entrada = tk.Entry(
            self.tabela,
            justify="center"
        )

        entrada.place(

            x=x,

            y=y,

            width=largura,

            height=altura

        )

        entrada.insert(
            0,
            str(
                produto.get(
                    "quantidade",
                    0
                )
            )
        )

        entrada.select_range(
            0,
            tk.END
        )

        entrada.focus()

        def salvar_quantidade(
            evento=None
        ):

            valor = entrada.get().strip()

            if not valor:

                valor = "0"

            try:

                quantidade = int(valor)

                if quantidade < 0:
                    raise ValueError

            except:

                messagebox.showwarning(

                    "Quantidade inválida",

                    "Digite uma quantidade inteira maior ou igual a zero."

                )

                entrada.focus()

                return

            produto["quantidade"] = (
                quantidade
            )

            entrada.destroy()

            self.mostrar_produtos(
                self.itens_exibidos
            )

            self.atualizar_total_geral()

        entrada.bind(
            "<Return>",
            salvar_quantidade
        )

        entrada.bind(
            "<FocusOut>",
            salvar_quantidade
        )

    # ========================================================
    # PRODUTOS SELECIONADOS
    # ========================================================

    def produtos_selecionados(self):

        return [

            produto

            for produto in self.produtos

            if id(produto)
            in self.selecionados

        ]

    # ========================================================
    # MARCAR TODOS
    # ========================================================

    def marcar_todos(self):

        self.selecionados = {

            id(produto)

            for produto in self.produtos

        }

        self.mostrar_produtos(

            self.itens_exibidos
            or self.produtos

        )

        self.atualizar_total_geral()

    # ========================================================
    # DESMARCAR TODOS
    # ========================================================

    def desmarcar_todos(self):

        self.selecionados.clear()

        self.mostrar_produtos(

            self.itens_exibidos
            or self.produtos

        )

        self.atualizar_total_geral()

    # ========================================================
    # FILTRO
    # ========================================================

    def filtrar(
        self,
        evento=None
    ):

        texto = (
            self.pesquisa
            .get()
            .strip()
            .lower()
        )

        if not texto:

            self.mostrar_produtos(
                self.produtos
            )

            self.atualizar_total_geral()

            return

        filtrados = []

        for produto in self.produtos:

            linha = " ".join([

                str(
                    produto.get(
                        "pagina",
                        ""
                    )
                ),

                produto.get(
                    "categoria",
                    ""
                ),

                produto.get(
                    "codigo",
                    ""
                ),

                produto.get(
                    "produto",
                    ""
                ),

                produto.get(
                    "embalagem",
                    ""
                ),

                produto.get(
                    "preco",
                    ""
                )

            ]).lower()

            if texto in linha:

                filtrados.append(
                    produto
                )

        self.mostrar_produtos(
            filtrados
        )

        self.atualizar_total_geral()

        self.info.config(

            text=(
                f"{len(filtrados)} "
                f"produto(s) encontrados."
            )

        )

    # ========================================================
    # SALVAR CSV COMPLETO
    # ========================================================

    def salvar_csv(self):

        if not self.produtos:

            messagebox.showwarning(

                "Atenção",

                "Abra um PDF primeiro."

            )

            return

        caminho = (
            filedialog.asksaveasfilename(

                title="Salvar produtos",

                defaultextension=".csv",

                filetypes=[
                    (
                        "Arquivo CSV",
                        "*.csv"
                    )
                ]

            )
        )

        if not caminho:
            return

        try:

            with open(

                caminho,

                "w",

                newline="",

                encoding="utf-8-sig"

            ) as arquivo:

                escritor = csv.DictWriter(

                    arquivo,

                    fieldnames=[

                        "pagina",

                        "categoria",

                        "codigo",

                        "produto",

                        "embalagem",

                        "preco",

                        "quantidade"

                    ],

                    delimiter=";"

                )

                escritor.writeheader()

                escritor.writerows(
                    self.produtos
                )

            messagebox.showinfo(

                "Concluído",

                "Arquivo CSV salvo com sucesso."

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    f"salvar o CSV.\n\n{erro}"
                )

            )

    # ========================================================
    # SALVAR TXT COMPLETO
    # ========================================================

    def salvar_txt(self):

        if not self.produtos:

            messagebox.showwarning(

                "Atenção",

                "Abra um PDF primeiro."

            )

            return

        caminho = (
            filedialog.asksaveasfilename(

                title="Salvar texto",

                defaultextension=".txt",

                filetypes=[
                    (
                        "Arquivo de texto",
                        "*.txt"
                    )
                ]

            )
        )

        if not caminho:
            return

        try:

            with open(

                caminho,

                "w",

                encoding="utf-8"

            ) as arquivo:

                categoria_anterior = ""

                for produto in self.produtos:

                    if (
                        produto["categoria"]
                        != categoria_anterior
                    ):

                        arquivo.write(
                            "\n"
                        )

                        arquivo.write(
                            "=" * 70
                            + "\n"
                        )

                        arquivo.write(
                            produto["categoria"]
                            + "\n"
                        )

                        arquivo.write(
                            "=" * 70
                            + "\n"
                        )

                        categoria_anterior = (
                            produto["categoria"]
                        )

                    arquivo.write(

                        f"{produto['codigo']} | "

                        f"{produto['produto']} | "

                        f"{produto['embalagem']} | "

                        f"R$ {produto['preco']} | "

                        f"Qtd: "
                        f"{produto.get('quantidade', 0)}\n"

                    )

            messagebox.showinfo(

                "Concluído",

                "Arquivo TXT salvo com sucesso."

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    f"salvar o TXT.\n\n{erro}"
                )

            )

    # ========================================================
    # SALVAR CSV DOS SELECIONADOS
    # ========================================================

    def salvar_csv_selecionados(
        self
    ):

        selecionados = (
            self.produtos_selecionados()
        )

        if not selecionados:

            messagebox.showwarning(

                "Atenção",

                (
                    "Marque pelo menos "
                    "um produto antes de salvar."
                )

            )

            return

        caminho = (
            filedialog.asksaveasfilename(

                title="Salvar produtos selecionados",

                defaultextension=".csv",

                filetypes=[
                    (
                        "Arquivo CSV",
                        "*.csv"
                    )
                ]

            )
        )

        if not caminho:
            return

        try:

            with open(

                caminho,

                "w",

                newline="",

                encoding="utf-8-sig"

            ) as arquivo:

                escritor = csv.DictWriter(

                    arquivo,

                    fieldnames=[

                        "pagina",

                        "categoria",

                        "codigo",

                        "produto",

                        "embalagem",

                        "preco",

                        "quantidade"

                    ],

                    delimiter=";"

                )

                escritor.writeheader()

                escritor.writerows(
                    selecionados
                )

            messagebox.showinfo(

                "Concluído",

                (
                    f"CSV salvo com "
                    f"{len(selecionados)} "
                    "produto(s) selecionado(s)."
                )

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    f"salvar o CSV.\n\n{erro}"
                )

            )

    # ========================================================
    # SALVAR TXT DOS SELECIONADOS
    # ========================================================

    def salvar_txt_selecionados(
        self
    ):

        selecionados = (
            self.produtos_selecionados()
        )

        if not selecionados:

            messagebox.showwarning(

                "Atenção",

                (
                    "Marque pelo menos "
                    "um produto antes de salvar."
                )

            )

            return

        caminho = (
            filedialog.asksaveasfilename(

                title=(
                    "Salvar texto dos "
                    "produtos selecionados"
                ),

                defaultextension=".txt",

                filetypes=[
                    (
                        "Arquivo de texto",
                        "*.txt"
                    )
                ]

            )
        )

        if not caminho:
            return

        try:

            with open(

                caminho,

                "w",

                encoding="utf-8"

            ) as arquivo:

                categoria_anterior = ""

                for produto in selecionados:

                    if (
                        produto["categoria"]
                        != categoria_anterior
                    ):

                        arquivo.write(
                            "\n"
                        )

                        arquivo.write(
                            "=" * 70
                            + "\n"
                        )

                        arquivo.write(
                            produto["categoria"]
                            + "\n"
                        )

                        arquivo.write(
                            "=" * 70
                            + "\n"
                        )

                        categoria_anterior = (
                            produto["categoria"]
                        )

                    total = (
                        self.calcular_total_produto(
                            produto
                        )
                    )

                    arquivo.write(

                        f"{produto['codigo']} | "

                        f"{produto['produto']} | "

                        f"{produto['embalagem']} | "

                        f"R$ {produto['preco']} | "

                        f"Qtd: "
                        f"{produto.get('quantidade', 0)} | "

                        f"Total: "
                        f"{self.formatar_moeda(total)}\n"

                    )

            messagebox.showinfo(

                "Concluído",

                (
                    f"TXT salvo com "
                    f"{len(selecionados)} "
                    "produto(s) selecionado(s)."
                )

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    f"salvar o TXT.\n\n{erro}"
                )

            )

    # ========================================================
    # GERAR PEDIDO CSV
    # ========================================================

    def gerar_pedido_csv(self):

        # ----------------------------------------------------
        # Pega somente os produtos selecionados
        # ----------------------------------------------------

        selecionados = (
            self.produtos_selecionados()
        )

        # ----------------------------------------------------
        # Filtra somente quantidade > 0
        # ----------------------------------------------------

        pedido = []

        for produto in selecionados:

            try:

                quantidade = int(
                    produto.get(
                        "quantidade",
                        0
                    )
                )

            except:

                quantidade = 0

            if quantidade > 0:

                pedido.append(
                    produto
                )

        # ----------------------------------------------------
        # Verifica se existe produto
        # ----------------------------------------------------

        if not pedido:

            messagebox.showwarning(

                "Pedido vazio",

                (
                    "Selecione produtos e "
                    "informe uma quantidade "
                    "maior que zero."
                )

            )

            return

        # ----------------------------------------------------
        # Escolher arquivo
        # ----------------------------------------------------

        caminho = (
            filedialog.asksaveasfilename(

                title="Gerar pedido",

                defaultextension=".csv",

                filetypes=[
                    (
                        "Arquivo CSV",
                        "*.csv"
                    )
                ]

            )
        )

        if not caminho:
            return

        try:

            total_pedido = 0.0

            quantidade_total = 0

            with open(

                caminho,

                "w",

                newline="",

                encoding="utf-8-sig"

            ) as arquivo:

                escritor = csv.writer(

                    arquivo,

                    delimiter=";"

                )

                # ------------------------------------------------
                # Cabeçalho
                # ------------------------------------------------

                escritor.writerow([

                    "CÓDIGO",

                    "PRODUTO",

                    "EMBALAGEM",

                    "PREÇO UNITÁRIO",

                    "QUANTIDADE",

                    "TOTAL"

                ])

                # ------------------------------------------------
                # Produtos
                # ------------------------------------------------

                for produto in pedido:

                    quantidade = int(
                        produto.get(
                            "quantidade",
                            0
                        )
                    )

                    preco = (
                        self.converter_preco(
                            produto.get(
                                "preco",
                                "0,00"
                            )
                        )
                    )

                    total_item = (
                        preco * quantidade
                    )

                    total_pedido += (
                        total_item
                    )

                    quantidade_total += (
                        quantidade
                    )

                    escritor.writerow([

                        produto.get(
                            "codigo",
                            ""
                        ),

                        produto.get(
                            "produto",
                            ""
                        ),

                        produto.get(
                            "embalagem",
                            ""
                        ),

                        f"{preco:.2f}".replace(
                            ".",
                            ","
                        ),

                        quantidade,

                        f"{total_item:.2f}".replace(
                            ".",
                            ","
                        )

                    ])

                # ------------------------------------------------
                # Linha vazia
                # ------------------------------------------------

                escritor.writerow([])

                # ------------------------------------------------
                # Total de quantidades
                # ------------------------------------------------

                escritor.writerow([

                    "",

                    "",

                    "",

                    "QUANTIDADE TOTAL",

                    quantidade_total,

                    ""

                ])

                # ------------------------------------------------
                # TOTAL DO PEDIDO
                # ------------------------------------------------

                escritor.writerow([

                    "",

                    "",

                    "",

                    "",

                    "TOTAL DO PEDIDO",

                    f"{total_pedido:.2f}".replace(
                        ".",
                        ","
                    )

                ])

            # ----------------------------------------------------
            # Mensagem final
            # ----------------------------------------------------

            messagebox.showinfo(

                "Pedido gerado",

                (

                    "Pedido gerado com sucesso!\n\n"

                    f"Produtos: {len(pedido)}\n"

                    f"Quantidade total: "
                    f"{quantidade_total}\n\n"

                    f"TOTAL DO PEDIDO: "
                    f"{self.formatar_moeda(total_pedido)}"

                )

            )

        except Exception as erro:

            messagebox.showerror(

                "Erro",

                (
                    "Não foi possível "
                    f"gerar o pedido.\n\n{erro}"
                )

            )


# ============================================================
# INÍCIO DO PROGRAMA
# ============================================================

if __name__ == "__main__":

    janela = tk.Tk()

    app = Aplicacao(
        janela
    )

    janela.mainloop()