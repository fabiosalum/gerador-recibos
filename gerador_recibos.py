import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
from num2words import num2words
import os
import difflib

def selecionar_arquivo():
    ano_selecionado = ano_combobox.get()
    if not ano_selecionado:
        messagebox.showerror("Erro", "Por favor, selecione um ano.")
        return
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar os recibos")
        if not pasta_destino:
            messagebox.showwarning("Aviso", "Seleção de pasta cancelada. O processo foi interrompido.")
            return
        gerar_recibos(arquivo, ano_selecionado, pasta_destino)

def formatar_valor(valor):
    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def encontrar_coluna(df, nome_procurado, cutoff=0.7):
    colunas = list(df.columns)
    # Tenta correspondência exata primeiro
    for col in colunas:
        if col == nome_procurado:
            return col
    # Se não encontrar, faz fuzzy matching
    correspondencia = difflib.get_close_matches(nome_procurado, colunas, n=1, cutoff=cutoff)
    if correspondencia:
        return correspondencia[0]
    return None

# Lista de possíveis nomes para o campo nome
possiveis_nomes = [
    'nome',
    'nome do(a) comprador(a)',
    'nome do comprador',
    'comprador(a)'
]

def encontrar_coluna_nome(df):
    colunas = list(df.columns)
    # Tenta correspondência exata
    for nome in possiveis_nomes:
        for col in colunas:
            if col == nome:
                return col
    # Fuzzy matching se não encontrar exato
    for nome in possiveis_nomes:
        correspondencia = difflib.get_close_matches(nome, colunas, n=1, cutoff=0.7)
        if correspondencia:
            return correspondencia[0]
    return None

def gerar_recibos(caminho_arquivo, ano, pasta_destino):
    try:
        df = pd.read_excel(caminho_arquivo)

        # Exibe os nomes das colunas lidos para o usuário
        colunas_lidas = list(df.columns)
        colunas_normalizadas = [col.strip().lower() for col in colunas_lidas]
        messagebox.showinfo(
            "Colunas encontradas",
            "Colunas lidas da planilha:\n" + "\n".join(colunas_lidas) + "\n\nColunas normalizadas:\n" + "\n".join(colunas_normalizadas)
        )

        # Normaliza os nomes das colunas: remove espaços e converte para minúsculas
        df.columns = colunas_normalizadas

        # Campos obrigatórios normalizados
        campos_obrigatorios = [
            'data da transação', 'produto', 'valor de compra com impostos',
            'quantidade de cobranças', 'nome', 'documento'
        ]
        mapeamento = {}
        mapeamento_detalhado = []
        for campo in campos_obrigatorios:
            if campo == 'nome':
                coluna_encontrada = encontrar_coluna_nome(df)
            else:
                coluna_encontrada = encontrar_coluna(df, campo)
            mapeamento_detalhado.append(f"'{campo}' => '{coluna_encontrada}'")
            if not coluna_encontrada:
                messagebox.showerror(
                    "Erro",
                    f"Campo obrigatório ausente na planilha: '{campo}'\n\nColunas normalizadas:\n" + "\n".join(colunas_normalizadas) + f"\n\nMapeamento realizado até agora:\n" + "\n".join(mapeamento_detalhado)
                )
                return
            mapeamento[campo] = coluna_encontrada

        # Exibe o mapeamento realizado
        mapeamento_str = '\n'.join([f"{campo} => {coluna}" for campo, coluna in mapeamento.items()])
        messagebox.showinfo("Mapeamento de colunas", f"Mapeamento realizado:\n{mapeamento_str}")

        # Converter a coluna de data para datetime
        df[mapeamento['data da transação']] = pd.to_datetime(
            df[mapeamento['data da transação']],
            format='%d/%m/%Y',
            errors='coerce'
        )

        # Filtrar apenas os pagamentos do ano selecionado
        ano_int = int(ano)
        df_filtrado = df[df[mapeamento['data da transação']].dt.year == ano_int]
        if df_filtrado.empty:
            messagebox.showwarning("Aviso", f"Nenhum pagamento encontrado para o ano {ano}.")
            return

        print("Todas as datas lidas:")
        print(df[mapeamento['data da transação']])
        print("Datas após o filtro do ano:")
        print(df_filtrado[mapeamento['data da transação']])

        # Carrega o template
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template('template_recibo.html')

        # Criar pasta de destino, se não existir
        os.makedirs(pasta_destino, exist_ok=True)

        # Agrupar por documento (CPF) apenas no DataFrame filtrado
        for documento, grupo in df_filtrado.groupby(mapeamento['documento']):
            if grupo.empty:
                continue  # Não gera recibo para grupos vazios (por segurança extra)
            nome = grupo[mapeamento['nome']].iloc[0]
            curso = ', '.join(sorted(grupo[mapeamento['produto']].unique()))
            valor_total_num = grupo[mapeamento['valor de compra com impostos']].sum()
            valor_total = formatar_valor(valor_total_num)
            valor_extenso = num2words(valor_total_num, lang='pt_BR', to='currency')

            # Ordena as parcelas pela data (mais antiga primeiro)
            grupo_ordenado = grupo.sort_values(by=mapeamento['data da transação'])
            parcelas = []
            for idx, (_, row) in enumerate(grupo_ordenado.iterrows(), 1):
                data_pagamento = row[mapeamento['data da transação']]
                if isinstance(data_pagamento, pd.Timestamp):
                    data_pagamento = data_pagamento.strftime("%d/%m/%Y")
                else:
                    data_pagamento = str(data_pagamento)
                valor_parcela = row[mapeamento['valor de compra com impostos']]
                parcelas.append({
                    'numero': idx,
                    'data': data_pagamento,
                    'valor': formatar_valor(valor_parcela)
                })

            # Determina o período do recibo (primeiro ao último pagamento, apenas mês e ano)
            def mes_extenso(mes):
                meses = [
                    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
                ]
                return meses[mes - 1]

            datas_pagamento = grupo_ordenado[mapeamento['data da transação']].dropna().sort_values()
            if not datas_pagamento.empty:
                data_inicio = datas_pagamento.iloc[0]
                data_fim = datas_pagamento.iloc[-1]
                if not isinstance(data_inicio, pd.Timestamp):
                    data_inicio = pd.to_datetime(data_inicio)
                if not isinstance(data_fim, pd.Timestamp):
                    data_fim = pd.to_datetime(data_fim)
                mes_inicio = mes_extenso(data_inicio.month)
                ano_inicio = data_inicio.year
                mes_fim = mes_extenso(data_fim.month)
                ano_fim = data_fim.year
                if ano_inicio == ano_fim:
                    periodo = f"período de {mes_inicio} a {mes_fim} de {ano_fim}"
                else:
                    periodo = f"período de {mes_inicio} de {ano_inicio} a {mes_fim} de {ano_fim}"
            else:
                periodo = ""

            # Quantidade de parcelas pagas e percentual do curso
            qtd_parcelas_pagas = len(parcelas)
            total_parcelas_curso = 18
            percentual_pago = (qtd_parcelas_pagas / total_parcelas_curso) * 100
            percentual_pago_str = f"{percentual_pago:.2f}".replace('.', ',')

            def caminho_weasyprint(path):
                return 'file:///' + os.path.abspath(path).replace('\\', '/')

            logo_esperancar_path = caminho_weasyprint("logo-esperancar.jpg")
            print("Logo Esperançar:", logo_esperancar_path)
            print("Existe?", os.path.exists("logo-esperançar.jpg"))
            logo_unita_path = caminho_weasyprint("logo-unita.jpg")
            assinatura_path = caminho_weasyprint("assinatura.png")

            # Renderiza o HTML
            html_out = template.render(
                nome=nome,
                cpf=documento,
                valor_total=valor_total,
                valor_extenso=valor_extenso,
                curso=curso,
                periodo=periodo,
                parcelas=parcelas,
                data_recibo=datetime.today().strftime("%d/%m/%Y"),
                percentual_pago=percentual_pago_str,
                logo_esperancar_path=logo_esperancar_path,
                logo_unita_path=logo_unita_path,
                assinatura_path=assinatura_path
            )

            # Garante que o nome é string e remove caracteres inválidos para nome de arquivo
            nome_str = str(nome).strip()
            nome_str = ''.join(c for c in nome_str if c.isalnum() or c in (' ', '_', '-')).replace(' ', '_')
            if not nome_str:
                nome_str = str(documento)

            # Salva PDF
            nome_arquivo = os.path.join(pasta_destino, f"{nome_str}_{ano}.pdf")
            HTML(string=html_out).write_pdf(nome_arquivo)

        messagebox.showinfo("Sucesso", f"Recibos gerados na pasta selecionada: {pasta_destino}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Interface
root = tk.Tk()
root.title("Gerador de Recibos")

# Frame principal
main_frame = tk.Frame(root, padx=20, pady=20)
main_frame.pack(expand=True, fill='both')

# Observação sobre os campos obrigatórios
obs_text = (
    "Observação: A planilha deve conter obrigatoriamente as seguintes colunas:\n"
    "- Data da transação\n"
    "- produto\n"
    "- Valor de compra com impostos\n"
    "- Quantidade de cobranças\n"
    "- nome\n"
    "- documento"
)
obs_label = tk.Label(main_frame, text=obs_text, fg="red", justify="left")
obs_label.pack(pady=(0, 15), anchor="w")

# Label e Combobox para seleção do ano
ano_label = tk.Label(main_frame, text="Selecione o ano:")
ano_label.pack(pady=(0, 5))

# Criar lista de anos (do ano atual até 5 anos atrás)
ano_atual = datetime.now().year
anos = [str(ano) for ano in range(ano_atual - 5, ano_atual + 1)]

ano_combobox = ttk.Combobox(main_frame, values=anos, state="readonly", width=10)
ano_combobox.set(str(ano_atual))  # Define o ano atual como padrão
ano_combobox.pack(pady=(0, 20))

# Label e botão para seleção do arquivo
label = tk.Label(main_frame, text="Clique no botão abaixo para selecionar a planilha:")
label.pack(pady=10)

botao = tk.Button(main_frame, text="Selecionar Planilha", command=selecionar_arquivo)
botao.pack(pady=20)

root.mainloop()
