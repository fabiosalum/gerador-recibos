import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
from num2words import num2words
import os
import difflib
import yagmail
from email_utils import enviar_emails
from tkinter.simpledialog import askstring
import sys

# Variáveis globais para armazenar os caminhos
arquivo_selecionado = None
pasta_destino = None

def selecionar_arquivo():
    global arquivo_selecionado
    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
    )
    if arquivo:
        arquivo_selecionado = arquivo
        label_arquivo.config(text=f"Arquivo selecionado: {os.path.basename(arquivo)}")
        botao_gerar.config(state='normal')

def selecionar_pasta():
    global pasta_destino
    pasta = filedialog.askdirectory(title="Selecione a pasta para salvar os recibos")
    if pasta:
        pasta_destino = pasta
        label_pasta.config(text=f"Pasta selecionada: {os.path.basename(pasta)}")

def gerar_recibos():
    global arquivo_selecionado, pasta_destino
    
    ano_selecionado = ano_combobox.get()
    mes_inicial = mes_para_numero(mes_inicial_combobox.get())
    mes_final = mes_para_numero(mes_final_combobox.get())
    
    if not ano_selecionado:
        messagebox.showerror("Erro", "Por favor, selecione um ano.")
        return
        
    if not mes_inicial or not mes_final:
        messagebox.showerror("Erro", "Por favor, selecione o período inicial e final.")
        return
        
    if int(mes_inicial) > int(mes_final):
        messagebox.showerror("Erro", "O mês inicial não pode ser maior que o mês final.")
        return
        
    if not arquivo_selecionado:
        messagebox.showerror("Erro", "Por favor, selecione um arquivo.")
        return
        
    if not pasta_destino:
        messagebox.showerror("Erro", "Por favor, selecione uma pasta de destino.")
        return

    try:
        df = pd.read_excel(arquivo_selecionado)

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

        # Filtrar por ano e período
        ano_int = int(ano_selecionado)
        mes_inicial_int = int(mes_inicial)
        mes_final_int = int(mes_final)
        
        # Criar data inicial e final do período
        data_inicial = pd.Timestamp(year=ano_int, month=mes_inicial_int, day=1)
        if mes_final_int == 12:
            data_final = pd.Timestamp(year=ano_int, month=mes_final_int, day=31)
        else:
            data_final = pd.Timestamp(year=ano_int, month=mes_final_int + 1, day=1) - pd.Timedelta(days=1)
        
        # Aplicar filtro
        df_filtrado = df[
            (df[mapeamento['data da transação']].dt.year == ano_int) &
            (df[mapeamento['data da transação']].dt.month >= mes_inicial_int) &
            (df[mapeamento['data da transação']].dt.month <= mes_final_int)
        ]
        
        if df_filtrado.empty:
            messagebox.showwarning(
                "Aviso", 
                f"Nenhum pagamento encontrado para o período de {mes_inicial}/{ano_selecionado} a {mes_final}/{ano_selecionado}."
            )
            return

        print("Todas as datas lidas:")
        print(df[mapeamento['data da transação']])
        print("Datas após o filtro do período:")
        print(df_filtrado[mapeamento['data da transação']])

        # Carrega o template
        env = Environment(loader=FileSystemLoader(resource_path('.')))
        template = env.get_template('template_recibo.html')

        # Criar pasta de destino, se não existir
        os.makedirs(pasta_destino, exist_ok=True)

        # Agrupar por documento (CPF) apenas no DataFrame filtrado
        for documento, grupo in df_filtrado.groupby(mapeamento['documento']):
            if (
                grupo.empty
                or not documento
                or pd.isna(documento)
                or str(documento).strip() == ""
                or str(documento).strip().lower() in ["(none)", "none"]
            ):
                continue  # Não gera recibo para grupos sem CPF/documento
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
                return 'file:///' + resource_path(path).replace('\\', '/')

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
            nome_arquivo = os.path.join(pasta_destino, f"{nome_str}_{ano_selecionado}.pdf")
            HTML(string=html_out).write_pdf(nome_arquivo)

        messagebox.showinfo("Sucesso", f"Recibos gerados na pasta selecionada: {pasta_destino}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def formatar_valor(valor):
    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def encontrar_coluna(df, nome_procurado, cutoff=0.7):
    import difflib
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

possiveis_nomes = [
    'nome',
    'nome do(a) comprador(a)',
    'nome do comprador',
    'comprador(a)'
]

def encontrar_coluna_nome(df):
    import difflib
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

def disparar_emails():
    global pasta_destino, arquivo_selecionado
    if not pasta_destino or not arquivo_selecionado:
        messagebox.showerror("Erro", "Selecione a planilha e a pasta de destino antes de disparar os e-mails.")
        return
    try:
        usuario = askstring("E-mail remetente", "Digite o e-mail remetente (SMTP):")
        senha = askstring("Senha de aplicativo", "Digite a senha de aplicativo do e-mail:", show='*')
        smtp_host = askstring("Servidor SMTP", "Servidor SMTP (ex: smtp.gmail.com) [deixe em branco para automático]:")
        smtp_porta = askstring("Porta SMTP", "Porta SMTP (ex: 587) [deixe em branco para automático]:")
        if not usuario or not senha:
            messagebox.showerror("Erro", "Usuário e senha são obrigatórios.")
            return
        ano = ano_combobox.get()
        enviados, sucessos, erros = enviar_emails(
            planilha=arquivo_selecionado,
            pasta_destino=pasta_destino,
            ano=ano,
            caminho_template="email_template.html",
            usuario=usuario,
            senha=senha,
            smtp_host=smtp_host,
            smtp_porta=smtp_porta
        )
        resumo = ""
        if sucessos:
            resumo += "E-mails enviados com sucesso:\n" + "\n".join(sucessos) + "\n\n"
        if erros:
            resumo += "Falhas no envio:\n" + "\n".join(erros)
        if resumo:
            messagebox.showinfo("Resumo do envio", resumo)
        else:
            messagebox.showinfo("Resumo do envio", "Nenhum e-mail foi enviado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao disparar os e-mails: {e}")

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

# Frame para seleção de período
periodo_frame = tk.Frame(main_frame)
periodo_frame.pack(pady=(0, 15))

# Label e Combobox para seleção do ano
ano_label = tk.Label(periodo_frame, text="Ano:")
ano_label.grid(row=0, column=0, padx=5)

# Criar lista de anos (do ano atual até 5 anos atrás)
ano_atual = datetime.now().year
anos = [str(ano) for ano in range(ano_atual - 5, ano_atual + 1)]

ano_combobox = ttk.Combobox(periodo_frame, values=anos, state="readonly", width=10)
ano_combobox.set(str(ano_atual))  # Define o ano atual como padrão
ano_combobox.grid(row=0, column=1, padx=5)

# Label e Combobox para seleção do mês inicial
mes_inicial_label = tk.Label(periodo_frame, text="Mês inicial:")
mes_inicial_label.grid(row=0, column=2, padx=5)

meses = [
    ("Janeiro", "1"), ("Fevereiro", "2"), ("Março", "3"), ("Abril", "4"),
    ("Maio", "5"), ("Junho", "6"), ("Julho", "7"), ("Agosto", "8"),
    ("Setembro", "9"), ("Outubro", "10"), ("Novembro", "11"), ("Dezembro", "12")
]

mes_inicial_combobox = ttk.Combobox(periodo_frame, values=[m[0] for m in meses], state="readonly", width=15)
mes_inicial_combobox.set("Janeiro")
mes_inicial_combobox.grid(row=0, column=3, padx=5)

# Label e Combobox para seleção do mês final
mes_final_label = tk.Label(periodo_frame, text="Mês final:")
mes_final_label.grid(row=0, column=4, padx=5)

mes_final_combobox = ttk.Combobox(periodo_frame, values=[m[0] for m in meses], state="readonly", width=15)
mes_final_combobox.set("Dezembro")
mes_final_combobox.grid(row=0, column=5, padx=5)

# Frame para botões de seleção
botoes_frame = tk.Frame(main_frame)
botoes_frame.pack(pady=15)

# Botão para selecionar arquivo
botao_arquivo = tk.Button(botoes_frame, text="Selecionar Planilha", command=selecionar_arquivo)
botao_arquivo.pack(side=tk.LEFT, padx=5)

# Label para mostrar arquivo selecionado
label_arquivo = tk.Label(botoes_frame, text="Nenhum arquivo selecionado")
label_arquivo.pack(side=tk.LEFT, padx=5)

# Botão para selecionar pasta
botao_pasta = tk.Button(botoes_frame, text="Selecionar Pasta de Destino", command=selecionar_pasta)
botao_pasta.pack(side=tk.LEFT, padx=5)

# Label para mostrar pasta selecionada
label_pasta = tk.Label(botoes_frame, text="Nenhuma pasta selecionada")
label_pasta.pack(side=tk.LEFT, padx=5)

# Botão para gerar recibos (inicialmente desabilitado)
botao_gerar = tk.Button(main_frame, text="Gerar Recibos", command=gerar_recibos, state='disabled')
botao_gerar.pack(pady=20)

# Adicionar botão para disparar e-mails
botao_email = tk.Button(main_frame, text="Disparar E-mails", command=disparar_emails)
botao_email.pack(pady=10)

# Função para converter nome do mês para número
def mes_para_numero(nome_mes):
    for mes in meses:
        if mes[0] == nome_mes:
            return mes[1]
    return "1"

# Função para atualizar o mês final quando o mês inicial for alterado
def atualizar_mes_final(event):
    mes_inicial = mes_inicial_combobox.get()
    mes_final = mes_final_combobox.get()
    
    # Encontrar os índices dos meses selecionados
    idx_inicial = next(i for i, m in enumerate(meses) if m[0] == mes_inicial)
    idx_final = next(i for i, m in enumerate(meses) if m[0] == mes_final)
    
    # Se o mês inicial for maior que o final, atualizar o final
    if idx_inicial > idx_final:
        mes_final_combobox.set(mes_inicial)

# Adicionar evento de mudança ao combobox do mês inicial
mes_inicial_combobox.bind('<<ComboboxSelected>>', atualizar_mes_final)

def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, compatível com PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

root.mainloop()
