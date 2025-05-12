import yagmail
import os
import pandas as pd
import sys

def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, compatível com PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

def ler_template_email(caminho_template, nome, curso):
    with open(resource_path(caminho_template), 'r', encoding='utf-8') as f:
        template = f.read()
    return template.format(nome=nome, curso=curso)

def enviar_emails(planilha, pasta_destino, ano, caminho_template, usuario, senha, smtp_host=None, smtp_porta=None):
    df = pd.read_excel(planilha)
    colunas_normalizadas = [col.strip().lower() for col in df.columns]
    df.columns = colunas_normalizadas
    if 'email' not in df.columns:
        raise Exception("A planilha deve conter uma coluna chamada 'email'.")
    if 'nome' not in df.columns:
        raise Exception("A planilha deve conter uma coluna chamada 'nome'.")
    if 'documento' not in df.columns:
        raise Exception("A planilha deve conter uma coluna chamada 'documento'.")
    if 'produto' not in df.columns:
        raise Exception("A planilha deve conter uma coluna chamada 'produto'.")
    if smtp_host and smtp_porta:
        yag = yagmail.SMTP(usuario, senha, host=smtp_host, port=int(smtp_porta))
    elif smtp_host:
        yag = yagmail.SMTP(usuario, senha, host=smtp_host)
    else:
        yag = yagmail.SMTP(usuario, senha)
    enviados = 0
    erros = []
    sucessos = []
    df_unicos = df.drop_duplicates(subset=['email'])
    for idx, row in df_unicos.iterrows():
        email = str(row['email']).strip()
        nome = str(row['nome']).strip()
        documento = str(row['documento']).strip()
        curso = str(row['produto']).strip() if 'produto' in row else ''
        if not email or email.lower() in ['nan', 'none', '(none)']:
            continue
        pdf_nome = f"{''.join(c for c in nome if c.isalnum() or c in (' ', '_', '-')).replace(' ', '_')}_{ano}.pdf"
        pdf_path = os.path.join(pasta_destino, pdf_nome)
        if not os.path.exists(pdf_path):
            continue  # Apenas ignora se não existir o PDF
        corpo = ler_template_email(caminho_template, nome, curso)
        assunto = "Recibo de Pagamento - Esperançar"
        try:
            yag.send(to=email, subject=assunto, contents=corpo, attachments=pdf_path)
            enviados += 1
            print(f"E-mail enviado com sucesso para {nome} <{email}>")
            sucessos.append(f"{nome} <{email}>: enviado com sucesso")
        except Exception as e:
            print(f"Erro ao enviar para {nome} <{email}>: {e}")
            erros.append(f"{nome} <{email}>: erro ao enviar - {e}")
    return enviados, sucessos, erros 