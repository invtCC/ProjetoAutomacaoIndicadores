# Importar bibliotecas
import pandas as pd
import smtplib
import email.message
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
import os
import pathlib


# Importar base de dados
emails = pd.read_excel(r"Bases de Dados\Emails.xlsx")
lojas = pd.read_csv(r"Bases de Dados\Lojas.csv", encoding="latin1", sep=";")
vendas = pd.read_excel(r"Bases de Dados\Vendas.xlsx")

# Incluir nome da loja em vendas
vendas = vendas.merge(lojas, on="ID Loja")


dicionario_lojas = {}

for loja in lojas["Loja"]:
    dicionario_lojas[loja] = vendas.loc[vendas["Loja"] ==loja, :]


dia_indicador = vendas["Data"].max()

print(dia_indicador)
print("{}/{}".format(dia_indicador.day, dia_indicador.month))


# Identificar se a pasta já existe
caminho_backup = pathlib.Path(r"Backup Arquivos Lojas")

arquivos_pasta_backup = caminho_backup.iterdir()


lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()


    # Salvar dentro da pasta
    nome_arquivo = "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day, loja)

    local_arquivo = caminho_backup / loja / nome_arquivo

    dicionario_lojas[loja].to_excel(local_arquivo)


# Definição de metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500




for loja in dicionario_lojas:

    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja["Data"]==dia_indicador, :]

    # Faturamento
    faturamento_ano = vendas_loja["Valor Final"].sum()

    faturamento_dia = vendas_loja_dia["Valor Final"].sum()


    # Diversidade de produtos
    qtde_produtos_ano = len(vendas_loja["Produto"].unique())


    qtde_produtos_dia = len(vendas_loja_dia["Produto"].unique())


    # ticket médio
    valor_venda = vendas_loja.groupby("Código Venda").sum(numeric_only=True)
    ticket_medio_ano = valor_venda["Valor Final"].mean()


    valor_venda_dia = vendas_loja_dia.groupby("Código Venda").sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia["Valor Final"].mean()


        # Enviar Email via SMTP com anexos

    def enviar_email():

        nome = emails.loc[emails["Loja"]==loja, "Gerente"].values[0]

        # Criar mensagem multipart
        msg = MIMEMultipart()
        msg["Subject"] = f"OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"
        msg["From"] = "youremail@gmail.com"
        msg["To"] = emails.loc[emails["Loja"]==loja, "E-mail"].values[0]

        # Adicionar corpo do e-mail como parte HTML

        if faturamento_dia >= meta_faturamento_dia:
            cor_fat_dia = "green"
        else:
            cor_fat_dia = "red"

        if faturamento_ano >= meta_faturamento_ano:
            cor_fat_ano = "green"
        else:
            cor_fat_ano = "red"          
        
        if qtde_produtos_dia >= meta_qtdeprodutos_dia:
            cor_qtde_dia = "green"
        else:
            cor_qtde_dia = "red"

        if qtde_produtos_ano >= meta_qtdeprodutos_ano:
            cor_qtde_ano = "green"
        else:
            cor_qtde_ano = "red"

        if ticket_medio_dia >= meta_faturamento_dia:
            cor_ticket_dia = "green"
        else:
            cor_ticket_dia = "red"

        if ticket_medio_ano >= meta_ticketmedio_ano:
            cor_ticket_ano = "green"
        else:
            cor_ticket_ano = "red"
            
            

        corpo_email = f"""
        <p>Bom dia, <sttrong>{nome}</strong></p>

        <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

        <table>
            <tr>
                <th>Indicador</th>
                <th>Valor Dia</th>
                <th>Meta Dia</th>
                <th>Cenário Dia</th>
            </tr>
            <tr>
                <td style="text-align: center">Faturamento</td>
                <td style="text-align: center">R${faturamento_dia:.2f}</td>
                <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
                <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Diversidade de Produtos</td>
                <td style="text-align: center">{qtde_produtos_dia}</td>
                <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
                <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Ticlet Médio</td>
                <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
                <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
                <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
            </tr>
        </table>
        <br>
        <table>
            <tr>
                <th>Indicador</th>
                <th>Valor Ano</th>
                <th>Meta Ano</th>
                <th>Cenário Ano</th>
            </tr>
            <tr>
                <td style="text-align: center">Faturamento</td>
                <td style="text-align: center">R${faturamento_ano:.2f}</td>
                <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
                <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Diversidade de Produtos</td>
                <td style="text-align: center">{qtde_produtos_ano}</td>
                <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
                <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
            </tr>
            <tr>
                <td style="text-align: center">Ticlet Médio</td>
                <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
                <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
                <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
            </tr>
        </table>

        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

        <p>Qualquer dúvida estou à disposição</p>

        <p>Att., Your_Name</p>
        """

        parte_html = MIMEText(corpo_email, "html", "utf-8")
        msg.attach(parte_html)

        attachment = pathlib.Path.cwd() / caminho_backup / loja / f"{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx"

        if os.path.exists(attachment):
                # Determinar o tipo MIME automaticamente (opcional, mas recomendado)
                tipo_mime, _ = mimetypes.guess_type(attachment)
                if tipo_mime is None:
                    tipo_mime = 'application/octet-stream' # Tipo padrão para arquivos desconhecidos
                
                tipo_principal, subtipo_mime = tipo_mime.split('/', 1)

                with open(attachment, 'rb') as f:
                    # Criar um objeto MIMEBase para o anexo
                    anexo = MIMEBase(tipo_principal, subtipo_mime)
                    anexo.set_payload(f.read())
                
                # Codificar o anexo em Base64 para envio
                encoders.encode_base64(anexo)
                
                # Adicionar cabeçalho Content-Disposition para que o cliente de e-mail saiba que é um anexo
                anexo.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
                
                # Anexar o objeto anexo à mensagem principal
                msg.attach(anexo)
                print(f"Arquivo {os.path.basename(attachment)} anexado com sucesso.")
        else:
            print(f"Erro: Arquivo '{attachment}' não encontrado.")
            return

        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(msg["From"], "mpua quty abjk ysjb")
        servidor.send_message(msg)
        servidor.quit()
        print("Email enviado")


    enviar_email()



faturamento_lojas = vendas.groupby("Loja")[["Loja", "Valor Final"]].sum(numeric_only=True)

faturamento_lojas_ano = faturamento_lojas.sort_values(by="Valor Final", ascending=False)



nome_arquivo = "{}_{}_Ranking Anual.xlsx".format(dia_indicador.month, dia_indicador.day)

faturamento_lojas_ano.to_excel(fr"Backup Arquivos Lojas\{nome_arquivo}")


vendas_dia = vendas.loc[vendas["Data"]==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby("Loja")[["Loja", "Valor Final"]].sum(numeric_only=True)

faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by="Valor Final", ascending=False)


nome_arquivo = "{}_{}_Ranking Dia.xlsx".format(dia_indicador.month, dia_indicador.day)

faturamento_lojas_dia.to_excel(fr"Backup Arquivos Lojas\{nome_arquivo}")



# Enviar Email via SMTP com anexos

def enviar_email():

    # Criar mensagem multipart
    msg = MIMEMultipart()
    msg["Subject"] = f"Ranking Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"
    msg["From"] = "youremail@gmail.com"
    msg["To"] = emails.loc[emails["Loja"]=="Diretoria", "E-mail"].values[0]

    # Adicionar corpo do e-mail como parte HTML
        

    corpo_email = f"""
    Prezados, bom dia
    <br>
    <br>
    Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
    <br>
    Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, -1]:.2f}

    <br>
    <br>

    Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
    <br>
    Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, -1]:.2f}
    <br>
    <br>
    Segue em anexo os rankings do ano e do dia de todas as lojas
    <br>
    <br>
    Qualquer Dúvida estou à disposição
    <br>
    <br>
    att., Your_name
    """

    parte_html = MIMEText(corpo_email, "html", "utf-8")
    msg.attach(parte_html)

    attachment_ano = pathlib.Path.cwd() / caminho_backup / f"{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx"

    attachment_dia = pathlib.Path.cwd() / caminho_backup / f"{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx"

    if os.path.exists(attachment_dia):
            # Determinar o tipo MIME automaticamente (opcional, mas recomendado)
            tipo_mime, _ = mimetypes.guess_type(attachment_ano)
            if tipo_mime is None:
                tipo_mime = 'application/octet-stream' # Tipo padrão para arquivos desconhecidos
            
            tipo_principal, subtipo_mime = tipo_mime.split('/', 1)

            with open(attachment_ano, 'rb') as f:
                # Criar um objeto MIMEBase para o anexo
                anexo = MIMEBase(tipo_principal, subtipo_mime)
                anexo.set_payload(f.read())
            
            # Codificar o anexo em Base64 para envio
            encoders.encode_base64(anexo)
            
            # Adicionar cabeçalho Content-Disposition para que o cliente de e-mail saiba que é um anexo
            anexo.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_ano)}')
            
            # Anexar o objeto anexo à mensagem principal
            msg.attach(anexo)
            print(f"Arquivo {os.path.basename(attachment_ano)} anexado com sucesso.")
    else:
        print(f"Erro: Arquivo '{attachment_ano}' não encontrado.")
        return
    
    if os.path.exists(attachment_dia):
            # Determinar o tipo MIME automaticamente (opcional, mas recomendado)
            tipo_mime, _ = mimetypes.guess_type(attachment_dia)
            if tipo_mime is None:
                tipo_mime = 'application/octet-stream' # Tipo padrão para arquivos desconhecidos
            
            tipo_principal, subtipo_mime = tipo_mime.split('/', 1)

            with open(attachment_dia, 'rb') as f:
                # Criar um objeto MIMEBase para o anexo
                anexo = MIMEBase(tipo_principal, subtipo_mime)
                anexo.set_payload(f.read())
            
            # Codificar o anexo em Base64 para envio
            encoders.encode_base64(anexo)
            
            # Adicionar cabeçalho Content-Disposition para que o cliente de e-mail saiba que é um anexo
            anexo.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_dia)}')
            
            # Anexar o objeto anexo à mensagem principal
            msg.attach(anexo)
            print(f"Arquivo {os.path.basename(attachment_dia)} anexado com sucesso.")
    else:
        print(f"Erro: Arquivo '{attachment_dia}' não encontrado.")
        return    

    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(msg["From"], "mpua quty abjk ysjb")
    servidor.send_message(msg)
    servidor.quit()
    print("Email enviado")


enviar_email()
