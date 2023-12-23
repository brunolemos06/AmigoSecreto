import smtplib
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
import pandas as pd
from dotenv import load_dotenv
import os
import sys

load_dotenv()

# Configurações
df = pd.read_excel('participants.xlsx')
user = os.getenv("USER")
token = os.getenv("TOKEN")

# Função para sortear o Amigo Secreto
def sortear_amigo_secreto(df):
    sorteio = []
    copia = df.copy()

    for _, sorteado in df.iterrows():
        amigo_secreto = sorteado
        while amigo_secreto.equals(sorteado):
            amigo_secreto = copia.sample(n=1).iloc[0]

        sorteio.append([sorteado, amigo_secreto])
        copia = copia[copia.index != amigo_secreto.index[0]]

    return sorteio

# Função para enviar o email
def send_email(user, token, destinatario_nome, destinatario_email,data):
    # Conectar ao servidor SMTP
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(user, token)

    # Criar a mensagem
    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = destinatario_email
    msg['Subject'] = "Jantar de Natal"

    # Corpo do email
    body = f"Ola {destinatario_nome},\n"
    body += (data + "\n\n")
    body += "Cumprimentos"
    msg.attach(MIMEText(body, 'plain'))

    # Enviar email
    server.sendmail(user, destinatario_email, msg.as_string())
    server.quit()


if __name__ == '__main__':
    print("Numero de participantes: " + str(len(df)))
    print("A sortear...")
    sorteio = sortear_amigo_secreto(df)
    # if exists arg with --test just print the result and don't send the emails
    if len(sys.argv) > 1 and (sys.argv[1] == "--test" or sys.argv[1] == "--t" or sys.argv[1] == "-test" or sys.argv[1] == "-t"):
        for (sorteado) in sorteio:
            print(sorteado[0]['Name'] + " -> " + sorteado[1]['Name'])
    else:
        for (sorteado) in sorteio:
            data = "O teu amigo secreto é " + str(sorteado[1]['Name'])
            # send_email(user, token, str(sorteado[0]['Name']), str(sorteado[0]['Email']), data)
            print("[OK] Email enviado a " + str(sorteado[0]['Name']) + " com sucesso!")

    print("Sorteio concluido!")