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

df = pd.read_excel('participants.xlsx')
user = str(os.getenv("EMAIL"))
token = str(os.getenv("TOKEN"))


# -------- FUNÇÃO SEGURA: DERANGEMENT --------
def sortear_amigo_secreto(df):
    # Tem de ter mais de 1 participante
    if len(df) < 2:
        raise ValueError("É necessário ter pelo menos 2 participantes para o sorteio.")

    indices = list(df.index)
    sorteados = indices.copy()

    while True:
        random.shuffle(sorteados)
        # verifica se alguém calhou a si próprio
        if all(i != sorteados[idx] for idx, i in enumerate(indices)):
            break

    # Criar pares (participante -> amigo secreto)
    resultado = []
    for idx, i in enumerate(indices):
        participante = df.loc[i]
        amigo = df.loc[sorteados[idx]]
        resultado.append([participante, amigo])

    return resultado


# -------- VALIDAR --------
def validar_sorteio(df, sorteio):
    if len(sorteio) != len(df):
        return False

    nomes = set(df["Name"])

    atribuidores = set()
    atribuidos = set()

    for s, a in sorteio:
        if s['Name'] == a['Name']:
            return False
        atribuidores.add(s['Name'])
        atribuidos.add(a['Name'])

    return atribuidores == nomes and atribuidos == nomes


# -------- EMAIL --------
def send_email(user, token, destinatario_nome, destinatario_email, subject, data):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(user, token)

    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = destinatario_email
    msg['Subject'] = subject

    body = f"Ola {destinatario_nome},\n{data}\n\nCumprimentos"
    msg.attach(MIMEText(body, 'plain'))

    server.sendmail(user, destinatario_email, msg.as_string())
    server.quit()


# -------- MAIN --------
if __name__ == '__main__':
    print("Numero de participantes:", len(df))
    print("A sortear...")
    if( len(df) < 2 ):
        print("Erro: É necessário ter pelo menos 2 participantes para o sorteio.")
        sys.exit(1)
    
    # TEST MODE
    numeroTestes = 10
    if len(sys.argv) > 1 and sys.argv[1] in ["--test", "-t"]:
        print("MDODO TESTE ATIVADO - A executar", numeroTestes, "testes...")
        ok = 0
        fail = 0

        for i in range(1, numeroTestes + 1):
            sorteio = sortear_amigo_secreto(df)
            if validar_sorteio(df, sorteio):
                ok += 1
            else:
                fail += 1
                print(f"[FALHA] Teste {i}")
                for s, a in sorteio:
                    print(f"{s['Name']} -> {a['Name']}")

        print("\n========== RESULTADOS ==========")
        print("✔️ Sucessos:", ok)
        print("❌ Falhas:", fail)
        # Imprimir 1 resultado de exemplo
        print("\nExemplo de sorteio:")
        exemplo = sortear_amigo_secreto(df)
        for s, a in exemplo:
            print(f"{s['Name']} -> {a['Name']}")
    else:
        # MODO NORMAL
        sorteio = sortear_amigo_secreto(df)
        if not validar_sorteio(df, sorteio):
            print("Erro: Sorteio inválido.")
            sys.exit(1)

        for s, a in sorteio:
            data = f"O teu amigo secreto é {a['Name']}"
            send_email(user, token, s['Name'], s['Email'], "Jantar de Natal", data)
            print("[OK] Email enviado a", s['Name'])

        print("Sorteio concluido!")
