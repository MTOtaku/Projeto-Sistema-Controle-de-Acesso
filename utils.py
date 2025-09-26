import sqlite3
import cv2
import win32com.client
from datetime import datetime
import os
from dotenv import load_dotenv
load_dotenv("dadosSensiveis.env")

"""
    Parte de vozes do windows
"""

speaker = win32com.client.Dispatch("SAPI.SpVoice")
voices = speaker.GetVoices()

for i, voice in enumerate(speaker.GetVoices()):
    print(i,voice.getDescription())

speaker.Rate=2 #-10 a 10


while True:
    try:
        idioma = int(input("Escolha seu idioma preferido para TTS: "))
        if 0 <= idioma < voices.Count:
            break  # Se estiver dentro do intervalo, sai do loop
        else:
            print(f"Erro: escolha inválida. Digite um número entre 0 e {voices.Count - 1}.")
    except ValueError:
        print("Erro, coloque uma das opções.")



speaker.Voice=speaker.GetVoices().Item(idioma)
print(f"Idioma escolhido: {speaker.GetVoices().Item(idioma).GetDescription()}")

#Código do e-mail

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurações do servidor
smtp_server = "smtp.gmail.com"
smtp_port = 587
email_usuario = os.getenv("EMAIL_USUARIO") #Coloque aqui o email em que é pra ser logado
email_senha = os.getenv("EMAIL_SENHA") #Coloque aqui a chave de API do envio de email

if not email_usuario :
    print("EMAIL_USUARIO não carregados!")
if not email_senha:
    print("EMAIL_SENHA não carregado!")
    exit()


#
# Inicialização
#

camera = cv2.VideoCapture(0)
if not camera.isOpened():
    print("Erro ao abrir a câmera")
    exit()

conexao = sqlite3.connect("sca.db")
cursor = conexao.cursor()
cursor.execute("""
               CREATE TABLE IF NOT EXISTS usuarios(
                   CPF VARCHAR(20) PRIMARY KEY,
                   nome VARCHAR(255) NOT NULL,
                   cartao VARCHAR (15) NOT NULL,
                   foto VARCHAR(50),
                   email VARCHAR(50) NOT NULL,
                   telefone VARCHAR(20) NOT NULL)
               """)

cursor.execute("""
               CREATE TABLE IF NOT EXISTS registros(
                   ID INTEGER PRIMARY KEY,
                   CPF VARCHAR(20) NOT NULL,
                   data DATE,
                   hora TIME,
                   FOREIGN KEY (CPF) REFERENCES usuarios(CPF)
               )
               """)

conexao.commit()

#
# Cadastrar
#
def cadastrar():
    print("\n=== Cadastro de Usuário ===")
    cpf = input("CPF: ")
    nome = input("Nome: ")
    cartao = input("Cartão: ")
    print("Pressione 'g' para salvar a foto.")

    while True:
        ret, frame = camera.read()
        if not ret:
            print("Erro ao capturar da câmera")
            break

        cv2.imshow("Câmera", frame)
        tecla = cv2.waitKey(1)

        if tecla == ord('g'):
            filename = f"f{cpf}.png"
            cv2.imwrite(filename, frame)
            break

    camera.release()
    cv2.destroyAllWindows()
    foto = filename
    email = input("Email: ")
    telefone = input("Telefone: ")

    try:
        cursor.execute("""
                       INSERT INTO usuarios (CPF, nome, cartao, foto, email, telefone)
                       VALUES (?, ?, ?, ?, ?, ?)
                       """, (cpf, nome, cartao, foto, email, telefone))
        conexao.commit()
        print("Usuário cadastrado com sucesso!")
    except sqlite3.Error as e:
        print("Ocorreu um erro: ", e)


#
# Buscar por CPF
#
def buscar_cpf():
    print("=== Buscar por CPF ===")
    cpf = input("Informe o CPF: ")
    cursor.execute("SELECT * FROM usuarios WHERE CPF = ?", (cpf,))
    resultado = cursor.fetchone()

    if resultado:
        print("CPF: ", resultado[0])
        print("Nome: ", resultado[1])
        print("Cartão: ", resultado[2])
        print("Email: ", resultado[4])
        print("Telefone: ", resultado[5])

        if resultado[3]:
            img = cv2.imread(resultado[3])
            cv2.imshow(f"Foto de {resultado[1]}", img)
            cv2.waitKey(2000)
            cv2.destroyAllWindows()

    else:
        print("CPF não cadastrado!")


#
# Buscar por Cartão
#
def buscar_cartao(cartao):
    cursor.execute("SELECT * FROM usuarios WHERE cartao = ?", (cartao,))
    resultado = cursor.fetchone()

    if resultado:
        print("CPF: ", resultado[0])
        print("Nome: ", resultado[1])
        print("Cartão: ", resultado[2])
        print("Email: ", resultado[4])
        print("Telefone: ", resultado[5])

        if resultado[3]:
            img = cv2.imread(resultado[3])
            cv2.imshow(f"Foto de {resultado[1]}", img)
            cv2.waitKey(2000)

            hora = datetime.now().hour
            if 6 <= hora < 12:
                speaker.Speak("Bom dia, " + resultado[1])
            elif 12 <= hora < 18:
                speaker.Speak("Boa tarde, " + resultado[1])
            else:
                speaker.Speak("Boa noite, " + resultado[1])

            cv2.destroyAllWindows()

            # bloco de registro
            agora = datetime.now()
            data = agora.date()
            hora = agora.strftime("%H:%M:%S")

            cursor.execute("INSERT INTO registros (CPF, data, hora) VALUES (?, ?, ?)",(resultado[0], data, hora))
            conexao.commit()

            #Bloco de email
            remetente = email_usuario
            destinatario = resultado[4]
            assunto = f"Acesso: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
            corpo = f"{resultado[1]} | {assunto}"

            msg = MIMEMultipart()
            msg['From'] = remetente
            msg['To'] = destinatario
            msg['Subject'] = assunto
            msg.attach(MIMEText(corpo, 'plain'))

            try:
                # Conectar ao servidor
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()  # segurança
                server.login(email_usuario, email_senha)

                # Enviar e-mail
                server.sendmail(remetente, destinatario, msg.as_string())
                print("E-mail enviado com sucesso!")

                server.quit()
            except Exception as e:
                print(f"Erro ao enviar e-mail: {e}")

    else:
        print("Cartão não cadastrado!")
        speaker.Speak("Acesso bloqueado!")


#
# Listar todos os usuarios
#
def listar_usuarios():
    cursor.execute("SELECT * FROM usuarios")
    resultados = cursor.fetchall()
    if resultados:
        print("Lista de usuários cadastrados")
        for usuario in resultados:

            print(f"CPF: {usuario[0]}")
            print(f"Nome: {usuario[1]}")
            print(f"Cartão: {usuario[2]}")
            print(f"Email: {usuario[4]}")
            print(f"Telefone: {usuario[5]}")
            print("*"*50)
    else:
        print("Não há usuários cadastrados!")

#
# EXCLUIR
#

def excluir_usuario():
    print("*** Excluir Usuário ****")
    cpf = input("Informe o CPF: ")

    cursor.execute("SELECT * FROM usuarios WHERE CPF = ?", (cpf,))
    resultado = cursor.fetchone()

    if resultado:
        confirma = input(f"deseja REALMENTE apagar? S/N {resultado[1]}: ")
        if confirma == "S":
            cursor.execute("DELETE FROM usuarios WHERE CPF = ?", (cpf,))
            print("Apagado com sucesso!")
            conexao.commit()
        else:
            print("Comando não executado")
    else:
        print("Usuario não cadastrado!")


#
# Editar usuário
#
def editar_usuario():
    print("==== Edição de usuário ====")
    cpf = input("Informe o CPF: ")
    cursor.execute("SELECT * FROM usuarios WHERE CPF = ?", (cpf,))
    resultado = cursor.fetchone()

    if resultado:
        print("Pressione ENTER para não mudar nada")
        novo_nome = input(f"Nome [{resultado[1]}]: ").strip()
        novo_cartao = input(f"Cartão [{resultado[2]}]: ").strip()
        novo_email = input(f"Email [{resultado[4]}]: ").strip()
        novo_tel = input(f"Telefone [{resultado[5]}]: ").strip()

        if not novo_nome:
            novo_nome = resultado[1]
        if not novo_cartao:
            novo_cartao = resultado[2]
        if not novo_email:
            novo_email = resultado[4]
        if not novo_tel:
            novo_tel = resultado[5]

        cursor.execute("""
                       UPDATE usuarios
                       SET nome=?,
                           cartao=?,
                           email=?,
                           telefone=?
                       WHERE cpf = ?
                       """, (novo_nome, novo_cartao, novo_email, novo_tel, cpf))
        conexao.commit()
    else:
        print("CPF não cadastrado!")

#
# Menu
#
def menu():
    while True:
        print("=== SCA - Sistema de Controle de Acesso")
        print("1. Cadastrar")
        print("2. Buscar")
        print("3. Listar Usuários")
        print("4. Excluir Usuário")
        print("5. Editar Usuário")
        print("6. Sair")
        opcao = input("Escolha uma opção: ")

        if opcao == "1":
            cadastrar()
        elif opcao == "2":
            buscar_cpf()
        elif opcao == "3":
            listar_usuarios()
        elif opcao == "4":
            excluir_usuario()
        elif opcao == "5":
            editar_usuario()
        elif opcao == "6":
            print("Saindo do sistema")
            conexao.close()
            break
        else:
            buscar_cartao(opcao)