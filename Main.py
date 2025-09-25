# SCA - Sistema de Controle de Acesso
# V1 - 19/09/2025 - Implementação inicial do projeto
# V2 - 25/09/2025 - Implementação de fala, consulta ao DB
import sqlite3
import cv2
import win32com.client
import datetime

"""
    Parte de vozes do windows
"""

speaker = win32com.client.Dispatch("SAPI.SpVoice")

for i, voice in enumerate(speaker.GetVoices()):
    print(i,voice.getDescription())

speaker.Rate=2 #-10 a 10


while True:
    try:
        idioma = int(input("Escolha seu idioma preferido para TTS: "))
        break
    except ValueError:
        print("Erro, coloque um número inteiro válido.")


speaker.Voice=speaker.GetVoices().Item(idioma)
print(f"Idioma escolhido: {speaker.GetVoices().Item(idioma).GetDescription()}")


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
# Buscar
#
def buscar():
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
# Menu
#
def menu():
    while True:
        print("=== SCA - Sistema de Controle de Acesso")
        print("1. Cadastrar")
        print("2. Buscar")
        print("6. Sair")
        opcao = input("Escolha uma opção: ")

        if opcao == "1":
            cadastrar()
        if opcao == "2":
            buscar()
        elif opcao == "6":
            print("Saindo do sistema")
            conexao.close()
            break
        else:
            print("Opção inválida")


#
# Programa principal
#
menu()


