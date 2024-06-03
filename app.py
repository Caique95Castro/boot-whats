import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Abrir o WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Exibir o diretório atual (para debug)
print(os.getcwd())

# Carregar a planilha de clientes
try:
    workbook = openpyxl.load_workbook('clientes.xlsx')
    pagina_clientes = workbook['Planilha1']
except FileNotFoundError:
    print("Erro: O arquivo 'clientes.xlsx' não foi encontrado.")
    exit()

# Processar cada linha na planilha
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Mensagem personalizada
    mensagem = (f'Olá {nome}, é o Caique e estou fazendo um teste, não responda. '
                f'Seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. '
                'Pague na data, Obrigado')

    # Link da mensagem personalizada
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)

    try:
        # Localizar a seta e clicar
        seta = pyautogui.locateCenterOnScreen('image.png')
        if seta:
            sleep(5)
            pyautogui.click(seta[0], seta[1])
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            sleep(5)
        else:
            raise Exception("Imagem da seta não encontrada")

    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}\n')


