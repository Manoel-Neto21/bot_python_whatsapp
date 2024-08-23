import openpyxl
import pyautogui
from urllib.parse import quote
import webbrowser
from time import sleep

# abre a o wahtsapp web para login e depois fecha
webbrowser.open('https://web.whatsapp.com/')
sleep(20)
pyautogui.hotkey('ctrl', 'w')
sleep(5)

# ler a planilha Excel
workbook = openpyxl.load_workbook('planilha_teste_bot_Whatsapp.xlsx')

# Defina qual planilha dentro da planilha será pega
pagina_clientes = workbook['Lista1']

# Pega os dados a partir da linha definida
for linha in pagina_clientes.iter_rows(min_row=2):
    
# Extrai os dados de linha e coluna
    nome = linha[0].value
    telefone = linha[1].value

# Caso a coluna tenha um campo de data, pode criar uma variavel e inserir a compo exemplo.strftime('%D
# Envia os dados da mensagem,  campo pode ser personalizado com a mensagens e liks desejados
    mensagem = f'Olá {nome} , sou o Gestorzinho, um bot muito esperto e estou aqui para ajudar. Veja nosso site https://bussoladagestao.com.br/'

    # Link para o envio das mensagem e numeros https://web.whatsapp.com/send?phone=&text=, quote serve para estruturar a mensagem

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    # Abre o navegador para o envio das mensagens aos contatos da planilha
        webbrowser.open(link_mensagem_whatsapp)
        sleep(30)
        # Localizada o botão de enviar do whatsapp
        seta = pyautogui.locateCenterOnScreen('Foto.png')
        sleep(10)
        pyautogui.click(seta[0],seta[1])
        sleep(10)
        # Fecha a guia depois de enviar a mensagem
        pyautogui.hotkey('ctrl', 'w')
        sleep(10)
    # Trata as mensagens que não foram enviadas
    except:
        print(f'Não foi possível enviar a mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
            
# Avisa quando a lista de contatos foi finalizada
print("Planilha finalizada!")