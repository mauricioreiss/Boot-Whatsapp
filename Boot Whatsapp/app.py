#Boot para mensagens no whatsapp seguindo uma planilha com contatos

#importar todas as dependencias do projeto
import openpyxl 
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#abrir o whatsweb
webbrowser.open('https://web.whatsapp.com/')
sleep(30)
#sleep serve para ter um delay de segundos entre uma açao e outra 


#leio a planilha "Numeros.xlsx"
Planilha = openpyxl.load_workbook('Numeros.xlsx')
#pego a pagina que esta os contatos
pagina = Planilha['Folha1']

#pego para cada linha sem contar a primeira que esta com os titulos e atribuo a variaveis  
for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value

#monto a mensagem
    mensagem = f'eae {nome} me paga caloteiro, seu prazo e ate dia {data.strftime('%d/%m/%Y')}'

    
    #tento entrar no link com a mensagem personalizada 
    try:
        link_message = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_message)
        sleep[10]
        #pega a imagem seta que eu adicionei e procura uma igual no site no caso o botao de enviar do whatsapp é a seta
        seta = pyautogui.locateCenterOnScreen('Seta.png')
        sleep(5)
        #clica na seta
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        #fecha a pagina para poder abrir outra seguindo a planilha e monstando nova mensagem para novo numero 
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        #trata erro e salva em arquivo
        print(f'nao foi possivel enviar mensagem a {nome}')
        with open('erro.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome} ,{telefone}')