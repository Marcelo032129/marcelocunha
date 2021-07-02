import pyscreenshot as s
from time import sleep
import rpa as r
import pyautogui as p
import win32com.client as w32

# Abrir o Navegador, entrar no link BI e maximizar a tela
r.init()
r.url('http://bi.rumolog.com/pbirs/powerbi/Manuten%C3%A7%C3%A3o%20Sul/Via/Engenharia/Detec%C3%A7%C3%A3o%20Rondas?rs:embed=true')
p.sleep(2)
janela = p.getActiveWindow()
janela.maximize()
p.sleep(2)

# Digitar usuário e senha 
p.typewrite('cs261742')
p.hotkey('Tab')
p.sleep(2)
p.typewrite('rumo@2021')
p.hotkey('Tab')
p.press('enter')

# Clicar para não salvar usuário e senha da página
p.doubleClick(1240, 385)

p.sleep(15)

# Relizar a captura da tabela do relatório
image = s.grab(bbox=(165, 165, 1250, 670))
image.save("DeteccaoRonda.png")

# Criar a integração com outolook
outlook = w32.Dispatch('outlook.application')

# Criar um email
email = outlook.CreateItem(0)

# Configurar as informações do seu email
email.To = "marcelo.cunha@rumolog.com"
email.Subject = "Relatório de Detecção de Rondas"
email.HTMLBody = """
<html>
    <head></head>
    <body>
    </p> Bom dia, segue relatório de aderência a Detecção de Via </p>

    <img src="C://Users//cs261742//Desktop//Bot//DeteccaoRonda.png">

    <br/>

    <img src="C://Users//cs261742//Desktop//Bot//AssinaturaEmail.png">

    </body>
</html>
"""

email.send()

