####Script para inicializar automaticamente programas que geralmente
####sao abertos assim que inicializa-se o Sistema Operacional.
####Autor: Lucas Ribeiro
####Python Version: 2.7.12
####Criacao: 07/02/2017
####Ultima modificacao: 17/05/2017


#Importando bibliotecas necessarias----
import os, time #os para incializar arquivos e time para manipular tempo
import win32com.client as win #win32com.client para usar uma funcao para simular o teclado
import win32api, win32con #win32apu e win32con para as funcoes que imitam o click esquerdo e direito do mouse

#Definindo funcoes para simular clicks do mouse----
def click_left(x,y):
    win32api.SetCursorPos((x,y)) #seta o cursor no monitor segundo as coordenadas que voce passar -> 15, 250 por exemplo
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0) #faz com que o comando que o click utiliza seja enviado
    time.sleep(0.05)#faz o programa dormir por 0.05 segundos -> medida preventiva para possiveis bugs
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0) #faz com que o comando que o click utiliza seja enviado

def click_right(x,y):
    win32api.SetCursorPos((x,y)) #seta o cursor no monitor segundo as coordenadas que voce passar -> 15, 250 por exemplo
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0) #faz com que o comando que o click utiliza seja enviado
    time.sleep(0.05) #faz o programa dormir por 0.05 segundos -> medida preventiva para possiveis bugs
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0) #faz com que o comando que o click utiliza seja enviado

#Inicio----
shell = win.Dispatch('WScript.Shell') #a variavel shell sera uma instancia da classe WScript.Shell podendo enviar comandos do teclado

#Abaixo tem um exemplo de inicializacao do OUTLOOK. Note que apos a abertura o programa e colocado para dormir por 30 sec
#isso eh feito pois esse comando nao aguarda um retorno informando que o programa foi aberto com sucesso. Ele simplesmente o starta
'''
os.startfile(r'C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE')
time.sleep(20)
'''

#Alimentando as variaveis que serao utilizadas para login, etc
email_senha = #coloque aqui a senha
facebook_senha = #coloque aqui a senha
email = #coloque aqui o email

#Abrindo o Chrome----
os.startfile(r'C:\Program Files\Google\Chrome\Application\chrome.exe')
time.sleep(20)

#Abrindo minhas paginas----
shell.SendKeys('^(+n)') #comandos do WScript.Shell na web. Nesse caso ^ eh CONTROL, + eh SHIFT e n eh n mesmo, logo CTRL+SHIFT+n
time.sleep(5) #esperando abrir a nova janela
#Deezer----
shell.SendKeys('www.deezer.com')
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(1)
#Abaixo esta um exemplo de como simular o click do mouse. Cada monitor eh diferente, entao deixei apenas como exemplo
'''
#Eu realizo o login pelo facebook, mas abro o deezer primeiro pois ele costuma levar um tempo para carregar a pagina
#entao eu abro para deixar carregando e, enquanto isso, vou abrindo e logando em outras paginas
#esse pedaco de codigo deveria estar apos o login no facebook
click_left(500,500) #os numeros sao hipoteticos, o teste eh realizado na tentativa e erro
#o click seria exatamente no botao para logar com facebook
'''
#Gmail----
shell.SendKeys('^t') #abrir nova aba. CTRL+t
time.sleep(1)
shell.SendKeys('www.gmail.com')
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(7) #esperando carregar o gmail
shell.SendKeys(email) #preenche o campo email com a string da variavel email
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(4) #espera carregar pagina para senha
shell.SendKeys(email_senha)
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(1)
#Facebook----
shell.SendKeys('^t')
time.sleep(1)
shell.SendKeys('www.facebook.com')
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(10)
shell.SendKeys(facebook_senha)
time.sleep(0.5)
shell.SendKeys('{TAB}')
shell.SendKeys(bla2)
time.sleep(0.5)
shell.SendKeys('{ENTER}')
time.sleep(1)
#Logando no deezer----
shell.SendKeys('^{TAB}') #vai para a aba do deezer
time.sleep(1)



####Os conceitos apresentados acima sao crus e a nivel de inciante. Existem frameworks para a automacao decente de tarefas
####um exemplo eh o selenium, que eu utilizo atualmente.
####Mas para entendimento inicial, recomendo construir um script besta para automatizar tarefas suas do dia a dia :)
####Espero que goste!

####Viva o OpenSource!!

