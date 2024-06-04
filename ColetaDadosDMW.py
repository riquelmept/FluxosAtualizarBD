import pyautogui
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from screeninfo import get_monitors
import shutil
import os

time.sleep(2)
pyautogui.PAUSE = 0.5
selected_monitor = None
siteDMW = "SITE QUE DESEJA EXTRAIR OS DADOS"

def get_screen_width_height():
    monitors = get_monitors()
    current_monitor = None
    for monitor in monitors:
        if monitor.x == 0 and monitor.y == 0:
            current_monitor = monitor
            break
    return current_monitor.width, current_monitor.height

"""Passo 1: Acessar DMW
Nesta etapa há a necessidade da identificação da tela para ajuste das coordenadas do mouse
"""
#Abrir o navegador
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(2)
pyautogui.write(siteDMW)
pyautogui.press("enter")

#Confirmar em qual monitor foi aberto a janela
# Função para selecionar o monitor    
def select_monitor(monitor):
    global selected_monitor
    selected_monitor = monitor
    root.destroy()

# Função para evitar fechar a janela sem selecionar um monitor
def on_closing():
    global selected_monitor
    if selected_monitor is None:
        messagebox.showinfo("Aviso", "Por favor, selecione um monitor.")
    else:
        root.destroy()

# Criar a janela principal
root = tk.Tk()
screen_width, screen_height = get_screen_width_height()
root.geometry(f"+{screen_width // 2}+{screen_height // 2}")  # Centralizar na tela atual
root.title("Seleção de Monitor")
root.protocol("WM_DELETE_WINDOW", on_closing)  # Chama on_closing quando a janela é fechada

# Criar os botões para selecionar o monitor
monitors = ['Notebook Internal Display 14"', 'Monitor Len S24e-03', 'Monitor AOC 27G2G5']
for monitor in monitors:
    button = tk.Button(root, text=monitor, command=lambda m=monitor: select_monitor(m))
    button.pack(pady=5)

# Iniciar a janela
root.mainloop()
Tela = selected_monitor

def identyMonitor(monitor):
    if monitor == 'Monitor Len S24e-03':
        return {
            "CampoLogin": (875, -701),
            "CampoSenha": (886, -655),
            "BotaoDashboard": (951, -447),
            "CampoIMMO": (542, -786),
            "CampoFIMMO": (1085, -783),
            "BotaoSubmit": (1394, -735),
            "BotaoExport": (1513, -730)
        }
    elif monitor == 'Monitor AOC 27G2G5':
        return { 
            "CampoLogin": (-1021, 199),
            "CampoSenha": (-993, 237),
            "BotaoDashboard": (-964, 440),
            "CampoIMMO": (-1387, 116),
            "CampoFIMMO": (-832, 112),
            "BotaoSubmit": (-531, 153),
            "BotaoExport": (-417, 158)
        }
    else:
        return {
            "CampoLogin": (930, 385),
            "CampoSenha": (890, 423),
            "BotaoDashboard": (924, 632),
            "CampoIMMO": (663, 305),
            "CampoFIMMO": (1124, 302),
            "BotaoSubmit": (1390, 345),
            "BotaoExport": (1501, 352)
        }
        
user = "USUÁRIO"
pin = "SENHA"
campos = identyMonitor(Tela)
#pyautogui.click(x=-1869, y=-19)
#time.sleep(1)

pyautogui.click(*campos["CampoLogin"])
pyautogui.write(user)
pyautogui.press("tab")
pyautogui.write(pin)
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(3)

"""Passo 2: Selecionar janela de visualizar KPIs 
"""
pyautogui.click(*campos["BotaoDashboard"])
time.sleep(3)

"""Passo 3:  Abrir dados de cada planta do CSV
Estrutura CSV: IMMO, FIMMO, Year, Table
IMMO: Define a Planta
FIMMO: Define o processo(Código de cada planta - Crushing)
"""
tabela = pd.read_csv("Endereços.csv")
print(tabela)

#Laço para puxar arquivo de cada linha do csv
for linha in tabela.index:
    pyautogui.click(*campos["CampoIMMO"])
    pyautogui.write(str(tabela.loc[linha, "IMMO"]))
    pyautogui.press("enter")
    pyautogui.press("tab")
    pyautogui.press("enter")
    pyautogui.write(str(tabela.loc[linha, "FIMMO"]))
    pyautogui.press("enter")
    pyautogui.click(*campos["BotaoSubmit"])
    time.sleep(4)
    pyautogui.click(*campos["BotaoExport"])
    time.sleep(18)

# Mover arquivos baixados para a pasta de destino com base no CSV
relacao_arquivos = pd.read_csv("dmwloc.csv")# Caminho para o seu arquivo CSV com a relação dos arquivos
print(relacao_arquivos.head())

for _, row in relacao_arquivos.iterrows():
    origem = row['dir_atual']
    destino = row['dir_dest']
    
    # Criar a pasta de destino se ela não existir
    destino_dir = os.path.dirname(destino)
    if not os.path.exists(destino_dir):
        os.makedirs(destino_dir)
    
    # Mover o arquivo
    if os.path.exists(origem):
        shutil.move(origem, destino)
    else:
        print(f"Arquivo não encontrado: {origem}")

# Criar a janela principal
janela = tk.Tk()
janela.title("Atualização Finalizada!")
# Criar o rótulo com a mensagem
mensagem = tk.Label(janela, text="Atualização Finalizada!", font=("Arial", 14))
mensagem.pack(padx=10, pady=10)

# Executar o loop principal da interface gráfica
janela.mainloop()
