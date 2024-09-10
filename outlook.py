import time
import openpyxl
import win32com.client
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import re

pyautogui.FAILSAFE = False

def enviar_whatsapp(telefone, mensagem):
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(9)
    pyautogui.typewrite('\n')
    sleep(8)
    fechar_aba_navegador()

def fechar_aba_navegador():
    pyautogui.hotkey('ctrl', 'w')

def ler_planilha(planilha_caminho, palavra_chave):
    workbook = openpyxl.load_workbook(planilha_caminho)
    pagina = workbook['Planilha1']  # Certifique-se de que este nome corresponde à aba real

    dados = []
    for linha in pagina.iter_rows(min_row=2):
        Grupo = linha[0].value
        Telefone = linha[2].value
        if Grupo and palavra_chave.lower() == Grupo.lower():
            dados.append((Grupo, Telefone))
    return dados

def marcar_como_lido(email):
    email.Unread = False
    email.Save()

def filtrar_corpo(corpo):
    palavras_chave = ['TI', 'AUM', 'AUE', 'MMC', 'MEC', 'MEL', 'QUA', 'LOG', 'PRD', 'RTB', 'EXC', 'FER', 'GEF']

    for palavra_chave in palavras_chave:
        if re.search(rf'\b{re.escape(palavra_chave)}\b', corpo, re.IGNORECASE):
            return palavra_chave
    return None

def determinar_grupo_por_atraso(tempo_atraso):
    if tempo_atraso <= 10:
        return 'Grupo Encarregado', 'encarregado.xlsx'
    elif tempo_atraso <= 20:
        return 'Grupo Supervisor', 'supervisor.xlsx'
    elif tempo_atraso <= 30:
        return 'Grupo Gerente', 'gerente.xlsx'
    else:
        return 'Grupo Presidencia', 'presidencia.xlsx'

def monitorar_outlook(planilha1_caminho):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    caixa_entrada = outlook.GetDefaultFolder(6)
    mensagens_pendentes = []
    
    while True:
        pyautogui.click(x=0, y=0)
        
        if mensagens_pendentes:
            mensagem = mensagens_pendentes.pop(0)
            enviar_whatsapp(mensagem['telefone'], mensagem['mensagem'])
            print(f"Mensagem enviada para {mensagem['telefone']}.")
        else:
            itens_ordenados = sorted(caixa_entrada.Items, key=lambda x: x.ReceivedTime, reverse=True)
            novos_nao_lidos = [msg for msg in itens_ordenados if msg.Unread and msg.SenderEmailAddress == 'alertas@directaautomacao.com.br']
            
            if novos_nao_lidos:
                print(f"Novo(s) e-mail(s) não lido(s) detectado(s) às {time.strftime('%H:%M:%S')}")

                for nova_mensagem in novos_nao_lidos:
                    print(f"E-mail recebido de: {nova_mensagem.SenderEmailAddress} em {nova_mensagem.ReceivedTime}")

                    corpo = nova_mensagem.Body
                    if 'Abertura de Ciclo de Ajuda' in corpo:
                        planilha = planilha1_caminho
                        palavra_chave = filtrar_corpo(corpo)
                    elif 'Aviso de Escalonamento de Card do Ciclo de Ajuda' in corpo:
                        tempo_atraso = int(re.search(r'Tempo Atraso: (\d+) minutos', corpo).group(1))
                        grupo, planilha = determinar_grupo_por_atraso(tempo_atraso)
                        palavra_chave = filtrar_corpo(corpo)
                    else:
                        planilha = None
                        palavra_chave = None

                    if planilha and palavra_chave:
                        dados = ler_planilha(planilha, palavra_chave)
                        for Grupo, Telefone in dados:
                            mensagem = f'{Grupo}, \n {corpo}'
                            mensagens_pendentes.append({'telefone': Telefone, 'mensagem': mensagem})
                        
                    marcar_como_lido(nova_mensagem)
                    print("E-mail marcado como lido.")
                    
                print("Verificação de novos e-mails concluída.")
            else:
                print(f"Aguardando novos e-mails... {time.strftime('%H:%M:%S')}")
                time.sleep(5)

if __name__ == "__main__":
    # Defina os caminhos completos para as planilhas
    planilha1_caminho = os.path.join(os.getcwd(), 'planilha1.xlsx')
    planilha2_caminho = os.path.join(os.getcwd(), 'encarregado.xlsx')
    planilha3_caminho = os.path.join(os.getcwd(), 'supervisor.xlsx')
    planilha4_caminho = os.path.join(os.getcwd(), 'gerente.xlsx')
    planilha5_caminho = os.path.join(os.getcwd(), 'presidencia.xlsx')

    webbrowser.open('https://web.whatsapp.com/')
    sleep(5)
    
    while True:
        monitorar_outlook(planilha1_caminho)
        time.sleep(100)
