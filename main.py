# BIBLIOTECAS
# ============================================================================= #
import requests
import shutil
import os
import pythoncom
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
from datetime import datetime
from pywinauto.application import Application
from time import sleep
from win32com.client import Dispatch

pyautogui.PAUSE = 0.1

from dotenv import load_dotenv
# load_dotenv(dotenv_path=r'C:\Python\GitHub\fin-rpa_baixa_arquivo_cobranca\.gitignore\.env')
load_dotenv()
url_sistema = os.getenv('url_sistema')
user = os.getenv('user')
password = os.getenv('password')
path_servidor = os.getenv('path_servidor')
path_chromedriver = os.getenv('path_chromedriver')

print(f'Automatização Iniciada!\n\n\
Localiza os arquivos ".RET" e renomeia no padrão de importação do ERP!\n\
Buscando arquivos na pasta Retorno, banco 237\n')

# FUNÇÃO E-MAIL
# ============================================================================= #
def enviar_email_outlook(assunto, corpo, anexos=None):
    pythoncom.CoInitialize()

    try:
        outlook = Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        mail.To = "; ".join(TO_EMAILS)
        mail.Cc = "; ".join(CC_EMAILS)
        mail.Subject = assunto
        mail.Body = corpo

        if anexos:
            for arquivo in anexos:
                mail.Attachments.Add(arquivo)

        mail.Send()
    finally:
        pythoncom.CoUninitialize()

# FUNÇÃO PARA LOG DE ERROS
# ============================================================================= #
rel_log = []
pasta_log = Path(f'{path_servidor}\\LOG')

def log(log_message):
    rel_log.append(log_message)

# INICIO DO CÓDIGO
# ============================================================================= #
try:
    # CAMINHO DAS PASTAS E ARQUIVOS
    # ============================================================================= #
    pasta_raiz = Path(path_servidor)
    pasta_662 = Path(f'{path_servidor}\\662')
    pasta_663 = Path(f'{path_servidor}\\663')

    arquivos_raiz = [f for f in os.listdir(pasta_raiz) if '.RET' in f and not f.endswith('.RET') and not f.endswith('.pdf')]
    arquivos_662 = [f for f in os.listdir(pasta_662) if '.RET' in f and not f.endswith('.RET') and not f.endswith('.pdf')]
    arquivos_663 = [f for f in os.listdir(pasta_663) if '.RET' in f and not f.endswith('.RET') and not f.endswith('.pdf')]
    log(f"Arquivos encontrados - Pasta Raiz: {arquivos_raiz} - Pasta 662: {arquivos_662} - Pasta 663: {arquivos_663}\n")

    # RENOMEIA OS ARQUIVOS
    # ============================================================================= #
    nome_antigo_raiz = []
    nome_antigo_662 = []
    nome_antigo_663 = []

    if arquivos_raiz != [] or arquivos_662 != [] or arquivos_663 != []:
        # Arquivos 661
        if arquivos_raiz != []:
            print(f"[🔄] Renomeando arquivos (pasta raiz):")
            for nome in arquivos_raiz:
                novo_nome = nome.split('.RET')[0] + f'-{datetime.now():%d%m%y}' + '.RET'
                caminho_antigo = pasta_raiz / nome
                caminho_novo = pasta_raiz / novo_nome
                nome_antigo_raiz.append(nome)
                try:
                    os.rename(caminho_antigo, caminho_novo)
                    print(f"[✅] Renomeado: {nome} -> {novo_nome}\n")
                    log(f"Arquivo renomeado - Pasta Raiz: {nome} -> {novo_nome}\n")
                except Exception as e:
                    print(f"[❌] Erro ao renomear {nome}: {e}")
                    log(f"Erro ao renomear {nome}: {e}\n")
                continue
        # Arquivos 662
        if arquivos_662 != []:
            print(f"[🔄] Renomeando arquivos (pasta 662):")
            for nome in arquivos_662:
                novo_nome = nome.split('.RET')[0] + f'-{datetime.now():%d%m%y}' + '.RET'
                caminho_antigo = pasta_662 / nome
                caminho_novo = pasta_662 / novo_nome
                nome_antigo_662.append(nome)
                try:
                    os.rename(caminho_antigo, caminho_novo)
                    print(f"[✅] Renomeado: {nome} -> {novo_nome}\n")
                    log(f"Arquivo renomeado - Pasta 662: {nome} -> {novo_nome}\n")
                except Exception as e:
                    print(f"[❌] Erro ao renomear {nome}: {e}")
                    log(f"Erro ao renomear {nome}: {e}\n")
                continue
        # Arquivos 663
        if arquivos_663 != []:
            print(f"[🔄] Renomeando arquivos (pasta 663):")
            for nome in arquivos_663:
                novo_nome = nome.split('.RET')[0] + f'-{datetime.now():%d%m%y}' + '.RET'
                caminho_antigo = pasta_663 / nome
                caminho_novo = pasta_663 / novo_nome
                nome_antigo_663.append(nome)
                try:
                    os.rename(caminho_antigo, caminho_novo)
                    print(f"[✅] Renomeado: {nome} -> {novo_nome}\n")
                    log(f"Arquivo renomeado - Pasta 663: {nome} -> {novo_nome}\n")
                except Exception as e:
                    print(f"[❌] Erro ao renomear {nome}: {e}")
                    log(f"Erro ao renomear {nome}: {e}\n")
                continue
    else:
        print('Arquivos .RET não encontrados.')

    arquivos_raiz = [f for f in os.listdir(pasta_raiz) if f.endswith('.RET')]
    arquivos_662 = [f for f in os.listdir(pasta_662) if f.endswith('.RET')]
    arquivos_663 = [f for f in os.listdir(pasta_663) if f.endswith('.RET')]

    carteira_661 = arquivos_raiz + arquivos_662
    carteira_663 = arquivos_663
    todos_arquivos = carteira_661 + carteira_663

    # PARÂMETROS CHROME
    # ============================================================================= #
    chromedriver_path = Path(path_chromedriver)
    service = Service(executable_path=chromedriver_path, service_log_path='/dev/null')
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    driver = webdriver.Chrome(service=service, options=chrome_options)
    actions = ActionChains(driver)
    log("Configurações do Chrome definidas e WebDriver iniciado.\n")

    # ACESSO AO PORTAL
    # ============================================================================= #
    driver.get(f'{url_sistema}/app/')
    WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.ID, "username"))).send_keys(user)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "Password"))).send_keys(password)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "btnAcessar"))).click()
    sleep(3)
    try:
        WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.ID, "erp"))).click()
    except:
        WebDriverWait(driver, 120).until(EC.element_to_be_clickable((By.XPATH, "/html/body/main/div[2]/div[1]"))).click()
    log("Login realizado com sucesso.\n")

    # ACESSO AO ERP
    # ============================================================================= #
    WebDriverWait(driver, 120).until(lambda d: len(d.window_handles) > 1)
    log(f'lambda d: len(d.window_handles) > 1\n')
    handles = driver.window_handles
    driver.switch_to.window(handles[1])
    print(f"Acesso em: {driver.title}\n")
    log(f"Acesso em: {driver.title}\n")

    # ITERA SOBRE ARQUIVOS PARA IMPORTAÇÃO
    # ============================================================================= #
    log(f"Arquivos para importação:\n - Carteira 661: {carteira_661}\n - Carteira 663: {carteira_663}\n")

    for arq_ret in todos_arquivos:

        # MÓDULO FINANCEIRO | COBRANÇA → RETORNO
        # ============================================================================= #
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "Financeiro-btnIconEl"))).click()
        WebDriverWait(driver, 30).until(lambda e: e.execute_script("return document.readyState") == "complete")
        log(f'Acesso ao módulo financeiro\n')
        sleep(2)
        log(f'Importando arquivo: {arq_ret}\n')
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "campoDePesquisa-1118-inputEl"))).clear()

        if arq_ret in arquivos_raiz:
            carteira_selecionada = '661'
            print(f'\nAcessando carteira {carteira_selecionada}...')
            print(f'Importando arquivo: {arq_ret}')
            caminho_completo = pasta_raiz / arq_ret
            log(f'elif:\n -arq_ret:{arq_ret}\n -cart_sel:{carteira_selecionada}\n -cam_compl:{caminho_completo}\n')

        elif arq_ret in arquivos_662:
            carteira_selecionada = '661'
            print(f'\nAcessando carteira {carteira_selecionada}...')
            print(f'Importando arquivo: {arq_ret}')
            caminho_completo = pasta_662 / arq_ret
            log(f'elif:\n -arq_ret:{arq_ret}\n -cart_sel:{carteira_selecionada}\n -cam_compl:{caminho_completo}\n')

        elif arq_ret in arquivos_663:
            carteira_selecionada = '663'
            print(f'\nAcessando carteira {carteira_selecionada}...')
            print(f'Importando arquivo: {arq_ret}')
            caminho_completo = pasta_663 / arq_ret
            log(f'elif:\n -arq_ret:{arq_ret}\n -cart_sel:{carteira_selecionada}\n -cam_compl:{caminho_completo}\n')
        
        log(f'for:\n -arq_ret:{arq_ret}\n -cart_sel:{carteira_selecionada}\n -cam_compl:{caminho_completo}\n')
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "campoDePesquisa-1118-inputEl"))).send_keys(carteira_selecionada)
        log(f'send_key\n')
        sleep(2)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@id,'displayfieldATAK') and @type='text']"))).click()
        sleep(2)
        WebDriverWait(driver, 30).until(lambda f: (f.find_element(By.XPATH, "//input[contains(@id,'displayfieldATAK') and @type='text']").get_attribute("value").strip() != ""))
        sleep(2)
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//li[normalize-space(text())='Retorno (arquivo)']"))).click()
        log(f'XPATH: //li[normalize-space(text())="Retorno (arquivo)"]\n')

        # IMPORTAÇÃO DOS ARQUIVOS
        # ============================================================================= #
        sleep(10)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Selecionar arquivo ']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        log(f'XPATH: //span[text()="Selecionar arquivo "]/ancestor::a[contains(@class, "x-btn")]\n')

        input_file = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@type='file' and contains(@class, 'x-form-file-input')]")))
        driver.execute_script("""
            arguments[0].style.display = 'block';
            arguments[0].style.visibility = 'visible';
            arguments[0].style.height = 'auto';
            arguments[0].style.width = 'auto';
            arguments[0].style.opacity = 1;
        """, input_file)
        driver.execute_script("arguments[0].click();", input_file)

        sleep(2)
        try:
            app = Application().connect(title_re="Abrir")
            dlg = app.window(title_re="Abrir")
            dlg["Edit"].type_keys(caminho_completo, with_spaces=True)
            dlg["Abrir"].click()
            log(f'driver.execute_script\n')
        except:
            pyautogui.write(str(caminho_completo))
            pyautogui.press('enter')
            log(f'pyautogui.write\n')

        sleep(2)
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Importar']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        log('XPATH: //span[text()="Importar"]/ancestor::a[contains(@class, "x-btn")]\n')
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        log('XPATH: //span[text()="OK"]/ancestor::a[contains(@class, "x-btn")]\n')
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Próximo']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        log('XPATH: //span[text()="Próximo"]/ancestor::a[contains(@class, "x-btn")]\n')

        # BAIXA RELATÓRIO DAS CRÍTICAS
        # ============================================================================= #
        sleep(20)
        try:
            try:
                WebDriverWait(driver, 60).until(  
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[2]/span/div/div/div[2]/div/div/div/div[2]/div/table/tbody/tr/td/a"))
                ).click()
                log('criticas XPATH\n')
            except:
                WebDriverWait(driver, 60).until(  
                    EC.element_to_be_clickable((By.XPATH, "//span[text()='Lançamentos criticas']/ancestor::a[contains(@class, 'x-btn')]"))
                ).click()
                log('criticas text()\n')
                
            WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) > 2)
            sleep(10)
            response = requests.get(f"{url_sistema}//CORE/relatorios/ObterStreamArquivo/?RWREC038.pdf.txt")
            log('requests criticas\n')
            pasta_relatorios = (f"{path_servidor}\\Relatórios Arquivos de Retorno\\criticas - {arq_ret} - {carteira_selecionada}.pdf")

            os.path.join(pasta_relatorios, arq_ret)
            with open(pasta_relatorios, 'wb') as wb:
                wb.write(response.content)
            print(f'Arquivo criticas - {arq_ret} - {carteira_selecionada}.pdf Salvo!')
            driver.switch_to.window(driver.window_handles[-1])
            driver.close()
        except:
            sleep(2)
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']/ancestor::a[contains(@class, 'x-btn')]"))).click()
            log('criticas except\n')
                        
        # BAIXA RELATÓRIO DE CONFERÊNCIA
        # ============================================================================= #
        driver.switch_to.window(driver.window_handles[1])
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Próximo']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        sleep(10)
        try:
            try:
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div[2]/span/div/div/div[2]/div/div/div/div[1]/div/table/tbody/tr/td[3]/a/span/span/span[2]"))).click()
                log('conferência XPATH\n')
            except:
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Conferência da Importação']/ancestor::a[contains(@class, 'x-btn')]"))).click()
                log('conferência text()\n')

            WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) > 2)
            sleep(10)
            response = requests.get(f"{url_sistema}//CORE/relatorios/ObterStreamArquivo/?RWREC038.pdf.txt")
            log('requests conferência\n')
            pasta_relatorios = (f"{path_servidor}\\Relatórios Arquivos de Retorno\\conferência - {arq_ret} - {carteira_selecionada}.pdf")

            os.path.join(pasta_relatorios, arq_ret)
            with open(pasta_relatorios, 'wb') as wb:
                wb.write(response.content)
            print(f'Arquivo conferência - {arq_ret} - {carteira_selecionada}.pdf Salvo!')
            driver.switch_to.window(driver.window_handles[-1])
            driver.close()

        except:
            log('conferência except\n')
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']/ancestor::a[contains(@class, 'x-btn')]"))).click()

        # BAIXA COMPROVANTE
        # ============================================================================= #
        driver.switch_to.window(driver.window_handles[1])
        sleep(10)
                        
        try:
            try:
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Concluir']/ancestor::a[contains(@class, 'x-btn')]"))).click()
                log('concluir XPATH\n')
            except:
                WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "buttonGMF-1527-btnInnerEl"))).click()    
                log('concluir ID\n')

            WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']/ancestor::a[contains(@class, 'x-btn')]"))).click()

            # Tenta gerar comprovante, caso não tenha, retorna ao menu de selecionar arquivo
            sleep(10)
            try:
                WebDriverWait(driver, 60).until( EC.element_to_be_clickable((By.XPATH, "//span[text()='Imprimir']/ancestor::a[contains(@class, 'x-btn')]"))).click()
                log('comprovante imprimir\n')
                WebDriverWait(driver, 60).until( EC.element_to_be_clickable((By.XPATH, "//span[text()='Gerar PDF']/ancestor::a[contains(@class, 'x-btn')]"))).click()
                log('comprovante gerar PDF\n')

                # Altera para o relatório aberto em PDF - comprovante
                WebDriverWait(driver, 60).until(lambda d: len(d.window_handles) > 2)
                sleep(10)
                driver.switch_to.window(driver.window_handles[-1])
                response = requests.get(f"{url_sistema}//CORE/relatorios/ObterStreamArquivo/?.txt")
                log('requests comprovante\n')
                pasta_relatorios = (f"{path_servidor}\\Relatórios Arquivos de Retorno\\comprovante - {arq_ret} - {carteira_selecionada}.pdf")

                # Baixa PDF
                os.path.join(pasta_relatorios, arq_ret)
                with open(pasta_relatorios, 'wb') as wb:
                    wb.write(response.content)
                print(f'Arquivo comprovante - {arq_ret} - {carteira_selecionada}.pdf Salvo!')
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()

                # Retorna ao início da importação
                sleep(1)
                driver.switch_to.window(driver.window_handles[1])
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Novo retorno remessa']/ancestor::a[contains(@class, 'x-btn')]"))).click()

            except:
                log('conferência except\n')
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Anterior']/ancestor::a[contains(@class, 'x-btn')]"))).click()
                sleep(2)
                WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Anterior']/ancestor::a[contains(@class, 'x-btn')]"))).click()
        except:
            log('concluir except\n')
            WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OK']/ancestor::a[contains(@class, 'x-btn')]"))).click()

        sleep(2)
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Fechar Aba']/ancestor::a[contains(@class, 'x-btn')]"))).click()

        # MOVE ARQUIVOS PARA AS PASTAS
        # ============================================================================= #
        if arq_ret in arquivos_raiz:
            pasta_destino_raiz = pasta_raiz / '2026'
            shutil.move(pasta_raiz/arq_ret, pasta_destino_raiz/arq_ret)
            print(f'[📂] Arquivo movido: {arq_ret} para a pasta {pasta_destino_raiz}')
            log(f'Shutil - pasta_destino_raiz: {pasta_destino_raiz} - pasta_raiz: {pasta_raiz} - arq_ret: {arq_ret}\n')

        if arq_ret in arquivos_662:
            pasta_destino_662 = pasta_662 / '2026'
            shutil.move(pasta_662/arq_ret, pasta_destino_662/arq_ret)
            print(f'[📂] Arquivo movido: {arq_ret} para a pasta {pasta_destino_662}')
            log(f'Shutil - pasta_destino_662: {pasta_destino_662} - pasta_662: {pasta_662} - arq_ret: {arq_ret}\n')

        if arq_ret in arquivos_663:
            pasta_destino_663 = pasta_663 / 'COBRANÇA'
            shutil.move(pasta_663/arq_ret, pasta_destino_663/arq_ret)
            print(f'[📂] Arquivo movido: {arq_ret} para a pasta {pasta_destino_663}')
            log(f'Shutil - pasta_destino_663: {pasta_destino_663} - pasta_663: {pasta_663} - arq_ret: {arq_ret}\n')

    sleep(2)
    driver.quit()

    # ENVIA E-MAIL DE SUCESSO
    # ============================================================================= #
    print('\nEnviando e-mail de Sucesso')
    TO_EMAILS = [
        email.strip()
        for email in os.getenv("TO_EMAILS", "").split(",")
        if email.strip()
    ]

    CC_EMAILS = [
        email.strip()
        for email in os.getenv("CC_EMAILS", "").split(",")
        if email.strip()
    ]
    Pasta_PDFS = Path(f"{path_servidor}\\Relatórios Arquivos de Retorno")
    assunto = f"Cobranças baixadas {datetime.now():%d/%m/%Y}"
    corpo = f"Cobranças baixadas em {datetime.now():%d/%m/%Y - %H:%M:%S} com Sucesso!"
                    
    enviar_email_outlook(assunto, corpo)

    # PRINTS FINAIS
    # ============================================================================= #
    # Salva log de Sucesso
    print('Salvando Log de Sucesso...')
    with open(f"{pasta_log}\\log_sucesso_{datetime.now():%d%m%y}.txt", "w") as file:
        file.writelines(rel_log)

    print('\n✅ Processo Finalizado!')
    sleep(5)
    print('\nDesenvolvido por RPA Palmali!')
    sleep(5)

# ROLLBACK
# ============================================================================= #
except:
    print('Erro!\nRenomeando arquivos para os originais')
    try:
        for renomear in todos_arquivos:
            if renomear in arquivos_raiz:
                caminho_completo = pasta_destino_raiz / renomear
                log(f"Rollback - Pasta Raiz: {caminho_completo} -> {nome_antigo_raiz[0]}\n")
                caminho_antigo = os.path.join(pasta_raiz, nome_antigo_raiz[0])
                os.rename(caminho_completo, caminho_antigo)
                print(f"[🔁] Renomeado: {renomear} -> {nome_antigo_raiz[0]}")
                shutil.move(pasta_destino_raiz/arq_ret, pasta_raiz/arq_ret)
                log(f'Shutil - pasta_destino_raiz: {pasta_destino_raiz} - pasta_raiz: {pasta_raiz} - arq_ret: {arq_ret}\n\n')
                print(f'[📂] Arquivo retornado: {arq_ret} para a pasta {pasta_raiz}')
            if renomear in arquivos_662:
                caminho_completo = pasta_destino_662 / renomear
                log(f"Rollback - Pasta 662: {caminho_completo} -> {nome_antigo_662[0]}\n")
                caminho_antigo = os.path.join(pasta_662, nome_antigo_662[0])
                os.rename(caminho_completo, caminho_antigo)
                print(f"[🔁] Renomeado: {renomear} -> {nome_antigo_662[0]}")
                shutil.move(pasta_destino_662/arq_ret, pasta_662/arq_ret)
                log(f'Shutil - pasta_destino_662: {pasta_destino_662} - pasta_662: {pasta_662} - arq_ret: {arq_ret}\n\n')
                print(f'[📂] Arquivo retornado: {arq_ret} para a pasta {pasta_662}')
            if renomear in arquivos_663:
                caminho_completo = pasta_destino_663 / renomear
                log(f"Rollback - Pasta 663: {caminho_completo} -> {nome_antigo_663[0]}\n")
                caminho_antigo = os.path.join(pasta_663, nome_antigo_663[0])
                os.rename(caminho_completo, caminho_antigo)
                print(f"[🔁] Renomeado: {renomear} -> {nome_antigo_663[0]}")
                shutil.move(pasta_destino_663/arq_ret, pasta_663/arq_ret)
                log(f'Shutil - pasta_destino_663: {pasta_destino_663} - pasta_663: {pasta_663} - arq_ret: {arq_ret}\n\n')
                print(f'[📂] Arquivo retornado: {arq_ret} para a pasta {pasta_663}')
    except:
        print("Erro ao renomear os arquivos")

    # Salva log de erro
    print('Salvando Log de Erro...')
    with open(f"{pasta_log}\\log_erro_{datetime.now():%d%m%y}.txt", "w") as file:
        file.writelines(rel_log)

sleep(10)