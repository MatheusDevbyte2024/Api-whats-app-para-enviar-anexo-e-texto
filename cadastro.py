import os
import sys
import time
import logging
import pandas as pd
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote

# CONFIGURAÇÕES
EXCEL_PATH = r'C:\OFICIO\OFICIO NOVO KSK.xlsx'
CHROMEDRIVER_PATH = r'C:\OFICIO\chromedriver.exe'
CAMINHO_PDF = r''  # caminho completo com nome e extensão

logging.basicConfig(
    filename='log_envio.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def validar_telefone(telefone: str) -> bool:
    tel_str = telefone.strip().replace("+", "").replace(" ", "").replace("-", "")
    return tel_str.isdigit() and 10 <= len(tel_str) <= 13

def abrir_whatsapp_web(driver, timeout=120):
    driver.get('https://web.whatsapp.com')
    print("[INFO] Escaneie o QR Code do WhatsApp Web...")
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, 'pane-side'))
        )
        print("[INFO] WhatsApp Web pronto.")
    except:
        print("[ERRO] Tempo excedido para escanear QR Code.")
        driver.quit()
        sys.exit(1)

def verificar_numero_valido(driver) -> bool:
    wait = WebDriverWait(driver, 5)
    erros_xpath = [
        '//div[contains(text(), "phone number shared via url is invalid")]',
        '//div[contains(text(), "não está no WhatsApp")]',
        '//span[contains(text(), "não está no WhatsApp")]',
        '//div[contains(@class,"_2Nr6U")]//span[contains(text(),"phone number")]',
    ]
    for xpath in erros_xpath:
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            return False
        except:
            continue
    return True

def enviar_mensagem(driver, telefone: str, mensagem: str) -> bool:
    texto_codificado = quote(mensagem)
    url = f"https://web.whatsapp.com/send?phone={telefone}&text={texto_codificado}"
    driver.get(url)
    time.sleep(5)

    if not verificar_numero_valido(driver):
        logging.warning(f"Número inválido: {telefone}")
        print(f"[AVISO] Número inválido: {telefone}")
        return False

    try:
        wait = WebDriverWait(driver, 15)
        caixa = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[3]/div[1]/p')))
        caixa.send_keys(Keys.ENTER)
        logging.info(f"Mensagem enviada: {telefone}")
        print(f"[INFO] Mensagem enviada: {telefone}")
        time.sleep(10)
        return True
    except Exception as e:
        logging.error(f"Erro ao enviar mensagem: {e}")
        print(f"[ERRO] Falha ao enviar mensagem: {e}")
        return False

def controlar_janela_arquivo_pyautogui(caminho_arquivo):
    time.sleep(2)  # espera a janela abrir
    pyautogui.write(caminho_arquivo)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    print("[INFO] Arquivo selecionado e enviado via pyautogui.")

def enviar_arquivo(driver, caminho_arquivo: str) -> bool:
    wait = WebDriverWait(driver, 20)
    try:
        botao_anexar = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[1]')))
        botao_anexar.click()

        botao_arquivo = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="app"]/div/span[6]/div/ul/div/div/div[1]/li/div/span')))
        botao_arquivo.click()

        controlar_janela_arquivo_pyautogui(caminho_arquivo)

        time.sleep(15)

        botao_enviar = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'span[data-icon="send"]')))
        botao_enviar.click()

        logging.info(f"Arquivo enviado: {caminho_arquivo}")
        print(f"[INFO] Arquivo enviado: {caminho_arquivo}")

        caixa_mensagem = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[3]/div[1]/p')))
        caixa_mensagem.click()
        time.sleep(2)

        return True

    except Exception as e:
        logging.error(f"Erro ao enviar arquivo: {e}")
        print(f"[ERRO] Falha ao enviar arquivo: {e}")
        return False

def carregar_dados_excel(caminho_arquivo: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(caminho_arquivo)
    except Exception as e:
        print(f"[ERRO] Falha ao ler Excel: {e}")
        sys.exit(1)

    df.columns = df.columns.str.strip().str.upper()
    colunas_esperadas = {'NOME', 'TELEFONE', 'MENSAGEM'}
    if not colunas_esperadas.issubset(set(df.columns)):
        print(f"[ERRO] Excel deve conter colunas: {colunas_esperadas}")
        sys.exit(1)

    return df

def main():
    df = carregar_dados_excel(EXCEL_PATH)

    # Inicializa a coluna STATUS, se ainda não existir
    if 'STATUS' not in df.columns:
        df['STATUS'] = ''

    service = Service(CHROMEDRIVER_PATH)
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)

    try:
        abrir_whatsapp_web(driver)

        for index, row in df.iterrows():
            nome = str(row['NOME']).strip()
            telefone = str(row['TELEFONE']).strip().replace(" ", "").replace("-", "")
            mensagem = str(row['MENSAGEM']).strip()

            if not validar_telefone(telefone):
                print(f"[AVISO] Telefone inválido: {telefone}")
                df.at[index, 'STATUS'] = 'FALHA - Telefone inválido'
                continue

            telefone_formatado = '+55' + telefone

            enviado = enviar_mensagem(driver, telefone_formatado, mensagem)
            if enviado:
                arquivo_enviado = enviar_arquivo(driver, CAMINHO_PDF)
                if arquivo_enviado:
                    df.at[index, 'STATUS'] = 'ENVIADO'
                else:
                    df.at[index, 'STATUS'] = 'FALHA - Erro ao enviar arquivo'
            else:
                df.at[index, 'STATUS'] = 'FALHA - Erro ao enviar mensagem'

        # Salvar nova planilha com o status de envio
        resultado_path = os.path.splitext(EXCEL_PATH)[0] + '_resultado.xlsx'
        df.to_excel(resultado_path, index=False)
        print(f"[INFO] Processo finalizado. Resultado salvo em: {resultado_path}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
