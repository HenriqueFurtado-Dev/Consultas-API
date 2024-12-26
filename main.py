import os
import io
import openpyxl
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

import logging
from dotenv import load_dotenv

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    ElementNotInteractableException,
)

load_dotenv()

# Configuração do logging
logging.basicConfig(level=logging.INFO)

# Credenciais
USERNAME_AXA = os.getenv('USUARIO_AXA')
PASSWORD_AXA = os.getenv('PASSWORD_AXA')

app = FastAPI()

# Configuração de CORS
origins = [
    "http://localhost",
    "http://localhost:5173",  # Porta usada pelo Vite
    "http://localhost:3000",  # Porta padrão do React (se estiver usando)
    # Adicione outros domínios conforme necessário
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,  # Permitir apenas esses domínios
    allow_credentials=True,
    allow_methods=["*"],  # Permitir todos os métodos (GET, POST, etc.)
    allow_headers=["*"],  # Permitir todos os headers
)

@app.get("/")
def hello_root():
    return {
        "message": "Hello World"
    }

def configurar_navegador():
    # Configuração do ChromeDriver no modo headless
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless")  # Ativa o modo headless
    chrome_options.add_argument("--disable-gpu")  # Desativa GPU (compatibilidade em servidores)
    chrome_options.add_argument("--no-sandbox")  # Desativa o sandboxing para evitar problemas de permissão
    chrome_options.add_argument("--disable-dev-shm-usage")  # Usa /dev/shm para evitar erros de espaço em disco
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920,1080")  # Define o tamanho da janela no modo headless
    
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

# Função para realizar login na AXA
def login_axa(driver):
    try:
        driver.get('https://portaldocorretor.axa.com.br/login/admin')
        wait = WebDriverWait(driver, 10)

        # Preenchimento do login
        wait.until(EC.presence_of_element_located((By.NAME, 'login'))).send_keys(USERNAME_AXA)
        logging.info("Campo de login preenchido.")
        wait.until(EC.presence_of_element_located((By.NAME, 'pwd'))).send_keys(PASSWORD_AXA)
        logging.info("Campo de senha preenchido.")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))).click()
        logging.info("Botão de submit clicado.")

        # Verificar se o login foi bem-sucedido
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.info-pessoa")))
        logging.info("Login realizado com sucesso.")
    except TimeoutException:
        raise Exception("Falha no login. Verifique as credenciais ou a estrutura da página.")

# Função para processar CNPJs e retornar dados
def consultar_dados_axa(driver, df):
    resultados = []
    wait = WebDriverWait(driver, 10)

    for _, row in df.iterrows():
        cnpj = str(row['CPF/CNPJ']).zfill(14)
        logging.info(f"Processando CNPJ: {cnpj}")

        try:
            campo_cnpj = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[ng-model='filtros.cpfCnpjEstipulante']")))
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)
            logging.info(f"CNPJ '{cnpj}' inserido no campo 'CPF/CNPJ' na AXA.")

            # Botão de filtrar
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button/span[text()='Filtrar']"))).click()
            logging.info(f"Botão 'Filtrar' clicado para CNPJ '{cnpj}' na AXA.")

            # Espera a tabela carregar
            time.sleep(2)  # Pode ajustar ou usar espera explícita

            # Extrair informações
            table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.tb-parcelas")))
            rows_table = table.find_elements(By.XPATH, ".//tbody/tr")

            if len(rows_table) > 0:
                logging.info(f"Encontradas {len(rows_table)} faturas pendentes para CNPJ '{cnpj}'.")
                for row_table in rows_table:
                    cells = row_table.find_elements(By.TAG_NAME, 'td')
                    if len(cells) >= 7:
                        resultados.append({
                            'CNPJ': cnpj,
                            'Vencimento': cells[0].text.strip(),
                            'Apólice/Endosso': cells[1].text.strip(),
                            'Segurado': cells[2].text.strip(),
                            'Parcela': cells[4].text.strip(),
                            'Valor do Prêmio': cells[6].text.strip(),
                        })
            else:
                logging.info(f"Nenhuma fatura pendente encontrada para CNPJ '{cnpj}'.")
        except Exception as e:
            logging.error(f"Erro ao processar CNPJ {cnpj}: {e}")

    # Define as colunas para garantir que existam mesmo que vazio
    colunas = ['CNPJ', 'Vencimento', 'Apólice/Endosso', 'Segurado', 'Parcela', 'Valor do Prêmio']
    return pd.DataFrame(resultados, columns=colunas)

@app.post("/upload/")
async def processar_planilha(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Arquivo deve ser no formato .xlsx")

    try:
        logging.info("Recebendo arquivo para processamento.")
        # Carregar planilha em DataFrame
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        logging.info(f"Planilha '{file.filename}' carregada com sucesso.")
        if 'CPF/CNPJ' not in df.columns:
            raise HTTPException(status_code=400, detail="Coluna 'CPF/CNPJ' ausente no arquivo enviado.")
        
        # Configurar navegador
        logging.info("Configurando navegador Selenium.")
        driver = configurar_navegador()
        logging.info("Navegador configurado. Realizando login.")
        login_axa(driver)

        # Processar dados da AXA
        logging.info("Iniciando consulta de dados da AXA.")
        invoices_data = consultar_dados_axa(driver, df)
        logging.info("Consulta de dados da AXA concluída.")
        driver.quit()

        # Atualizar o DataFrame original com o status
        logging.info("Atualizando o DataFrame com os resultados.")
        if not invoices_data.empty and 'CNPJ' in invoices_data.columns:
            df['STATUS'] = df.apply(
                lambda row: 'FATURAS-PENDENTES' if invoices_data['CNPJ'].str.contains(row['CPF/CNPJ'].zfill(14)).any() else 'SEM PENDENCIA',
                axis=1
            )
            logging.info("Coluna 'STATUS' atualizada com 'FATURAS-PENDENTES' ou 'SEM PENDENCIA'.")
        else:
            logging.info("Nenhuma fatura pendente encontrada para todos os CNPJs.")
            df['STATUS'] = 'SEM PENDENCIA'

        # Salvar resultados em uma planilha
        file_path = "comissoes_pendentes_corretora_atualizado.xlsx"
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Salva a aba 'Status'
            df.to_excel(writer, sheet_name='Status', index=False)
            logging.info("Aba 'Status' salva com sucesso.")

            # Salva as faturas pendentes da AXA, se existirem
            if not invoices_data.empty:
                invoices_data.to_excel(writer, sheet_name='AXA Faturas Pendentes', index=False)
                logging.info("Aba 'AXA Faturas Pendentes' salva com as faturas pendentes da AXA.")
            else:
                # Cria uma aba com uma mensagem padrão
                df_placeholder = pd.DataFrame({'Mensagem': ['Não há faturas pendentes para a AXA.']})
                df_placeholder.to_excel(writer, sheet_name='AXA Faturas Pendentes', index=False)
                logging.info("Aba 'AXA Faturas Pendentes' criada com uma mensagem padrão.")

        # Verifique se o arquivo foi criado
        if not os.path.exists(file_path):
            logging.error(f"Arquivo gerado não encontrado: {file_path}")
            raise HTTPException(status_code=500, detail="Erro ao gerar o arquivo de resposta.")
        
        logging.info(f"Arquivo gerado com sucesso: {file_path}")

        # Retorna a planilha gerada como resposta
        return FileResponse(
            path=file_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="comissoes_pendentes_corretora_atualizado.xlsx"
        )
    except Exception as e:
        logging.error(f"Erro durante o processamento: {e}")
        raise HTTPException(status_code=500, detail=f"Erro no processamento: {e}")
