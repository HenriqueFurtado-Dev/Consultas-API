import sys
import asyncio

# Se estiver em Windows, force a policy (pode tentar remover se não ajudar)
if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

import os
import io
import time
import datetime
import logging

import pandas as pd
import openpyxl

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

from fastapi.concurrency import run_in_threadpool

from dotenv import load_dotenv

# Selenium e webdriver_manager
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from playwright.sync_api import sync_playwright



# Em vez de async_playwright, usaremos sync_playwright para ESSOR
from playwright.sync_api import sync_playwright

# Configuração do logging
logging.basicConfig(level=logging.INFO)
load_dotenv()

# Credenciais
USERNAME_AXA = os.getenv('USUARIO_AXA')
PASSWORD_AXA = os.getenv('PASSWORD_AXA')

USERNAME_ESSOR = os.getenv('USUARIO_ESSOR')
PASSWORD_ESSOR = os.getenv('PASSWORD_ESSOR')

app = FastAPI()

origins = ["http://localhost", "http://localhost:5173", "http://localhost:3000"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def hello_root():
    return {"message": "Hello World"}


def configurar_navegador_selenium(headless: bool = True):
    chrome_options = webdriver.ChromeOptions()
    if headless:
       chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--window-size=1920,1080")

    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=chrome_options)
    return driver

def login_axa(driver):
    try:
        driver.get('https://portaldocorretor.axa.com.br/login/admin')
        wait = WebDriverWait(driver, 20)

        # Aguarda 3 segundos
        time.sleep(3)

        campo_login = wait.until(EC.presence_of_element_located((By.NAME, 'login')))
        campo_login.clear()
        campo_login.send_keys(USERNAME_AXA)

        time.sleep(3)

        campo_senha = wait.until(EC.presence_of_element_located((By.NAME, 'pwd')))
        campo_senha.clear()
        campo_senha.send_keys(PASSWORD_AXA)

        botao_submit = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
        botao_submit.click()

        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.info-pessoa")))
        logging.info("Login realizado com sucesso na AXA.")
    except TimeoutException:
        raise Exception("Login não foi realizado com sucesso na AXA. Verifique credenciais ou fluxo de login.")

def consultar_dados_axa(driver, df_axa: pd.DataFrame) -> pd.DataFrame:
    driver.get('https://e-solutions.axa.com.br/#!/lista-parcelas')
    logging.info("Navegou para a página de boletos na AXA.")

    wait = WebDriverWait(driver, 20)
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Apolices e Endossos')]")))
        logging.info("Página de 'Apolices e Endossos' carregada com sucesso na AXA.")
    except TimeoutException:
        logging.warning("A página de 'Apolices e Endossos' pode não ter carregado na AXA.")

    time.sleep(3)

    # Inserir datas
    try:
        campo_dt_ini = wait.until(EC.presence_of_element_located((By.ID, 'dt_ini')))
        campo_dt_ini.clear()
        campo_dt_ini.send_keys('01/01/2024')
        logging.info("Data inicial '01/01/2024' inserida no campo dt_ini.")
    except Exception as e:
        raise Exception(f"Erro ao inserir data inicial na AXA: {e}")

    try:
        data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
        campo_dt_ter = wait.until(EC.presence_of_element_located((By.ID, 'dt_ter')))
        campo_dt_ter.clear()
        campo_dt_ter.send_keys(data_atual)
        logging.info(f"Data atual '{data_atual}' inserida no campo dt_ter.")
    except Exception as e:
        raise Exception(f"Erro ao inserir data final na AXA: {e}")

    time.sleep(2)

    resultados = []

    for index, row in df_axa.iterrows():
        cnpj = str(row['CPF/CNPJ']).strip().zfill(14)
        logging.info(f"Processando cliente AXA {index + 1} — CNPJ = {cnpj}")

        try:
            campo_cnpj = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[ng-model='filtros.cpfCnpjEstipulante']")))
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)

            botao_filtrar = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class, 'button') and contains(@class, 'custom-icon') and contains(@class, 'ghost-blue')]/span[text()='Filtrar']"))
            )
            botao_filtrar.click()
            logging.info(f"Botão 'Filtrar' clicado para CNPJ '{cnpj}'.")

            time.sleep(5)
            try:
                tabela = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.tb-parcelas")))
                linhas = tabela.find_elements(By.XPATH, ".//tbody/tr")
                if len(linhas) > 0:
                    for linha in linhas:
                        celulas = linha.find_elements(By.TAG_NAME, 'td')
                        if len(celulas) >= 7:
                            resultados.append({
                                'CNPJ': cnpj,
                                'Vencimento': celulas[0].text.strip(),
                                'Apólice/Endosso': celulas[1].text.strip(),
                                'Segurado': celulas[2].text.strip(),
                                'Parcela': celulas[4].text.strip(),
                                'Valor do Prêmio': celulas[6].text.strip(),
                            })
                else:
                    logging.info(f"Sem linhas na tabela para {cnpj}.")
            except TimeoutException:
                logging.info(f"Nenhuma parcela encontrada para {cnpj}.")
        except Exception as e:
            logging.error(f"Erro ao processar CNPJ {cnpj}: {e}")
        finally:
            try:
                campo_cnpj.clear()
            except:
                pass

    if resultados:
        return pd.DataFrame(resultados)
    else:
        return pd.DataFrame([])
    

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from playwright.async_api import async_playwright


def consultar_dados_essor_sync(df_essor: pd.DataFrame) -> pd.DataFrame:
    resultados_essor = []

    # Se não houver registros para ESSOR, retorne um DataFrame vazio
    if len(df_essor) == 0:
        logging.info("DataFrame ESSOR vazio - retornando DataFrame vazio")
        return pd.DataFrame([])

    with sync_playwright() as p:
        # Lança o browser (chromium)
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            # 1) Ir até a página de login
            page.goto('https://portal.essor.com.br/')
            logging.info("Página de login ESSOR acessada.")

            # 2) Fazer login
            page.fill('input[name="Login"]', USERNAME_ESSOR)
            page.fill('input[name="Senha"]', PASSWORD_ESSOR)
            page.click('button[type="submit"]')
            page.wait_for_load_state('networkidle')
            logging.info("Login realizado no ESSOR.")

            # 3) Clicar em 'Consultas' e em seguida 'Parcelas Pendentes'
            page.click("text=Consultas")
            page.click("text=Parcelas Pendentes")
            page.wait_for_load_state('networkidle')
            logging.info("Página 'Parcelas Pendentes' no ESSOR carregada.")

            # 4) Verificar se o campo #nr_apolice (ou #nr_cnpj) está na página principal ou em um iframe
            search_context = None

            try:
                # Tenta achar o seletor na página principal
                page.wait_for_selector('#nr_apolice', timeout=10000)
                logging.info("Campo '#nr_apolice' encontrado na página principal.")
                search_context = page
            except:
                # Se não estiver na página principal, procurar nos frames
                logging.warning("Não foi possível encontrar '#nr_apolice' na página principal. Tentando iframes...")

                for frame in page.frames:
                    try:
                        frame.wait_for_selector('#nr_apolice', timeout=5000)
                        search_context = frame
                        logging.info("Campo '#nr_apolice' encontrado dentro de um iframe.")
                        break
                    except:
                        continue

            # Se mesmo assim não achar, aborta
            if not search_context:
                logging.error("Não foi possível encontrar o campo '#nr_apolice' em nenhum contexto.")
                browser.close()
                return pd.DataFrame([])

            # 5) Iterar sobre cada linha do df_essor, pegando o valor que deseja pesquisar
            for index, row in df_essor.iterrows():

                # Se for CNPJ, troque row['CNPJ'] ou row['CPF/CNPJ']. 
                # Se for Apólice, mantenha row['Apólice'].  
                # (No seu script “funcional”, você usava row['Apólice']. Ajuste conforme a real necessidade.)
                apolice = str(row['Apólice']).strip()
                logging.info(f"Consultando apólice ESSOR: {apolice}")

                # 5.1) Preencher o campo (via JavaScript)
                try:
                    search_context.evaluate(f'''() => {{
                        document.getElementById('nr_apolice').value = "{apolice}";
                    }}''')
                except Exception as e:
                    logging.error(f"Erro ao preencher o campo '#nr_apolice' para a apólice '{apolice}': {e}")
                    continue

                # 5.2) Clicar no botão 'Pesquisar'
                try:
                    search_context.evaluate('''() => {
                        document.getElementById('btnPesquisar').click();
                    }''')
                except Exception as e:
                    logging.error(f"Erro ao clicar no botão 'Pesquisar' para apólice '{apolice}': {e}")
                    continue

                # 5.3) Aguardar alguns segundos para que a tabela seja carregada
                search_context.wait_for_timeout(3000)

                # 5.4) Verificar se a tabela existe
                has_table = False
                try:
                    has_table = search_context.evaluate('''() => {
                        return document.getElementById('dataTableParcelas') !== null;
                    }''')
                except Exception as e:
                    logging.error(f"Erro ao verificar a existência da tabela para '{apolice}': {e}")

                # 5.5) Se existe, extrair dados
                if has_table:
                    table_data = []
                    try:
                        table_data = search_context.evaluate('''() => {
                            const data = [];
                            const table = document.getElementById('dataTableParcelas');
                            if(!table) return data;

                            const rows = table.querySelectorAll('tbody tr');
                            rows.forEach(row => {
                                const cells = Array.from(row.children).map(td => td.innerText.trim());
                                data.push(cells);
                            });
                            return data;
                        }''')
                    except Exception as e:
                        logging.error(f"Erro ao extrair dados da tabela para '{apolice}': {e}")

                    # Verifica se há alguma linha significativa (evitando 'Nenhum registro encontrado')
                    if len(table_data) == 1 and "Nenhum registro encontrado" in table_data[0][0]:
                        logging.info(f"Nenhuma pendência para apólice {apolice}")
                        # Aqui, se quiser, já pode marcar 'SEM PENDENCIA' em algum DF
                    else:
                        # Adiciona linhas ao `resultados_essor`
                        for data_row in table_data:
                            if len(data_row) >= 8:
                                resultados_essor.append({
                                    'Apólice': apolice,
                                    'Corretor Líder': data_row[0],
                                    'Segurado': data_row[1],
                                    'Apólice (2)': data_row[2],
                                    'Endosso': data_row[3],
                                    'Nº Parcela': data_row[4],
                                    'Valor da Parcela': data_row[5],
                                    'Data de vencimento': data_row[6],
                                    'Dias em atraso': data_row[7],
                                })
                            else:
                                logging.warning(f"Dados insuficientes para apólice {apolice}: {data_row}")
                else:
                    logging.info(f"Nenhuma tabela encontrada para a apólice {apolice}.")

                # 5.6) Limpar o campo para a próxima iteração
                try:
                    search_context.evaluate('''() => {
                        document.getElementById('nr_apolice').value = '';
                    }''')
                except:
                    pass

        except Exception as e:
            logging.error(f"Erro durante a consulta ESSOR: {e}")
        finally:
            browser.close()
            logging.info("Navegador ESSOR fechado.")

    # Retorna o DataFrame final
    return pd.DataFrame(resultados_essor) if resultados_essor else pd.DataFrame([])



@app.post("/upload/")
async def processar_planilha(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Arquivo deve ser no formato .xlsx")

    try:
        logging.info("Recebendo arquivo para processamento (AXA + ESSOR).")

        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        logging.info(f"Planilha '{file.filename}' carregada com sucesso.")

        if 'CPF/CNPJ' not in df.columns:
            raise HTTPException(status_code=400, detail="Coluna 'CPF/CNPJ' ausente no arquivo enviado.")
        if 'Seg.' not in df.columns:
            raise HTTPException(status_code=400, detail="Coluna 'Seg.' ausente no arquivo enviado.")

        if 'STATUS' not in df.columns:
            df['STATUS'] = 'NAO VERIFICADO'

        # AXA via Selenium
        df_axa = df[df['Seg.'] == 'AXA'].copy()
        invoices_data_axa = pd.DataFrame()

        if not df_axa.empty:
            driver = configurar_navegador_selenium(headless=True)
            try:
                login_axa(driver)
                invoices_data_axa = consultar_dados_axa(driver, df_axa)
            except Exception as e:
                logging.error(f"Erro durante consulta AXA: {e}")
            finally:
                driver.quit()
        else:
            logging.info("Nenhuma linha para 'AXA' na planilha.")

        # ESSOR via função síncrona + threadpool
        df_essor = df[df['Seg.'] == 'ESSO'].copy()  # ou 'ESSOR'
        invoices_data_essor = pd.DataFrame()

        if not df_essor.empty:
            try:
                # Chamamos a função síncrona via run_in_threadpool
                invoices_data_essor = await run_in_threadpool(consultar_dados_essor_sync, df_essor)
            except Exception as e:
                logging.error(f"Erro durante consulta ESSOR: {e}")
        else:
            logging.info("Nenhuma linha para 'ESSOR' na planilha.")

        # Atualiza STATUS no DF principal
        # 1) AXA
        if not invoices_data_axa.empty:
            for cnpj_pendente in invoices_data_axa['CNPJ'].unique():
                df.loc[
                    (df['Seg.'] == 'AXA') &
                    (df['CPF/CNPJ'].astype(str).str.zfill(14) == cnpj_pendente),
                    'STATUS'
                ] = 'FATURAS-PENDENTES'

            cnpjs_axa_pendentes = invoices_data_axa['CNPJ'].unique().tolist()
            df.loc[
                (df['Seg.'] == 'AXA') &
                (~df['CPF/CNPJ'].astype(str).str.zfill(14).isin(cnpjs_axa_pendentes)),
                'STATUS'
            ] = 'SEM PENDENCIA'
        else:
            df.loc[df['Seg.'] == 'AXA', 'STATUS'] = 'SEM PENDENCIA'

        # 2) ESSOR
        if not invoices_data_essor.empty:
            for apolice_pendente in invoices_data_essor['Apólice'].unique():
                df.loc[
                    (df['Seg.'] == 'ESSO') &
                    (df['Apólice'].astype(str) == apolice_pendente),
                    'STATUS'
                ] = 'FATURAS-PENDENTES'

            apolices_essor_pendentes = invoices_data_essor['Apólice'].unique().tolist()
            df.loc[
                (df['Seg.'] == 'ESSO') &
                (~df['Apólice'].astype(str).isin(apolices_essor_pendentes)),
                'STATUS'
            ] = 'SEM PENDENCIA'
        else:
            df.loc[df['Seg.'] == 'ESSO', 'STATUS'] = 'SEM PENDENCIA'

        # Salva e retorna arquivo
        nome_arquivo_saida = "comissoes_pendentes_corretora_atualizado.xlsx"
        try:
            with pd.ExcelWriter(nome_arquivo_saida, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Status', index=False)

                if not invoices_data_axa.empty:
                    invoices_data_axa.to_excel(writer, sheet_name='AXA Faturas Pendentes', index=False)
                else:
                    pd.DataFrame(
                        {'Mensagem': ['Não há faturas pendentes para a AXA.']}
                    ).to_excel(writer, sheet_name='AXA Faturas Pendentes', index=False)

                if not invoices_data_essor.empty:
                    invoices_data_essor.to_excel(writer, sheet_name='ESSOR Faturas Pendentes', index=False)
                else:
                    pd.DataFrame(
                        {'Mensagem': ['Não há faturas pendentes para a ESSOR.']}
                    ).to_excel(writer, sheet_name='ESSOR Faturas Pendentes', index=False)

            if not os.path.exists(nome_arquivo_saida):
                raise HTTPException(status_code=500, detail="Erro ao gerar o arquivo de resposta.")

            return FileResponse(
                path=nome_arquivo_saida,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=nome_arquivo_saida
            )
        except Exception as e:
            logging.error(f"Erro ao salvar a planilha: {e}")
            raise HTTPException(status_code=500, detail=f"Erro ao salvar a planilha: {e}")

    except Exception as e:
        logging.error(f"Erro durante o processamento: {e}")
        raise HTTPException(status_code=500, detail=f"Erro no processamento: {e}")