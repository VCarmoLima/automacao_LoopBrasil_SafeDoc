import pandas as pd
import os
import time
import json
import base64
import datetime
import logging
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

load_dotenv()

logging.basicConfig(
    filename='automacao.log',
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- Configurações de Arquivos ---
NOME_ARQUIVO_EXCEL = "Base_TESTE.xlsm"
NOME_ABA = "Calculos"
PASTA_DOWNLOADS = os.getenv("PASTA_DOWNLOADS")

# --- Arquivo de Log ---
NOME_ARQUIVO_LOG = "log_processados.xlsx"

# --- Nomes das Colunas ---
COLUNA_PLACA = "Placa"
COLUNA_END1 = "Endereço transportadora"
COLUNA_END2 = "Endereço Pátio"
COLUNA_END3 = "Cidade convertida"
COLUNA_STATUS = "SAFE DOC PRINT"
COLUNA_TESTE = "Teste"
COLUNA_STATUS_SAFE_DOC = "STATUS SAFE DOC"
COLUNA_CONTRATO = "Contrato"
COLUNA_CATEGORIA = "Categoria"

# --- Configurações do Portal ---
URL_BANCO = os.getenv("URL_BANCO")
USUARIO_BANCO = os.getenv("USUARIO_BANCO")
SENHA_BANCO = os.getenv("SENHA_BANCO")

VALOR_RANGES = {
    "leve": [
        (200, 241),
        (500, 468),
        (700, 620),
        (1000, 900),
        (9999, 1320)
    ],
    "moto": [
        (200, 230),
        (500, 438),
        (700, 580),
        (1000, 795),
        (9999, 880)
    ],
    "pesado": [
        (200, 665),
        (500, 1045),
        (700, 2020),
        (1000, 3235),
        (9999, 4175)
    ]
}

SELECTORS = {
    "google_maps": {
        "km_xpath": "/html/body/div[1]/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[5]/div[1]/div[1]/div/div[1]/div[2]/div"
    },
    "login": {
        "usuario": "/html/body/div/main/div/div[2]/form/div/div[2]/div/input",
        "senha": "/html/body/div/main/div/div[2]/form/div/div[3]/div/input",
        "botao": "/html/body/div/main/div/div[2]/form/div/div[4]/input[1]"
    },
    "gca_menu": {
        "link_1": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[1]/span[1]/img",
        "link_2": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[2]/div[3]/span[1]/img[2]",
        "link_3": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[2]/div[4]/div[2]/img"
    },
    "iframes": {
        "externo": "ifrmForm",
        "interno": "ifrmObject"
    },
    "form_upload": {
        "input_arquivo": "_ctl0_ContentPlaceHolder_IDX_FILE",
        "select_status": "/html/body/form/div/div/div/div/div[2]/div[3]/select",
        "input_data": "/html/body/form/div/div/div/div/div[2]/div[4]/div/input",
        "input_contrato": "/html/body/form/div/div/div/div/div[2]/div[5]/input",
        "input_placa": "/html/body/form/div/div/div/div/div[2]/div[6]/input",
        "select_tipo_despesa": "/html/body/form/div/div/div/div/div[2]/div[7]/select",
        "input_valor": "/html/body/form/div/div/div/div/div[2]/div[8]/input",
        "input_caixa_arquivo": "/html/body/form/div/div/div/div/div[2]/div[9]/input",
        "input_observacao": "/html/body/form/div/div/div/div/div[2]/div[10]/input",
        "botao_salvar": "/html/body/form/div/div/div/div/div[3]/input"
    }
}

def configurar_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    app_state = {
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    chrome_options.add_experimental_option("prefs", {
        "printing.print_preview_sticky_settings.appState": json.dumps(app_state),
        "savefile.default_directory": PASTA_DOWNLOADS
    })

    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        logging.critical(f"Falha ao iniciar o Selenium: {e}", exc_info=True)
        return None


def extrair_km_do_mapa(driver):
    try:
        xpath_km = SELECTORS["google_maps"]["km_xpath"]
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_km)))
        km_element = driver.find_element(By.XPATH, xpath_km)
        km_bruto_texto = km_element.text
        km_bruto_num_str = km_bruto_texto.split(' ')[0]
        km_sem_milhar = km_bruto_num_str.replace('.', '')
        km_para_float = km_sem_milhar.replace(',', '.')
        km_num = float(km_para_float)
        km_str_arquivo = str(int(km_num))

        return km_num, km_str_arquivo

    except Exception as e:
        logging.warning(f"Não foi possível extrair o KM da página com o XPATH fornecido: {e}")
        return None, "KM_NAO_ENCONTRADO"

def get_valor_por_range(categoria, km_numerico):
    if km_numerico is None:
        return "VALOR_PENDENTE"

    categoria_limpa = categoria.strip().lower()

    ranges_da_categoria = VALOR_RANGES.get(categoria_limpa, [])

    for limite_km, valor in ranges_da_categoria:
        if km_numerico <= limite_km:
            return valor

    logging.warning(f"Valor não encontrado para Categoria: {categoria_limpa}, KM: {km_numerico}")
    return "VALOR_NAO_ENCONTRADO"

def gerar_pdf_mapa(driver, nome_arquivo_pdf):
    try:
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "landscape": False, "printBackground": True, "displayHeaderFooter": True,
            "marginTop": 1, "marginBottom": 1, "marginLeft": 0.5, "marginRight": 0.5
        })
        caminho_completo = os.path.join(PASTA_DOWNLOADS, nome_arquivo_pdf)
        with open(caminho_completo, "wb") as f:
            f.write(base64.b64decode(result['data']))

        return caminho_completo
    except Exception as e:
        logging.error(f"ERRO ao tentar gerar PDF do mapa: {e}", exc_info=True)
        return None


def fazer_login_banco(driver):
    try:
        driver.get(URL_BANCO)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, SELECTORS["login"]["usuario"])))
        driver.find_element(By.XPATH, SELECTORS["login"]["usuario"]).send_keys(USUARIO_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["senha"]).send_keys(SENHA_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["botao"]).click()

        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"]))
        )
        return True
    except Exception as e:
        logging.error(f"--- ERRO NA ETAPA DE LOGIN ---: {e}", exc_info=True)
        return False


def navegar_menu_gca(driver):
    try:
        el1 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"])))
        el1.click()
        el2 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_2"])))
        el2.click()
        el3 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_3"])))
        el3.click()
        return True
    except Exception as e:
        logging.error(f"--- ERRO NA NAVEGAÇÃO GCA ---: {e}", exc_info=True)
        return False


def preencher_formulario_com_upload(driver, dados_upload):
    try:
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, SELECTORS["form_upload"]["select_status"]))
        )

        upload_element = driver.find_element(By.ID, SELECTORS["form_upload"]["input_arquivo"])
        caminho_pdf = dados_upload['caminho_pdf']

        driver.execute_script(
            "arguments[0].style.display = 'block'; " +
            "arguments[0].style.visibility = 'visible'; " +
            "arguments[0].style.opacity = 1; " +
            "arguments[0].style.height = '1px'; " +
            "arguments[0].style.width = '1px';",
            upload_element
        )
        time.sleep(1)

        upload_element.send_keys(caminho_pdf)

        driver.execute_script(
            "validate(arguments[0]);",
            upload_element
        )
        time.sleep(2)

        Select(driver.find_element(By.XPATH, SELECTORS["form_upload"]["select_status"])).select_by_visible_text("Cadastrar")

        data_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_data"])
        data_field.clear()
        data_field.send_keys(dados_upload['data'])

        driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_contrato"]).send_keys(dados_upload['contrato'])

        placa_sem_hifen = dados_upload['placa'].replace('-', '')
        driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_placa"]).send_keys(placa_sem_hifen)

        Select(driver.find_element(By.XPATH, SELECTORS["form_upload"]["select_tipo_despesa"])).select_by_visible_text("018 - GUINCHO")

        valor_formatado = f"{dados_upload['valor']},00"
        driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_valor"]).send_keys(valor_formatado)

        driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_caixa_arquivo"]).send_keys("0")
        driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_observacao"]).send_keys(dados_upload['tipo_str'])

        time.sleep(1)

        driver.find_element(By.XPATH, SELECTORS["form_upload"]["botao_salvar"]).click()
        time.sleep(4)

        return True

    except Exception as e:
        logging.error("--- ERRO NA ETAPA DE UPLOAD ---")
        logging.error(f"Falha ao fazer upload ou preencher o formulário: {e}", exc_info=True)
        return False


def iniciar_automacao_completa():
    logging.info("--- Iniciando Automação Completa (Maps + Banco) ---")

    lista_placas_log = []
    try:
        logging.info(f"Lendo log de placas já processadas: {NOME_ARQUIVO_LOG}")
        df_log = pd.read_excel(NOME_ARQUIVO_LOG)
        lista_placas_log = df_log[COLUNA_PLACA].astype(str).tolist()
    except FileNotFoundError:
        logging.warning("Arquivo de log não encontrado. Será criado um novo no final.")
    except Exception as e:
        logging.warning(f"Erro ao ler o arquivo de log: {e}. O script continuará.")

    try:
        logging.info(f"Lendo o arquivo: {NOME_ARQUIVO_EXCEL} (Aba: {NOME_ABA})")
        df = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA)
    except FileNotFoundError:
        logging.critical(f"ERRO: Arquivo {NOME_ARQUIVO_EXCEL} não encontrado! Encerrando.")
        return
    except Exception as e:
        logging.critical(f"ERRO ao ler a planilha: {e}", exc_info=True)
        return

    df[COLUNA_STATUS_SAFE_DOC] = df[COLUNA_STATUS_SAFE_DOC].astype(str)
    df[COLUNA_STATUS] = df[COLUNA_STATUS].astype(str)
    df[COLUNA_CATEGORIA] = df[COLUNA_CATEGORIA].astype(str)

    driver = None
    novas_placas_sucesso = []
    data_hoje_formatada = datetime.date.today().strftime("%d-%m-%Y")

    try:
        driver = configurar_driver(headless=True)
        if driver is None:
            logging.critical("Driver do Selenium não pôde ser iniciado. Encerrando.")
            return

        for index, linha in df.iterrows():
            try:
                placa = str(linha[COLUNA_PLACA]).strip()
                status_ak = str(linha[COLUNA_STATUS]).strip()

                if status_ak == "OK":
                    continue
                if placa in lista_placas_log:
                    continue

                contrato = str(linha[COLUNA_CONTRATO]).strip()
                categoria = str(linha[COLUNA_CATEGORIA]).strip()
                end1 = str(linha[COLUNA_END1])
                end2 = str(linha[COLUNA_END2])
                end3 = str(linha[COLUNA_END3])
                status_safe_doc = str(linha[COLUNA_STATUS_SAFE_DOC]).strip()
                teste_val = linha[COLUNA_TESTE]

            except KeyError as e:
                logging.error("--- ERRO (KeyError) ---")
                logging.error(f"Não encontrei a coluna com o nome: {e}")
                break

            end1_url = end1.replace(" ", "+")
            end2_url = end2.replace(" ", "+")
            end3_url = end3.replace(" ", "+")

            logging.info(f"--- Processando Linha {index + 2} (Placa: {placa}, Categoria: {categoria}) ---")

            url_remocao = f"https://www.google.com/maps/dir/{end1_url}/{end2_url}/{end3_url}/{end1_url}"
            url_restituicao = f"https://www.google.com/maps/dir/{end1_url}/{end3_url}/{end1_url}/{end2_url}"

            run_rem = False
            run_rest = False

            if status_safe_doc == "Pendente restituição":
                run_rest = True
            elif status_safe_doc == "Pendente remoção":
                run_rem = True
            elif teste_val == 1:
                run_rem = True
                run_rest = True
            else:
                run_rem = True

            sucesso_geral = True

            caminho_pdf_rem = None
            valor_rem = "N/A"

            if run_rem:
                driver.get(url_remocao)
                km_num_rem, km_str_rem = extrair_km_do_mapa(driver)
                valor_rem = get_valor_por_range(categoria, km_num_rem)
                nome_arquivo = f"{placa}_{contrato}_{data_hoje_formatada}_{km_str_rem}_{valor_rem}_REMO.pdf"

                caminho_pdf_rem = gerar_pdf_mapa(driver, nome_arquivo)
                if not caminho_pdf_rem:
                    sucesso_geral = False
                    logging.warning(f"Falha ao gerar PDF de Remoção para {placa}.")

            caminho_pdf_rest = None
            valor_rest = "N/A"

            if run_rest and sucesso_geral:
                driver.get(url_restituicao)
                km_num_rest, km_str_rest = extrair_km_do_mapa(driver)
                valor_rest = get_valor_por_range(categoria, km_num_rest)
                nome_arquivo = f"{placa}_{contrato}_{data_hoje_formatada}_{km_str_rest}_{valor_rest}_REST.pdf"

                caminho_pdf_rest = gerar_pdf_mapa(driver, nome_arquivo)
                if not caminho_pdf_rest:
                    sucesso_geral = False
                    logging.warning(f"Falha ao gerar PDF de Restituição para {placa}.")

            if sucesso_geral and (run_rem or run_rest):

                upload_sucesso_final = True

                if run_rem:
                    dados_rem = {
                        "placa": placa, "contrato": contrato, "data": data_hoje_formatada,
                        "valor": valor_rem, "tipo_str": "Remocao",
                        "caminho_pdf": caminho_pdf_rem
                    }

                    login_ok = fazer_login_banco(driver)
                    if login_ok:
                        nav_ok = navegar_menu_gca(driver)
                        if nav_ok:
                            try:
                                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
                                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

                                upload_sucesso_rem = preencher_formulario_com_upload(driver, dados_rem)
                                if not upload_sucesso_rem:
                                    upload_sucesso_final = False

                            except Exception as e:
                                logging.error(f"ERRO ao focar no iframe ou preencher formulário: {e}", exc_info=True)
                                upload_sucesso_final = False
                            finally:
                                driver.switch_to.default_content()
                        else:
                            upload_sucesso_final = False
                    else:
                        upload_sucesso_final = False

                if run_rest and upload_sucesso_final:
                    dados_rest = {
                        "placa": placa, "contrato": contrato, "data": data_hoje_formatada,
                        "valor": valor_rest, "tipo_str": "Restituicao",
                        "caminho_pdf": caminho_pdf_rest
                    }

                    login_ok = fazer_login_banco(driver)
                    if login_ok:
                        nav_ok = navegar_menu_gca(driver)
                        if nav_ok:
                            try:
                                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
                                WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

                                upload_sucesso_rest = preencher_formulario_com_upload(driver, dados_rest)
                                if not upload_sucesso_rest:
                                    upload_sucesso_final = False

                            except Exception as e:
                                logging.error(f"ERRO ao focar no iframe ou preencher formulário: {e}", exc_info=True)
                                upload_sucesso_final = False
                            finally:
                                driver.switch_to.default_content()
                        else:
                            upload_sucesso_final = False
                    else:
                        upload_sucesso_final = False

                if upload_sucesso_final:
                    novas_placas_sucesso.append(placa)
                    logging.info(f"Linha {index + 2} (Placa: {placa}) Processada e marcada para o log.")
                else:
                    logging.error(f"Falha no UPLOAD. Linha {index + 2} (Placa: {placa}) NÃO será logada.")
                    logging.critical("Interrompendo script devido a falha no upload.")
                    break

            elif not (run_rem or run_rest):
                logging.info(f"Nenhuma ação necessária para a linha {index + 2}.")
            else:
                logging.warning(f"Falha no processo da linha {index + 2} (provavelmente PDF). NÃO será logada.")

    except Exception as e:
        logging.critical(f"Ocorreu um erro inesperado no loop principal: {e}", exc_info=True)

    finally:
        if driver:
            driver.quit()

        if not novas_placas_sucesso:
            logging.info("Nenhuma placa nova foi processada. Log não precisa ser atualizado.")
            return

        try:
            logging.info(f"Salvando {len(novas_placas_sucesso)} novas placas no log...")

            lista_placas_final = lista_placas_log + novas_placas_sucesso
            lista_placas_final_sem_duplicatas = list(dict.fromkeys(lista_placas_final))
            df_para_salvar = pd.DataFrame(lista_placas_final_sem_duplicatas, columns=[COLUNA_PLACA])

            df_para_salvar.to_excel(NOME_ARQUIVO_LOG, index=False)

            logging.info(f"Arquivo de log '{NOME_ARQUIVO_LOG}' salvo com sucesso!")

        except Exception as e:
            logging.critical(f"ERRO CRÍTICO AO SALVAR O LOG: {e}", exc_info=True)
            logging.critical("Suas placas processadas NÃO foram salvas no log.")

if __name__ == "__main__":
    iniciar_automacao_completa()