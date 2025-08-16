import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

def run_scraper():
    # Solicitar inputs do usuário
    code = input("Digite o CNPJ: ")
    user = input("Digite o usuário: ")
    password = input("Digite a senha: ")
    data_inicio = input("Digite a data de início (DD/MM/AAAA): ")
    data_fim = input("Digite a data de fim (DD/MM/AAAA): ")

    # Configurar opções do Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")  # Novo modo headless do Chrome
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.6613.137 Safari/537.36")

    # Caminho padrão do Chrome no Windows
    chrome_options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    try:
        # Inicializar o WebDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("Chrome iniciado com sucesso!")

        # Navegar para a página de login
        driver.get("https://saude.sulamericaseguros.com.br/prestador/login/?accessError=2")
        print("Página carregada!")

        # Aceitar cookies (se aparecer)
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Continuar')]"))
            ).click()
            print("Cookies aceitos!")
        except:
            print("Nenhum aviso de cookies encontrado.")

        # Preencher login
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "code"))).send_keys(code)
        driver.find_element(By.ID, "user").send_keys(user)
        driver.find_element(By.ID, "senha").send_keys(password)
        print("Campos de login preenchidos!")

        # Entrar
        driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()
        time.sleep(3)

        # Fechar pop-up se aparecer
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "p.close-btn"))
            ).click()
            print("Pop-up fechado!")
        except:
            print("Nenhum pop-up encontrado, prosseguindo...")

        # Serviços Médicos > RGE
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-label='Serviços Médicos']"))).click()
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-label='RGE']"))).click()
        print("RGE selecionado!")

        # Espera a nova aba abrir
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])
        print("Mudamos para a nova aba!")

        # Itens
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "formLink:menuItens"))).click()
        print("Itens selecionados!")

        # Datas
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "formFiltros:dataPagamentoConta_input"))).send_keys(data_inicio)
        driver.find_element(By.ID, "formFiltros:dataPagamentoContaFim_input").send_keys(data_fim)
        print("Datas preenchidas!")

        # Desmarcar checkbox
        checkbox = driver.find_element(By.ID, "formFiltros:somenteGuiasDisponiveis")
        if checkbox.is_selected():
            checkbox.click()
            print("Checkbox desmarcado!")

        # Pesquisar
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Pesquisar')]//span[contains(., 'Pesquisar')]"))
        ).click()
        print("Pesquisado! indo pra etapa de extrair dados da tabela!")

        # Esperar tabela
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "formGrid:formGrid:gridTable_data")))

        all_extracted_data = []
        page_count = 0

        while True:
            page_count += 1
            print(f"Processando página {page_count}...")

            main_table_body = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "formGrid:formGrid:gridTable_data"))
            )
            num_rows = len(main_table_body.find_elements(By.TAG_NAME, "tr"))
            if num_rows == 0:
                print("Nenhuma linha encontrada na página. Finalizando.")
                break
            print(f"Encontradas {num_rows} linhas na página {page_count}.")

            for i in range(num_rows):
                try:
                    print(f"Processando linha {i+1}/{num_rows}...")
                    
                    # --- INÍCIO DA CORREÇÃO FINAL ---
                    # Construa um XPath dinâmico que vai direto ao checkbox da linha 'i'.
                    # O índice do XPath é 1-based, então usamos i + 1.
                    checkbox_xpath = f"//tbody[@id='formGrid:formGrid:gridTable_data']/tr[{i+1}]//div[contains(@class, 'ui-chkbox-box')]"
                    
                    # Espere o elemento estar presente e clicável em uma única ação
                    checkbox_box = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
                    )
                    # --- FIM DA CORREÇÃO FINAL ---

                    # Usar JavaScript para o clique continua sendo a abordagem mais segura
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", checkbox_box)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", checkbox_box)

                    # ... (resto do código para abrir modal, extrair dados e fechar) ...
                    modal_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "formGrid:formButtons:modalDialogButton"))
                    )
                    modal_button.click()

                    detail_table = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.ID, "formGrid:formRecursos:gridTableRecursos_data"))
                    )
                    
                    headers = [th.text for th in detail_table.find_element(By.TAG_NAME, "thead").find_elements(By.TAG_NAME, "th")]
                    detail_rows = detail_table.find_elements(By.TAG_NAME, "tbody tr")

                    for detail_row in detail_rows:
                        row_data = [td.text.strip() for td in detail_row.find_elements(By.TAG_NAME, "td")]
                        if len(row_data) == len(headers):
                            all_extracted_data.append(dict(zip(headers, row_data)))
                    
                    print(f"Linha {i+1}: Dados do modal extraídos.")

                    close_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Fechar')]"))
                    )
                    close_btn.click()

                    WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.ID, "formGrid:formRecursos:gridTableRecursos_data"))
                    )
                    time.sleep(0.5)

                except (TimeoutException, StaleElementReferenceException) as e:
                    print(f"Erro de timing na linha {i+1}: {type(e).__name__}. A página pode ter atualizado durante a operação. Pulando linha.")
                    continue
                except Exception as e:
                    print(f"Erro inesperado ao processar linha {i+1}: {e}")
                    continue

            # ... (lógica de paginação) ...
            try:
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'ui-paginator-next') and not(contains(@class, 'ui-state-disabled'))]"))
                )
                main_table_id = driver.find_element(By.ID, "formGrid:formGrid:gridTable_data")
                next_button.click()
                WebDriverWait(driver, 20).until(EC.staleness_of(main_table_id))
                print("Navegando para a próxima página...")
            except (NoSuchElementException, TimeoutException):
                print("Nenhuma próxima página encontrada, finalizando scraping.")
                break

    except Exception as e:
        print(f"Ocorreu um erro geral na execução: {e}")
    finally:
        if driver:
            driver.quit()
            print("Driver finalizado.")
        
        if all_extracted_data:
            df = pd.DataFrame(all_extracted_data)
            df.to_excel("dados_extraidos.xlsx", index=False)
            print(f"Dados salvos com sucesso em 'dados_extraidos.xlsx'. Total de {len(all_extracted_data)} registros.")

if __name__ == "__main__":
    run_scraper()