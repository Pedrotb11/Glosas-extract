import pandas as pd
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (TimeoutException, 
                                        NoSuchElementException, 
                                        StaleElementReferenceException, 
                                        WebDriverException)
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
    
    all_extracted_data = []
    page_count = 0

    # Configurar opções do Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")  # Novo modo headless do Chrome
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:129.0) Gecko/20100101 Firefox/129.0")

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
        time.sleep(10)

        # Serviços Médicos > RGE
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-label='Serviços Médicos']"))).click()
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//a[@data-label='RGE']"))).click()
        print("RGE selecionado!")
        time.sleep(10)

        # Espera a nova aba abrir
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[-1])
        print("Mudamos para a nova aba!")
        time.sleep(10)

        # Itens
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, "formLink:menuItens"))).click()
        print("Itens selecionados!")
        time.sleep(10)

        # Datas
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "formFiltros:dataPagamentoConta_input"))).send_keys(data_inicio)
        driver.find_element(By.ID, "formFiltros:dataPagamentoContaFim_input").send_keys(data_fim)
        print("Datas preenchidas!")
        time.sleep(10)

        time.sleep(5)

        # Desmarcar checkbox
        checkbox = driver.find_element(By.ID, "formFiltros:somenteGuiasDisponiveis")
        if checkbox.is_selected():
            checkbox.click()
            print("Checkbox desmarcado!")
        time.sleep(10)

        time.sleep(5)

        # Pesquisar
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Pesquisar')]//span[contains(., 'Pesquisar')]"))
        ).click()
        print("Pesquisado! indo pra etapa de extrair dados da tabela!")
        time.sleep(10)

        # Esperar tabela
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "formGrid:formGrid:gridTable_data")))
        time.sleep(10)

        while True:
            page_count += 1
            print(f"--- Processando página {page_count} ---")

            # Espera o corpo da tabela estar presente e ter pelo menos uma linha
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//tbody[@id='formGrid:formGrid:gridTable_data']/tr"))
            )
            time.sleep(10) # Pausa para o JS renderizar

            # --- INÍCIO DA NOVA ABORDAGEM ---
            # 1. Conta o número total de linhas na página atual pelo atributo data-ri
            rows_on_page = driver.find_elements(By.XPATH, "//tbody[@id='formGrid:formGrid:gridTable_data']/tr")
            num_rows = len(rows_on_page)
            print(f"Encontradas {num_rows} linhas na página.")
            time.sleep(10)            

            # 2. Itera usando o índice `data-ri`
            for i in range(num_rows):
                try:
                    print(f"Processando linha com data-ri='{i}'...")
                    
                    # 3. Localiza a linha específica pelo data-ri
                    row_selector = f"//tbody[@id='formGrid:formGrid:gridTable_data']/tr[@data-ri='{i}']"
                    row_to_process = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, row_selector))
                    )
                    time.sleep(10)
                    # Rola até a linha para garantir que ela esteja na tela
                    #driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", row_to_process)
                    #time.sleep(0.5)

                    # 4. Modifica o atributo 'aria-selected' e dispara o evento de clique na linha
                    # Primeiro, desmarca qualquer outra linha que possa estar selecionada
                    try:
                        driver.execute_script("var rows = document.querySelectorAll('#formGrid\\\\:formGrid\\\\:gridTable_data tr[aria-selected=\"true\"]'); rows.forEach(r => r.setAttribute('aria-selected', 'false'));")
                        print("Todos os campos desmarcados")
                    except:
                        print("não coisou")
                    time.sleep(10)
                    # Agora, seleciona a linha atual e clica nela para notificar a página
                    try:
                        driver.execute_script("arguments[0].setAttribute('aria-selected', 'true'); arguments[0].click();", row_to_process)
                        print("linha clickada, Vamos abrir o modal!!!")
                    except:
                        print("não coisou")
                    time.sleep(10)

                    # 5. Abre o modal
                    try:
                        modal_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "formGrid:formButtons:modalDialogButton")))
                        modal_button.click()
                    except:
                        print("não abriu o modal")
                    time.sleep(10)

                    # 6. Extrai os dados do modal (lógica existente)
                    detail_table = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "formGrid:formRecursos:gridTableRecursos_data")))
                    try:
                        headers = [th.text for th in detail_table.find_element(By.TAG_NAME, "thead").find_elements(By.TAG_NAME, "th")]
                        detail_rows = detail_table.find_elements(By.TAG_NAME, "tbody tr")
                        for detail_row in detail_rows:
                            row_data = [td.text.strip() for td in detail_row.find_elements(By.TAG_NAME, "td")]
                            if len(row_data) == len(headers):
                                all_extracted_data.append(dict(zip(headers, row_data)))
                    except NoSuchElementException:
                        print("Aviso: Tabela de detalhes no modal não continha dados ou estrutura esperada.")

                    time.sleep(10)

                    close_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Fechar')]")))
                    close_btn.click()
                    WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "formGrid:formRecursos:gridTableRecursos_data")))
                    time.sleep(10)
                    time.sleep(0.5)

                except Exception as e:
                    print(f"Erro ao processar linha com data-ri='{i}': {e}. Pulando para a próxima.")
                    continue
            # --- FIM DA NOVA ABORDAGEM ---
            time.sleep(10)
            # Lógica de paginação
            try:
                next_button = driver.find_element(By.XPATH, "//a[contains(@class, 'ui-paginator-next') and not(contains(@class, 'ui-state-disabled'))]")
                main_table_id = driver.find_element(By.ID, "formGrid:formGrid:gridTable_data")
                next_button.click()
                WebDriverWait(driver, 20).until(EC.staleness_of(main_table_id))
                print("Navegando para a próxima página...")
            except NoSuchElementException:
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
