from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
import os
import time

def excluir_ultimo_arquivo(diretorio):
    try:
        # Listar arquivos no diretório
        arquivos = os.listdir(diretorio)
        # Filtrar apenas arquivos (ignorar diretórios)
        arquivos = [f for f in arquivos if os.path.isfile(os.path.join(diretorio, f))]
        
        if arquivos:
            # Ordenar arquivos por data de modificação
            arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)))
            # Obter o último arquivo (o mais antigo)
            ultimo_arquivo = arquivos[0]  # O primeiro na lista é o mais antigo
            # Excluir o último arquivo
            os.remove(os.path.join(diretorio, ultimo_arquivo))
            print(f"Arquivo excluido: {ultimo_arquivo}")
        else:
            print("Nenhum arquivo para excluir.")
    except Exception as e:
        print(f"Ocorreu um erro ao tentar excluir o arquivo: {e}")

def executar_extracao():

    # Caminho para a pasta de downloads desejada
    download_folder = os.path.expanduser('G:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_19')

    # Configurações do Edge
    edge_options = Options()
    edge_options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,  # Define o diretório de download
        "download.prompt_for_download": False,           # Não solicitar confirmação para download
        "download.directory_upgrade": True,               # Permitir a atualização do diretório
        "safebrowsing.enabled": True                       # Habilitar navegação segura
    })

    print("Aguarde, a importacao de dados esta sendo realizada")

    # Inicializa o WebDriver
    driver = webdriver.Edge(options=edge_options)

    # Defina o número máximo de tentativas
    max_tentativas = 3
    tentativa = 0

    while tentativa < max_tentativas:
        try:
            # Navega até a página do formulário
            driver.get("https://sistema.ssw.inf.br/bin/ssw0422")  # Substitua pela URL do seu formulário

            # Atraso para garantir que a página carregue completamente
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

            # Preenche os campos de login
            driver.find_element(By.NAME, "f1").send_keys("LDI")
            driver.find_element(By.NAME, "f2").send_keys("12373493977")  # Exemplo de CPF
            driver.find_element(By.NAME, "f3").send_keys("gustavo")
            driver.find_element(By.NAME, "f4").send_keys("12032006")

            # Envia a tecla "+" e aguarda 1 segundo
            driver.find_element(By.NAME, "f4").send_keys("+")  # Pressiona a tecla "+"
            time.sleep(1)  # Atraso de 1 segundo

            # Clica no botão de login diretamente
            login_button = driver.find_element(By.ID, "5")
            
            # Usando JavaScript para clicar no botão
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(5)

            # Preenche os campos de Unidade e Opção
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
            driver.find_element(By.NAME, "f2").clear()
            time.sleep(0.5)
            driver.find_element(By.NAME, "f2").send_keys("CTA")

            driver.find_element(By.NAME, "f3").send_keys("19+")
            
            time.sleep(5)  # Atraso para carregar a nova aba
            
            abas = driver.window_handles  # Lista o número de abas abertas.

            # Muda o foco para a última aba (a nova aba)
            driver.switch_to.window(abas[-1])

            data_atual = datetime.now().strftime('%d%m%y')
            hora_atual = datetime.now().strftime('%H:%M:%S')

            driver.find_element(By.NAME, "data_prev_man").clear()
            time.sleep(0.2)
            driver.find_element(By.NAME, "data_prev_man").send_keys(data_atual)
            time.sleep(0.2)
            driver.find_element(By.NAME, "hora_prev_man").clear()
            time.sleep(0.2)
            driver.find_element(By.NAME, "hora_prev_man").send_keys("0001")

            driver.find_element(By.NAME, "data_emit_ctrc").clear()
            time.sleep(0.2)
            driver.find_element(By.NAME, "hora_emit_ctrc").clear()

            driver.find_element(By.NAME, "relatorio_excel").clear()
            time.sleep(0.2)
            driver.find_element(By.NAME, "relatorio_excel").send_keys("S")

            # Clica no botão de envio
            login_button = driver.find_element(By.ID, "btn_envia")
            
            # Usando JavaScript para clicar no botão
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(15)   

            print("O arquivo foi baixado para a pasta destino -", hora_atual)

            excluir_ultimo_arquivo(download_folder) # Puxa a função de apagar o último arquivo.

            break # Sai do loop
        except (WebDriverWait, TimeoutException) as e:
            print(f"Ocorreu um erro: {e}. Tentativa {tentativa + 1} de {max_tentativas}")
            tentativa += 1
            time.sleep(600) #Espera 10 minutos caso ocorra algum erro com a página.

        finally:
            # Fecha o navegador
            driver.quit()

while True:
   executar_extracao()
   time.sleep(3600)