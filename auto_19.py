import tkinter as tk
from tkinter import scrolledtext
from threading import Thread, Event
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os

# Variáveis globais para controle de pausa e parada
pause_event = Event()
stop_event = Event()

# Função para excluir o último arquivo
def excluir_ultimo_arquivo(diretorio):
    try:
        arquivos = os.listdir(diretorio)
        arquivos = [f for f in arquivos if os.path.isfile(os.path.join(diretorio, f))]
        
        if arquivos:
            arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)))
            ultimo_arquivo = arquivos[0]
            os.remove(os.path.join(diretorio, ultimo_arquivo))
            return f"Arquivo excluído: {ultimo_arquivo}"
        else:
            return "Nenhum arquivo para excluir."
    except Exception as e:
        return f"Ocorreu um erro ao tentar excluir o arquivo: {e}"

# Função para executar a extração
def executar_extracao(log_area):
    download_folder = os.path.expanduser('G:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_19')
    edge_options = Options()
    edge_options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    while not stop_event.is_set():  # Loop infinito para execução a cada hora
        log_area.insert(tk.END, "Aguarde, a importação de dados está sendo realizada...\n")
        log_area.yview(tk.END)  # Rolagem automática para o final

        driver = webdriver.Edge(options=edge_options)
        max_tentativas = 3
        tentativa = 0

        while tentativa < max_tentativas and not stop_event.is_set():
            try:
                driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

                driver.find_element(By.NAME, "f1").send_keys("LDI")
                driver.find_element(By.NAME, "f2").send_keys("12373493977")
                driver.find_element(By.NAME, "f3").send_keys("gustavo")
                driver.find_element(By.NAME, "f4").send_keys("12032006")
                driver.find_element(By.NAME, "f4").send_keys("+")
                time.sleep(1)

                login_button = driver.find_element(By.ID, "5")
                driver.execute_script("arguments[0].click();", login_button)
                time.sleep(5)

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
                driver.find_element(By.NAME, "f2").clear()
                driver.find_element(By.NAME, "f2").send_keys("CTA")
                driver.find_element(By.NAME, "f3").send_keys("19+")
                time.sleep(5)

                abas = driver.window_handles
                driver.switch_to.window(abas[-1])

                data_atual = datetime.now().strftime('%d%m%y')
                driver.find_element(By.NAME, "data_prev_man").clear()
                driver.find_element(By.NAME, "data_prev_man").send_keys(data_atual)
                driver.find_element(By.NAME, "hora_prev_man").clear()
                driver.find_element(By.NAME, "hora_prev_man").send_keys("0001")
                driver.find_element(By.NAME, "relatorio_excel").clear()
                driver.find_element(By.NAME, "relatorio_excel").send_keys("S")

                login_button = driver.find_element(By.ID, "btn_envia")
                driver.execute_script("arguments[0].click();", login_button)
                time.sleep(15)

                log_area.insert(tk.END, "O arquivo foi baixado para a pasta destino.\n")
                log_area.yview(tk.END)

                log_area.insert(tk.END, excluir_ultimo_arquivo(download_folder) + "\n")
                log_area.yview(tk.END)

                break  # Sai do loop de tentativas
            except (TimeoutException) as e:
                log_area.insert(tk.END, f"Ocorreu um erro: {e}. Tentativa {tentativa + 1} de {max_tentativas}\n")
                log_area.yview(tk.END)
                tentativa += 1
                time.sleep(600)  # Espera 10 minutos caso ocorra algum erro com a página.
            finally:
                driver.quit()

        log_area.insert(tk.END, "Aguarde 01 hora até a próxima execução...\n")
        log_area.insert(tk.END, "-------------------------------------------\n")
        log_area.yview(tk.END)

        # Espera 1 hora antes de repetir a extração, mas verifica se a extração deve ser parada
        for _ in range(3600):  # 3600 segundos
            if stop_event.is_set():  # Se o evento de parada estiver definido
                log_area.insert(tk.END, "Extração parada. Saindo...\n")
                log_area.yview(tk.END)
                return  # Sai da função
            if pause_event.is_set():  # Se o evento de pausa estiver definido
                log_area.insert(tk.END, "Extração pausada. Pressione 'Continuar' para retomar.\n")
                log_area.yview(tk.END)
                while pause_event.is_set():  # Aguarda até que o evento de pausa seja removido
                    time.sleep(1)
            time.sleep(1)  # Espera 1 segundo

# Função para iniciar a extração em uma thread separada
def iniciar_extracao(log_area):
    thread = Thread(target=executar_extracao, args=(log_area,))
    thread.start()

# Função para pausar a extração
def pausar_extracao():
    pause_event.set()  # Define o evento de pausa

# Função para continuar a extração
def continuar_extracao():
    pause_event.clear()  # Limpa o evento de pausa

# Função para parar a extração
def parar_extracao():
    stop_event.set()  # Define o evento de parada

# Configuração da interface gráfica
def criar_interface():
    root = tk.Tk()
    root.title("Extrator de Dados")

    # Área de texto para logs
    log_area = scrolledtext.ScrolledText(root, width=80, height=20)
    log_area.pack(pady=10)

    # Botão para iniciar a extração
    btn_iniciar = tk.Button(root, text="Iniciar Extração", command=lambda: iniciar_extracao(log_area))
    btn_iniciar.pack(pady=5)

    # Botão para pausar a extração
    btn_pausar = tk.Button(root, text="Pausar Extração", command=pausar_extracao)
    btn_pausar.pack(pady=5)

    # Botão para continuar a extração
    btn_continuar = tk.Button(root, text="Continuar Extração", command=continuar_extracao)
    btn_continuar.pack(pady=5)

    # Botão para parar a extração
    btn_parar = tk.Button(root, text="Parar Extração", command=parar_extracao)
    btn_parar.pack(pady=5)

    # Botão para sair
    btn_sair = tk.Button(root, text="Sair", command=root.quit)
    btn_sair.pack(pady=5)

    root.mainloop()

# Chamada para criar a interface
if __name__ == "__main__":
    criar_interface()