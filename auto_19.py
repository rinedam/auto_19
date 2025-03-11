import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox
from tkinter import font as tkFont
from ttkthemes import ThemedTk
from threading import Thread, Event
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, time as dt_time, timedelta
import os
import requests
import logging

# Configuração de logging
logging.basicConfig(
    filename='extracao.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Variáveis globais para controle de pausa e parada
pause_event = Event()
stop_event = Event()

# Função para verificar a conectividade com o site específico
def verificar_conexao():
    try:
        response = requests.get("https://www.google.com/", timeout=5)
        return response.status_code == 200
    except requests.ConnectionError:
        return False

# Função para excluir o último arquivo, ignorando "desktop.ini"
def excluir_ultimo_arquivo(diretorio):
    try:
        arquivos = [f for f in os.listdir(diretorio) if os.path.isfile(os.path.join(diretorio, f)) and f != "desktop.ini"]
        if arquivos:
            arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)))
            ultimo_arquivo = arquivos[0]
            os.remove(os.path.join(diretorio, ultimo_arquivo))
            return f"Arquivo excluído: {ultimo_arquivo}"
        else:
            return "Nenhum arquivo para excluir."
    except Exception as e:
        return f"Ocorreu um erro ao tentar excluir o arquivo: {e}"

# Função para verificar se está no horário comercial
def esta_no_horario_comercial():
    agora = datetime.now().time()
    horario_inicio = dt_time(8, 0)
    horario_fim = dt_time(18, 0)
    return horario_inicio <= agora <= horario_fim

# Função para calcular o tempo restante até as 08:00 do próximo dia
def tempo_ate_proxima_extracao():
    agora = datetime.now()
    if agora.time() > dt_time(18, 0):
        proximo_dia = agora + timedelta(days=1)
        proximo_dia = proximo_dia.replace(hour=8, minute=0, second=0, microsecond=0)
        tempo_restante = (proximo_dia - agora).total_seconds()
    else:
        proxima_extracao = agora.replace(hour=8, minute=0, second=0, microsecond=0)
        if agora >= proxima_extracao:
            tempo_restante = 3600
        else:
            tempo_restante = (proxima_extracao - agora).total_seconds()
    return tempo_restante

# Função para executar a extração
def executar_extracao(log_area, status_label):
    download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\CTRCs Disponíveis para Transferência - 19\\DB_CTRCs Disponíveis para Transferência\\CTA')
    edge_options = Options()
    edge_options.add_argument("--headless")
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--window-size=1920,1080")
    edge_options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    while not stop_event.is_set():
        if pause_event.is_set():
            atualizar_status(status_label, "Extração pausada.", "amarelo")
            log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Extração pausada.\n")
            log_area.yview(tk.END)
            pause_event.wait()  # Aguarda até que o evento de pausa seja limpo
            continue

        if not esta_no_horario_comercial():
            tempo_restante = tempo_ate_proxima_extracao()
            atualizar_status(status_label, f"Fora do horário comercial. Próxima extração em {int(tempo_restante // 3600)} horas e {int((tempo_restante % 3600) // 60)} minutos.", "amarelo")
            log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Fora do horário comercial. Próxima extração em {int(tempo_restante // 3600)} horas e {int((tempo_restante % 3600) // 60)} minutos.\n")
            log_area.yview(tk.END)
            time.sleep(tempo_restante)  # Aguarda até a próxima extração
            continue

        if not verificar_conexao():
            atualizar_status(status_label, "Sem conexão com o site. Tentando reconectar...", "vermelho")
            log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Sem conexão com o site. Tentando reconectar...\n")
            log_area.yview(tk.END)
            time.sleep(60)
            continue

        atualizar_status(status_label, "Iniciando extração...", "azul")
        log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Iniciando extração...\n")
        log_area.yview(tk.END)

        try:
            with webdriver.Edge(options=edge_options) as driver:
                max_tentativas = 3
                tentativa = 0

                while tentativa < max_tentativas and not stop_event.is_set():
                    try:
                        driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
                        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "f1")))

                        driver.find_element(By.NAME, "f1").send_keys("LDI")
                        driver.find_element(By.NAME, "f2").send_keys("41968069020")
                        driver.find_element(By.NAME, "f3").send_keys("botlogdi")
                        driver.find_element(By.NAME, "f4").send_keys("logbotdi")
                        driver.find_element(By.NAME, "f4").send_keys("+")
                        time.sleep(1)

                        login_button = driver.find_element(By.ID, "5")
                        driver.execute_script("arguments[0].click();", login_button)
                        time.sleep(5)

                        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "f2")))
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

                        log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Arquivo baixado com sucesso.\n")
                        log_area.yview(tk.END)

                        log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {excluir_ultimo_arquivo(download_folder)}\n")
                        log_area.yview(tk.END)

                        break  # Sai do loop de tentativas
                    except (TimeoutException, WebDriverException) as e:
                        logging.error(f"Erro: {e}. Tentativa {tentativa + 1} de {max_tentativas}")
                        log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Erro: {e}. Tentativa {tentativa + 1} de {max_tentativas}\n")
                        log_area.yview(tk.END)
                        tentativa += 1
                        time.sleep(600)  # Espera 10 minutos antes de tentar novamente
        except Exception as e:
            logging.error(f"Erro crítico: {e}")
            log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Erro crítico: {e}\n")
            log_area.yview(tk.END)

        atualizar_status(status_label, "Aguardando próxima execução...", "verde")
        log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Aguardando 1 hora para próxima execução...\n")
        log_area.insert(tk.END, "-" * 80 + "\n")
        log_area.yview(tk.END)

        # Espera 1 hora ou até que o evento de parada seja acionado
        for _ in range(3600):  # 3600 segundos = 1 hora
            if stop_event.is_set():
                atualizar_status(status_label, "Extração parada.", "vermelho")
                log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Extração parada.\n")
                log_area.yview(tk.END)
                return
            if pause_event.is_set():
                atualizar_status(status_label, "Extração pausada.", "amarelo")
                log_area.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - Extração pausada.\n")
                log_area.yview(tk.END)
                pause_event.wait()  # Aguarda até que o evento de pausa seja limpo
            time.sleep(1)

# Função para atualizar o status na interface
def atualizar_status(label, texto, cor):
    cores = {
        "vermelho": "#FF0000",
        "verde": "#00FF00",
        "azul": "#0000FF",
        "amarelo": "#FFFF00"
    }
    label.config(text=texto, foreground=cores.get(cor, "#000000"))

# Função para iniciar a extração em uma thread separada
def iniciar_extracao(log_area, status_label):
    if not stop_event.is_set() and not pause_event.is_set():
        thread = Thread(target=executar_extracao, args=(log_area, status_label))
        thread.start()
        # Habilitar botões de controle enquanto a extração está em andamento
        pausar_btn.config(state=tk.NORMAL)
        continuar_btn.config(state=tk.DISABLED)  # Desabilitar botão de continuar inicialmente
        parar_btn.config(state=tk.NORMAL)

# Função para pausar a extração
def pausar_extracao():
    pause_event.set()
    pausar_btn.config(state=tk.DISABLED)  # Desabilitar botão de pausar
    continuar_btn.config(state=tk.NORMAL)  # Habilitar botão de continuar

# Função para continuar a extração
def continuar_extracao():
    pause_event.clear()  # Limpa o evento de pausa
    logging.info("Extração retomada.")
    pausar_btn.config(state=tk.NORMAL)  # Habilitar botão de pausar
    continuar_btn.config(state=tk.DISABLED)  # Desabilitar botão de continuar

# Função para parar a extração
def parar_extracao():
    stop_event.set()
    pause_event.clear()
    pausar_btn.config(state=tk.DISABLED)  # Desabilitar botão de pausar
    continuar_btn.config(state=tk.DISABLED)  # Desabilitar botão de continuar
    parar_btn.config(state=tk.DISABLED)  # Desabilitar botão de parar

# Configuração da interface gráfica
def criar_interface():
    root = ThemedTk(theme="arc")  # Tema moderno
    root.title("Extrator de Dados Automático")
    root.geometry("900x700")  # Tamanho da janela

    # Fontes personalizadas
    fonte_titulo = tkFont.Font(family="Helvetica", size=18, weight="bold")
    fonte_botoes = tkFont.Font(family="Helvetica", size=12)

    # Frame principal
    main_frame = ttk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    # Título
    titulo = ttk.Label(main_frame, text="Extrator de Dados Automático", font=fonte_titulo)
    titulo.grid(row=0, column=0, columnspan=4, pady=(0, 20))

    # Área de texto para logs
    log_area = scrolledtext.ScrolledText(main_frame, width=100, height=20, wrap=tk.WORD)
    log_area.grid(row=1, column=0, columnspan=4, pady=(0, 20))

    # Label de status
    status_label = ttk.Label(main_frame, text="Status: Ocioso", font=fonte_botoes)
    status_label.grid(row=2, column=0, columnspan=4, pady=(0, 20))

    # Botões
    global pausar_btn, continuar_btn, parar_btn  # Tornar os botões globais
    pausar_btn = ttk.Button(main_frame, text="Pausar Extração", command=pausar_extracao)
    continuar_btn = ttk.Button(main_frame, text="Continuar Extração", command=continuar_extracao, state=tk.DISABLED)
    parar_btn = ttk.Button(main_frame, text="Parar Extração", command=parar_extracao)

    # Centralizando os botões
    pausar_btn.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
    continuar_btn.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
    parar_btn.grid(row=3, column=2, padx=5, pady=5, sticky="ew")

    # Ajustar o layout para centralizar os botões
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_columnconfigure(2, weight=1)

    # Botão para sair
    btn_sair = ttk.Button(main_frame, text="Sair", command=root.quit)
    btn_sair.grid(row=4, column=0, columnspan=4, pady=10, sticky="ew")

    # Ajustar o layout
    for i in range(4):
        main_frame.grid_columnconfigure(i, weight=1)

    # Iniciar a extração automaticamente ao criar a interface
    iniciar_extracao(log_area, status_label)

    root.mainloop()

# Chamada para criar a interface
if __name__ == "__main__":
    criar_interface()