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
import sys
from theme import ModernTheme
from assets.styles import (
    APP_NAME, APP_VERSION, APP_DESCRIPTION, PADDING, 
    STATUS_MESSAGES, STATUS_STYLES, BUTTONS, 
    get_status_style, get_status_message
)

# Configuração de logging
logging.basicConfig(
    filename='extracao.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# Handler para mostrar logs também no console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console_handler.setFormatter(formatter)
logging.getLogger().addHandler(console_handler)

# Variáveis globais para controle de pausa e parada
pause_event = Event()
stop_event = Event()

class Application:
    """Aplicação principal para extração automatizada de dados SSW"""
    
    def __init__(self, root):
        """Inicializa a aplicação"""
        self.root = root
        self.setup_window()
        self.create_ui()
        self.initialize_state()
        
    def setup_window(self):
        """Configura a janela principal"""
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("950x700")
        self.root.minsize(800, 600)
        
        # Centraliza a janela na tela
        window_width = 950
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")
        
        # Aplica o tema moderno
        self.style = ModernTheme.apply(self.root)
        
        # Configura o ícone da aplicação
        try:
            icon_path = os.path.join("assets", "logo.svg")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            logging.warning("Não foi possível carregar o ícone da aplicação.")
        
        # Configura comportamento ao fechar
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def create_ui(self):
        """Cria a interface de usuário"""
        # Frame principal com padding
        self.main_frame = ttk.Frame(self.root, style='TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=PADDING['large'], pady=PADDING['large'])
        
        # Título e subtítulo
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, PADDING['medium']))
        
        self.title_label = ttk.Label(
            self.header_frame, 
            text=APP_NAME, 
            style='Title.TLabel'
        )
        self.title_label.pack(anchor=tk.W)
        
        self.subtitle_label = ttk.Label(
            self.header_frame, 
            text=APP_DESCRIPTION, 
            style='Subtitle.TLabel'
        )
        self.subtitle_label.pack(anchor=tk.W)
        
        # Separador
        ttk.Separator(self.main_frame, orient='horizontal').pack(fill=tk.X, pady=PADDING['medium'])
        
        # Frame para status e controles no topo
        self.status_frame = ttk.Frame(self.main_frame, style='Surface.TFrame')
        self.status_frame.pack(fill=tk.X, pady=PADDING['medium'])
        
        # Status com ícone
        self.status_container = ttk.Frame(self.status_frame, style='Surface.TFrame')
        self.status_container.pack(side=tk.LEFT, padx=PADDING['medium'], pady=PADDING['medium'])
        
        self.status_label = ttk.Label(
            self.status_container, 
            text="Status: Ocioso", 
            style='StatusIdle.TLabel'
        )
        self.status_label.pack(side=tk.LEFT)
        
        # Frame para os botões de controle
        self.controls_frame = ttk.Frame(self.status_frame, style='Surface.TFrame')
        self.controls_frame.pack(side=tk.RIGHT, padx=PADDING['medium'], pady=PADDING['medium'])
        
        # Botões de controle
        self.start_btn = ttk.Button(
            self.controls_frame,
            text=BUTTONS['start']['text'],
            style=BUTTONS['start']['style'],
            command=self.iniciar_extracao
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, PADDING['small']))
        
        self.pause_btn = ttk.Button(
            self.controls_frame,
            text=BUTTONS['pause']['text'],
            style=BUTTONS['pause']['style'],
            command=self.pausar_extracao,
            state=tk.DISABLED
        )
        self.pause_btn.pack(side=tk.LEFT, padx=(0, PADDING['small']))
        
        self.resume_btn = ttk.Button(
            self.controls_frame,
            text=BUTTONS['resume']['text'],
            style=BUTTONS['resume']['style'],
            command=self.continuar_extracao,
            state=tk.DISABLED
        )
        self.resume_btn.pack(side=tk.LEFT, padx=(0, PADDING['small']))
        
        self.stop_btn = ttk.Button(
            self.controls_frame,
            text=BUTTONS['stop']['text'],
            style=BUTTONS['stop']['style'],
            command=self.parar_extracao,
            state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT)
        
        # Área de log
        self.log_frame = ttk.Frame(self.main_frame, style='Surface.TFrame')
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=PADDING['medium'])
        
        self.log_header = ttk.Label(
            self.log_frame, 
            text="Registro de Atividades", 
            style='Heading.TLabel'
        )
        self.log_header.pack(anchor=tk.W, padx=PADDING['medium'], pady=PADDING['small'])
        
        # Estilização da área de texto
        text_font = tkFont.Font(family="Consolas", size=10)
        
        self.log_area = scrolledtext.ScrolledText(
            self.log_frame, 
            wrap=tk.WORD,
            font=text_font, 
            background="#f8f9fa", 
            foreground="#333333",
            borderwidth=1,
            relief="solid"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=PADDING['medium'], pady=PADDING['small'])
        
        # Footer com informações
        self.footer_frame = ttk.Frame(self.main_frame)
        self.footer_frame.pack(fill=tk.X, pady=PADDING['small'])
        
        # Versão da aplicação no rodapé à direita
        self.version_label = ttk.Label(
            self.footer_frame, 
            text=f"v{APP_VERSION}", 
            font=("Helvetica", 8)
        )
        self.version_label.pack(side=tk.RIGHT)
        
    def initialize_state(self):
        """Inicializa os estados da aplicação"""
        # Limpa os eventos de pausa e parada
        stop_event.clear()
        pause_event.clear()
        
        # Configura os botões iniciais
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.resume_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        # Adiciona mensagem inicial ao log
        self.adicionar_log("Sistema iniciado e pronto para extração.")
        self.atualizar_status('idle')
    
    def atualizar_status(self, status_key, **kwargs):
        """Atualiza o status na interface"""
        message = get_status_message(status_key, **kwargs)
        style = get_status_style(status_key)
        
        # Atualiza o texto e o estilo
        self.status_label.config(text=f"Status: {message}", style=style)
        
        # Atualiza o log com a mesma mensagem se necessário
        if status_key in ['initializing', 'extracting', 'waiting', 'paused', 'stopped', 'error', 'success']:
            self.adicionar_log(message)
    
    def adicionar_log(self, mensagem, nivel='info'):
        """Adiciona uma mensagem ao log com formatação de hora"""
        hora_atual = datetime.now().strftime('%H:%M:%S')
        
        # Definição de cores para diferentes níveis de log
        cores = {
            'info': '#333333',     # Preto padrão
            'success': '#2ecc71',  # Verde
            'warning': '#f39c12',  # Laranja
            'error': '#e74c3c'     # Vermelho
        }
        
        cor = cores.get(nivel, cores['info'])
        
        self.log_area.tag_config(nivel, foreground=cor)
        
        # Adiciona a mensagem com a tag de cor apropriada
        self.log_area.insert(tk.END, f"[{hora_atual}] ", 'time')
        self.log_area.insert(tk.END, f"{mensagem}\n", nivel)
        
        # Configura a tag de tempo para ser cinza
        self.log_area.tag_config('time', foreground='#7f8c8d')
        
        # Rola para a última linha
        self.log_area.yview(tk.END)
        
        # Registra no log do sistema
        if nivel == 'info':
            logging.info(mensagem)
        elif nivel == 'success':
            logging.info(f"SUCESSO: {mensagem}")
        elif nivel == 'warning':
            logging.warning(mensagem)
        elif nivel == 'error':
            logging.error(mensagem)
    
    def iniciar_extracao(self):
        """Inicia a extração em uma thread separada"""
        if not stop_event.is_set() and not pause_event.is_set():
            # Atualiza os botões
            self.start_btn.config(state=tk.DISABLED)
            self.pause_btn.config(state=tk.NORMAL)
            self.resume_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            
            # Atualiza o status
            self.atualizar_status('initializing')
            
            # Inicia a thread de extração
            self.extraction_thread = Thread(target=self.executar_extracao)
            self.extraction_thread.daemon = True
            self.extraction_thread.start()
    
    def pausar_extracao(self):
        """Pausa a extração"""
        pause_event.set()
        self.atualizar_status('paused')
        
        # Atualiza os botões
        self.pause_btn.config(state=tk.DISABLED)
        self.resume_btn.config(state=tk.NORMAL)
    
    def continuar_extracao(self):
        """Continua a extração"""
        pause_event.clear()
        self.atualizar_status('extracting')
        
        # Atualiza os botões
        self.pause_btn.config(state=tk.NORMAL)
        self.resume_btn.config(state=tk.DISABLED)
    
    def parar_extracao(self):
        """Para a extração"""
        stop_event.set()
        pause_event.clear()
        
        self.atualizar_status('stopped')
        
        # Atualiza os botões
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.resume_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
        # Adiciona separador no log
        self.log_area.insert(tk.END, "-" * 80 + "\n")
        self.log_area.yview(tk.END)
    
    def on_closing(self):
        """Manipulador para quando a janela é fechada"""
        if messagebox.askokcancel("Sair", "Deseja realmente sair? A extração será interrompida."):
            stop_event.set()
            self.root.destroy()
    
    def verificar_conexao(self):
        """Verifica a conectividade com o site específico"""
        try:
            response = requests.get("https://www.google.com/", timeout=5)
            return response.status_code == 200
        except requests.ConnectionError:
            return False
    
    def excluir_ultimo_arquivo(self, diretorio):
        """Exclui o último arquivo, ignorando "desktop.ini" """
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
    
    def esta_no_horario_comercial(self):
        """Verifica se está no horário comercial"""
        agora = datetime.now().time()
        horario_inicio = dt_time(8, 0)
        horario_fim = dt_time(18, 0)
        return horario_inicio <= agora <= horario_fim
    
    def tempo_ate_proxima_extracao(self):
        """Calcula o tempo restante até as 08:00 do próximo dia"""
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
    
    def executar_extracao(self):
        """Executa a extração de dados"""
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
            # Verifica se está pausado
            if pause_event.is_set():
                time.sleep(1)
                continue
            
            # Verifica se está no horário comercial
            if not self.esta_no_horario_comercial():
                tempo_restante = self.tempo_ate_proxima_extracao()
                horas = int(tempo_restante // 3600)
                minutos = int((tempo_restante % 3600) // 60)
                
                self.atualizar_status('outside_hours')
                self.adicionar_log(
                    f"Fora do horário comercial. Próxima extração em {horas} horas e {minutos} minutos.",
                    nivel='warning'
                )
                
                # Aguarda até a próxima extração, verificando a cada segundo se foi pausado ou parado
                for _ in range(int(min(tempo_restante, 3600))):  # Verifica de novo após no máximo 1 hora
                    if stop_event.is_set() or pause_event.is_set():
                        break
                    time.sleep(1)
                
                continue
            
            # Verifica conexão com o site
            if not self.verificar_conexao():
                self.atualizar_status('no_connection')
                self.adicionar_log(
                    "Sem conexão com o site. Tentando reconectar em 60 segundos...", 
                    nivel='error'
                )
                
                # Aguarda 60 segundos antes de tentar novamente
                for _ in range(60):
                    if stop_event.is_set() or pause_event.is_set():
                        break
                    time.sleep(1)
                
                continue
            
            # Inicia a extração
            self.atualizar_status('extracting')
            
            try:
                with webdriver.Edge(options=edge_options) as driver:
                    max_tentativas = 3
                    tentativa = 0
                    
                    while tentativa < max_tentativas and not stop_event.is_set():
                        if pause_event.is_set():
                            break
                        
                        try:
                            # Acessa o site
                            driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "f1")))
                            
                            # Preenche o formulário de login
                            driver.find_element(By.NAME, "f1").send_keys("LDI")
                            driver.find_element(By.NAME, "f2").send_keys("41968069020")
                            driver.find_element(By.NAME, "f3").send_keys("botlogdi")
                            driver.find_element(By.NAME, "f4").send_keys("logbotdi")
                            driver.find_element(By.NAME, "f4").send_keys("+")
                            time.sleep(1)
                            
                            # Clica no botão de login
                            login_button = driver.find_element(By.ID, "5")
                            driver.execute_script("arguments[0].click();", login_button)
                            time.sleep(5)
                            
                            # Acessa a área de relatórios
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "f2")))
                            driver.find_element(By.NAME, "f2").clear()
                            driver.find_element(By.NAME, "f2").send_keys("CTA")
                            driver.find_element(By.NAME, "f3").send_keys("19+")
                            time.sleep(5)
                            
                            # Muda para a nova aba/janela
                            abas = driver.window_handles
                            driver.switch_to.window(abas[-1])
                            
                            # Preenche os campos do relatório
                            data_atual = datetime.now().strftime('%d%m%y')
                            driver.find_element(By.NAME, "data_prev_man").clear()
                            driver.find_element(By.NAME, "data_prev_man").send_keys(data_atual)
                            driver.find_element(By.NAME, "hora_prev_man").clear()
                            driver.find_element(By.NAME, "hora_prev_man").send_keys("0001")
                            driver.find_element(By.NAME, "relatorio_excel").clear()
                            driver.find_element(By.NAME, "relatorio_excel").send_keys("S")
                            
                            # Envia o formulário
                            login_button = driver.find_element(By.ID, "btn_envia")
                            driver.execute_script("arguments[0].click();", login_button)
                            time.sleep(15)
                            
                            # Registra sucesso
                            self.adicionar_log("Arquivo baixado com sucesso.", nivel='success')
                            
                            # Exclui o último arquivo
                            resultado_exclusao = self.excluir_ultimo_arquivo(download_folder)
                            self.adicionar_log(resultado_exclusao)
                            
                            break  # Sai do loop de tentativas
                            
                        except (TimeoutException, WebDriverException) as e:
                            tentativa += 1
                            erro_msg = f"Erro: {e}. Tentativa {tentativa} de {max_tentativas}"
                            self.adicionar_log(erro_msg, nivel='error')
                            logging.error(erro_msg)
                            
                            if tentativa < max_tentativas:
                                self.adicionar_log("Aguardando 10 minutos antes da próxima tentativa...", nivel='warning')
                                
                                # Espera 10 minutos antes de tentar novamente, verificando a pausa/parada
                                for _ in range(600):  # 600 segundos = 10 minutos
                                    if stop_event.is_set() or pause_event.is_set():
                                        break
                                    time.sleep(1)
            
            except Exception as e:
                erro_msg = f"Erro crítico: {e}"
                self.adicionar_log(erro_msg, nivel='error')
                logging.error(erro_msg)
            
            # Extração concluída ou com erro
            self.atualizar_status('waiting')
            self.adicionar_log("Aguardando 1 hora para próxima execução...", nivel='info')
            self.log_area.insert(tk.END, "-" * 80 + "\n")
            self.log_area.yview(tk.END)
            
            # Espera 1 hora ou até que o evento de parada/pausa seja acionado
            for _ in range(3600):  # 3600 segundos = 1 hora
                if stop_event.is_set():
                    self.atualizar_status('stopped')
                    return
                if pause_event.is_set():
                    break
                time.sleep(1)


if __name__ == "__main__":
    try:
        root = ThemedTk(theme="arc")  # Tema base
        app = Application(root)
        root.mainloop()
    except Exception as e:
        logging.critical(f"Erro fatal na aplicação: {e}")
        messagebox.showerror("Erro Fatal", f"Ocorreu um erro fatal na aplicação: {e}")
