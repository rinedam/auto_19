

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
import sys
import logging
from pathlib import Path
import requests

# Configuração de diretórios
BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"

# Criar diretórios se não existirem
LOGS_DIR.mkdir(exist_ok=True)

# Configuração de logging
LOG_FILE = LOGS_DIR / "extracao.log"
logging.basicConfig(
    filename=LOG_FILE,
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

# Constantes da aplicação
APP_NAME = "Extrator SSW Automático"
APP_VERSION = "2.0"
APP_DESCRIPTION = "Extrai dados do sistema SSW em horário comercial"

# Constantes de estilo
PADDING = {
    'small': 5,
    'medium': 10,
    'large': 20,
    'xlarge': 30
}

# Cores para a interface
COLORS = {
    'primary': '#3498db',
    'primary_dark': '#2980b9',
    'secondary': '#2ecc71',
    'secondary_dark': '#27ae60',
    'background': '#f8f9fa',
    'surface': '#ffffff',
    'text': '#333333',
    'text_light': '#7f8c8d',
    'error': '#e74c3c',
    'warning': '#f39c12',
    'success': '#2ecc71',
    'info': '#3498db',
    'border': '#dcdfe6'
}

# Cores de status
STATUS_COLORS = {
    'idle': '#7f8c8d',       # Cinza para ocioso
    'running': '#3498db',    # Azul para em execução
    'success': '#2ecc71',    # Verde para sucesso
    'warning': '#f39c12',    # Amarelo/laranja para aviso
    'error': '#e74c3c'       # Vermelho para erro
}

# Mapeamento de mensagens de status
STATUS_MESSAGES = {
    'idle': "Sistema ocioso",
    'initializing': "Inicializando extração...",
    'extracting': "Extração em andamento...",
    'waiting': "Aguardando próxima execução...",
    'paused': "Extração pausada",
    'stopped': "Extração interrompida",
    'error': "Erro na extração",
    'success': "Extração concluída com sucesso",
    'outside_hours': "Fora do horário comercial",
    'no_connection': "Sem conexão com o site"
}

# Mapeamento de estilos de status
STATUS_STYLES = {
    'idle': 'StatusIdle.TLabel',
    'initializing': 'StatusRunning.TLabel',
    'extracting': 'StatusRunning.TLabel',
    'waiting': 'StatusSuccess.TLabel',
    'paused': 'StatusWarning.TLabel',
    'stopped': 'StatusError.TLabel',
    'error': 'StatusError.TLabel',
    'success': 'StatusSuccess.TLabel',
    'outside_hours': 'StatusWarning.TLabel',
    'no_connection': 'StatusError.TLabel'
}

# Configurações dos botões
BUTTONS = {
    'start': {
        'text': "Iniciar Extração",
        'style': 'Success.TButton'
    },
    'pause': {
        'text': "Pausar Extração",
        'style': 'Warning.TButton'
    },
    'resume': {
        'text': "Continuar Extração",
        'style': 'Success.TButton'
    },
    'stop': {
        'text': "Parar Extração",
        'style': 'Danger.TButton'
    }
}


# ===== FUNÇÕES AUXILIARES =====

def get_status_style(status_key):
    """Obtém o estilo ttk para uma label de status com base em sua chave"""
    return STATUS_STYLES.get(status_key, 'StatusIdle.TLabel')


def get_status_message(status_key, **kwargs):
    """Obtém a mensagem de status com formatação opcional"""
    message = STATUS_MESSAGES.get(status_key, "Status desconhecido")
    if kwargs:
        return message.format(**kwargs)
    return message


def verificar_conexao(url="https://www.google.com/"):
    try:
        response = requests.get(url, timeout=5)
        return response.status_code == 200
    except requests.ConnectionError:
        logging.error(f"Falha de conexão com {url}")
        return False
    except Exception as e:
        logging.error(f"Erro ao verificar conexão: {e}")
        return False


def excluir_penultimo_arquivo(diretorio, arquivos_ignorados=None):
    if arquivos_ignorados is None:
        arquivos_ignorados = ["desktop.ini"]
        
    try:
        # Verificar se o diretório existe
        if not os.path.exists(diretorio):
            return f"O diretório {diretorio} não existe."
            
        # Obter todos os arquivos, excluindo os ignorados
        arquivos = [
            f for f in os.listdir(diretorio) 
            if os.path.isfile(os.path.join(diretorio, f)) and f not in arquivos_ignorados
        ]
        
        # Verificar se há pelo menos 2 arquivos
        if len(arquivos) < 2:
            return "Não há arquivos suficientes para excluir o penúltimo (precisa de pelo menos 2)."
            
        # Ordenar por data de modificação (mais antigos primeiro)
        arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(diretorio, x)))
        
        # Pegar o penúltimo arquivo (o segundo mais antigo)
        penultimo_arquivo = arquivos[0]
        caminho_completo = os.path.join(diretorio, penultimo_arquivo)
        
        # Excluir o arquivo
        os.remove(caminho_completo)
        logging.info(f"Arquivo excluído: {penultimo_arquivo}")
        return f"Arquivo excluído: {penultimo_arquivo}"
        
    except PermissionError:
        erro_msg = f"Permissão negada ao tentar excluir arquivo em {diretorio}"
        logging.error(erro_msg)
        return erro_msg
    except Exception as e:
        erro_msg = f"Ocorreu um erro ao tentar excluir o arquivo: {e}"
        logging.error(erro_msg)
        return erro_msg


def esta_no_horario_comercial(hora_inicio=8, hora_fim=18):
    """
    Verifica se o momento atual está dentro do horário comercial definido
    
    Args:
        hora_inicio (int): Hora de início do horário comercial (24h)
        hora_fim (int): Hora de término do horário comercial (24h)
        
    Returns:
        bool: True se está no horário comercial, False caso contrário
    """
    agora = datetime.now().time()
    horario_inicio = dt_time(hora_inicio, 0)
    horario_fim = dt_time(hora_fim, 0)
    return horario_inicio <= agora <= horario_fim


def tempo_ate_proxima_extracao(hora_inicio=8, intervalo_horas=1):
    """
    Calcula o tempo restante até a próxima extração
    
    Args:
        hora_inicio (int): Hora de início do horário comercial (24h)
        intervalo_horas (int): Intervalo entre extrações em horas
        
    Returns:
        float: Tempo em segundos até a próxima extração
    """
    agora = datetime.now()
    
    # Se estiver fora do horário comercial (após as 18h), calcular para o próximo dia
    if agora.time() > dt_time(18, 0):
        proximo_dia = agora + timedelta(days=1)
        proxima_extracao = proximo_dia.replace(hour=hora_inicio, minute=0, second=0, microsecond=0)
    else:
        # Se antes do horário comercial, calcular para o início do dia atual
        if agora.time() < dt_time(hora_inicio, 0):
            proxima_extracao = agora.replace(hour=hora_inicio, minute=0, second=0, microsecond=0)
        else:
            # Durante o horário comercial, calcular próximo intervalo
            hora_atual = agora.hour
            # Calcular quantas horas completas passaram desde o início
            horas_passadas = hora_atual - hora_inicio
            # Calcular próximo múltiplo de intervalo_horas
            proxima_hora = hora_inicio + ((horas_passadas // intervalo_horas) + 1) * intervalo_horas
            
            # Se a próxima hora calculada for após o horário comercial
            if proxima_hora >= 18:
                proximo_dia = agora + timedelta(days=1)
                proxima_extracao = proximo_dia.replace(hour=hora_inicio, minute=0, second=0, microsecond=0)
            else:
                proxima_extracao = agora.replace(hour=proxima_hora, minute=0, second=0, microsecond=0)
    
    # Se a próxima extração já passou (devido a atrasos), usar o intervalo padrão
    if proxima_extracao <= agora:
        proxima_extracao = agora + timedelta(hours=intervalo_horas)
    
    tempo_restante = (proxima_extracao - agora).total_seconds()
    return max(tempo_restante, 60)  # Mínimo de 60 segundos


def formatar_tempo_restante(segundos):
    """
    Formata um tempo em segundos para formato legível
    
    Args:
        segundos (float): Tempo em segundos
        
    Returns:
        str: Tempo formatado como "X horas e Y minutos"
    """
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    
    if horas > 0:
        return f"{horas} {'hora' if horas == 1 else 'horas'} e {minutos} {'minuto' if minutos == 1 else 'minutos'}"
    else:
        return f"{minutos} {'minuto' if minutos == 1 else 'minutos'}"


# ===== CLASSE DE TEMA MODERNO =====

class ModernTheme:
    """
    Configurador de tema moderno para Tkinter
    """
    
    @classmethod
    def apply(cls, root):
        """Aplica o tema moderno para toda a aplicação"""
        style = ttk.Style()
        
        # Configura estilos básicos
        style.configure('TFrame', background=COLORS['background'])
        style.configure('Surface.TFrame', background=COLORS['surface'])
        
        style.configure('TLabel', 
                        background=COLORS['background'], 
                        foreground=COLORS['text'], 
                        font=('Helvetica', 11))
        
        style.configure('Title.TLabel', 
                        font=('Helvetica', 20, 'bold'), 
                        foreground=COLORS['primary_dark'])
        
        style.configure('Subtitle.TLabel', 
                        font=('Helvetica', 16, 'bold'), 
                        foreground=COLORS['primary'])
        
        style.configure('Heading.TLabel', 
                        font=('Helvetica', 14, 'bold'), 
                        foreground=COLORS['text'])
        
        # Configura botões - com texto preto para todos os botões
        style.configure('TButton', 
                        font=('Helvetica', 11),
                        background=COLORS['primary'],
                        foreground=COLORS['text'])  # Texto preto
        
        style.map('TButton',
                  background=[('active', COLORS['primary_dark']), 
                              ('disabled', COLORS['text_light'])],
                  foreground=[('disabled', COLORS['text'])])
        
        # Estilo de botão de sucesso
        style.configure('Success.TButton', 
                        background=COLORS['success'], 
                        foreground=COLORS['text'])  # Texto preto
        
        style.map('Success.TButton',
                  background=[('active', COLORS['secondary_dark']), 
                              ('disabled', COLORS['text_light'])],
                  foreground=[('disabled', COLORS['text'])])
        
        # Estilo de botão de aviso
        style.configure('Warning.TButton', 
                        background=COLORS['warning'], 
                        foreground=COLORS['text'])  # Texto preto
        
        style.map('Warning.TButton',
                  background=[('active', '#e67e22'), 
                              ('disabled', COLORS['text_light'])],
                  foreground=[('disabled', COLORS['text'])])
        
        # Estilo de botão de perigo
        style.configure('Danger.TButton', 
                        background=COLORS['error'], 
                        foreground=COLORS['text'])  # Texto preto
        
        style.map('Danger.TButton',
                  background=[('active', '#c0392b'), 
                              ('disabled', COLORS['text_light'])],
                  foreground=[('disabled', COLORS['text'])])
        
        # Estilos de label de status
        style.configure('Status.TLabel', 
                        font=('Helvetica', 14, 'bold'), 
                        foreground=COLORS['text'])
        
        style.configure('StatusIdle.TLabel', 
                        foreground=STATUS_COLORS['idle'])
        
        style.configure('StatusRunning.TLabel', 
                        foreground=STATUS_COLORS['running'])
        
        style.configure('StatusSuccess.TLabel', 
                        foreground=STATUS_COLORS['success'])
        
        style.configure('StatusWarning.TLabel', 
                        foreground=STATUS_COLORS['warning'])
        
        style.configure('StatusError.TLabel', 
                        foreground=STATUS_COLORS['error'])
        
        # Estilo de separador
        style.configure('TSeparator', background=COLORS['border'])
        
        # Define o fundo para toda a aplicação
        root.configure(background=COLORS['background'])
        
        return style


# ===== CLASSE PRINCIPAL DA APLICAÇÃO =====

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
            text="Status: Sistema ocioso", 
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
        
        # Data da última execução bem-sucedida
        self.last_run_label = ttk.Label(
            self.footer_frame,
            text="Última execução: Nunca",
            font=("Helvetica", 8)
        )
        self.last_run_label.pack(side=tk.LEFT)
        
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
    
    def atualizar_ultima_execucao(self):
        """Atualiza a label de última execução"""
        agora = datetime.now().strftime('%d/%m/%Y %H:%M')
        self.last_run_label.config(text=f"Última execução: {agora}")
    
    def mostrar_mensagem_erro(self, titulo, mensagem):
        """Exibe uma caixa de diálogo de erro"""
        messagebox.showerror(titulo, mensagem)
        self.adicionar_log(mensagem, nivel='error')
    
    def executar_extracao(self):
        """Executa a extração de dados"""
        # Caminho para o diretório de download
        download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\CTRCs Disponíveis para Transferência - 19\\DB_CTRCs Disponíveis para Transferência\\CTA')
        
        # Verificar se o diretório de download existe
        if not os.path.exists(download_folder):
            mensagem = f"Diretório de download não encontrado: {download_folder}"
            self.mostrar_mensagem_erro("Erro de Diretório", mensagem)
            self.parar_extracao()
            return
        
        # Configurações do Edge
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
            if not esta_no_horario_comercial():
                tempo_restante = tempo_ate_proxima_extracao()
                tempo_formatado = formatar_tempo_restante(tempo_restante)
                
                self.atualizar_status('outside_hours')
                self.adicionar_log(
                    f"Fora do horário comercial. Próxima extração em {tempo_formatado}.",
                    nivel='warning'
                )
                
                # Aguarda até a próxima extração, verificando a cada segundo se foi pausado ou parado
                for _ in range(int(min(tempo_restante, 3600))):  # Verifica de novo após no máximo 1 hora
                    if stop_event.is_set() or pause_event.is_set():
                        break
                    time.sleep(1)
                
                continue
            
            # Verifica conexão com o site
            if not verificar_conexao():
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
            
            # Inicia o processo de extração
            self.atualizar_status('extracting')
            self.adicionar_log("Iniciando processo de extração de dados...", nivel='info')
            
            try:
                with webdriver.Edge(options=edge_options) as driver:
                    max_tentativas = 3
                    tentativa = 0
                    
                    while tentativa < max_tentativas and not stop_event.is_set():
                        try:
                            # Abre o site do SSW
                            driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
                            
                            # Aguarda a página de login
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
                            
                            # Aguarda a próxima tela e preenche
                            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "f2")))
                            driver.find_element(By.NAME, "f2").clear()
                            driver.find_element(By.NAME, "f2").send_keys("CTA")
                            driver.find_element(By.NAME, "f3").send_keys("19+")
                            time.sleep(5)
                            
                            # Troca para a nova aba aberta
                            abas = driver.window_handles
                            driver.switch_to.window(abas[-1])
                            
                            # Preenche a data e configurações
                            data_atual = datetime.now().strftime('%d%m%y')
                            driver.find_element(By.NAME, "data_prev_man").clear()
                            driver.find_element(By.NAME, "data_prev_man").send_keys(data_atual)
                            driver.find_element(By.NAME, "hora_prev_man").clear()
                            driver.find_element(By.NAME, "hora_prev_man").send_keys("0001")
                            driver.find_element(By.NAME, "relatorio_excel").clear()
                            driver.find_element(By.NAME, "relatorio_excel").send_keys("S")
                            
                            # Clica no botão para enviar
                            login_button = driver.find_element(By.ID, "btn_envia")
                            driver.execute_script("arguments[0].click();", login_button)
                            
                            # Aguarda o download
                            time.sleep(30)
                            
                            # Registra sucesso no log
                            self.adicionar_log("Arquivo baixado com sucesso.", nivel='success')
                            
                            # Exclui o último arquivo (caso necessário)
                            resultado = excluir_penultimo_arquivo(download_folder)
                            self.adicionar_log(resultado, nivel='info')
                            
                            # Atualiza status e tempo da última execução
                            self.atualizar_status('success')
                            self.atualizar_ultima_execucao()
                            
                            # Sai do loop de tentativas
                            break
                            
                        except (TimeoutException, WebDriverException) as e:
                            # Registra erro no log
                            erro_msg = f"Erro: {e}. Tentativa {tentativa + 1} de {max_tentativas}"
                            logging.error(erro_msg)
                            self.adicionar_log(erro_msg, nivel='error')
                            
                            # Incrementa contador de tentativas
                            tentativa += 1
                            
                            # Aguarda antes de tentar novamente (10 min)
                            if tentativa < max_tentativas:
                                self.adicionar_log("Aguardando 10 minutos antes de tentar novamente...", nivel='warning')
                                
                                for _ in range(600):  # 10 minutos
                                    if stop_event.is_set() or pause_event.is_set():
                                        break
                                    time.sleep(1)
                            else:
                                self.adicionar_log("Número máximo de tentativas excedido.", nivel='error')
                                
            except Exception as e:
                # Registra erro crítico no log
                erro_msg = f"Erro crítico: {e}"
                logging.error(erro_msg)
                self.adicionar_log(erro_msg, nivel='error')
            
            # Configura para aguardar próxima execução
            self.atualizar_status('waiting')
            self.adicionar_log("Aguardando 1 hora para próxima execução...", nivel='info')
            self.adicionar_log("-" * 80, nivel='info')
            
            # Espera 1 hora ou até que o evento de parada seja acionado
            for _ in range(3600):  # 3600 segundos = 1 hora
                if stop_event.is_set() or pause_event.is_set():
                    break
                time.sleep(1)

def main():
    """Função principal que inicia a aplicação"""
    # Usar o tema ThemedTk
    root = ThemedTk(theme="arc")
    app = Application(root)
    root.mainloop()

if __name__ == "__main__":
    main()