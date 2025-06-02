# Auto_19 - Extrator SSW Automático

## Visão Geral

O Auto_19 é uma aplicação desktop desenvolvida em Python que automatiza o processo de extração de dados do sistema SSW (Sistema de Soluções Web). Este software realiza extrações periódicas de relatórios durante o horário comercial, utilizando uma interface gráfica moderna e intuitiva que permite ao usuário controlar todo o processo de extração.

A aplicação combina a biblioteca Tkinter para a interface gráfica com o Selenium WebDriver para automação de navegador, permitindo o login automático no sistema SSW, navegação entre telas, preenchimento de formulários e download de relatórios sem intervenção manual do usuário.

## Principais Funcionalidades

### Interface Gráfica Moderna

- **Design Responsivo**: Interface construída com ttk e ttkthemes, oferecendo um visual moderno e agradável
- **Tema Personalizado**: Utiliza o tema "arc" do ThemedTk com cores e estilos personalizados
- **Painel de Status**: Exibe o estado atual da aplicação com códigos de cores intuitivos
- **Console de Logs**: Área de texto rolável que exibe em tempo real todas as operações realizadas
- **Controles Intuitivos**: Botões para iniciar, pausar e parar o processo de extração

### Automação de Navegador

- **Login Automático**: Realiza login no sistema SSW utilizando credenciais armazenadas em arquivo .env
- **Navegação entre Telas**: Navega automaticamente entre as diferentes telas do sistema
- **Preenchimento de Formulários**: Preenche automaticamente os campos necessários para extração
- **Download de Relatórios**: Realiza o download dos relatórios em formato Excel
- **Tratamento de Erros**: Sistema de tentativas múltiplas em caso de falha na extração

### Controle de Execução

- **Execução Periódica**: Realiza extrações em intervalos regulares durante o horário comercial
- **Verificação de Horário**: Executa apenas durante o horário comercial configurado (padrão: 7h às 18h)
- **Sistema de Pausa**: Permite pausar e retomar a extração a qualquer momento
- **Verificação de Conexão**: Verifica a conectividade com a internet antes de tentar extrações

### Registro e Monitoramento

- **Logs Detalhados**: Registra todas as operações em arquivo de log e na interface
- **Níveis de Log**: Diferentes níveis de log (info, warning, error, success) com formatação visual
- **Histórico de Execuções**: Mantém registro da última execução bem-sucedida
- **Gerenciamento de Arquivos**: Opção para excluir arquivos antigos automaticamente

## Funcionamento

Ao iniciar a aplicação, o usuário é apresentado a uma interface gráfica com botões de controle para iniciar, pausar e parar o processo de extração. Quando iniciado, o sistema:

1. Verifica se está dentro do horário comercial (7h às 18h por padrão)
2. Confirma a conexão com a internet
3. Abre o navegador Edge automaticamente
4. Realiza login no sistema SSW usando credenciais armazenadas
5. Navega até o relatório CTA 19
6. Configura os parâmetros necessários (data atual, formato Excel)
7. Baixa o relatório para a pasta de downloads
8. Gerencia os arquivos baixados, excluindo versões antigas quando necessário
9. Aguarda o próximo intervalo de extração (1 hora por padrão)

Todo o processo é exibido em tempo real na área de logs da interface, permitindo ao usuário acompanhar cada etapa da extração. O sistema também é capaz de lidar com erros, realizando até três tentativas em caso de falha na extração.

## Integração

O software pode ser integrado a diversas outras aplicações, como o Power BI, possibilitando a extração automática de dados do sistema para a criação de dashboards, o que facilita a gestão visual, otimiza a análise de informações e apoia a tomada de decisões.

## Estrutura do Projeto

- **auto_19.py**: Arquivo principal contendo toda a lógica da aplicação
- **logs/**: Diretório onde são armazenados os arquivos de log
- **build/**: Diretório contendo arquivos de compilação
- **dist/**: Diretório contendo o executável compilado
- **favicon.ico**: Ícone da aplicação
