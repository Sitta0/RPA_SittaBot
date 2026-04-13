# RPA_SittaBot
Automação de download de relatórios CSV com interface gráfica e suporte a MFA

SittaBot Reports
Bot de automação para exportação e download de relatórios CSV em sistemas web com autenticação SSO Microsoft. Roda em background com uma interface gráfica que mostra o progresso em tempo real.

O que ele faz
Faz login via SSO, navega até o relatório, solicita a exportação, aguarda o processamento e baixa o arquivo CSV — sem precisar de interação manual. Se a conta usar MFA ou se a senha expirar, ele trata os dois casos automaticamente com popups de suporte.

Requisitos

Python 3.9+
Windows 10 ou 11
Microsoft Edge instalado


Instalação
bashgit clone https://github.com/seu-usuario/sittabot-reports.git
cd sittabot-reports

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium

Configuração
bashcopy config.example.ini config.ini
Abra o config.ini e preencha:
ini[credenciais]
email = seu_email@empresa.com
senha = sua_senha

[sistema]
url_login       = https://seudominio.com/#/login
url_parcelas    = https://seudominio.com/#/relatorio/parcelas
url_emitidos    = https://seudominio.com/#/relatorio/meus-relatorios/emitidos
sso_button_text = SSO Login
nome_relatorio  = Relatório
# edge_path     = C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
O config.ini está no .gitignore e nunca deve ser versionado.

Uso
bashpython bot.py
A janela abre e o processo segue quatro etapas: login SSO, solicitação do relatório, aguardo do processamento e download do CSV. O arquivo é salvo na mesma pasta do script com o nome relatorio_parcelas_YYYY_MM_DD.csv.

MFA e senha expirada
Quando o login exige aprovação no Microsoft Authenticator, o bot exibe o número de aprovação em um popup — basta abrir o app no celular e confirmar.
Se a senha da conta expirar, um popup permite informar a nova senha sem cancelar o processo. O config.ini é atualizado automaticamente e o login é retentado.

Gerar um .exe
bashpip install pyinstaller
pyinstaller --onefile --windowed --name SittaBot bot.py
O executável fica em dist/SittaBot.exe. Para distribuir, inclua o config.example.ini junto e peça para renomear para config.ini antes de rodar.

Estrutura
sittabot-reports/
├── bot.py
├── config.example.ini
├── requirements.txt
├── .gitignore
└── README.md
Arquivos gerados localmente (não versionados):
config.ini                    # credenciais — nunca commitar
erro.log
relatorio_parcelas_*.csv
sittabot_edge_profile/

Licença
MIT
