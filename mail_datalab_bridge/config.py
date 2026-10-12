from pathlib import Path

# Заполнить своими значениями. Не храните пароль в исходном коде.
BASE_URL = "https://jupyterhub-datalab.orf.ru"
FULL_USERNAME = ""  # логин DataLab, используемый в /user/<логин>/lab
BROWSER_PATH = r"C:\Program Files (x86)\Browser\Application\browser.exe"
DRIVER_PATH = r"C:\tools\chromedriver.exe"
TECH_MAIL_FOLDER = "TEST MAIL"  # точное имя дерева папок слева в Outlook
TECH_MAIL_SMTP = "technical@company.ru"  # SMTP-адрес технического ящика
SUBJECT_PHRASE = "вопрос помощнику"
REMOTE_INPUT_DIR = "email_test/input"
REMOTE_OUTPUT_DIR = "email_test/output"
SCAN_LIMIT = 20
POLL_SECONDS = 15
LOGIN_RETRIES = 3
STATE_FILE = Path(__file__).with_name("state.json")
