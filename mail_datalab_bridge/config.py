from pathlib import Path


# ARM: ящик, открытый в настольном Outlook.
TECH_MAIL_FOLDER = "TEST MAIL"  # имя корня ящика в дереве Outlook
TECH_MAIL_SMTP = "technical@company.ru"

# Ящик-посредник, доступный из VARM по POP3/SMTP.
RELAY_MAIL_SMTP = "relay@company.ru"

# Обычное обращение пользователя обрабатывается, если эта фраза есть в теме.
REQUEST_SUBJECT_PHRASE = "вопрос помощнику"

# Служебные письма ARM <-> VARM. Не используйте эту фразу для обычных писем.
BRIDGE_SUBJECT = "json_handlers"

# Имена стандартных папок могут отличаться в другой локализации Outlook.
INBOX_FOLDER = "Входящие"
SENT_FOLDER = "Отправленные"

SCAN_LIMIT = 100
POLL_SECONDS = 20

# Если процесс завершился точно во время Send(), мост сначала ищет письмо в
# «Отправленных» и только затем разрешает повторную отправку.
SEND_RECOVERY_SECONDS = 300

STATE_FILE = Path(__file__).with_name("state.json")
LOG_FILE = Path(__file__).with_name("mail_agent.log")
