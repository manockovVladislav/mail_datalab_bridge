import logging

import config
from mail_processor import MailProcessor
from outlook_client import OutlookClient


def validate_config():
    required = {
        "TECH_MAIL_FOLDER": config.TECH_MAIL_FOLDER,
        "TECH_MAIL_SMTP": config.TECH_MAIL_SMTP,
        "RELAY_MAIL_SMTP": config.RELAY_MAIL_SMTP,
        "REQUEST_SUBJECT_PHRASE": config.REQUEST_SUBJECT_PHRASE,
        "BRIDGE_SUBJECT": config.BRIDGE_SUBJECT,
    }
    missing = [name for name, value in required.items() if not str(value).strip()]
    if missing:
        raise ValueError("Заполните параметры config.py: " + ", ".join(missing))

    for name in ("SCAN_LIMIT", "POLL_SECONDS", "SEND_RECOVERY_SECONDS"):
        value = getattr(config, name)
        if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
            raise ValueError(f"{name} должен быть положительным целым числом")

    for name in ("TECH_MAIL_SMTP", "RELAY_MAIL_SMTP"):
        if "@" not in getattr(config, name):
            raise ValueError(f"{name} должен содержать почтовый адрес")


def configure_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(config.LOG_FILE, encoding="utf-8"),
        ],
        force=True,
    )


def main():
    configure_logging()
    validate_config()
    logging.info("Запуск ARM-почтового моста")

    outlook = OutlookClient(
        folder_name=config.TECH_MAIL_FOLDER,
        sender_smtp=config.TECH_MAIL_SMTP,
        inbox_name=config.INBOX_FOLDER,
        sent_name=config.SENT_FOLDER,
    )
    processor = MailProcessor(outlook, config)
    try:
        processor.run()
    except KeyboardInterrupt:
        logging.info("Остановлено пользователем")


if __name__ == "__main__":
    main()
