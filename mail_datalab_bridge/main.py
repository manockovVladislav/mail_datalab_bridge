import getpass
import logging
from urllib.parse import urlparse

import config
from datalab_client import DataLabClient
from mail_processor import MailProcessor
from outlook_client import OutlookClient


def validate_config():
    required = {
        "BASE_URL": config.BASE_URL,
        "FULL_USERNAME": config.FULL_USERNAME,
        "BROWSER_PATH": config.BROWSER_PATH,
        "DRIVER_PATH": config.DRIVER_PATH,
        "TECH_MAIL_FOLDER": config.TECH_MAIL_FOLDER,
        "TECH_MAIL_SMTP": config.TECH_MAIL_SMTP,
        "SUBJECT_PHRASE": config.SUBJECT_PHRASE,
        "REMOTE_INPUT_DIR": config.REMOTE_INPUT_DIR,
        "REMOTE_OUTPUT_DIR": config.REMOTE_OUTPUT_DIR,
    }
    missing = [name for name, value in required.items() if not str(value).strip()]
    if missing:
        raise ValueError("Заполните параметры config.py: " + ", ".join(missing))

    parsed_url = urlparse(config.BASE_URL)
    if parsed_url.scheme not in ("http", "https") or not parsed_url.netloc:
        raise ValueError("BASE_URL должен быть полным HTTP(S)-адресом")
    if "@" not in config.TECH_MAIL_SMTP:
        raise ValueError("TECH_MAIL_SMTP должен содержать SMTP-адрес")

    for name in ("SCAN_LIMIT", "POLL_SECONDS", "LOGIN_RETRIES"):
        value = getattr(config, name)
        if not isinstance(value, int) or isinstance(value, bool) or value <= 0:
            raise ValueError(f"{name} должен быть положительным целым числом")


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("mail_agent.log", encoding="utf-8"),
        ],
        force=True,
    )
    logging.info("=== Запуск почтового агента ===")
    validate_config()
    logging.info("Запрос пароля DataLab")
    password = getpass.getpass("Пароль DataLab: ")
    datalab = DataLabClient(config.BASE_URL, config.FULL_USERNAME,
                            config.BROWSER_PATH, config.DRIVER_PATH)
    try:
        logging.info("Подключение к DataLab")
        datalab.connect(password, config.LOGIN_RETRIES)
        logging.info("Подключение к Outlook")
        outlook = OutlookClient(config.TECH_MAIL_FOLDER, config.TECH_MAIL_SMTP)
        logging.info("Запуск цикла обработки")
        MailProcessor(outlook, datalab, config).run()
    except KeyboardInterrupt:
        logging.info("Остановлено пользователем")
    finally:
        logging.info("Закрытие соединений")
        datalab.close()


if __name__ == "__main__":
    main()
