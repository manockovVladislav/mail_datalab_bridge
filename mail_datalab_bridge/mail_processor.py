import json
import logging
import os
import time
from pathlib import Path


class MailProcessor:
    def __init__(self, outlook, datalab, config):
        self.outlook = outlook
        self.datalab = datalab
        self.cfg = config
        self.state = self._load()
        logging.info("Загружено записей состояния: %s", len(self.state))

    def _load(self):
        path = self.cfg.STATE_FILE
        if not path.exists():
            return {}
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(
                f"Не удалось прочитать файл состояния {path}: {exc}"
            ) from exc
        self._validate_state(state, path)
        return state

    @staticmethod
    def _validate_state(state, path):
        if not isinstance(state, dict):
            raise RuntimeError(f"Файл состояния {path} должен содержать JSON-объект")
        for uid, record in state.items():
            if not isinstance(uid, str) or not isinstance(record, dict):
                raise RuntimeError(f"Неверная запись в файле состояния {path}")
            if record.get("status") not in {"waiting", "sending", "sent"}:
                raise RuntimeError(f"Неверный статус запроса {uid} в {path}")
            if not isinstance(record.get("email"), str) or not isinstance(
                record.get("subject"), str
            ):
                raise RuntimeError(f"В записи {uid} нет строк email/subject")

    def _save(self):
        path = self.cfg.STATE_FILE
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, path)

    def scan(self):
        logging.info("Проверка последних %s входящих писем", self.cfg.SCAN_LIMIT)
        matches = self.outlook.recent_requests(self.cfg.SUBJECT_PHRASE, self.cfg.SCAN_LIMIT)
        logging.info("Найдено подходящих писем: %s", len(matches))
        for mail in matches:
            uid = mail["id"]
            if uid in self.state:
                logging.info(
                    "Письмо %s уже учтено, пропуск (статус: %s)",
                    uid,
                    self.state[uid]["status"],
                )
                continue
            remote_path = f"{self.cfg.REMOTE_INPUT_DIR}/request_{uid}.txt"
            # JSON внутри .txt сохраняет переносы строк, кавычки и все поля без неоднозначности.
            content = json.dumps({"id": uid, "email": mail["email"],
                                  "subject": mail["subject"], "body": mail["body"]},
                                 ensure_ascii=False, indent=2)
            logging.info("Загрузка запроса %s -> %s", uid, remote_path)
            self.datalab.upload_text(remote_path, content)
            self.state[uid] = {"status": "waiting", "email": mail["email"],
                               "subject": mail["subject"]}
            self._save()
            logging.info("Запрос %s загружен, ожидаем ответ", uid)

    def check_responses(self):
        logging.info("Проверка готовых ответов DataLab")
        for uid, record in list(self.state.items()):
            if record["status"] not in ("waiting", "sending"):
                continue
            if record["status"] == "sending":
                logging.warning(
                    "Запрос %s был в процессе отправки. Проверьте "
                    "'Отправленные' вручную; автоматический повтор отключён",
                    uid,
                )
                continue
            path = f"{self.cfg.REMOTE_OUTPUT_DIR}/response_{uid}.txt"
            logging.info("Проверяем файл ответа %s", path)
            response = self.datalab.read_text_if_exists(path)
            if response is None:
                logging.info("Ответ %s ещё не готов", uid)
                continue
            logging.info("Получен ответ %s (%s символов)", uid, len(response))
            # Сохраняем статус до Send(), чтобы после сбоя не отправить дубль автоматически.
            record["status"] = "sending"
            self._save()
            logging.info("Передача ответа %s в Outlook для отправки", uid)
            self.outlook.send(record["email"], record["subject"], response)
            record["status"] = "sent"
            self._save()
            logging.info("Ответ %s передан Outlook на отправку", uid)

    def run(self):
        while True:
            logging.info("--- Начало цикла ---")
            try:
                self.scan()
                self.check_responses()
                logging.info("--- Цикл завершён ---")
            except Exception:
                logging.exception(
                    "Ошибка цикла. Следующая попытка через %s сек",
                    self.cfg.POLL_SECONDS,
                )
            logging.info("Ожидание %s секунд", self.cfg.POLL_SECONDS)
            time.sleep(self.cfg.POLL_SECONDS)
