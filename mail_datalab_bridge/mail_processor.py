import json
import logging
import os
import re
import time


STATUSES = {
    "new",
    "forwarding",
    "forwarded",
    "answer_ready",
    "replying",
    "replied",
    "cleaning",
    "completed",
}

SIGNATURE_LINE = re.compile(
    r"^\s*(?:--\s*|с\s+уважением\b.*|с\s+наилучшими\s+пожеланиями\b.*|"
    r"best\s+regards\b.*|kind\s+regards\b.*|sent\s+from\s+my\b.*)\s*$",
    re.IGNORECASE,
)
REPLY_SEPARATOR = re.compile(
    r"^\s*-{2,}\s*(?:original message|исходное сообщение)\s*-*\s*$",
    re.IGNORECASE,
)
REQUEST_ID = re.compile(r"^[0-9a-f]{24}$")


def clean_question(body):
    """Убирает подпись и процитированную переписку, сохраняя абзацы."""
    cleaned = []
    for line in str(body or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if line.lstrip().startswith(">"):
            continue
        if SIGNATURE_LINE.match(line) or REPLY_SEPARATOR.match(line):
            break
        cleaned.append(line.rstrip())

    # Не превращаем несколько вопросов в одну строку. Убираем лишь лишние
    # пустые строки по краям и серии из трёх и более пустых строк.
    text = "\n".join(cleaned).strip()
    return re.sub(r"\n{3,}", "\n\n", text)


class MailProcessor:
    def __init__(self, outlook, config):
        self.outlook = outlook
        self.cfg = config
        self.state = self._load()
        logging.info("Загружено запросов из состояния: %s", len(self.state))

    def _load(self):
        path = self.cfg.STATE_FILE
        if not path.exists():
            return {}
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Не удалось прочитать {path}: {exc}") from exc
        self._validate_state(state, path)
        return state

    @staticmethod
    def _validate_state(state, path):
        if not isinstance(state, dict):
            raise RuntimeError(f"Файл состояния {path} должен содержать JSON-объект")
        for request_id, record in state.items():
            if not isinstance(request_id, str) or not isinstance(record, dict):
                raise RuntimeError(f"Неверная запись в {path}")
            if record.get("status") not in STATUSES:
                raise RuntimeError(f"Неверный статус запроса {request_id} в {path}")
            for field in ("email", "subject", "question", "relay_subject"):
                if not isinstance(record.get(field), str):
                    raise RuntimeError(
                        f"Поле {field} запроса {request_id} должно быть строкой"
                    )

    def _save(self):
        path = self.cfg.STATE_FILE
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(path.suffix + ".tmp")
        tmp.write_text(
            json.dumps(self.state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(tmp, path)

    def _is_request_subject(self, subject):
        return self.cfg.REQUEST_SUBJECT_PHRASE.casefold() in subject.casefold()

    def _is_response_subject(self, subject):
        prefix = f"{self.cfg.BRIDGE_SUBJECT} response ".casefold()
        return subject.casefold().startswith(prefix)

    def _register_request(self, message):
        request_id = message["id"]
        if request_id in self.state:
            return

        question = clean_question(message["body"])
        if not question:
            logging.warning("Письмо %s не содержит вопроса после очистки", request_id)
            return

        relay_subject = f"{self.cfg.BRIDGE_SUBJECT} request {request_id}"
        self.state[request_id] = {
            "status": "new",
            "email": message["email"],
            "subject": message["subject"],
            "question": question,
            "relay_subject": relay_subject,
            "source_entry_id": message["entry_id"],
            "source_store_id": message["store_id"],
            "response_messages": [],
        }
        self._save()
        logging.info("Зарегистрирован новый запрос %s", request_id)

    def _register_response(self, message):
        if message["email"].casefold() != self.cfg.RELAY_MAIL_SMTP.casefold():
            logging.warning(
                "Служебное письмо отклонено: неожиданный отправитель %s",
                message["email"],
            )
            return

        try:
            payload = json.loads(message["body"])
        except (TypeError, json.JSONDecodeError):
            logging.warning("В служебном письме нет корректного JSON")
            return

        if (
            not isinstance(payload, dict)
            or payload.get("protocol") != 1
            or payload.get("type") != "response"
        ):
            logging.warning("Получен неизвестный тип служебного сообщения")
            return
        request_id = payload.get("request_id")
        answer = payload.get("answer")
        if not isinstance(request_id, str) or not REQUEST_ID.fullmatch(request_id):
            logging.warning("В ответе нет корректного request_id")
            return
        expected_subject = f"{self.cfg.BRIDGE_SUBJECT} response {request_id}"
        if message["subject"].casefold() != expected_subject.casefold():
            logging.warning("request_id в теме и JSON не совпадают")
            return
        if not isinstance(answer, str) or not answer.strip():
            logging.warning("Ответ %s пуст", request_id)
            return

        record = self.state.get(request_id)
        if record is None:
            logging.warning("Ответ относится к неизвестному запросу %s", request_id)
            return
        if record["status"] == "completed":
            # VARM использует доставку как минимум один раз. Поздний дубль
            # служебного ответа можно удалить: пользовательский ответ уже ушёл.
            self.outlook.delete_message(message["entry_id"], message["store_id"])
            return

        message_ref = {
            "entry_id": message["entry_id"],
            "store_id": message["store_id"],
        }
        refs = record.setdefault("response_messages", [])
        if message_ref not in refs:
            refs.append(message_ref)

        if record["status"] in {"new", "forwarding", "forwarded", "answer_ready"}:
            record["answer"] = answer.strip()
            record["status"] = "answer_ready"
            logging.info("Получен ответ VARM для запроса %s", request_id)
        self._save()

    def scan(self):
        messages = self.outlook.incoming(self.cfg.SCAN_LIMIT)
        for message in messages:
            subject = message["subject"]
            if self._is_response_subject(subject):
                self._register_response(message)
            elif self._is_request_subject(subject):
                self._register_request(message)

    def _forward(self, request_id, record):
        payload = {
            "protocol": 1,
            "type": "request",
            "request_id": request_id,
            "reply_to": record["email"],
            "original_subject": record["subject"],
            "question": record["question"],
        }
        record["status"] = "forwarding"
        record["send_started_at"] = time.time()
        self._save()
        self.outlook.send_json(
            self.cfg.RELAY_MAIL_SMTP,
            record["relay_subject"],
            payload,
        )
        record["status"] = "forwarded"
        record.pop("send_started_at", None)
        self._save()
        logging.info("Запрос %s отправлен в relay-ящик", request_id)

    def _recover_forwarding(self, request_id, record):
        if self.outlook.sent_message_exists(record["relay_subject"]):
            record["status"] = "forwarded"
            record.pop("send_started_at", None)
            self._save()
            logging.info("Отправка запроса %s подтверждена по папке Отправленные", request_id)
            return
        if self._recovery_timeout_passed(record):
            logging.warning("Запрос %s не найден в Отправленных, повторяем", request_id)
            record["status"] = "new"
            record.pop("send_started_at", None)
            self._save()

    def _reply(self, request_id, record):
        record["status"] = "replying"
        record["send_started_at"] = time.time()
        self._save()
        self.outlook.send_answer(
            to=record["email"],
            original_subject=record["subject"],
            question=record["question"],
            answer=record["answer"],
            request_id=request_id,
        )
        # Статус остаётся replying до появления письма в «Отправленных».
        # Только это подтверждение разрешает очистку служебной цепочки.
        logging.info("Ответ по запросу %s передан Outlook; ждём подтверждение", request_id)

    def _recover_replying(self, request_id, record):
        subject = f"Re: {record['subject']}"
        if self.outlook.sent_message_exists(subject, marker=request_id):
            record["status"] = "replied"
            record.pop("send_started_at", None)
            self._save()
            logging.info("Ответ %s подтверждён по папке Отправленные", request_id)
            return
        if self._recovery_timeout_passed(record):
            logging.warning("Ответ %s не найден в Отправленных, повторяем", request_id)
            record["status"] = "answer_ready"
            record.pop("send_started_at", None)
            self._save()

    def _recovery_timeout_passed(self, record):
        started = record.get("send_started_at")
        if not isinstance(started, (int, float)):
            return True
        return time.time() - started >= self.cfg.SEND_RECOVERY_SECONDS

    def _cleanup(self, request_id, record):
        record["status"] = "cleaning"
        self._save()

        self.outlook.delete_message(
            record.get("source_entry_id"), record.get("source_store_id")
        )
        for response in record.get("response_messages", []):
            self.outlook.delete_message(
                response.get("entry_id"), response.get("store_id")
            )
        self.outlook.delete_sent_by_subject(record["relay_subject"])

        record["status"] = "completed"
        record.pop("answer", None)
        record.pop("send_started_at", None)
        self._save()
        logging.info("Цепочка %s завершена и очищена", request_id)

    def process_pending(self):
        for request_id, record in list(self.state.items()):
            status = record["status"]
            if status == "new":
                self._forward(request_id, record)
            elif status == "forwarding":
                self._recover_forwarding(request_id, record)
            elif status == "answer_ready":
                self._reply(request_id, record)
            elif status == "replying":
                self._recover_replying(request_id, record)
            elif status in {"replied", "cleaning"}:
                self._cleanup(request_id, record)

    def cycle(self):
        # Сохранённые операции продолжаются даже до нового чтения входящих.
        # После сканирования второй проход сразу запускает только что найденное.
        self.process_pending()
        self.scan()
        self.process_pending()

    def run(self):
        while True:
            try:
                self.cycle()
            except Exception:
                logging.exception(
                    "Ошибка цикла; повтор через %s секунд", self.cfg.POLL_SECONDS
                )
            time.sleep(self.cfg.POLL_SECONDS)
