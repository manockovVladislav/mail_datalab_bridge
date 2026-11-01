import inspect
import json
import logging
import os
import poplib
import re
import smtplib
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import parseaddr
from html.parser import HTMLParser
from pathlib import Path


logger = logging.getLogger(__name__)
REQUEST_ID = re.compile(r"^[0-9a-f]{24}$")


class _HTMLText(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)


class EmailConfig:
    """Сохраняет прежнюю инициализацию и точку входа обработчика."""

    def __init__(self, email, username, password, codeword="Помощнику:", state_file=None):
        # Ваши существующие почтовые настройки оставлены внутри класса.
        self.POP3_SERVER = "pop..ru"
        self.POP3_PORT = 1
        self.SMTP_SERVER = "smtp..ru"
        self.SMTP_PORT = 11
        self.EMAIL = email
        self.USERNAME = username
        self.PASSWORD = password
        self.CODEWORD = codeword
        self.REPLY_SUBJECT = f"Ответ {codeword} на вопрос: "

        self.CHECK_INTERVAL = 20
        self.MODEL_REQUEST_DELAY = 5
        self.MAX_MESSAGES = 20

        # Служебная тема новой цепочки ARM <-> VARM.
        self.BRIDGE_SUBJECT = "json_handlers"
        self.STATE_FILE = Path(state_file or Path(__file__).with_name("state.json"))
        self.state = self._load_state()

    def _load_state(self):
        if not self.STATE_FILE.exists():
            return {}
        try:
            state = json.loads(self.STATE_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"Не удалось прочитать {self.STATE_FILE}: {exc}") from exc
        if not isinstance(state, dict):
            raise RuntimeError(f"{self.STATE_FILE} должен содержать JSON-объект")
        return state

    def _save_state(self):
        self.STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.STATE_FILE.with_suffix(self.STATE_FILE.suffix + ".tmp")
        temporary.write_text(
            json.dumps(self.state, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temporary, self.STATE_FILE)

    def _extract_text_only(self, message):
        """Извлекает текст JSON-письма без вложений."""
        plain = None
        html_body = None
        parts = message.walk() if message.is_multipart() else (message,)
        for part in parts:
            if part.get_content_maintype() == "multipart":
                continue
            if part.get_content_disposition() == "attachment":
                continue
            content_type = part.get_content_type()
            if content_type not in {"text/plain", "text/html"}:
                continue
            try:
                content = str(part.get_content())
            except Exception:
                raw = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                content = raw.decode(charset, errors="replace")
            if content_type == "text/plain" and plain is None:
                plain = content
            elif content_type == "text/html" and html_body is None:
                html_body = content

        if plain is not None:
            return plain.strip()
        if html_body is not None:
            parser = _HTMLText()
            parser.feed(html_body)
            return " ".join(" ".join(parser.parts).split())
        return ""

    def _read_requests(self, pop_connection):
        _, message_list, _ = pop_connection.list()
        for message_info in message_list[-self.MAX_MESSAGES :]:
            message_number = int(message_info.split()[0])
            _, lines, _ = pop_connection.retr(message_number)
            message = BytesParser(policy=policy.default).parsebytes(b"\n".join(lines))

            subject = str(message.get("Subject", ""))
            prefix = f"{self.BRIDGE_SUBJECT} request "
            if not subject.casefold().startswith(prefix.casefold()):
                continue

            try:
                payload = json.loads(self._extract_text_only(message))
            except json.JSONDecodeError:
                continue
            if not isinstance(payload, dict):
                continue

            request_id = payload.get("request_id")
            question = payload.get("question")
            expected_subject = f"{prefix}{request_id}"
            if (
                payload.get("protocol") != 1
                or payload.get("type") != "request"
                or not isinstance(request_id, str)
                or not REQUEST_ID.fullmatch(request_id)
                or subject.casefold() != expected_subject.casefold()
                or not isinstance(question, str)
                or not question.strip()
            ):
                continue

            return_to = parseaddr(str(message.get("From", "")))[1]
            if return_to:
                yield request_id, question.strip(), return_to

    @staticmethod
    def _call_answer(answer_function, question, request_id):
        """Поддерживает старый answer_function(question) и вариант с request_id."""
        try:
            inspect.signature(answer_function).bind(question, request_id)
        except (TypeError, ValueError):
            return answer_function(question)
        return answer_function(question, request_id)

    def _send_json_response(self, to, request_id, question, answer):
        payload = {
            "protocol": 1,
            "type": "response",
            "request_id": request_id,
            "question": question,
            "answer": answer,
        }
        message = EmailMessage()
        message["Subject"] = f"{self.BRIDGE_SUBJECT} response {request_id}"
        message["From"] = self.EMAIL
        message["To"] = to
        message.set_content(json.dumps(payload, ensure_ascii=False, indent=2))

        with smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT) as server:
            server.starttls()
            server.login(self.USERNAME, self.PASSWORD)
            refused = server.send_message(message)
            if refused:
                raise RuntimeError(f"SMTP отклонил получателей: {refused}")

    def _process_pending(self, answer_function):
        for request_id, record in list(self.state.items()):
            if record.get("status") == "completed":
                continue
            try:
                if not isinstance(record.get("answer"), str):
                    record["status"] = "processing"
                    self._save_state()
                    answer = self._call_answer(
                        answer_function, record["question"], request_id
                    )
                    if not isinstance(answer, str) or not answer.strip():
                        raise ValueError("Помощник вернул пустой ответ")
                    record["answer"] = answer.strip()
                    record["status"] = "ready"
                    self._save_state()

                record["status"] = "sending"
                self._save_state()
                self._send_json_response(
                    record["return_to"],
                    request_id,
                    record["question"],
                    record["answer"],
                )
                record["status"] = "completed"
                self._save_state()
            except Exception:
                logger.exception("Ошибка обработки письма %s", request_id)

    def process_incoming_emails(self, answer_function):
        """Один проход почты; внешний существующий цикл запуска остаётся прежним."""
        # Сначала возобновляем сохранённую работу после перезапуска.
        self._process_pending(answer_function)

        pop_connection = None
        try:
            pop_connection = poplib.POP3_SSL(self.POP3_SERVER, self.POP3_PORT)
            pop_connection.user(self.USERNAME)
            pop_connection.pass_(self.PASSWORD)

            for request_id, question, return_to in self._read_requests(pop_connection):
                if request_id in self.state:
                    continue
                self.state[request_id] = {
                    "status": "received",
                    "question": question,
                    "return_to": return_to,
                }
                self._save_state()
        finally:
            if pop_connection is not None:
                try:
                    pop_connection.quit()
                except Exception:
                    logger.exception("Ошибка закрытия POP3")

        self._process_pending(answer_function)
        # POP3 DELE не вызывается: удаляет служебные письма только ARM.
