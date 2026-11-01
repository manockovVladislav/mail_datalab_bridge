import json
import sys
import tempfile
import unittest
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch


PROJECT_DIR = Path(__file__).resolve().parents[1] / "my_server_halper"
sys.path.insert(0, str(PROJECT_DIR))

import mails_handler  # noqa: E402
from mails_handler import EmailConfig  # noqa: E402


REQUEST_ID = "c" * 24


class FakePop3:
    def __init__(self, raw_message=b""):
        self.raw_message = raw_message
        self.deleted = []

    def user(self, username):
        pass

    def pass_(self, password):
        pass

    def list(self):
        messages = [b"1 100"] if self.raw_message else []
        return b"+OK", messages, 0

    def retr(self, number):
        return b"+OK", self.raw_message.splitlines(), len(self.raw_message)

    def dele(self, number):
        self.deleted.append(number)

    def quit(self):
        pass


class EmailConfigTests(unittest.TestCase):
    def test_old_initialization_and_one_argument_answer_function_still_work(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            payload = {
                "protocol": 1,
                "type": "request",
                "request_id": REQUEST_ID,
                "question": "Первый вопрос?\n\nВторой вопрос?",
            }
            message = EmailMessage()
            message["From"] = "technical@example.com"
            message["To"] = "relay@example.com"
            message["Subject"] = f"json_handlers request {REQUEST_ID}"
            message.set_content(json.dumps(payload, ensure_ascii=False))
            pop3 = FakePop3(message.as_bytes())

            # Первые четыре аргумента совпадают с прежней инициализацией.
            handler = EmailConfig(
                "relay@example.com",
                "login",
                "password",
                "Помощнику:",
                state_file=Path(temp_dir) / "state.json",
            )
            responses = []
            handler._send_json_response = lambda *args: responses.append(args)
            calls = []

            def answer(question):
                calls.append(question)
                return "Общий ответ"

            with patch.object(mails_handler.poplib, "POP3_SSL", return_value=pop3):
                handler.process_incoming_emails(answer)
                handler.process_incoming_emails(answer)

            self.assertEqual(calls, ["Первый вопрос?\n\nВторой вопрос?"])
            self.assertEqual(len(responses), 1)
            self.assertEqual(responses[0][0], "technical@example.com")
            self.assertEqual(pop3.deleted, [])
            self.assertEqual(handler.state[REQUEST_ID]["status"], "completed")

    def test_saved_response_is_resumed_before_reading_pop3(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            state_file = Path(temp_dir) / "state.json"
            state_file.write_text(
                json.dumps(
                    {
                        REQUEST_ID: {
                            "status": "sending",
                            "question": "Вопрос",
                            "answer": "Сохранённый ответ",
                            "return_to": "technical@example.com",
                        }
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            handler = EmailConfig(
                "relay@example.com",
                "login",
                "password",
                state_file=state_file,
            )
            responses = []
            handler._send_json_response = lambda *args: responses.append(args)

            with patch.object(
                mails_handler.poplib, "POP3_SSL", return_value=FakePop3()
            ):
                handler.process_incoming_emails(
                    lambda question: self.fail("Помощник не должен вызываться")
                )

            self.assertEqual(len(responses), 1)
            self.assertEqual(handler.state[REQUEST_ID]["status"], "completed")


if __name__ == "__main__":
    unittest.main()
