import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace


PROJECT_DIR = Path(__file__).resolve().parents[1] / "mail_datalab_bridge"
sys.path.insert(0, str(PROJECT_DIR))

from mail_processor import MailProcessor, clean_question  # noqa: E402


REQUEST_ID = "a" * 24


class FakeOutlook:
    def __init__(self, messages=None):
        self.messages = messages or []
        self.json_sent = []
        self.answers_sent = []
        self.sent_subjects = set()
        self.sent_markers = set()
        self.deleted = []
        self.deleted_sent = []

    def incoming(self, limit):
        return self.messages[:limit]

    def send_json(self, to, subject, payload):
        self.json_sent.append((to, subject, payload))
        self.sent_subjects.add(subject)

    def send_answer(self, **kwargs):
        self.answers_sent.append(kwargs)
        self.sent_markers.add((f"Re: {kwargs['original_subject']}", kwargs["request_id"]))

    def sent_message_exists(self, subject, marker=None):
        if marker is None:
            return subject in self.sent_subjects
        return (subject, marker) in self.sent_markers

    def delete_message(self, entry_id, store_id):
        self.deleted.append((entry_id, store_id))

    def delete_sent_by_subject(self, subject):
        self.deleted_sent.append(subject)


class MailProcessorTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.state_file = Path(self.temp_dir.name) / "state.json"
        self.cfg = SimpleNamespace(
            STATE_FILE=self.state_file,
            SCAN_LIMIT=100,
            POLL_SECONDS=1,
            SEND_RECOVERY_SECONDS=1,
            REQUEST_SUBJECT_PHRASE="вопрос помощнику",
            BRIDGE_SUBJECT="json_handlers",
            RELAY_MAIL_SMTP="relay@example.com",
        )

    @staticmethod
    def request_message():
        return {
            "id": REQUEST_ID,
            "entry_id": "source-entry",
            "store_id": "store",
            "email": "user@example.com",
            "subject": "Вопрос помощнику: отчёт",
            "body": "1. Первый вопрос?\n\n2. Второй вопрос?\n\nС уважением,\nИван",
        }

    def test_question_keeps_multiple_questions_and_removes_signature(self):
        self.assertEqual(
            clean_question(self.request_message()["body"]),
            "1. Первый вопрос?\n\n2. Второй вопрос?",
        )

    def test_new_request_is_forwarded_as_one_json_and_saved(self):
        outlook = FakeOutlook([self.request_message()])
        processor = MailProcessor(outlook, self.cfg)

        processor.cycle()
        processor.cycle()

        self.assertEqual(len(outlook.json_sent), 1)
        to, subject, payload = outlook.json_sent[0]
        self.assertEqual(to, "relay@example.com")
        self.assertEqual(subject, f"json_handlers request {REQUEST_ID}")
        self.assertEqual(payload["question"], "1. Первый вопрос?\n\n2. Второй вопрос?")
        self.assertEqual(processor.state[REQUEST_ID]["status"], "forwarded")

    def test_response_is_sent_to_original_address_then_cleaned(self):
        response = {
            "id": "b" * 24,
            "entry_id": "response-entry",
            "store_id": "store",
            "email": "relay@example.com",
            "subject": f"json_handlers response {REQUEST_ID}",
            "body": json.dumps(
                {
                    "protocol": 1,
                    "type": "response",
                    "request_id": REQUEST_ID,
                    # Этот адрес ARM обязан игнорировать.
                    "reply_to": "attacker@example.com",
                    "answer": "Ответ на оба вопроса",
                },
                ensure_ascii=False,
            ),
        }
        outlook = FakeOutlook([self.request_message()])
        processor = MailProcessor(outlook, self.cfg)
        processor.cycle()
        outlook.messages = [response]

        processor.cycle()
        self.assertEqual(processor.state[REQUEST_ID]["status"], "replying")
        self.assertEqual(outlook.answers_sent[0]["to"], "user@example.com")
        self.assertEqual(outlook.answers_sent[0]["question"], "1. Первый вопрос?\n\n2. Второй вопрос?")

        processor.cycle()
        self.assertEqual(processor.state[REQUEST_ID]["status"], "completed")
        self.assertIn(("source-entry", "store"), outlook.deleted)
        self.assertIn(("response-entry", "store"), outlook.deleted)
        self.assertEqual(
            outlook.deleted_sent, [f"json_handlers request {REQUEST_ID}"]
        )

    def test_restart_recovers_forwarding_from_sent_folder(self):
        relay_subject = f"json_handlers request {REQUEST_ID}"
        self.state_file.write_text(
            json.dumps(
                {
                    REQUEST_ID: {
                        "status": "forwarding",
                        "email": "user@example.com",
                        "subject": "Вопрос помощнику",
                        "question": "Вопрос",
                        "relay_subject": relay_subject,
                        "send_started_at": 0,
                    }
                }
            ),
            encoding="utf-8",
        )
        outlook = FakeOutlook()
        outlook.sent_subjects.add(relay_subject)
        processor = MailProcessor(outlook, self.cfg)

        processor.process_pending()

        self.assertEqual(processor.state[REQUEST_ID]["status"], "forwarded")
        self.assertEqual(outlook.json_sent, [])

    def test_response_scan_does_not_break_reply_recovery(self):
        outlook = FakeOutlook([self.request_message()])
        processor = MailProcessor(outlook, self.cfg)
        processor.cycle()
        record = processor.state[REQUEST_ID]
        record["status"] = "replying"
        record["answer"] = "Готовый ответ"
        record["send_started_at"] = 0
        processor._save()
        outlook.sent_markers.add((f"Re: {record['subject']}", REQUEST_ID))
        outlook.messages = [
            {
                "id": "d" * 24,
                "entry_id": "response-entry",
                "store_id": "store",
                "email": "relay@example.com",
                "subject": f"json_handlers response {REQUEST_ID}",
                "body": json.dumps(
                    {
                        "protocol": 1,
                        "type": "response",
                        "request_id": REQUEST_ID,
                        "answer": "Готовый ответ",
                    }
                ),
            }
        ]

        processor.cycle()

        self.assertEqual(processor.state[REQUEST_ID]["status"], "completed")
        self.assertEqual(outlook.answers_sent, [])

    def test_invalid_state_is_rejected(self):
        self.state_file.write_text("[]", encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "JSON-объект"):
            MailProcessor(FakeOutlook(), self.cfg)


if __name__ == "__main__":
    unittest.main()
