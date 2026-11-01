import hashlib
import html
import json
import logging
from pathlib import Path

import markdown
import win32com.client


MAIL_ITEM_CLASS = 43
INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"


class OutlookClient:
    def __init__(self, folder_name, sender_smtp, inbox_name, sent_name):
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self.root = self.ns.Folders.Item(folder_name)
        self.inbox = self.root.Folders.Item(inbox_name)
        self.sent = self.root.Folders.Item(sent_name)
        self.sender_smtp = sender_smtp
        logging.info("Outlook подключён к ящику %s", folder_name)

    @staticmethod
    def sender_email(mail):
        try:
            if str(mail.SenderEmailType or "").upper() != "EX":
                return str(mail.SenderEmailAddress or "")
        except Exception:
            return ""

        try:
            sender = mail.Sender
            user = sender.GetExchangeUser() if sender else None
            if user and user.PrimarySmtpAddress:
                return str(user.PrimarySmtpAddress)
        except Exception:
            logging.debug("ExchangeUser недоступен", exc_info=True)

        try:
            value = mail.PropertyAccessor.GetProperty(SMTP_ADDRESS)
            return str(value or "")
        except Exception:
            logging.debug("SMTP MAPI-свойство недоступно", exc_info=True)
            return ""

    @staticmethod
    def _request_id(mail):
        try:
            stable_value = mail.PropertyAccessor.GetProperty(INTERNET_MESSAGE_ID)
        except Exception:
            stable_value = None
        stable_value = stable_value or mail.EntryID
        return hashlib.sha256(str(stable_value).encode("utf-8")).hexdigest()[:24]

    def incoming(self, limit):
        items = self.inbox.Items
        items.Sort("[ReceivedTime]", True)
        result = []
        for index in range(1, min(limit, items.Count) + 1):
            try:
                mail = items.Item(index)
                if mail.Class != MAIL_ITEM_CLASS:
                    continue
                sender = self.sender_email(mail)
                if not sender or "@" not in sender:
                    logging.warning("Пропущено письмо без SMTP-адреса отправителя")
                    continue
                result.append(
                    {
                        "id": self._request_id(mail),
                        "entry_id": str(mail.EntryID),
                        "store_id": str(mail.Parent.StoreID),
                        "email": sender,
                        "subject": str(mail.Subject or ""),
                        "body": str(mail.Body or ""),
                    }
                )
            except Exception:
                logging.exception("Не удалось прочитать письмо Outlook №%s", index)
        return result

    def _new_mail(self, to, subject):
        mail = self.app.CreateItem(0)
        # Для этого контура подтверждён именно этот способ отправки.
        mail.SentOnBehalfOfName = self.sender_smtp
        mail.To = to
        mail.Subject = subject.replace("\r", " ").replace("\n", " ")
        mail.SaveSentMessageFolder = self.sent
        if not mail.Recipients.ResolveAll():
            raise ValueError(f"Outlook не смог разрешить адрес получателя: {to}")
        return mail

    def send_json(self, to, subject, payload):
        mail = self._new_mail(to, subject)
        mail.BodyFormat = 1
        mail.Body = json.dumps(payload, ensure_ascii=False, indent=2)
        mail.Send()
        logging.info("JSON-письмо передано Outlook: %s", subject)

    def send_answer(self, to, original_subject, question, answer, request_id):
        subject = f"Re: {original_subject}"
        mail = self._new_mail(to, subject)

        template = Path(__file__).with_name("template.html").read_text(encoding="utf-8")
        question_html = html.escape(question).replace("\n", "<br>\n")
        # Экранирование запрещает LLM вставить активный HTML, но Markdown-разметка
        # (списки, таблицы и блоки кода) остаётся читаемой.
        answer_html = markdown.markdown(
            html.escape(answer),
            extensions=["tables", "fenced_code", "sane_lists"],
        )
        mail.BodyFormat = 2
        mail.HTMLBody = (
            template.replace("{question}", question_html)
            .replace("{answers_html}", answer_html)
            .replace("{request_id}", request_id)
        )
        mail.Send()
        logging.info("Ответ передан Outlook для %s", to)

    def sent_message_exists(self, subject, marker=None):
        items = self.sent.Items
        items.Sort("[SentOn]", True)
        for index in range(1, min(200, items.Count) + 1):
            try:
                mail = items.Item(index)
                if mail.Class != MAIL_ITEM_CLASS or str(mail.Subject or "") != subject:
                    continue
                if marker is None:
                    return True
                body = str(getattr(mail, "HTMLBody", "") or "")
                if marker in body:
                    return True
            except Exception:
                logging.debug("Ошибка проверки Отправленных", exc_info=True)
        return False

    def delete_message(self, entry_id, store_id):
        if not entry_id or not store_id:
            return
        try:
            mail = self.ns.GetItemFromID(entry_id, store_id)
        except Exception:
            # Уже перемещённое письмо получает другой EntryID.
            return
        if str(mail.Parent.EntryID) != str(self.inbox.EntryID):
            return
        mail.Delete()
        logging.info("Входящее служебное письмо перемещено в Удалённые")

    def delete_sent_by_subject(self, subject):
        items = self.sent.Items
        entry_ids = []
        for index in range(1, items.Count + 1):
            try:
                mail = items.Item(index)
                if mail.Class == MAIL_ITEM_CLASS and str(mail.Subject or "") == subject:
                    entry_ids.append((str(mail.EntryID), str(mail.Parent.StoreID)))
            except Exception:
                logging.debug("Ошибка просмотра Отправленных", exc_info=True)

        for entry_id, store_id in entry_ids:
            try:
                mail = self.ns.GetItemFromID(entry_id, store_id)
                if str(mail.Parent.EntryID) == str(self.sent.EntryID):
                    mail.Delete()
            except Exception:
                logging.debug("Письмо уже удалено из Отправленных", exc_info=True)
        if entry_ids:
            logging.info("Служебный запрос перемещён из Отправленных в Удалённые")
