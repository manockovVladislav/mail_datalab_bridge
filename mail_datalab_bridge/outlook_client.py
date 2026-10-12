import hashlib
import logging
import win32com.client


class OutlookClient:
    def __init__(self, folder_name, sender_smtp):
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self.folder_name = folder_name
        self.sender_smtp = sender_smtp
        logging.info("Outlook подключён, папка ящика: %s", folder_name)

    @staticmethod
    def sender_email(mail):
        try:
            if mail.SenderEmailType == "EX":
                user = mail.Sender.GetExchangeUser() if mail.Sender else None
                if user and user.PrimarySmtpAddress:
                    return user.PrimarySmtpAddress
                # SMTP-адрес отправителя в MAPI
                return mail.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
                )
            return mail.SenderEmailAddress
        except Exception:
            return None

    def recent_requests(self, subject_phrase, limit):
        inbox = self.ns.Folders.Item(self.folder_name).Folders.Item("Входящие")
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        logging.info("Во входящих всего объектов: %s; проверяем первые %s", messages.Count, limit)
        result = []
        for index in range(1, min(limit, messages.Count) + 1):
            try:
                mail = messages.Item(index)
                if mail.Class != 43:
                    continue
                subject = str(mail.Subject or "")
                if subject_phrase.casefold() not in subject.casefold():
                    continue
                sender = self.sender_email(mail)
                if not sender or "@" not in sender:
                    logging.warning("Не удалось определить SMTP отправителя письма")
                    continue
                # Стабильный идентификатор письма, не раскрываем EntryID в имени файла.
                uid = hashlib.sha256(str(mail.EntryID).encode("utf-8")).hexdigest()[:24]
                result.append({
                    "id": uid,
                    "email": sender,
                    "subject": subject,
                    "body": str(mail.Body or ""),
                })
            except Exception:
                logging.exception("Не удалось обработать объект Outlook с индексом %s", index)
        logging.info("Отобрано писем по теме: %s", len(result))
        return result

    def send(self, to, subject, body):
        logging.info("Создание исходящего письма (тема: %s)", subject)
        mail = self.app.CreateItem(0)
        found = False
        for account in self.ns.Accounts:
            if str(account.SmtpAddress).casefold() == self.sender_smtp.casefold():
                mail.SendUsingAccount = account
                found = True
                break
        if not found:
            mail.SentOnBehalfOfName = self.sender_smtp
        mail.To = to
        mail.Subject = "Re: " + subject
        mail.Body = body
        mail.Send()
        logging.info("Outlook принял письмо на отправку")
