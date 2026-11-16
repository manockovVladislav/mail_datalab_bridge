import sys
import types
import unittest
from pathlib import Path


if "win32com.client" not in sys.modules:
    win32com = types.ModuleType("win32com")
    win32com_client = types.ModuleType("win32com.client")
    win32com.client = win32com_client
    sys.modules["win32com"] = win32com
    sys.modules["win32com.client"] = win32com_client

if "markdown" not in sys.modules:
    markdown = types.ModuleType("markdown")
    markdown.markdown = lambda value, extensions=None: value
    sys.modules["markdown"] = markdown

PROJECT_DIR = Path(__file__).resolve().parents[1] / "mail_datalab_bridge"
sys.path.insert(0, str(PROJECT_DIR))

from outlook_client import MAIL_ITEM_CLASS, OutlookClient  # noqa: E402


class FakeItems:
    def __init__(self, messages):
        self.messages = messages
        self.Count = len(messages)

    def Item(self, index):
        return self.messages[index - 1]


class FakeMail:
    def __init__(self, subject, entry_id, parent):
        self.Class = MAIL_ITEM_CLASS
        self.Subject = subject
        self.EntryID = entry_id
        self.Parent = parent
        self.deleted = False

    def Delete(self):
        self.deleted = True


class FakeFolder:
    def __init__(self, entry_id, store_id):
        self.EntryID = entry_id
        self.StoreID = store_id
        self.Items = None


class FakeNamespace:
    def __init__(self, messages):
        self.messages = {message.EntryID: message for message in messages}

    def GetItemFromID(self, entry_id, store_id):
        message = self.messages[entry_id]
        if message.Parent.StoreID != store_id:
            raise KeyError(store_id)
        return message


class OutlookClientTests(unittest.TestCase):
    def test_only_exact_request_is_deleted_from_personal_relay_inbox(self):
        relay_inbox = FakeFolder("relay-inbox", "personal-store")
        target = FakeMail(
            "json_handlers request " + "a" * 24,
            "target-entry",
            relay_inbox,
        )
        another = FakeMail(
            "json_handlers request " + "b" * 24,
            "another-entry",
            relay_inbox,
        )
        relay_inbox.Items = FakeItems([target, another])

        client = OutlookClient.__new__(OutlookClient)
        client.relay_inbox = relay_inbox
        client.ns = FakeNamespace([target, another])

        client.delete_relay_inbox_by_subject(target.Subject)

        self.assertTrue(target.deleted)
        self.assertFalse(another.deleted)


if __name__ == "__main__":
    unittest.main()
