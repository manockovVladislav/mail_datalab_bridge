import base64
import logging
import time
from pathlib import Path
from urllib.parse import quote, urlsplit

import requests
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class DataLabClient:
    def __init__(self, base_url, username, browser_path, driver_path):
        self.base_url = base_url.rstrip("/")
        self.username = username
        self.browser_path = browser_path
        self.driver_path = driver_path
        self.session = requests.Session()
        self.driver = None

    def connect(self, password, retries=3):
        options = Options()
        options.binary_location = self.browser_path
        options.add_argument("--window-size=1280,1000")
        # Не отключаем проверку сертификатов и защиту браузера.
        logging.info("Запуск корпоративного браузера")
        self.driver = webdriver.Chrome(
            service=Service(executable_path=self.driver_path), options=options
        )
        for attempt in range(1, retries + 1):
            try:
                logging.info("Авторизация DataLab: попытка %s/%s", attempt, retries)
                self.driver.get(f"{self.base_url}/hub/login")
                try:
                    user = WebDriverWait(self.driver, 12).until(
                        EC.presence_of_element_located((By.NAME, "username"))
                    )
                    user.clear()
                    user.send_keys(self.username)
                    pw = self.driver.find_element(By.NAME, "password")
                    pw.clear()
                    pw.send_keys(password)
                    pw.send_keys(Keys.RETURN)
                    logging.info("Форма авторизации отправлена")
                except TimeoutException:
                    logging.info("Форма логина не найдена: возможно, уже выполнен вход")

                # Если JupyterHub предлагает выбрать образ, выбираем его.
                try:
                    image = WebDriverWait(self.driver, 8).until(
                        EC.presence_of_element_located((By.NAME, "image"))
                    )
                    image.clear()
                    image.send_keys("Python 3.11 CUDA 12.1 (COIN-OR)")
                    image.send_keys(Keys.RETURN)
                    button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                "//button[.//span[text()='Создать'] "
                                "or contains(., 'Создать')]",
                            )
                        )
                    )
                    button.click()
                    logging.info("Выбран образ JupyterHub")
                except TimeoutException:
                    pass  # Образ может быть уже выбран

                deadline = time.monotonic() + 180
                while time.monotonic() < deadline:
                    if "/lab" in self.driver.current_url:
                        break
                    time.sleep(3)
                else:
                    raise TimeoutError("JupyterLab не открылся за 180 секунд")

                logging.info("JupyterLab открыт; перенос cookies в requests")
                self.session.cookies.clear()
                self.session.headers.pop("X-XSRFToken", None)
                browser_cookies = self.driver.get_cookies()
                for cookie in browser_cookies:
                    cookie_options = {
                        "path": cookie.get("path") or "/",
                        "secure": bool(cookie.get("secure", False)),
                    }
                    if cookie.get("domain"):
                        cookie_options["domain"] = cookie["domain"]
                    self.session.cookies.set(
                        cookie["name"], cookie["value"], **cookie_options
                    )

                api_path = urlsplit(self._url("")).path
                xsrf_cookies = [
                    cookie
                    for cookie in browser_cookies
                    if cookie.get("name") == "_xsrf"
                    and api_path.startswith(cookie.get("path") or "/")
                ]
                if xsrf_cookies:
                    xsrf_cookie = max(
                        xsrf_cookies, key=lambda cookie: len(cookie.get("path") or "/")
                    )
                    self.session.headers["X-XSRFToken"] = xsrf_cookie["value"].strip(
                        '"'
                    )
                # Проверяем API, а не только открытие страницы.
                logging.info("Проверка Jupyter Contents API")
                resp = self.session.get(self._url(""), timeout=30)
                resp.raise_for_status()
                if "application/json" not in resp.headers.get("Content-Type", ""):
                    raise RuntimeError("API вернул не JSON: проверьте авторизацию")
                logging.info("DataLab подключён")
                return
            except Exception as exc:
                logging.warning("Попытка подключения %s/%s: %s", attempt, retries, exc)
                if attempt == retries:
                    raise
                time.sleep(5)

    def _url(self, path):
        prefix = f"{self.base_url}/user/{quote(self.username, safe='')}/api/contents"
        return prefix + ("/" + quote(path.strip("/"), safe="/") if path else "")

    def ensure_directory(self, path):
        current = ""
        for segment in Path(path).parts:
            if segment in (".", ""):
                continue
            current = f"{current}/{segment}".strip("/")
            r = self.session.get(self._url(current), params={"content": 0}, timeout=30)
            if r.status_code == 404:
                logging.info("Создание директории DataLab: %s", current)
                r = self.session.put(self._url(current), json={"type": "directory"}, timeout=30)
            r.raise_for_status()

    def upload_text(self, remote_path, text):
        logging.info("PUT файл DataLab: %s (%s символов)", remote_path, len(text))
        self.ensure_directory(str(Path(remote_path).parent).replace("\\", "/"))
        # Как в коде коллег: base64, чтобы корректно передавать русский текст.
        encoded = base64.b64encode(text.encode("utf-8")).decode("ascii")
        r = self.session.put(
            self._url(remote_path),
            json={"type": "file", "format": "base64", "content": encoded},
            timeout=60,
        )
        r.raise_for_status()
        logging.info("Файл успешно загружен: %s (HTTP %s)", remote_path, r.status_code)

    def read_text_if_exists(self, remote_path):
        r = self.session.get(self._url(remote_path), params={"content": 1}, timeout=30)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        logging.info("Файл найден в DataLab: %s", remote_path)
        data = r.json()
        if data.get("type") != "file":
            raise ValueError("Ожидался файл, а не папка")
        if data.get("format") == "base64":
            return base64.b64decode(data["content"]).decode("utf-8")
        return data["content"]

    def close(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            logging.exception("Не удалось корректно закрыть браузер")
        finally:
            self.driver = None
            self.session.close()
