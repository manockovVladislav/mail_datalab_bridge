1. Windows: python -m pip install -r requirements.txt
2. Заполните config.py: BASE_URL, FULL_USERNAME, BROWSER_PATH, DRIVER_PATH, TECH_MAIL_FOLDER, TECH_MAIL_SMTP, SUBJECT_PHRASE.
3. Запустите в Windows: python main.py
4. В DataLab ваш обработчик читает email_test/input/request_<id>.txt (UTF-8 JSON) и создаёт email_test/output/response_<id>.txt (UTF-8 обычный текст).
5. Windows опрашивает папку ответов каждые 15 секунд и отправляет ответ с технического ящика.
6. Состояние сохраняется в state.json. Не удаляйте его, если не хотите повторно обработать письма.
7. Если статус 'sending' остался после сбоя, проверьте папку 'Отправленные' Outlook. После проверки вручную поменяйте статус на 'sent' или 'waiting' в state.json.
8. Запускайте только один экземпляр программы.
9. Из корня проекта тесты запускаются так: python -m unittest discover -s tests -v

Внимание: Windows-скрипт НЕ запускает обработчик DataLab. Его нужно запустить отдельно. Для первого теста создайте файл response_<id>.txt вручную.
