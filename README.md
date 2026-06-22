# GD AutoFill Cloud Free

Бесплатная веб-версия для Streamlit Community Cloud.

## Что делает

- Пользователь открывает ссылку.
- Загружает Excel.
- Программа ищет характеристики через Google/Serper.
- Gemini извлекает характеристики под колонки Excel.
- Пользователь скачивает готовый Excel.

## Как запустить локально

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

## Локальный .env

Создай файл `.env` рядом с app.py:

```text
GEMINI_API_KEY=твой_ключ
SERPER_API_KEY=твой_ключ
```

## Как выложить бесплатно

1. Создай репозиторий на GitHub.
2. Загрузи туда файлы из этой папки.
3. Открой https://share.streamlit.io/
4. Подключи GitHub.
5. Выбери репозиторий и файл `app.py`.
6. В Settings → Secrets добавь:

```toml
GEMINI_API_KEY = "твой_ключ"
SERPER_API_KEY = "твой_ключ"
```

7. Нажми Deploy.
8. Получишь ссылку, которую можно отдавать коллегам.
