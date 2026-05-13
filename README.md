# Report Generator

Flask-приложение для генерации Excel-отчётов из PostgreSQL по заранее подготовленным SQL-скриптам.

## Структура проекта

```
report_app/
├── app.py                    # Основное Flask-приложение
├── requirements.txt
├── .env.example              # Шаблон переменных окружения
├── scripts/
│   └── registry.json         # Реестр доступных скриптов
├── sql_scripts/              # SQL-файлы с отчётами
│   └── sales_report.sql      # Пример скрипта
└── templates/
    └── index.html            # Фронтенд
    └── login.html            # Страница авторизации
```

## Быстрый старт

```bash
# 1. Установите зависимости
pip install -r requirements.txt

# 2. Создайте .env
cp .env.example .env
# Отредактируйте .env — укажите параметры вашей БД

# 3. Запустите
python app.py
```

Приложение запустится на http://localhost:5000

## Добавление новых скриптов

### 1. Создайте SQL-файл в `sql_scripts/`

Используйте плейсхолдеры `:date_from` и `:date_to` для дат:

```sql
SELECT *
FROM your_table
WHERE created_at::date BETWEEN :date_from AND :date_to;
```

### 2. Зарегистрируйте в `scripts/registry.json`

```json
{
  "id":          "my_report",
  "name":        "Мой отчёт",
  "description": "Краткое описание",
  "file":        "my_report.sql"
}
```

## Логика формирования файла

| Ситуация | Результат |
|---|---|
| 1 скрипт, есть данные | Один `.xlsx` файл |
| Несколько скриптов, все с данными | Один `.xlsx` с листами по каждому скрипту |
| Скрипт без данных | Toast-уведомление с названием скрипта |
| Все скрипты без данных | Уведомление, файл не скачивается |

## Формат дат в SQL

Плейсхолдеры автоматически заменяются перед выполнением:

| Плейсхолдер | Результат |
|---|---|
| `:date_from` | `'2024-01-01'` |
| `:date_to` | `'2024-01-31'` |

## Переменные окружения

| Переменная | По умолчанию | Описание |
|---|---|---|
| `DB_HOST` | `localhost` | Хост PostgreSQL |
| `DB_PORT` | `5432` | Порт |
| `DB_NAME` | `your_database` | Имя базы данных |
| `DB_USER` | `your_user` | Пользователь |
| `DB_PASSWORD` | `your_password` | Пароль |

## Авторизация

Пользователи хранятся в `users.json` (хэши паролей, bcrypt через Werkzeug).

```bash
# Добавить пользователя
python manage_users.py add ivan SuperSecret123

# Список пользователей
python manage_users.py list

# Сменить пароль
python manage_users.py passwd ivan NewPassword456

# Удалить пользователя
python manage_users.py remove ivan
```

Добавьте `SECRET_KEY` в `.env` для безопасных сессий:
```
SECRET_KEY=ваш-длинный-случайный-ключ
```

## Журнал аудита

Файл `logs/audit.log` — создаётся автоматически. Формат:

```
дата-время    логин    IP-адрес    действие    детали
```

Пример:
```
2026-05-08 15:30:12    ivan    127.0.0.1    LOGIN       success
2026-05-08 15:31:05    ivan    127.0.0.1    GENERATE    period=2026-04-01/2026-04-30 requested=[Отчёт по продажам] with_data=[Отчёт по продажам] empty=[]
2026-05-08 15:45:00    ivan    127.0.0.1    LOGOUT
2026-05-08 15:46:10    —       127.0.0.1    LOGIN_FAIL  username=hacker
```
