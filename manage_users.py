#!/usr/bin/env python3
"""
Управление пользователями Report Generator.

Использование:
  python manage_users.py list
  python manage_users.py add <username> <password>
  python manage_users.py remove <username>
  python manage_users.py passwd <username> <new_password>
"""
import json
import sys
from pathlib import Path
from werkzeug.security import generate_password_hash

USERS_FILE = Path(__file__).parent / "users.json"


def load():
    if USERS_FILE.exists():
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    return {}


def save(users: dict):
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")


def cmd_list():
    users = load()
    if not users:
        print("Нет пользователей.")
        return
    print(f"{'Логин':<20} {'Хэш пароля'}")
    print("-" * 70)
    for name, phash in users.items():
        print(f"{name:<20} {phash[:40]}…")


def cmd_add(username: str, password: str):
    users = load()
    if username in users:
        print(f"Пользователь '{username}' уже существует. Используйте passwd для смены пароля.")
        sys.exit(1)
    users[username] = generate_password_hash(password)
    save(users)
    print(f"Пользователь '{username}' добавлен.")


def cmd_remove(username: str):
    users = load()
    if username not in users:
        print(f"Пользователь '{username}' не найден.")
        sys.exit(1)
    del users[username]
    save(users)
    print(f"Пользователь '{username}' удалён.")


def cmd_passwd(username: str, new_password: str):
    users = load()
    if username not in users:
        print(f"Пользователь '{username}' не найден.")
        sys.exit(1)
    users[username] = generate_password_hash(new_password)
    save(users)
    print(f"Пароль '{username}' изменён.")


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(0)

    cmd = args[0]
    if cmd == "list":
        cmd_list()
    elif cmd == "add" and len(args) == 3:
        cmd_add(args[1], args[2])
    elif cmd == "remove" and len(args) == 2:
        cmd_remove(args[1])
    elif cmd == "passwd" and len(args) == 3:
        cmd_passwd(args[1], args[2])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
