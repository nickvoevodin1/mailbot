#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Почтовый ИИ-бот.

Читает письма ТОЛЬКО от одного адреса (ALLOWED_SENDER), на которые ещё не отвечено,
задаёт вопрос GigaChat и отвечает письмом простым текстом. Если в письме
просят «документ» или «файл», прикладывает ответ в формате Word 97-2003 (.doc).

Запуск:
    python bot.py          - обработать почту
    python bot.py --peek   - только проверить, есть ли что обрабатывать
                             (нужно для экономного запуска в GitHub Actions)

Все секреты берутся из переменных окружения, в коде их нет.
"""
import html
import imaplib
import os
import re
import shutil
import smtplib
import subprocess
import sys
import tempfile
import uuid
from email import message_from_bytes, policy
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, parseaddr
from pathlib import Path

import requests

IMAP_HOST = "imap.mail.ru"
SMTP_HOST = "smtp.mail.ru"
OAUTH_URL = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
# Сначала пробуем новый адрес API, потом прежний.
CHAT_BASE_URLS = [
    "https://api.giga.chat/v1",
    "https://gigachat.devices.sberbank.ru/api/v1",
]

MAX_PER_RUN = 5          # не больше стольких писем за один запуск
MAX_BODY_CHARS = 6000    # длиннее письмо обрезаем
DOC_WORDS = ("документ", "файл", "ворд", "word", "docx")

SYSTEM_PROMPT = (
    "Ты помощник для незрячего человека, который читает ответы программой "
    "экранного чтения. Отвечай по-русски, ясно и по делу, обычным текстом. "
    "Не используй таблицы, markdown-разметку, смайлики и рисунки из символов. "
    "Перечисления оформляй короткими строками, начинающимися с дефиса. "
    "Если чего-то не знаешь, так и скажи."
)
DOC_PROMPT = (
    "Подготовь документ. Первая строка - название документа. Каждый раздел "
    "начинай строкой, которая начинается с «# ». Списки - строками, "
    "начинающимися с «- ». Всё остальное - обычные абзацы. "
    "Никаких таблиц и другой разметки."
)


# ---------- настройки ----------

def load_config(need_ai):
    def get(name, default=None, required=False):
        value = os.environ.get(name, default)
        if required and not value:
            sys.exit(f"Не задана переменная окружения {name}")
        return value

    cfg = {
        "user": get("MAIL_USER", required=True),
        "password": get("MAIL_APP_PASSWORD", required=True),
        "sender": get("ALLOWED_SENDER", required=True).strip().lower(),
        "healthcheck": get("HEALTHCHECK_URL"),
    }
    if need_ai:
        cfg.update({
            "auth_key": get("GIGACHAT_AUTH_KEY", required=True),
            "scope": get("GIGACHAT_SCOPE", "GIGACHAT_API_PERS"),
            "model": get("GIGACHAT_MODEL", "GigaChat-2-Pro"),
            "base_url": get("GIGACHAT_BASE_URL"),
            "ca": get("CA_BUNDLE", True),  # путь к файлу сертификатов или True
        })
    return cfg


# ---------- разбор письма ----------

def message_text(msg):
    """Текст письма: предпочитаем text/plain, при необходимости чистим HTML."""
    part = msg.get_body(preferencelist=("plain", "html"))
    if part is None:
        return ""
    try:
        text = part.get_content()
    except Exception:
        return ""
    if part.get_content_type() == "text/html":
        text = re.sub(r"(?is)<(script|style).*?</\1>", " ", text)
        text = re.sub(r"(?i)<br\s*/?>|</p>", "\n", text)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text)
    return text


def tidy(text):
    """Убираем значки цитирования «>», лишние пустые строки, ограничиваем длину."""
    lines = [re.sub(r"^\s*(>\s*)+", "", ln).rstrip() for ln in text.splitlines()]
    text = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
    return text[:MAX_BODY_CHARS]


def wants_doc(subject, text):
    """Просят документ? Смотрим тему и начало письма (новый текст идёт сверху)."""
    head = (subject + " " + text[:1000]).lower()
    return any(word in head for word in DOC_WORDS)


def sender_address(msg):
    return parseaddr(str(msg.get("From", "")))[1].strip().lower()


# ---------- почта ----------

def open_imap(cfg):
    imap = imaplib.IMAP4_SSL(IMAP_HOST, 993)
    imap.login(cfg["user"], cfg["password"])
    imap.select("INBOX")
    return imap


def collect(imap, cfg, mark_strangers):
    """Письма от нужного адреса, на которые бот ещё не ответил: [(номер, письмо), ...].

    Смотрим на пометку «отвечено», а не «прочитано»: так бот работает, даже если
    кто-то (например, Инна) открыл ящик в почтовой программе и письма стали прочитанными.
    """
    typ, data = imap.search(None, "UNANSWERED", "FROM", f'"{cfg["sender"]}"')
    nums = data[0].split() if typ == "OK" and data and data[0] else []
    found = []
    for num in nums:
        typ, parts = imap.fetch(num, "(BODY.PEEK[])")
        if typ != "OK" or not parts or not isinstance(parts[0], tuple):
            continue
        msg = message_from_bytes(parts[0][1], policy=policy.default)
        if sender_address(msg) != cfg["sender"]:
            # IMAP ищет по части адреса, поэтому проверяем точное совпадение
            if mark_strangers:
                mark_done(imap, num)
            continue
        found.append((num, msg))
    return found


def mark_done(imap, num):
    imap.store(num, "+FLAGS", "(\\Answered \\Seen)")


def send_reply(cfg, original, body, attachment=None):
    subject = str(original.get("Subject", "") or "").strip()
    if not re.match(r"(?i)^re:", subject):
        subject = "Re: " + (subject or "Ваш вопрос")
    reply = EmailMessage()
    reply["From"] = cfg["user"]
    reply["To"] = cfg["sender"]
    reply["Subject"] = subject
    reply["Date"] = formatdate(localtime=True)
    reply["Message-ID"] = make_msgid()
    if original.get("Message-ID"):
        reply["In-Reply-To"] = original["Message-ID"]
        reply["References"] = original["Message-ID"]
    reply.set_content(body)
    if attachment:
        reply.add_attachment(
            attachment.read_bytes(),
            maintype="application",
            subtype="msword",
            filename=attachment.name,
        )
    with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=60) as smtp:
        smtp.login(cfg["user"], cfg["password"])
        smtp.send_message(reply)


# ---------- ИИ ----------

def get_token(cfg):
    resp = requests.post(
        OAUTH_URL,
        headers={
            "Authorization": f"Basic {cfg['auth_key']}",
            "RqUID": str(uuid.uuid4()),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data={"scope": cfg["scope"]},
        timeout=30,
        verify=cfg["ca"],
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def ask_ai(cfg, question, as_document):
    token = get_token(cfg)
    system = SYSTEM_PROMPT + ((" " + DOC_PROMPT) if as_document else "")
    payload = {
        "model": cfg["model"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": question},
        ],
        "temperature": 0.5,
        "max_tokens": 3000,
    }
    last_error = None
    for base in ([cfg["base_url"]] if cfg.get("base_url") else CHAT_BASE_URLS):
        try:
            resp = requests.post(
                base.rstrip("/") + "/chat/completions",
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                json=payload,
                timeout=120,
                verify=cfg["ca"],
            )
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"].strip()
        except Exception as exc:  # пробуем следующий адрес
            last_error = exc
    raise RuntimeError(f"ИИ не ответил: {type(last_error).__name__}: {last_error}")


# ---------- документ .doc ----------

def plain_line(line):
    """Убираем остатки markdown, которые модель могла добавить."""
    return re.sub(r"\*\*|__|`", "", line).strip()


def build_docx(text, path):
    from docx import Document

    lines = [ln for ln in (plain_line(x) for x in text.splitlines()) if ln]
    if not lines:
        lines = ["Документ"]
    title = re.sub(r"^#+\s*", "", lines[0]) or "Документ"
    doc = Document()
    doc.add_heading(title, level=1)          # настоящие заголовки удобны для экранного чтения
    for line in lines[1:]:
        if re.match(r"^#{1,6}\s+", line):
            doc.add_heading(re.sub(r"^#+\s*", "", line), level=2)
        elif re.match(r"^[-*•]\s+", line):
            doc.add_paragraph(re.sub(r"^[-*•]\s+", "", line), style="List Bullet")
        else:
            doc.add_paragraph(line)
    doc.save(path)
    return title


def make_doc_file(text, workdir):
    """Возвращает (путь к .doc, название) или (None, название), если не получилось."""
    safe = "Документ"
    try:
        docx_path = Path(workdir) / "Документ.docx"
        title = build_docx(text, docx_path)
        cleaned = re.sub(r"[^\w\- ]", "", title).strip()[:40]
        safe = cleaned or "Документ"
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return None, title
        subprocess.run(
            [soffice, "--headless", "--convert-to", "doc:MS Word 97",
             "--outdir", str(workdir), str(docx_path)],
            check=True, timeout=180,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        produced = Path(workdir) / "Документ.doc"
        if not produced.exists():
            return None, title
        final = Path(workdir) / f"{safe}.doc"
        produced.rename(final)
        return final, title
    except Exception as exc:
        print(f"Не удалось создать .doc: {type(exc).__name__}")
        return None, safe


# ---------- основная логика ----------

def ping(cfg):
    if cfg.get("healthcheck"):
        try:
            requests.get(cfg["healthcheck"], timeout=10)
        except Exception:
            pass


def write_outputs(pending, doc):
    line = f"pending={pending}\ndoc={1 if doc else 0}\n"
    print(line.replace("\n", " ").strip())
    path = os.environ.get("GITHUB_OUTPUT")
    if path:
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(line)


def main():
    peek = "--peek" in sys.argv
    cfg = load_config(need_ai=not peek)
    imap = open_imap(cfg)
    try:
        items = collect(imap, cfg, mark_strangers=not peek)[:MAX_PER_RUN]

        if peek:
            need_doc = any(
                wants_doc(str(m.get("Subject", "") or ""), tidy(message_text(m)))
                for _, m in items
            )
            write_outputs(len(items), need_doc)
            ping(cfg)
            return

        for num, msg in items:
            subject = str(msg.get("Subject", "") or "")
            text = tidy(message_text(msg))
            if not text and not subject.strip():
                mark_done(imap, num)
                continue
            as_doc = wants_doc(subject, text)
            question = f"Тема письма: {subject}\n\n{text}" if text else subject
            answer = ask_ai(cfg, question, as_doc)

            attachment = None
            body = answer
            if as_doc:
                with tempfile.TemporaryDirectory() as tmp:
                    attachment, title = make_doc_file(answer, tmp)
                    if attachment:
                        body = (f"Документ «{title}» во вложении (формат Word 97-2003). "
                                f"Тот же текст ниже.\n\n{answer}")
                    else:
                        body = ("Файл создать не удалось, поэтому присылаю текст "
                                f"письмом.\n\n{answer}")
                    send_reply(cfg, msg, body, attachment)
            else:
                send_reply(cfg, msg, body)
            mark_done(imap, num)   # отмечаем только после успешной отправки
            print("Ответ отправлен")
        ping(cfg)
    finally:
        try:
            imap.logout()
        except Exception:
            pass


if __name__ == "__main__":
    main()
