import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
import smtplib
import csv
import time
import os
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.utils import formataddr
import requests
from io import BytesIO

TEMPLATES_FILE = "templates.json"

class MailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Рассылка приглашений — финальная версия")
        self.guests = []
        self.image_data = None
        self.sent_count = 0
        self.error_count = 0

        # Gmail + пароль
        tk.Label(root, text="Gmail:").grid(row=0, column=0, sticky="e")
        self.entry_email = tk.Entry(root, width=30)
        self.entry_email.grid(row=0, column=1, sticky="w")

        tk.Label(root, text="App password:").grid(row=1, column=0, sticky="e")
        self.entry_password = tk.Entry(root, show="*", width=30)
        self.entry_password.grid(row=1, column=1, sticky="w")

        # CSV
        tk.Button(root, text="📂 Загрузить CSV гостей", command=self.load_csv).grid(row=0, column=2, padx=5)
        self.label_guests = tk.Label(root, text="Гостей: 0")
        self.label_guests.grid(row=1, column=2, padx=5)

        # Шаблоны
        tk.Label(root, text="Шаблон письма:").grid(row=2, column=0, sticky="e")
        self.template_var = tk.StringVar()
        self.combo_templates = tk.OptionMenu(root, self.template_var, "")
        self.combo_templates.grid(row=2, column=1, sticky="w")
        self.load_templates()
        self.template_var.trace("w", lambda *a: self.apply_template())

        # Тема письма
        tk.Label(root, text="Тема письма:").grid(row=3, column=0, sticky="e")
        self.entry_subject = tk.Entry(root, width=60)
        self.entry_subject.grid(row=3, column=1, columnspan=2, sticky="we")

        # HTML поле
        tk.Label(root, text="HTML-содержимое:").grid(row=4, column=0, sticky="nw")
        self.html_widget = scrolledtext.ScrolledText(root, width=100, height=20)
        self.html_widget.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

        # Кнопки картинок
        self.frame_buttons = tk.Frame(root)
        self.frame_buttons.grid(row=6, column=0, columnspan=3, sticky="w", pady=5)

        tk.Button(self.frame_buttons, text="📎 Прикрепить картинку (файл)", command=self.load_image_file).grid(row=0, column=0, padx=5)
        tk.Button(self.frame_buttons, text="🌐 Картинка по URL", command=self.load_image_url).grid(row=0, column=1, padx=5)

        # Нижние кнопки
        self.frame_send = tk.Frame(root)
        self.frame_send.grid(row=7, column=0, columnspan=3, pady=10)

        tk.Button(self.frame_send, text="👀 Предпросмотр", bg="lightblue", command=self.preview_invite).grid(row=0, column=0, padx=5)
        tk.Button(self.frame_send, text="✉ Отправить тест", bg="orange", command=self.send_test_email).grid(row=0, column=1, padx=5)
        tk.Button(self.frame_send, text="📨 Отправить всем", bg="lightgreen", command=self.send_all).grid(row=0, column=2, padx=5)

        # Правая панель
        self.frame_right = tk.Frame(root)
        self.frame_right.grid(row=0, column=3, rowspan=8, padx=10, sticky="ns")

        tk.Label(self.frame_right, text="Настройка").pack()

        tk.Label(self.frame_right, text="Задержка (сек):").pack()
        self.delay_entry = tk.Entry(self.frame_right, width=5)
        self.delay_entry.insert(0, "3")
        self.delay_entry.pack()

        self.label_stats = tk.Label(self.frame_right, text="📤 Отправлено: 0 | ⚠ Ошибок: 0")
        self.label_stats.pack(pady=5)

        tk.Label(self.frame_right, text="Лог:").pack()
        self.log = scrolledtext.ScrolledText(self.frame_right, width=40, height=20, bg="black", fg="lime")
        self.log.pack()

    def log_message(self, text, error=False):
        self.log.insert(tk.END, f"{text}\n", "error" if error else "success")
        self.log.tag_config("error", foreground="red")
        self.log.tag_config("success", foreground="lime")
        self.log.see(tk.END)

    def load_csv(self):
        file = filedialog.askopenfilename(filetypes=[("CSV файлы", "*.csv")])
        if not file: return
        with open(file, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            self.guests = [row for row in reader]
        self.label_guests.config(text=f"Гостей: {len(self.guests)}")
        self.log_message(f"📂 Загружен CSV: {len(self.guests)} гостей")

    def load_templates(self):
        if not os.path.exists(TEMPLATES_FILE):
            default_templates = {
                "День рождения": "<h2>Привет, {name}!</h2><p>Приглашаю тебя на мой день рождения 🎉</p>",
                "Свадьба": "<h2>Дорогой {name},</h2><p>Приглашаем тебя на свадьбу 💍</p>"
            }
            with open(TEMPLATES_FILE, "w", encoding="utf-8") as f:
                json.dump(default_templates, f, ensure_ascii=False, indent=2)

        with open(TEMPLATES_FILE, "r", encoding="utf-8") as f:
            templates = json.load(f)

        menu = self.combo_templates["menu"]
        menu.delete(0, "end")
        for name in templates.keys():
            menu.add_command(label=name, command=lambda v=name: self.template_var.set(v))
        self.templates = templates

    def apply_template(self):
        name = self.template_var.get()
        if name in self.templates:
            self.html_widget.delete("1.0", tk.END)
            self.html_widget.insert(tk.END, self.templates[name])
            self.entry_subject.delete(0, tk.END)
            self.entry_subject.insert(0, f"Приглашение: {name}")

    def load_image_file(self):
        file = filedialog.askopenfilename(filetypes=[("Изображения", "*.png;*.jpg;*.jpeg;*.gif")])
        if file:
            with open(file, "rb") as f:
                self.image_data = f.read()
            self.log_message("📎 Картинка прикреплена (файл)")

    def load_image_url(self):
        url = simpledialog.askstring("URL картинки", "Введите ссылку на картинку:")
        if url:
            try:
                resp = requests.get(url)
                resp.raise_for_status()
                self.image_data = resp.content
                self.log_message("🌐 Картинка прикреплена (URL)")
            except Exception as e:
                self.log_message(f"Ошибка загрузки картинки: {e}", error=True)

    def preview_invite(self):
        preview_win = tk.Toplevel(self.root)
        preview_win.title("Предпросмотр письма")
        text = self.html_widget.get("1.0", tk.END)
        tk.Label(preview_win, text="Тема: " + self.entry_subject.get()).pack()
        sc = scrolledtext.ScrolledText(preview_win, width=100, height=20)
        sc.pack()
        sc.insert(tk.END, text)

    def send_email(self, to_email, to_name):
        try:
            msg = MIMEMultipart()
            msg["From"] = formataddr(("Приглашение", self.entry_email.get()))
            msg["To"] = to_email
            msg["Subject"] = self.entry_subject.get()

            html_content = self.html_widget.get("1.0", tk.END).replace("{name}", to_name)
            msg.attach(MIMEText(html_content, "html"))

            if self.image_data:
                image = MIMEImage(self.image_data)
                image.add_header("Content-ID", "<invitation>")
                msg.attach(image)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(self.entry_email.get(), self.entry_password.get())
                server.send_message(msg)

            self.sent_count += 1
            self.label_stats.config(text=f"📤 Отправлено: {self.sent_count} | ⚠ Ошибок: {self.error_count}")
            self.log_message(f"✅ Успешно: {to_email}")
        except Exception as e:
            self.error_count += 1
            self.label_stats.config(text=f"📤 Отправлено: {self.sent_count} | ⚠ Ошибок: {self.error_count}")
            self.log_message(f"❌ Ошибка {to_email}: {e}", error=True)

    def send_test_email(self):
        email = simpledialog.askstring("Тестовое письмо", "Введите email для теста:")
        if email:
            self.send_email(email, "Тест")

    def send_all(self):
        delay = float(self.delay_entry.get() or 0)
        for guest in self.guests:
            name = guest.get("name", "Гость")
            email = guest.get("email")
            if email:
                self.log_message(f"⏳ Отправка письма: {name} <{email}>")
                self.send_email(email, name)
                time.sleep(delay)
        self.log_message("📨 Рассылка завершена!")


if __name__ == "__main__":
    root = tk.Tk()
    app = MailApp(root)
    root.mainloop()
