import tkinter as tk
import customtkinter as ctk 
import requests 
import subprocess
import threading
import os
import sys
import config
import monthly_report
import re
import importlib

# Налаштування загального вигляду
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Налаштування")
        self.geometry("450x600") # Збільшено висоту для нового елемента
        
        self.transient(parent)
        self.grab_set()

        # База налаштувань для різних поштовиків
        self.providers = {
            "Ukr.net": {"SMTP_SERVER": "smtp.ukr.net", "SMTP_PORT": 465, "IMAP_SERVER": "imap.ukr.net", "IMAP_PORT": 993},
            "Gmail": {"SMTP_SERVER": "smtp.gmail.com", "SMTP_PORT": 465, "IMAP_SERVER": "imap.gmail.com", "IMAP_PORT": 993},
            "Meta.ua": {"SMTP_SERVER": "smtp.meta.ua", "SMTP_PORT": 465, "IMAP_SERVER": "imap.meta.ua", "IMAP_PORT": 993}
        }

        # --- 1. Вибір поштового провайдера ---
        lbl_prov = ctk.CTkLabel(self, text="Поштовий сервіс відправника", font=("Arial", 12, "bold"))
        lbl_prov.grid(row=0, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.combo_provider = ctk.CTkComboBox(self, values=list(self.providers.keys()), width=400)
        
        # Автоматично обираємо провайдера у списку на основі поточного конфігу
        current_smtp = getattr(config, "SMTP_SERVER", "smtp.ukr.net")
        for prov_name, prov_data in self.providers.items():
            if prov_data["SMTP_SERVER"] == current_smtp:
                self.combo_provider.set(prov_name)
                break
                
        self.combo_provider.grid(row=1, column=0, padx=20, pady=(0, 5), sticky="w")

        # --- 2. Інші текстові поля ---
        self.fields = {
            "EMAIL_USER": "Email відправника",
            "EMAIL_PASS": "Пароль додатку",
            "TARGET_SENDER": "Email для звітів (від кого чекаємо)",
            "RECIPIENTS": "Отримувачі (через кому)",
            "TG_TOKEN": "Telegram Token",
            "TG_CHAT_ID": "Telegram Chat ID",
        }
        self.entries = {}

        # Зміщення рядків через доданий ComboBox (починаємо з row=2)
        for i, (key, label_text) in enumerate(self.fields.items()):
            row_idx = i * 2 + 2 
            lbl = ctk.CTkLabel(self, text=label_text, font=("Arial", 12, "bold"))
            lbl.grid(row=row_idx, column=0, padx=20, pady=(10, 0), sticky="w")
            
            ent = ctk.CTkEntry(self, width=400)
            
            val = getattr(config, key, "")
            if isinstance(val, list):
                val = ", ".join(val)
                
            ent.insert(0, str(val))
            ent.grid(row=row_idx+1, column=0, padx=20, pady=(0, 5), sticky="w")
            self.entries[key] = ent

        btn_save = ctk.CTkButton(self, text="Зберегти", fg_color="green", hover_color="darkgreen", command=self.save_config)
        btn_save.grid(row=len(self.fields)*2 + 2, column=0, pady=20)

    def save_config(self):
        config_path = "config.py"
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 1. Зберігаємо текстові поля та списки
            for key, ent in self.entries.items():
                new_val = ent.get().strip()
                
                if key == "RECIPIENTS":
                    emails = [e.strip() for e in new_val.split(",") if e.strip()]
                    formatted_list = "[\n    " + ",\n    ".join([f'"{e}"' for e in emails]) + "\n]"
                    pattern = rf'^{key}\s*=\s*\[.*?\]'
                    content = re.sub(pattern, f'{key} = {formatted_list}', content, flags=re.DOTALL | re.MULTILINE)
                else:
                    pattern = rf'^({key}\s*=\s*)(["\']).*?(["\'])'
                    content = re.sub(pattern, rf'\g<1>"{new_val}"', content, flags=re.MULTILINE)

            # 2. Зберігаємо налаштування серверів (SMTP / IMAP)
            selected_provider = self.combo_provider.get()
            provider_settings = self.providers[selected_provider]
            
            for key, val in provider_settings.items():
                if isinstance(val, int):
                    # Регулярка для чисел (наприклад порти: 465)
                    pattern = rf'^({key}\s*=\s*)\d+'
                    content = re.sub(pattern, rf'\g<1>{val}', content, flags=re.MULTILINE)
                else:
                    # Регулярка для тексту (наприклад сервера: "smtp.gmail.com")
                    pattern = rf'^({key}\s*=\s*)(["\']).*?(["\'])'
                    content = re.sub(pattern, rf'\g<1>"{val}"', content, flags=re.MULTILINE)

            # Запис у файл та оновлення пам'яті
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            importlib.reload(config)
            self.destroy()
        except Exception as e:
            print(f"Помилка збереження: {e}")

class AZSControlPanel(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- НАЛАШТУВАННЯ ГОЛОВНОГО ВІКНА ---
        self.title("АЗС CONTROL PANEL")
        self.geometry("700x750")
        
        # Виклик функції, яка малює всі елементи на екрані
        self._setup_ui()

    def _setup_ui(self):
        """Функція для створення та розміщення всіх віджетів (кнопок, тексту)"""

        # ==========================================
        #              БЛОК 1: ПОГОДА
        # ==========================================
        self.frame_weather = ctk.CTkFrame(self)
        self.frame_weather.pack(pady=10, padx=20, fill="x")
        
        self.lbl_weather = ctk.CTkLabel(self.frame_weather, text="Погода в Ніжині", font=("Arial", 14, "bold"))
        self.lbl_weather.pack(pady=5)
        
        weather_text = self.get_weather()

        self.weather_label = ctk.CTkLabel(
            self.frame_weather, 
            text=weather_text, 
            font=("Arial", 14, "bold"),
            text_color="#96e911" 
        )
        self.weather_label.pack(pady=10)
        
        # ==========================================
        #            БЛОК 1.5: КУРСИ ВАЛЮТ
        # ==========================================
        self.frame_rates = ctk.CTkFrame(self)
        self.frame_rates.pack(pady=5, padx=20, fill="x") # pady=5 щоб було ближче до погоди
        
        self.lbl_rates = ctk.CTkLabel(self.frame_rates, text="Курси валют АТ ПриватБанк", font=("Arial", 14, "bold"))
        self.lbl_rates.pack(pady=5)
        
        rates_text = self.get_exchange_rates()
        
        self.lbl_rates = ctk.CTkLabel(
            self.frame_rates, 
            text=rates_text, 
            font=("Arial", 14, "bold"),
            text_color="#F39C12" # Золотистий колір для грошей
        )
        self.lbl_rates.pack(pady=10)
        
        # ==========================================
        #             НАЛАШТУВАННЯ
        # ==========================================
        self.btn_settings = ctk.CTkButton(
            self, 
            text="⚙️ Налаштування", 
            fg_color="#555555", 
            hover_color="#333333",
            command=self.open_settings
        )
        self.btn_settings.pack(pady=10)
        
        # ==========================================
        #            БЛОК 2: КНОПКИ КЕРУВАННЯ
        # ==========================================
        self.frame_evening = ctk.CTkFrame(self)
        self.frame_evening.pack(pady=10, padx=20, fill="x")
        
        self.lbl_evening = ctk.CTkLabel(self.frame_evening, text="Робота", font=("Arial", 14, "bold"))
        self.lbl_evening.pack(pady=5)
        
        # Зберігаємо всі кнопки у змінні (self.btn_...), щоб потім мати змогу змінювати їхній стан
        self.btn_start = ctk.CTkButton(
            self.frame_evening, text="Почати Зміну", text_color="black",
            command=self.run_morning, fg_color="#2CC985", hover_color="#209664", cursor="hand2"
        )
        self.btn_start.pack(pady=10, padx=20, fill="x")
        
        self.btn_mail = ctk.CTkButton(
            self.frame_evening, text="Отримати Пошту", text_color="black",
            command=self.run_mail, fg_color="#E5A00D", hover_color="#B37D0A", cursor="hand2"
        )
        self.btn_mail.pack(pady=10, padx=20, fill="x")
        
        self.btn_tg = ctk.CTkButton(
            self.frame_evening, text="Telegram ✈️", text_color="black",
            command=self.run_telegram, fg_color="#3390EC", hover_color="#2870B7", cursor="hand2"
        )
        self.btn_tg.pack(pady=5, padx=20, fill="x")

        self.btn_email = ctk.CTkButton(
            self.frame_evening, text="Email 📧", text_color="black",
            command=self.run_email, fg_color="#C92C2C", hover_color="#962020", cursor="hand2"
        )
        self.btn_email.pack(pady=10, padx=20, fill="x")

        self.btn_report = ctk.CTkButton(
            self.frame_evening, text="📊 Місячний звіт", text_color="black",
            command=self.run_monthly_report, fg_color="#08DCAB", hover_color="#78E48C", cursor="hand2"
        )
        self.btn_report.pack(pady=5, padx=18, fill="x") 

        # ==========================================
        #             БЛОК 3: КОНСОЛЬ
        # ==========================================
        self.log_box = ctk.CTkTextbox(self, height=200)
        self.log_box.pack(pady=20, padx=20, fill="x")
        self.log_box.insert("0.0", "Система готова.\n")
        self.log_box.configure(state="disabled")

    # ==========================================
    #            СИСТЕМНІ ФУНКЦІЇ
    # ==========================================
    def log(self, message):
        """Функція для безпечного виведення тексту у вбудовану консоль панелі"""
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end") # Автоматичне прокручування вниз
        self.log_box.configure(state="disabled")

    def _execute_subprocess(self, script_name):
        """
        Внутрішня функція для синхронного запуску інших Python-файлів.
        Вона викликається всередині окремих потоків, щоб не вішати інтерфейс.
        """
        path = os.path.join(config.BASE_DIR, script_name)
        self.log(f"🚀 Старт: {script_name}")
        
        si = None
        cf = 0
        if os.name == 'nt':
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            cf = subprocess.CREATE_NO_WINDOW

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"

        try:
            p = subprocess.Popen(
                [sys.executable, path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',
                cwd=config.BASE_DIR,
                startupinfo=si,
                creationflags=cf,
                env=env
            )
            
            for line in p.stdout:
                self.log(f" > {line.strip()}")
            
            for err in p.stderr:
                self.log(f"❌ {err.strip()}")

            p.wait()
            self.log(f"✅ {script_name} готово!\n---")
            
        except Exception as e:
            self.log(f"❌ Помилка запуску {script_name}: {e}")

    # ==========================================
    #            ФУНКЦІЇ КНОПОК (ПОТОКИ)
    # ==========================================
    
    def open_settings(self):
        SettingsWindow(self)
    
    def run_morning(self):
        # 1. Змінюємо вигляд кнопки
        self.btn_start.configure(state="disabled", text="⏳ Виконую...", fg_color="gray")
        
        # 2. Описуємо важку задачу
        def task():
            try:
                # ЕТАП 1: Безпека (Резервна копія)
                # self._execute_subprocess("backup.py")
                # time.sleep(2) 

                # ЕТАП 2: Аналітика (Збереження продажів)
                # self._execute_subprocess("save_sales.py")
                # time.sleep(2) 
                
                # ЕТАП 3: Новий день (Очищення і зміна дати)
                self._execute_subprocess("rollover_v2.py")
            finally:
                # 3. Повертаємо кнопку в нормальний стан після завершення
                # .after(0, ...) потрібен для безпечного оновлення UI з іншого потоку
                self.after(0, lambda: self.btn_start.configure(state="normal", text="Почати Зміну", fg_color="#2CC985"))

        # 4. Запускаємо задачу у фоні
        threading.Thread(target=task, daemon=True).start()

    def run_mail(self):
        self.btn_mail.configure(state="disabled", text="⏳ Отримую...", fg_color="gray")
        
        def task():
            try:
                self._execute_subprocess("get_mail.py")
            finally:
                self.after(0, lambda: self.btn_mail.configure(state="normal", text="Отримати Пошту", fg_color="#E5A00D"))
                
        threading.Thread(target=task, daemon=True).start()

    def run_telegram(self):
        self.btn_tg.configure(state="disabled", text="⏳ Відправляю...", fg_color="gray")
        
        def task():
            try:
                self._execute_subprocess("telegram_send.py")
            finally:
                self.after(0, lambda: self.btn_tg.configure(state="normal", text="Telegram ✈️", fg_color="#3390EC"))
                
        threading.Thread(target=task, daemon=True).start()

    def run_email(self):
        self.btn_email.configure(state="disabled", text="⏳ Відправляю...", fg_color="gray")
        
        def task():
            try:
                self._execute_subprocess("send_report.py")
            finally:
                self.after(0, lambda: self.btn_email.configure(state="normal", text="Email 📧", fg_color="#C92C2C"))
                
        threading.Thread(target=task, daemon=True).start()

    def run_monthly_report(self):
        # Отримуємо місяць від користувача
        dialog = ctk.CTkInputDialog(text="Введіть місяць та рік (наприклад: 03.2026):", title="Місячний звіт")
        target_month = dialog.get_input()

        if target_month:
            # Блокуємо кнопку
            self.btn_report.configure(state="disabled", text="⏳ Рахую...", fg_color="gray")
            
            def task():
                self.log(f"🚀 Старт: Збір даних за {target_month}")
                try:
                    monthly_report.generate(target_month)
                    self.log(f"✅ Місячний звіт успішно згенеровано!\n---")
                except Exception as e:
                    self.log(f"❌ Помилка побудови звіту: {e}\n---")
                finally:
                    # Розблоковуємо кнопку
                    self.after(0, lambda: self.btn_report.configure(state="normal", text="📊 Місячний звіт", fg_color="#08DCAB"))

            threading.Thread(target=task, daemon=True).start()
            
    # ==========================================
    #                 ПОГОДА
    # ==========================================
    def get_weather(self):  # sourcery skip: extract-method, inline-variable
        # Додано &timezone=auto для уникнення помилок із часом
        url = "https://api.open-meteo.com/v1/forecast?latitude=51.04&longitude=31.88&current=temperature_2m,relative_humidity_2m,surface_pressure,wind_speed_10m&timezone=auto"
        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            
            # Якщо Open-Meteo повертає помилку, виводимо її в консоль
            if "error" in data:
                print(f"Помилка API Open-Meteo: {data.get('reason', 'Невідома помилка')}")
                return "🌡 Погода: Помилка API"
            
            current = data['current']
            temp = current['temperature_2m']
            humidity = current['relative_humidity_2m']
            wind = current['wind_speed_10m']
            
            pressure_hpa = current['surface_pressure']
            pressure_mm = int(pressure_hpa * 0.75) 
            
            return f"🌡 {temp}°C | 💧 {humidity}% | ⏬ {pressure_mm} мм рт.ст. | 💨 {wind} км/год"
            
        except Exception as e:
            print(f"Помилка обробки погоди: {e}")
            return "🌡 Погода: Немає зв'язку"
        
    # ==========================================
    #                 КУРСИ ВАЛЮТ
    # ==========================================
    def get_exchange_rates(self):
        # sourcery skip: extract-method, remove-unnecessary-else, swap-if-else-branches
        # API ПриватБанку
        url = "https://api.privatbank.ua/p24api/pubinfo?json&exchange&coursid=5"
        try:
            response = requests.get(url, timeout=5)
            data = response.json()
            
            usd_data = next((item for item in data if item['ccy'] == 'USD'), None)
            eur_data = next((item for item in data if item['ccy'] == 'EUR'), None)
            
            if usd_data and eur_data:
                # Отримуємо значення купівлі та продажу
                usd_buy = float(usd_data['buy'])
                usd_sale = float(usd_data['sale'])
                
                eur_buy = float(eur_data['buy'])
                eur_sale = float(eur_data['sale'])
                
                # Форматуємо вивід: Купівля / Продаж
                return f"💵 USD: {usd_buy:.2f} / {usd_sale:.2f} ₴  |  💶 EUR: {eur_buy:.2f} / {eur_sale:.2f} ₴"
            else:
                return "Дані валют недоступні"
                
        except Exception as e:
            print(f"Помилка курсів валют: {e}")
            return "💵 Курс: Немає зв'язку"
# Запуск програми
if __name__ == "__main__":
    app = AZSControlPanel()
    app.mainloop()