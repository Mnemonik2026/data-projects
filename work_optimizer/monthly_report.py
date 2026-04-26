import xlwings as xw
import pandas as pd
import os
import sys
import glob
import re
import matplotlib.pyplot as plt

def generate(target_month):  # sourcery skip: extract-duplicate-method
    # Фіксація папки
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    SHEET_NAME = 'База'
    START_ROW = 2
    COL_AZS = 'A'
    COL_FUEL = 'B'
    COL_SALES = 'G'

    # Шукаємо всі файли за потрібний місяць
    files = glob.glob(f"*{target_month}.xls*")
    valid_files = []

    for f in files:
        if f.startswith("~$") or "rollover" in f or "Звіт" in f or "Зведена" in f:
            continue
        if match := re.search(r"(\d{2})\.(\d{2})\.(\d{4})", f):
            valid_files.append((match[0], f))

    if not valid_files:
        print(f"❌ Файлів за місяць {target_month} не знайдено!")
        return False

    print(f"📂 Знайдено файлів для обробки: {len(valid_files)}")
    print("⏳ Починаю збір даних...")

    all_dataframes = []
    app = xw.App(visible=False, add_book=False)

    try:
        for date_str, filename in valid_files:
            try:
                wb = app.books.open(filename)
                sheet = wb.sheets[SHEET_NAME]
                last_row = sheet.range(f'{COL_FUEL}1048576').end('up').row

                azs_data = sheet.range(f'{COL_AZS}{START_ROW}:{COL_AZS}{last_row}').options(ndim=1).value
                fuel_data = sheet.range(f'{COL_FUEL}{START_ROW}:{COL_FUEL}{last_row}').options(ndim=1).value
                sales_data = sheet.range(f'{COL_SALES}{START_ROW}:{COL_SALES}{last_row}').options(ndim=1).value
                wb.close()

                df = pd.DataFrame({
                    'Дата': date_str,
                    'АЗС': azs_data,
                    'Вид_пального': fuel_data,
                    'Продажі': sales_data
                })
                df['АЗС'] = df['АЗС'].astype(str).str.strip().replace(['', 'nan', 'None'], pd.NA)
                df['Вид_пального'] = df['Вид_пального'].astype(str).str.strip().replace(['', 'nan', 'None'], pd.NA)
                df['АЗС'] = df['АЗС'].ffill()
                df = df.dropna(subset=['Вид_пального'])

                # Беремо рівно 6 рядків для кожної знайденої АЗС
                df = df.groupby('АЗС').head(6)

                df['Продажі'] = df['Продажі'].astype(str).str.replace(',', '.').str.replace(' ', '')
                df['Продажі'] = pd.to_numeric(df['Продажі'], errors='coerce').fillna(0)

                all_dataframes.append(df)
                print(f"✅ Прочитано: {filename}")

            except Exception as e:
                print(f"⚠️ Помилка у файлі {filename}: {e}")

        app.quit()
    except Exception as e:
        print(f"❌ Критична помилка Excel: {e}")
        try: app.quit() 
        except: pass
        return False

    if not all_dataframes:
        return False

    master_df = pd.concat(all_dataframes, ignore_index=True)

    # Сортуємо дані по даті, щоб на графіках дні йшли по порядку
    master_df['Дата'] = pd.to_datetime(master_df['Дата'], format='%d.%m.%Y')
    master_df = master_df.sort_values(by='Дата')
    master_df['Дата_str'] = master_df['Дата'].dt.strftime('%d.%m')

    # Створюємо папку для збереження звітів
    report_folder = f"Звіт_{target_month}"
    os.makedirs(report_folder, exist_ok=True)

    excel_path = os.path.join(report_folder, f"Місячний_Звіт_{target_month}.xlsx")
    plt.rcParams['font.family'] = 'Segoe UI'

    # Зберігаємо дані в Excel та будуємо графіки
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:

        # 1. ЗАГАЛЬНА АНАЛІТИКА
        pivot_total = pd.pivot_table(
            master_df, values='Продажі', index='АЗС', columns='Вид_пального', aggfunc='sum', fill_value=0
        )
        pivot_total.to_excel(writer, sheet_name="Загальна_Аналітика")

        ax = pivot_total.plot(kind='bar', figsize=(12, 6), width=0.8)
        plt.title(f"Загальні обсяги реалізації пального по АЗС за {target_month}", fontsize=14, fontweight='bold', pad=15)
        plt.xlabel("Назва об'єкту", fontsize=12, labelpad=10)
        plt.ylabel("Продажі (літри / кг)", fontsize=12, labelpad=10)
        plt.xticks(rotation=0, fontsize=11)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.legend(title="Вид пального")
        plt.tight_layout()
        plt.savefig(os.path.join(report_folder, f"0_Загальний_графік_{target_month}.png"), dpi=300)
        plt.close()

        # 2. РОЗБИВКА ПО ДНЯХ ДЛЯ КОЖНОЇ АЗС
        unique_azs = master_df['АЗС'].unique()
        for azs in unique_azs:
            df_azs = master_df[master_df['АЗС'] == azs]

            # Зведена таблиця для конкретної АЗС по днях
            pivot_azs = pd.pivot_table(
                df_azs, values='Продажі', index='Дата_str', columns='Вид_пального', aggfunc='sum', fill_value=0
            )

            # Назва вкладки (обмеження Excel - 31 символ)
            safe_sheet_name = str(azs)[:31]
            pivot_azs.to_excel(writer, sheet_name=safe_sheet_name)

            # Графік для конкретної АЗС
            ax = pivot_azs.plot(kind='bar', figsize=(12, 6), width=0.8)
            plt.title(f"Динаміка продажів по днях: {azs} ({target_month})", fontsize=14, fontweight='bold', pad=15)
            plt.xlabel("Дата", fontsize=12, labelpad=10)
            plt.ylabel("Продажі (літри / кг)", fontsize=12, labelpad=10)
            plt.xticks(rotation=45, fontsize=10)
            plt.grid(axis='y', linestyle='--', alpha=0.7)
            plt.legend(title="Вид пального", fontsize=10)
            plt.tight_layout()

            # Збереження графіка (фільтруємо назву від заборонених символів Windows)
            safe_filename = "".join(c for c in str(azs) if c.isalnum() or c in " _-").strip()
            plt.savefig(os.path.join(report_folder, f"Графік_{safe_filename}.png"), dpi=300)
            plt.close()

    print(f"\n🎉 ГОТОВО! Звіт та графіки збережено у папку: {report_folder}")

    # Відкриваємо створену папку після завершення
    os.startfile(os.path.abspath(report_folder))

    return True

if __name__ == "__main__":
    generate("03.2026")