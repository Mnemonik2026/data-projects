# Interactive Brent Oil Price Analysis Dashboard (Python & Power BI)

Опис
I developed a comprehensive data visualization tool to analyze 30 years of Brent Oil historical price trends (1996–2026). This project demonstrates the full data lifecycle: raw data extraction from exchange API, cleaning with SQL and Python, and interactive reporting in Power BI.

Ключові можливості
- Історичні тренди цін (1996–2026)
- Місячні та річні агрегати, сезонність, волатильність
- Інтерактивні фільтри за періодом, регіоном та індикаторами
- Порівняння показників та експорт звітів

Джерело даних
- API нафтової біржі → CSV
- Попередня обробка: SQL (чистка/агрегація), Python (аналіз, додаткові трансформації)

Технології
- Python (pandas, matplotlib/seaborn, Jupyter)
- SQL (ETL / очищення)
- Power BI (інтеракт��вні звіти)
- CSV (файли даних)

Як відтворити локально
1. Отримайте сирі CSV-файли з API (або імпортуйте через надані скрипти).
2. Запустіть SQL-скрипти для очищення/агрегації.
3. Виконайте Python-ноутбук для додаткової обробки та візуалізацій.
4. Відкрийте Power BI Desktop і завантажте підготовлені CSV/таблиці.

Скріншот
![Brent Oil Dashboard](./images/brent-oil-dashboard-1.png)

Файли в репозиторії
- notebooks/ — Jupyter notebooks (аналіз)
- sql/ — SQL-скрипти для очищення/агрегації
- data/ — очищені CSV (якщо можна зберігати)
- powerbi/ — файл Power BI (.pbix) або експортні звіти

Ліцензія
- MIT