import pandas as pd

# Шляхи до файлів
DATA_FILE = "Data.csv"
OUTPUT_FILE = "akcii_clean.xlsx"

# Функція для очистки чисел
def clean_number(x):
    if pd.isna(x):
        return None
    x = str(x).replace(" ", "").replace(",", ".")
    if "%" in x:
        return float(x.replace("%", "")) / 100
    try:
        return float(x)
    except:
        return x

# Зчитуємо CSV
df = pd.read_csv(DATA_FILE, sep=None, engine="python")

# Прибираємо пусті рядки й колонки
df = df.dropna(how="all", axis=0)
df = df.dropna(how="all", axis=1)

# Очищаємо числа
df_clean = df.applymap(clean_number)

# Зберігаємо результат у Excel
df_clean.to_excel(OUTPUT_FILE, index=False)

print(f"Файл {OUTPUT_FILE} збережено")
