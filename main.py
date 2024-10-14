import pandas as pd
import numpy as np

import argparse

# Funkce pro převod čísla sloupce na písmeno
def get_column_letter(col_idx):
    letter = ''
    while col_idx >= 0:
        letter = chr(col_idx % 26 + 65) + letter
        col_idx = col_idx // 26 - 1
    return letter

# Nastavení argumentů příkazové řádky
parser = argparse.ArgumentParser(description='Vyhledání obsahu v Excelovém souboru.')
parser.add_argument('filename', type=str, help='Název Excelového souboru')
parser.add_argument('search_term', type=str, help='Hledaný obsah')
args = parser.parse_args()

# Načtení Excelového souboru
df = pd.read_excel(args.filename, engine='openpyxl')


# Hledaný obsah
hledany_obsah = args.search_term

# Vyhledání buňky
bunky = df.isin([hledany_obsah])

# Získání indexů řádků a sloupců, kde se obsah nachází
pozice = list(zip(*np.where(bunky)))

# Výpis pozic na obrazovku
for radek_idx, sloupec_idx in pozice:
    radek = radek_idx + 1  # Přidání 1, protože indexy jsou od 0
    sloupec = get_column_letter(sloupec_idx)
    print(f"Obsah nalezen na řádku: {radek}, sloupci: {sloupec}")