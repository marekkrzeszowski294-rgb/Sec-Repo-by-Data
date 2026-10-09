from pathlib import Path

import pandas as pd
import math
import numbers
import re
# Path pomaga zapisać plik wynikowy zawsze w tym samym folderze projektu.

# Definicja danych
# Dane produktów, które zostaną sprawdzone, a następnie zapisane w pliku SQL.
dane = {
    'id_produktu': ['P001', 'P002', 'P003', 'P004', 'P005', 'P006', 'P007', 'P008'],
    'nazwa_produktu': ['Laptop Dell Latitude 5420','Monitor LG 27QN600','Klawiatura mechaniczna Logitech','Mysz bezprzewodowa Logitech','Krzesło biurkowe Ergonomic','Biurko regulowane elektrycznie','Papier ksero A4 (pakiet 5 szt.)','Stacja dokująca USB-C'],
    'kategoria': ['Elektronika', 'Elektronika', 'Akcesoria', 'Akcesoria', 'Meble', 'Meble', 'Biuro', 'Akcesoria'],
    'ilosc_w_magazynie': [45, 120, 310, 450, 25, 14, 600, 85],
    'cena_netto_pln': [3500.00, 950.50, 420.00, 150.00, 899.99, 1850.00, 65.50, 620.00],
    'magazyn': ['Warszawa-Centralny', 'Warszawa-Centralny', 'Poznań-Logistyka', 'Poznań-Logistyka','Wrocław-Magazyn', 'Wrocław-Magazyn', 'Warszawa-Centralny', 'Poznań-Logistyka'],
    'data_ostatniej_aktualizacji': ['2026-10-09', '2026-10-09', '2026-10-08', '2026-10-08', '2026-10-07', '2026-10-07', '2026-10-09', '2026-10-08']
}

df = pd.DataFrame(dane)

# Zamienia pojedynczą wartość Pythona na bezpieczny literał SQL.
def wartosc_sql(wartosc):
    if isinstance(wartosc, str):
        # Znaki sterujące mogłyby utrudnić odczyt lub analizę importowanego SQL.
        if any(ord(znak) < 32 or ord(znak) == 127 for znak in wartosc):
            raise ValueError('Tekst zawiera niedozwolone znaki sterujące.')
        # Podwaja znaki specjalne tekstu, aby nie mogły zakończyć literału SQL.
        wartosc = wartosc.replace('\\', '\\\\').replace("'", "''")
        return f"'{wartosc}'"
    # Do SQL dopuszczamy tylko liczby, nigdy np. fragment kodu lub wartość logiczną.
    if isinstance(wartosc, bool) or not isinstance(wartosc, numbers.Real):
        raise TypeError(f'Niedozwolony typ wartości: {type(wartosc).__name__}')
    # NaN i nieskończoności nie są poprawnymi wartościami kolumn liczbowych SQL.
    if not math.isfinite(wartosc):
        raise ValueError('Wartości liczbowe muszą być skończone.')
    return str(wartosc)


# Sprawdza strukturę i zawartość danych, zanim zostaną umieszczone w zapytaniu.
def sprawdz_dane(ramka):
    # Dozwolony zestaw kolumn jest stały; nazwy kolumn nie pochodzą od użytkownika.
    oczekiwane_kolumny = (
        'id_produktu',
        'nazwa_produktu',
        'kategoria',
        'ilosc_w_magazynie',
        'cena_netto_pln',
        'magazyn',
        'data_ostatniej_aktualizacji',
    )
    if tuple(ramka.columns) != oczekiwane_kolumny:
        raise ValueError('Nieoczekiwane kolumny danych.')

    # Walidacja każdego produktu zapobiega błędnym danym i przekroczeniu limitów tabeli.
    for rekord in ramka.itertuples(index=False, name=None):
        id_produktu, nazwa, kategoria, ilosc, cena, magazyn, data = rekord
        # ID ma ustalony format i nie może zawierać składni SQL.
        if not isinstance(id_produktu, str) or not re.fullmatch(r'P[0-9]{3,9}', id_produktu):
            raise ValueError(f'Nieprawidłowy identyfikator produktu: {id_produktu!r}')
        # Teksty muszą być niepuste, mieścić się w kolumnie i przejść escapowanie SQL.
        for pole, wartosc, limit in (
            ('nazwa_produktu', nazwa, 255),
            ('kategoria', kategoria, 100),
            ('magazyn', magazyn, 150),
        ):
            if not isinstance(wartosc, str) or not wartosc or len(wartosc) > limit:
                raise ValueError(f'Nieprawidłowa wartość pola {pole}.')
            wartosc_sql(wartosc)
        # Ilość musi być nieujemną liczbą całkowitą.
        if isinstance(ilosc, bool) or not isinstance(ilosc, numbers.Integral) or ilosc < 0:
            raise ValueError('Ilość w magazynie musi być nieujemną liczbą całkowitą.')
        # Cena musi być skończoną, nieujemną liczbą.
        if isinstance(cena, bool) or not isinstance(cena, numbers.Real) or not math.isfinite(cena) or cena < 0:
            raise ValueError('Cena musi być nieujemną, skończoną liczbą.')
        # Najpierw sprawdzamy format daty, a potem czy taka data istnieje.
        if not isinstance(data, str) or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', data):
            raise ValueError('Data musi mieć format RRRR-MM-DD.')
        try:
            pd.Timestamp(data)
        except (TypeError, ValueError) as blad:
            raise ValueError(f'Nieprawidłowa data: {data!r}') from blad


# Zatrzymuje program od razu, jeśli choć jeden rekord nie spełnia reguł.
sprawdz_dane(df)

# Ścieżki są liczone od położenia tego skryptu, nie od katalogu terminala.
katalog_wynikowy = Path(__file__).resolve().parent.parent / 'data'
sciezka_sql = katalog_wynikowy / 'Plik_Magazyn.sql'
sciezka_csv = katalog_wynikowy / 'Plik_Magazyn.csv'
sciezka_xlsx = katalog_wynikowy / 'Plik_Magazyn.xlsx'
# Tworzy folder data, jeśli jeszcze go nie ma.
katalog_wynikowy.mkdir(parents=True, exist_ok=True)

# Przygotowuje stałe nazwy kolumn i wiersze danych do zapytania INSERT.
kolumny = list(df.columns)
lista_kolumn = ', '.join(f'`{kolumna}`' for kolumna in kolumny)
wiersze = [
    f"({', '.join(wartosc_sql(wartosc) for wartosc in wiersz)})"
    for wiersz in df.itertuples(index=False, name=None)
]
aktualizacje = ',\n'.join(
    f"    `{kolumna}` = VALUES(`{kolumna}`)"
    for kolumna in kolumny
    if kolumna != 'id_produktu'
)

# Składa skrypt importu: baza, tabela, dane oraz bezpieczna aktualizacja istniejących ID.
sql = f"""CREATE DATABASE IF NOT EXISTS `magazyn`
    CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE `magazyn`;

CREATE TABLE IF NOT EXISTS `produkty` (
    `id_produktu` VARCHAR(10) NOT NULL PRIMARY KEY,
    `nazwa_produktu` VARCHAR(255) NOT NULL,
    `kategoria` VARCHAR(100) NOT NULL,
    `ilosc_w_magazynie` INT NOT NULL,
    `cena_netto_pln` DECIMAL(10, 2) NOT NULL,
    `magazyn` VARCHAR(150) NOT NULL,
    `data_ostatniej_aktualizacji` DATE NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

INSERT INTO `produkty` ({lista_kolumn})
VALUES
{',\n'.join(wiersze)}
ON DUPLICATE KEY UPDATE
{aktualizacje};
"""

# Zapisuje eksporty w UTF-8, aby polskie znaki były poprawne w każdym formacie.
sciezka_sql.write_text(sql, encoding='utf-8')
df.to_csv(sciezka_csv, index=False, encoding='utf-8-sig')
df.to_excel(sciezka_xlsx, index=False)

print(f"Wygenerowano plik SQL: {sciezka_sql}")
print(f"Wygenerowano plik CSV: {sciezka_csv}")
print(f"Wygenerowano plik XLSX: {sciezka_xlsx}")