# Generator Numerów Nadawczych

**Autor:** Erwin Słabuszewski  
**Licencja:** MIT

## Opis

"Generator Numerów Nadawczych" to aplikacja napisana w języku Python z interfejsem graficznym opartym na bibliotece Tkinter. Umożliwia generowanie różnych typów numerów przesyłek pocztowych oraz identyfikatorów, takich jak:

- Numery przesyłek 13-znakowych (np. Pocztex, EMS)
- Numery przesyłek 20-znakowych zgodnych ze standardem GS1-128
- Numery PESEL, NIP, REGON i NRB

Aplikacja automatycznie oblicza cyfry kontrolne i zapewnia zgodność wygenerowanych numerów z obowiązującymi standardami. Użytkownik ma możliwość dostosowania parametrów generowania, takich jak prefiksy, sufiksy czy początkowe liczby.

## Funkcjonalności

- **Generowanie numerów 13-znakowych:** Wybór rodzaju przesyłki, ustawienie opcjonalnych parametrów, generowanie numerów z cyfrą kontrolną.
- **Generowanie numerów 20-znakowych:** Zgodność z GS1-128, możliwość wyboru parametrów takich jak S1 i IAC.
- **Generowanie numerów PESEL, NIP, REGON i NRB:** Automatyczne obliczanie cyfr kontrolnych, opcjonalne wprowadzanie danych (np. data urodzenia dla PESEL).
- **Eksport do CSV:** Zapisywanie wygenerowanych numerów do pliku CSV.
- **Drukowanie kodów kreskowych:** Generowanie plików PDF z kodami kreskowymi wygenerowanych numerów.
- **Kopiowanie do schowka:** Możliwość skopiowania wybranych numerów poprzez podwójne kliknięcie.
- **Wysyłanie numerów:** Integracja z zewnętrznymi aplikacjami (np. Poczta+), wysyłanie numerów po kliknięciu przycisku "Wyślij".

## Wymagania

- **Python:** wersja 3.x
- **Biblioteki Python:**
  - `tkinter` (standardowo dostępny w instalacji Pythona)
  - `pyperclip`
  - `python-barcode`
  - `Pillow`
  - `reportlab`
  - `pygetwindow`
  - `pyautogui`
  - `pywin32`

## Instalacja

1. **Klonowanie repozytorium**

   Skopiuj repozytorium na swój komputer:

   ```bash
   git clone https://github.com/husk007/generator_PX.git
   ```

2. **Instalacja wymaganych bibliotek**

   Zainstaluj wymagane biblioteki za pomocą polecenia:

   ```bash
   pip install -r requirements.txt
   ```

   *Jeśli plik `requirements.txt` nie jest dostępny, zainstaluj biblioteki ręcznie:*

   ```bash
   pip install pyperclip python-barcode Pillow reportlab pygetwindow pyautogui pywin32
   ```

## Uruchomienie aplikacji

1. Przejdź do katalogu z plikiem `generator_PX.py`:

   ```bash
   cd generator_PX
   ```

2. Uruchom aplikację:

   ```bash
   python generator_PX.py
   ```

## Instrukcja obsługi

Po uruchomieniu aplikacji pojawi się okno z trzema zakładkami:

### 1. Numery 13 znaków

- **Opis:** Generowanie numerów przesyłek 13-znakowych, takich jak Pocztex, EMS, przesyłki zagraniczne.
- **Użycie:**
  - Wybierz rodzaj przesyłki z dostępnych opcji.
  - Ustaw opcjonalne parametry, takie jak ilość numerów, prefiks, sufiks, początkowe liczby.
  - Kliknij przycisk **"Generuj"**, aby wygenerować numery.
  - Wygenerowane numery pojawią się na liście. Podwójne kliknięcie kopiuje numer do schowka.
  - Zaznaczenie numeru wyświetla jego kod kreskowy poniżej listy.

### 2. Numery 20 znaków

- **Opis:** Generowanie numerów przesyłek 20-znakowych zgodnych ze standardem GS1-128.
- **Użycie:**
  - Wybierz rodzaj przesyłki, a następnie odpowiednie wartości **S1** i **IAC**.
  - Ustaw parametry **Prefix GS1** i **Numer jednostki kodującej**.
  - Kliknij **"Generuj"**.
  - Numery pojawią się na liście, z możliwością kopiowania i podglądu kodu kreskowego.

### 3. Inne

- **Opis:** Generowanie numerów PESEL, NIP, REGON i NRB.
- **Użycie:**
  - Wybierz typ numeru (PESEL, NIP, REGON lub NRB).
  - Opcjonalnie wprowadź dodatkowe dane (np. data urodzenia dla PESEL).
  - Kliknij **"Generuj"**.
  - Wygenerowane numery zostaną wyświetlone na liście.

### Dodatkowe funkcje

- **Eksport do CSV:** Kliknij przycisk **"Eksport do CSV"**, aby zapisać listę numerów do pliku CSV.
- **Drukuj kody kreskowe:** Kliknij **"Drukuj kody kreskowe"**, aby wygenerować plik PDF z kodami kreskowymi.
- **Wyślij:** Przycisk **"Wyślij"** pozwala na przesłanie zaznaczonego numeru do zewnętrznej aplikacji.
- **Pomoc:** Kliknij przycisk **"?"** w prawym górnym rogu, aby wyświetlić informacje o aplikacji.

## Licencja

Ten projekt jest objęty licencją MIT. Pełny tekst licencji znajduje się w pliku [LICENSE](LICENSE).

---

# Shipment Number Generator

**Author:** Erwin Słabuszewski  
**License:** MIT

## Description

The "Shipment Number Generator" is a Python application with a graphical user interface built using Tkinter. It allows generating various types of postal shipment numbers and identifiers, such as:

- 13-character shipment numbers (e.g., Pocztex, EMS)
- 20-character shipment numbers compliant with the GS1-128 standard
- PESEL (Polish national identification number), NIP (tax identification number), REGON (statistical number), and NRB (bank account number)

The application automatically calculates control digits and ensures that the generated numbers comply with the applicable standards. Users can customize generation parameters, such as prefixes, suffixes, or initial numbers.

## Features

- **13-Character Number Generation:** Choose the type of shipment, set optional parameters, generate numbers with control digits.
- **20-Character Number Generation:** GS1-128 compliant, ability to select parameters like S1 and IAC.
- **PESEL, NIP, REGON, and NRB Number Generation:** Automatic control digit calculation, optional data input (e.g., date of birth for PESEL).
- **Export to CSV:** Save generated numbers to a CSV file.
- **Print Barcodes:** Generate PDF files with barcodes of the generated numbers.
- **Copy to Clipboard:** Ability to copy selected numbers by double-clicking.
- **Send Numbers:** Integration with external applications (e.g., Poczta+), send numbers by clicking the "Wyślij" (Send) button.

## Requirements

- **Python:** version 3.x
- **Python Libraries:**
  - `tkinter` (usually included with Python installation)
  - `pyperclip`
  - `python-barcode`
  - `Pillow`
  - `reportlab`
  - `pygetwindow`
  - `pyautogui`
  - `pywin32`

## Installation

1. **Clone the Repository**

   Clone the repository to your computer:

   ```bash
   git clone https://github.com/husk007/generator_PX.git
   ```

2. **Install Required Libraries**

   Install the required libraries using:

   ```bash
   pip install -r requirements.txt
   ```

   *If the `requirements.txt` file is not available, install the libraries manually:*

   ```bash
   pip install pyperclip python-barcode Pillow reportlab pygetwindow pyautogui pywin32
   ```

## Running the Application

1. Navigate to the directory containing `generator_PX.py`:

   ```bash
   cd generator_PX
   ```

2. Run the application:

   ```bash
   python generator_PX.py
   ```

## User Guide

After launching the application, a window with three tabs will appear:

### 1. 13-Character Numbers

- **Description:** Generate 13-character shipment numbers, such as Pocztex, EMS, international shipments.
- **Usage:**
  - Select the type of shipment from the available options.
  - Set optional parameters like the number of numbers, prefix, suffix, initial numbers.
  - Click the **"Generuj" (Generate)** button to generate numbers.
  - Generated numbers will appear in the list. Double-click to copy a number to the clipboard.
  - Selecting a number displays its barcode below the list.

### 2. 20-Character Numbers

- **Description:** Generate 20-character shipment numbers compliant with the GS1-128 standard.
- **Usage:**
  - Choose the type of shipment, then select appropriate **S1** and **IAC** values.
  - Set the **GS1 Prefix** and **Encoding Unit Number** parameters.
  - Click **"Generuj" (Generate)**.
  - Numbers will appear in the list, with options to copy and view barcodes.

### 3. Others

- **Description:** Generate PESEL, NIP, REGON, and NRB numbers.
- **Usage:**
  - Select the type of number (PESEL, NIP, REGON, or NRB).
  - Optionally enter additional data (e.g., date of birth for PESEL).
  - Click **"Generuj" (Generate)**.
  - Generated numbers will be displayed in the list.

### Additional Features

- **Export to CSV:** Click the **"Eksport do CSV"** button to save the list of numbers to a CSV file.
- **Print Barcodes:** Click **"Drukuj kody kreskowe"** to generate a PDF file with barcodes.
- **Send:** The **"Wyślij" (Send)** button allows you to send the selected number to an external application.
- **Help:** Click the **"?"** button in the top-right corner to display application information.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
