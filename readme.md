##Generator Numerów Nadawczych

##Opis

Ten program to graficzny generator numerów nadawczych dla usług pocztowych. Pozwala na generowanie numerów w formacie PX oraz APM, eksportowanie ich do pliku CSV, drukowanie kodów kreskowych w formacie PDF oraz wysyłanie numerów bezpośrednio do innych aplikacji.

##Funkcje

Generowanie numerów PX z cyfrą kontrolną.
Generowanie numerów pomocniczych APM o długości 9 lub 10 znaków.
Eksport wygenerowanych numerów do pliku CSV.
Drukowanie kodów kreskowych w standardzie Code 128 do pliku PDF.
Wysyłanie wybranych numerów do zewnętrznych aplikacji poprzez interfejs HID.
Kopiowanie wybranych numerów do schowka.
Intuicyjny interfejs graficzny z zakładkami dla różnych typów numerów.

##Wymagania

Python 3.x
Zależności wymienione w pliku requirements.txt

##Instalacja

Sklonuj repozytorium lub pobierz pliki.

Zainstaluj wymagane biblioteki:

pip install -r requirements.txt

##Uruchomienie

Uruchom skrypt za pomocą polecenia:

python nazwa_skryptu.py
Upewnij się, że masz zainstalowane wszystkie wymagane biblioteki oraz że używasz odpowiedniej wersji Pythona.

##Użycie

Zakładka PX:

Wprowadź ilość numerów do wygenerowania (1-999).
Opcjonalnie wprowadź trzeci znak prefixu oraz początkowe liczby.
Kliknij "Generuj", aby wygenerować numery.
Możesz eksportować numery do CSV lub drukować kody kreskowe.
Zakładka APM:

Wprowadź ilość numerów pomocniczych do wygenerowania (1-999).
Opcjonalnie wprowadź początkowe liczby.
Wybierz długość numeru (9 lub 10 znaków).
Kliknij "Generuj", aby wygenerować numery.
Możesz eksportować numery do CSV lub drukować kody kreskowe.
Wysyłanie numerów:

W zakładce PX, zaznacz numer na liście i kliknij "Wyślij", aby wysłać numer do zewnętrznej aplikacji.
Pomoc:

Kliknij przycisk "?" w prawym górnym rogu, aby uzyskać informacje o programie.

##Kontakt

Autor: Erwin Słabuszewski
Email: erwin.slabuszewski@gmail.com
GitHub: https://github.com/husk007/generator_PX

Tracking Number Generator

##Description

This program is a graphical tracking number generator for postal services. It allows you to generate numbers in PX and APM formats, export them to CSV files, print barcodes in PDF format, and send numbers directly to other applications.

##Features

Generate PX numbers with a check digit.
Generate auxiliary APM numbers with a length of 9 or 10 characters.
Export generated numbers to a CSV file.
Print barcodes in Code 128 standard to a PDF file.
Send selected numbers to external applications via the HID interface.
Copy selected numbers to the clipboard.
Intuitive graphical interface with tabs for different types of numbers.

##Requirements

Python 3.x
Dependencies listed in requirements.txt

##Installation

Clone the repository or download the files.

Install the required libraries:

pip install -r requirements.txt

##Running the Program

Run the script using the command:

python script_name.py
Make sure all required libraries are installed and you are using the appropriate version of Python.

##Usage

PX Tab:

Enter the number of tracking numbers to generate (1-999).
Optionally, enter the third prefix character and starting numbers.
Click "Generuj" (Generate) to create the numbers.
You can export the numbers to CSV or print barcodes.
APM Tab:

Enter the number of auxiliary numbers to generate (1-999).
Optionally, enter starting numbers.
Choose the length of the number (9 or 10 characters).
Click "Generuj" (Generate) to create the numbers.
You can export the numbers to CSV or print barcodes.
Sending Numbers:

In the PX tab, select a number from the list and click "Wyślij" (Send) to send the number to an external application.
Help:

Click the "?" button in the top right corner to get information about the program.

##Contact

Author: Erwin Słabuszewski
Email: erwin.slabuszewski@gmail.com
GitHub: https://github.com/husk007/generator_PX
