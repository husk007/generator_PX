import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyperclip
import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.graphics.barcode import code128
import webbrowser
import subprocess
import pygetwindow as gw
import pyautogui
import win32gui
import win32con
import time

def generuj_numery():
    """Generuje numery nadawcze z cyfrą kontrolną."""
    try:
        ilosc_numerow = int(ilosc_entry.get())
        if not 1 <= ilosc_numerow <= 999:
            raise ValueError
    except ValueError:
        error_label.config(text="Błędna ilość numerów (1-999)")
        numery_listbox.delete(0, tk.END)  # Wyczyść listbox
        return

    prefix = "PX"
    trzeci_znak = trzeci_znak_entry.get().upper()
    if trzeci_znak.isalpha() and len(trzeci_znak) == 1:
        prefix = prefix[:2] + trzeci_znak
    elif trzeci_znak != "":
        error_label.config(text="Błędny trzeci znak prefixu (A-Z)")
        numery_listbox.delete(0, tk.END)  # Wyczyść listbox
        return

    poczatkowe_liczby_str = poczatkowe_liczby_entry.get()
    poczatkowe_liczby_str = poczatkowe_liczby_str.zfill(9)  # Uzupełnij zerami do 9 znaków

    try:
        poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str if x.isdigit()]
        if len(poczatkowe_liczby) > 9:
            raise ValueError
    except ValueError:
        error_label.config(text="Błędne cyfry numeru początkowego")
        numery_listbox.delete(0, tk.END)  # Wyczyść listbox
        return

    poczatkowe_liczby = [0] * (9 - len(poczatkowe_liczby)) + poczatkowe_liczby

    numery_listbox.delete(0, tk.END)  # Wyczyść listbox
    error_label.config(text="")  # Wyczyść komunikat o błędzie

    for i in range(ilosc_numerow):
        # Oblicz cyfrę kontrolną dla bieżącego numeru
        suma_wazona = 0
        wagi = [8, 6, 4, 2, 3, 5, 9, 7, 8]
        for j in range(9):
            suma_wazona += poczatkowe_liczby[j] * wagi[j]
        reszta = suma_wazona % 11
        if reszta == 0:
            cyfra_kontrolna = 5
        elif reszta == 1:
            cyfra_kontrolna = 0
        else:
            cyfra_kontrolna = 11 - reszta

        # Utwórz numer nadawczy
        numer_nadawczy = f"{prefix}{''.join(map(str, poczatkowe_liczby))}{cyfra_kontrolna}"
        numery_listbox.insert(tk.END, numer_nadawczy)

        if i < ilosc_numerow - 1:
            # Zwiększ numer początkowy
            for j in range(8, -1, -1):
                poczatkowe_liczby[j] += 1
                if poczatkowe_liczby[j] < 10:
                    break
                poczatkowe_liczby[j] = 0

def kopiuj_do_schowka(event):
    """Kopiuje wybrany numer nadawczy do schowka."""
    try:
        indices = numery_listbox.curselection()
        if indices:
            numery = [numery_listbox.get(index) for index in indices]
            pyperclip.copy('\n'.join(numery))
            messagebox.showinfo("Skopiowano", "Wybrane numery zostały skopiowane do schowka.")
    except IndexError:
        pass

def eksport_do_csv():
    """Eksportuje numery do pliku CSV."""
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filepath:
        try:
            with open(filepath, mode='w', newline='') as file:
                writer = csv.writer(file)
                for i in range(numery_listbox.size()):
                    writer.writerow([numery_listbox.get(i)])
            messagebox.showinfo("Sukces", "Numery zostały zapisane do pliku CSV.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

def drukuj_kody_kreskowe():
    """Generuje kody kreskowe w pliku PDF w standardzie Code 128."""
    filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if filepath:
        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            x_left, x_right = 40, width / 2 + 20
            y = height - 60
            column = 0
            for i in range(numery_listbox.size()):
                numer = numery_listbox.get(i)
                barcode = code128.Code128(numer, barHeight=40, barWidth=1.5)  # Zwiększ rozmiar kodu kreskowego
                if column == 0:
                    barcode.drawOn(c, x_left, y)
                    c.drawCentredString(x_left + barcode.width / 2, y - 10, numer)
                    column = 1
                else:
                    barcode.drawOn(c, x_right, y)
                    c.drawCentredString(x_right + barcode.width / 2, y - 10, numer)
                    column = 0
                    y -= 80  # Przesuń w dół dla kolejnych kodów
                if y < 60:
                    c.showPage()
                    y = height - 60
            c.save()
            messagebox.showinfo("Sukces", "Kody kreskowe zostały zapisane do pliku PDF.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

def pomoc():
    """Wyświetla informacje o programie."""
    info = (
        "Generator Numerów Nadawczych\n"
        "Autor: Erwin Słabuszewski\n"
        "Licencja: MIT\n"
        "Projekt na GitHub: https://github.com/husk007/generator_PX"
    )
    pomoc_window = tk.Toplevel(root)
    pomoc_window.title("Pomoc")
    pomoc_window.geometry("400x200")
    ttk.Label(pomoc_window, text=info, justify="left").pack(padx=10, pady=10)
    ttk.Label(pomoc_window, text="mailto:erwin.slabuszewski@gmail.com", foreground="blue", cursor="hand2").pack()
    ttk.Label(pomoc_window, text="https://github.com/husk007/generator_PX", foreground="blue", cursor="hand2").pack()

    def open_link(event):
        webbrowser.open_new(event.widget.cget("text"))

    pomoc_window.children["!label2"].bind("<Button-1>", open_link)
    pomoc_window.children["!label3"].bind("<Button-1>", open_link)

def wyslij_numer():
    """Wysyła zaznaczony numer do odpowiedniej aplikacji poprzez interfejs HID."""
    try:
        selected_indices = numery_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Błąd", "Proszę wybrać numer z listy.")
            return

        numer = numery_listbox.get(selected_indices[0])
        prefix = prefix_entry.get() or "@"
        suffix = suffix_entry.get() or "@"
        data_to_send = f"{prefix}{numer}{suffix}"

        # Funkcja do aktywacji okna i wysyłania danych
        def activate_and_send(window_title):
            windows = gw.getWindowsWithTitle(window_title)
            if windows:
                window = windows[0]
                hwnd = window._hWnd
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.5)  # Poczekaj, aż okno stanie się aktywne
                pyautogui.typewrite(data_to_send)
                return True
            return False

        # Próbuj aktywować i wysłać do okna Poczta+
        if not activate_and_send("Poczta+"):
            # Jeśli nie znaleziono, próbuj okna Rejestracja
            if not activate_and_send("Rejestracja"):
                messagebox.showerror("Błąd", "Nie znaleziono odpowiedniej aplikacji.")
    except Exception as e:
        messagebox.showerror("Błąd", f"Wystąpił błąd: {e}")

# --- Nowe funkcje dla zakładki APM ---

def generuj_numery_apm():
    """Generuje numery pomocnicze APM."""
    try:
        ilosc_numerow = int(ilosc_apm_entry.get())
        if not 1 <= ilosc_numerow <= 999:
            raise ValueError
    except ValueError:
        error_apm_label.config(text="Błędna ilość numerów (1-999)")
        numery_apm_listbox.delete(0, tk.END)
        return

    dlugosc_numeru = var_dlugosc.get()
    prefix = "APM"
    liczba_cyfr = 7 if dlugosc_numeru == 10 else 6

    poczatkowe_liczby_str = poczatkowe_liczby_apm_entry.get()
    poczatkowe_liczby_str = poczatkowe_liczby_str.zfill(liczba_cyfr)

    try:
        poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str if x.isdigit()]
        if len(poczatkowe_liczby) > liczba_cyfr:
            raise ValueError
    except ValueError:
        error_apm_label.config(text="Błędne cyfry numeru początkowego")
        numery_apm_listbox.delete(0, tk.END)
        return

    poczatkowe_liczby = [0] * (liczba_cyfr - len(poczatkowe_liczby)) + poczatkowe_liczby

    numery_apm_listbox.delete(0, tk.END)
    error_apm_label.config(text="")

    for i in range(ilosc_numerow):
        numer_pomocniczy = f"{prefix}{''.join(map(str, poczatkowe_liczby))}"
        numery_apm_listbox.insert(tk.END, numer_pomocniczy)

        if i < ilosc_numerow - 1:
            # Zwiększ numer początkowy
            for j in range(liczba_cyfr - 1, -1, -1):
                poczatkowe_liczby[j] += 1
                if poczatkowe_liczby[j] < 10:
                    break
                poczatkowe_liczby[j] = 0

def eksport_do_csv_apm():
    """Eksportuje numery APM do pliku CSV."""
    filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filepath:
        try:
            with open(filepath, mode='w', newline='') as file:
                writer = csv.writer(file)
                for i in range(numery_apm_listbox.size()):
                    writer.writerow([numery_apm_listbox.get(i)])
            messagebox.showinfo("Sukces", "Numery zostały zapisane do pliku CSV.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

def drukuj_kody_kreskowe_apm():
    """Generuje kody kreskowe APM w pliku PDF w standardzie Code 128."""
    filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if filepath:
        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            x_left, x_right = 40, width / 2 + 20
            y = height - 60
            column = 0
            for i in range(numery_apm_listbox.size()):
                numer = numery_apm_listbox.get(i)
                barcode = code128.Code128(numer, barHeight=40, barWidth=1.5)
                if column == 0:
                    barcode.drawOn(c, x_left, y)
                    c.drawCentredString(x_left + barcode.width / 2, y - 10, numer)
                    column = 1
                else:
                    barcode.drawOn(c, x_right, y)
                    c.drawCentredString(x_right + barcode.width / 2, y - 10, numer)
                    column = 0
                    y -= 80
                if y < 60:
                    c.showPage()
                    y = height - 60
            c.save()
            messagebox.showinfo("Sukces", "Kody kreskowe zostały zapisane do pliku PDF.")
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

# --- Interfejs graficzny ---

root = tk.Tk()
root.title("Generator Numerów Nadawczych")
root.geometry("400x550")
root.resizable(False, False)

# Górny frame dla przycisku "Pomoc"
top_frame = ttk.Frame(root)
top_frame.pack(side=tk.TOP, fill=tk.X)

# Przycisk "?"
pomoc_button = ttk.Button(top_frame, text="?", command=pomoc)
pomoc_button.pack(side=tk.RIGHT, padx=10, pady=0)

# Notebook z zakładkami
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both", padx=10, pady=0)

px_frame = ttk.Frame(notebook)
notebook.add(px_frame, text="PX")

apm_frame = ttk.Frame(notebook)
notebook.add(apm_frame, text="APM")

# --- Zakładka PX ---

# Etykiety i pola wprowadzania
ilosc_label = ttk.Label(px_frame, text="Ilość numerów nadania (1-999):")
ilosc_label.grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
ilosc_entry = ttk.Entry(px_frame)
ilosc_entry.grid(row=0, column=1, sticky=tk.E, padx=10, pady=5)

trzeci_znak_label = ttk.Label(px_frame, text="Trzeci znak prefixu (A-Z, opcjonalnie):")
trzeci_znak_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
trzeci_znak_entry = ttk.Entry(px_frame)
trzeci_znak_entry.grid(row=1, column=1, sticky=tk.E, padx=10, pady=5)

poczatkowe_liczby_label = ttk.Label(px_frame, text="Pierwsze liczby (opcjonalnie):")
poczatkowe_liczby_label.grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
poczatkowe_liczby_entry = ttk.Entry(px_frame)
poczatkowe_liczby_entry.grid(row=2, column=1, sticky=tk.E, padx=10, pady=5)

prefix_label = ttk.Label(px_frame, text="Prefix (do wysyłania):")
prefix_label.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
prefix_entry = ttk.Entry(px_frame)
prefix_entry.insert(0, "@")
prefix_entry.grid(row=3, column=1, sticky=tk.E, padx=10, pady=5)

suffix_label = ttk.Label(px_frame, text="Suffix (do wysyłania):")
suffix_label.grid(row=4, column=0, sticky=tk.W, padx=10, pady=5)
suffix_entry = ttk.Entry(px_frame)
suffix_entry.insert(0, "@")
suffix_entry.grid(row=4, column=1, sticky=tk.E, padx=10, pady=5)

# Przycisk generowania
generuj_button = ttk.Button(px_frame, text="Generuj", command=generuj_numery)
generuj_button.grid(row=5, column=0, padx=10, pady=10, sticky=tk.W)

# Przycisk wysyłania numeru
wyslij_button = ttk.Button(px_frame, text="Wyślij", command=wyslij_numer)
wyslij_button.grid(row=5, column=1, padx=10, pady=10, sticky=tk.E)

# Etykieta błędu
error_label = ttk.Label(px_frame, text="", foreground="red")
error_label.grid(row=6, column=0, columnspan=2)

# Listbox z numerami
numery_listbox = tk.Listbox(px_frame, selectmode=tk.EXTENDED, width=50)
numery_listbox.grid(row=7, column=0, columnspan=2, pady=10)
numery_listbox.bind("<<ListboxSelect>>", kopiuj_do_schowka)

# Przycisk eksportu do CSV
eksport_csv_button = ttk.Button(px_frame, text="Eksport do CSV", command=eksport_do_csv)
eksport_csv_button.grid(row=8, column=0, pady=5, padx=10, sticky=tk.W)

# Przycisk drukowania kodów kreskowych
drukuj_kody_button = ttk.Button(px_frame, text="Drukuj kody kreskowe", command=drukuj_kody_kreskowe)
drukuj_kody_button.grid(row=8, column=1, pady=5, padx=10, sticky=tk.E)

# --- Zakładka APM ---

# Etykiety i pola wprowadzania
ilosc_apm_label = ttk.Label(apm_frame, text="Ilość numerów pomocniczych (1-999):")
ilosc_apm_label.grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
ilosc_apm_entry = ttk.Entry(apm_frame)
ilosc_apm_entry.grid(row=0, column=1, sticky=tk.E, padx=10, pady=5)

poczatkowe_liczby_apm_label = ttk.Label(apm_frame, text="Pierwsze liczby (opcjonalnie):")
poczatkowe_liczby_apm_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
poczatkowe_liczby_apm_entry = ttk.Entry(apm_frame)
poczatkowe_liczby_apm_entry.grid(row=1, column=1, sticky=tk.E, padx=10, pady=5)

# Radiobutton dla długości numeru
var_dlugosc = tk.IntVar(value=10)
dlugosc_label = ttk.Label(apm_frame, text="Długość numeru:")
dlugosc_label.grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)
radio_10 = ttk.Radiobutton(apm_frame, text="10 znaków", variable=var_dlugosc, value=10)
radio_10.grid(row=2, column=1, sticky=tk.W, padx=10, pady=5)
radio_9 = ttk.Radiobutton(apm_frame, text="9 znaków", variable=var_dlugosc, value=9)
radio_9.grid(row=3, column=1, sticky=tk.W, padx=10, pady=5)

# Przycisk generowania
generuj_apm_button = ttk.Button(apm_frame, text="Generuj", command=generuj_numery_apm)
generuj_apm_button.grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)

# Etykieta błędu
error_apm_label = ttk.Label(apm_frame, text="", foreground="red")
error_apm_label.grid(row=5, column=0, columnspan=2)

# Listbox z numerami APM
numery_apm_listbox = tk.Listbox(apm_frame, selectmode=tk.EXTENDED, width=50)
numery_apm_listbox.grid(row=6, column=0, columnspan=2, pady=10)

# Przycisk eksportu do CSV
eksport_csv_apm_button = ttk.Button(apm_frame, text="Eksport do CSV", command=eksport_do_csv_apm)
eksport_csv_apm_button.grid(row=7, column=0, pady=5, padx=10, sticky=tk.W)

# Przycisk drukowania kodów kreskowych
drukuj_kody_apm_button = ttk.Button(apm_frame, text="Drukuj kody kreskowe", command=drukuj_kody_kreskowe_apm)
drukuj_kody_apm_button.grid(row=7, column=1, pady=5, padx=10, sticky=tk.E)

root.mainloop()
