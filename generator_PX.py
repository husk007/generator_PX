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
import random
import datetime

# Dodane importy
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageTk

# --- Wspólne funkcje ---

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
    ttk.Label(
        pomoc_window, text="mailto:erwin.slabuszewski@gmail.com", foreground="blue", cursor="hand2").pack()
    ttk.Label(
        pomoc_window, text="https://github.com/husk007/generator_PX", foreground="blue", cursor="hand2").pack()

    def open_link(event):
        webbrowser.open_new(event.widget.cget("text"))

    pomoc_window.children["!label2"].bind("<Button-1>", open_link)
    pomoc_window.children["!label3"].bind("<Button-1>", open_link)

def kopiuj_do_schowka(event):
    """Kopiuje wybrany numer nadawczy do schowka po podwójnym kliknięciu."""
    listbox = event.widget
    try:
        indices = listbox.curselection()
        if indices:
            numery = [listbox.get(index) for index in indices]
            pyperclip.copy('\n'.join(numery))
            messagebox.showinfo("Skopiowano", "Wybrane numery zostały skopiowane do schowka.")
    except IndexError:
        pass

def activate_and_send(data_to_send):
    """Aktywuje okno aplikacji i wysyła dane."""
    def try_activate(window_title):
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

    if not try_activate("Poczta+"):
        if not try_activate("Rejestracja"):
            messagebox.showerror("Błąd", "Nie znaleziono odpowiedniej aplikacji.")

def eksport_do_csv(listbox):
    """Eksportuje numery z listboxa do pliku CSV."""
    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filepath:
        try:
            with open(filepath, mode='w', newline='') as file:
                writer = csv.writer(file)
                for i in range(listbox.size()):
                    writer.writerow([listbox.get(i)])
            messagebox.showinfo("Sukces", "Numery zostały zapisane do pliku CSV.")
        except Exception as e:
            messagebox.showerror(
                "Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

def drukuj_kody_kreskowe(listbox, numer_formatowany=False):
    """Generuje kody kreskowe w pliku PDF w standardzie Code 128."""
    filepath = filedialog.asksaveasfilename(
        defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if filepath:
        try:
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            x_left, x_right = 40, width / 2 + 20
            y = height - 60
            column = 0
            for i in range(listbox.size()):
                numer = listbox.get(i)
                numer_kod = numer[4:] if numer_formatowany else numer
                barcode = code128.Code128(
                    numer_kod, barHeight=40, barWidth=1.5)
                if column == 0:
                    barcode.drawOn(c, x_left, y)
                    c.drawCentredString(
                        x_left + barcode.width / 2, y - 10, numer)
                    column = 1
                else:
                    barcode.drawOn(c, x_right, y)
                    c.drawCentredString(
                        x_right + barcode.width / 2, y - 10, numer)
                    column = 0
                    y -= 80
                if y < 60:
                    c.showPage()
                    y = height - 60
            c.save()
            messagebox.showinfo(
                "Sukces", "Kody kreskowe zostały zapisane do pliku PDF.")
        except Exception as e:
            messagebox.showerror(
                "Błąd", f"Wystąpił problem podczas zapisywania pliku: {e}")

def create_entry_with_label(parent, label_text, default_value="", row=None):
    """Tworzy etykietę i pole wprowadzania."""
    label = ttk.Label(parent, text=label_text)
    entry = ttk.Entry(parent)
    entry.insert(0, default_value)
    if row is not None:
        label.grid(row=row, column=0, sticky=tk.W, padx=10, pady=2)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=10, pady=2)
    return entry

# --- Klasa dla zakładki "Numery 13 znaków" ---

class Numery13ZnakowTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.var_dlugosc = tk.IntVar(value=10)
        self.przesylki_13znakow_list = [
            'Pocztex 2.0',
            'Pocztex',
            'EMS',
            'Paczka zagraniczna (w tym z zadeklarowaną wartością)',
            'Paczka zagraniczna dla UE',
            'Paczka zagraniczna dla Ukrainy',
            'Paczka zagraniczna dla poczty niemieckiej',
            'List polecony zagraniczny',
            'List wartościowy zagraniczny',
            'Global Expres',
            'PUH',
            'numery pomocnicze APM'
        ]
        self.przesylka_option_13znakow = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        self.ilosc_entry = create_entry_with_label(self, "Ilość numerów (1-999):", row=0)
        self.poczatkowe_liczby_entry = create_entry_with_label(self, "Pierwsze liczby (opcjonalnie):", row=1)
        self.prefix_entry = create_entry_with_label(self, "Prefix (do wysyłania):", "@", row=2)
        self.suffix_entry = create_entry_with_label(self, "Suffix (do wysyłania):", "@", row=3)

        przesylka_label = ttk.Label(self, text="Wybierz rodzaj przesyłki:")
        przesylka_label.grid(row=4, column=0, sticky=tk.W, padx=10, pady=2)
        przesylka_frame = ttk.Frame(self)
        przesylka_frame.grid(row=5, column=0, columnspan=2, sticky=tk.W, padx=10)

        for idx, przesylka in enumerate(self.przesylki_13znakow_list):
            rb = ttk.Radiobutton(przesylka_frame, text=przesylka, variable=self.przesylka_option_13znakow, value=przesylka, command=self.on_przesylka_selected_13znakow)
            rb.grid(row=idx // 2, column=idx % 2, sticky=tk.W, padx=5, pady=2)

        # Pola konfiguracyjne
        self.trzeci_znak_label = ttk.Label(self, text="Trzeci znak prefixu (A-Z, opcjonalnie):")
        self.trzeci_znak_entry = ttk.Entry(self)

        self.rodzaj_przesylki_label = ttk.Label(self, text="Rodzaj przesyłki:")
        self.rodzaj_przesylki_entry = ttk.Entry(self)

        self.kraj_nadania_label = ttk.Label(self, text="Kraj nadania przesyłki:")
        self.kraj_nadania_entry = ttk.Entry(self)

        self.dlugosc_label = ttk.Label(self, text="Długość numeru:")
        self.radio_frame = ttk.Frame(self)
        radio_10 = ttk.Radiobutton(self.radio_frame, text="10 znaków", variable=self.var_dlugosc, value=10)
        radio_10.pack(side=tk.LEFT)
        radio_9 = ttk.Radiobutton(self.radio_frame, text="9 znaków", variable=self.var_dlugosc, value=9)
        radio_9.pack(side=tk.LEFT)

        self.dlugosc_puh_label = ttk.Label(self, text="Długość numeru (10-20):")
        self.dlugosc_puh_entry = ttk.Entry(self)

        self.rodzaj_przesylki_puh_label = ttk.Label(self, text="Rodzaj przesyłki:")
        self.rodzaj_przesylki_puh_entry = ttk.Entry(self)
        self.rodzaj_przesylki_puh_entry.insert(0, "PUH")

        # Przycisk generowania
        generuj_button = ttk.Button(self, text="Generuj", command=self.generuj_numery)
        generuj_button.grid(row=10, column=0, padx=10, pady=10, sticky=tk.W)

        # Przycisk wysyłania numeru
        wyslij_button = ttk.Button(self, text="Wyślij", command=self.wyslij_numer)
        wyslij_button.grid(row=10, column=1, padx=10, pady=10, sticky=tk.E)

        # Etykieta błędu
        self.error_label = ttk.Label(self, text="", foreground="red")
        self.error_label.grid(row=11, column=0, columnspan=2)

        # Listbox z numerami
        self.numery_listbox = tk.Listbox(self, selectmode=tk.EXTENDED, width=50)
        self.numery_listbox.grid(row=12, column=0, columnspan=2, pady=10)
        self.numery_listbox.bind("<Double-Button-1>", kopiuj_do_schowka)
        # Dodane bindowanie zdarzenia wyboru
        self.numery_listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        # Label do wyświetlania kodu kreskowego
        self.barcode_label = ttk.Label(self)
        self.barcode_label.grid(row=13, column=0, columnspan=2, pady=10)

        # Przyciski eksportu i drukowania
        eksport_csv_button = ttk.Button(self, text="Eksport do CSV", command=lambda: eksport_do_csv(self.numery_listbox))
        eksport_csv_button.grid(row=11, column=0, pady=5, padx=10, sticky=tk.W)

        drukuj_kody_button = ttk.Button(self, text="Drukuj kody kreskowe", command=lambda: drukuj_kody_kreskowe(self.numery_listbox))
        drukuj_kody_button.grid(row=11, column=1, pady=5, padx=10, sticky=tk.E)

    def on_przesylka_selected_13znakow(self):
        """Aktualizuje widok po wybraniu rodzaju przesyłki."""
        selected = self.przesylka_option_13znakow.get()
        # Ukryj wszystkie dodatkowe pola
        self.trzeci_znak_label.grid_remove()
        self.trzeci_znak_entry.grid_remove()
        self.rodzaj_przesylki_label.grid_remove()
        self.rodzaj_przesylki_entry.grid_remove()
        self.kraj_nadania_label.grid_remove()
        self.kraj_nadania_entry.grid_remove()
        self.var_dlugosc.set(10)
        self.dlugosc_label.grid_remove()
        self.radio_frame.grid_remove()
        self.dlugosc_puh_label.grid_remove()
        self.dlugosc_puh_entry.grid_remove()
        self.rodzaj_przesylki_puh_label.grid_remove()
        self.rodzaj_przesylki_puh_entry.grid_remove()

        if selected == 'Pocztex 2.0':
            self.trzeci_znak_label.grid(row=8, column=0, sticky=tk.W, padx=10, pady=2)
            self.trzeci_znak_entry.grid(row=8, column=1, sticky=tk.EW, padx=10, pady=2)
        elif selected in ['Pocztex', 'EMS', 'Paczka zagraniczna (w tym z zadeklarowaną wartością)',
                          'Paczka zagraniczna dla UE', 'Paczka zagraniczna dla Ukrainy',
                          'Paczka zagraniczna dla poczty niemieckiej', 'List polecony zagraniczny',
                          'List wartościowy zagraniczny', 'Global Expres']:
            self.rodzaj_przesylki_label.grid(row=8, column=0, sticky=tk.W, padx=10, pady=2)
            self.rodzaj_przesylki_entry.grid(row=8, column=1, sticky=tk.EW, padx=10, pady=2)
            self.kraj_nadania_label.grid(row=9, column=0, sticky=tk.W, padx=10, pady=2)
            self.kraj_nadania_entry.grid(row=9, column=1, sticky=tk.EW, padx=10, pady=2)
            if selected == 'Pocztex':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "EE")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'EMS':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "EE")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'Paczka zagraniczna (w tym z zadeklarowaną wartością)':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "CP")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'List polecony zagraniczny':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "RR")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'List wartościowy zagraniczny':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "VV")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'Paczka zagraniczna dla poczty niemieckiej':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "CR")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'Paczka zagraniczna dla UE':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "CZ")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'Paczka zagraniczna dla Ukrainy':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "CU")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
            elif selected == 'Global Expres':
                self.rodzaj_przesylki_entry.delete(0, tk.END)
                self.rodzaj_przesylki_entry.insert(0, "LX")
                self.kraj_nadania_entry.delete(0, tk.END)
                self.kraj_nadania_entry.insert(0, "PL")
        elif selected == 'numery pomocnicze APM':
            self.dlugosc_label.grid(row=8, column=0, sticky=tk.W, padx=10, pady=2)
            self.radio_frame.grid(row=8, column=1, sticky=tk.W, padx=10, pady=2)
        elif selected == 'PUH':
            self.dlugosc_puh_label.grid(row=8, column=0, sticky=tk.W, padx=10, pady=2)
            self.dlugosc_puh_entry.grid(row=8, column=1, sticky=tk.EW, padx=10, pady=2)
            self.rodzaj_przesylki_puh_label.grid(row=9, column=0, sticky=tk.W, padx=10, pady=2)
            self.rodzaj_przesylki_puh_entry.grid(row=9, column=1, sticky=tk.EW, padx=10, pady=2)

    def wyslij_numer(self):
        """Wysyła zaznaczony numer z listy do odpowiedniej aplikacji."""
        selected_indices = self.numery_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Błąd", "Proszę wybrać numer z listy.")
            return
        numer = self.numery_listbox.get(selected_indices[0])
        prefix = self.prefix_entry.get() or "@"
        suffix = self.suffix_entry.get() or "@"
        data_to_send = f"{prefix}{numer}{suffix}"
        activate_and_send(data_to_send)

    def generuj_numery(self):
        """Generuje numery nadawcze z cyfrą kontrolną."""
        try:
            ilosc_numerow = int(self.ilosc_entry.get())
            if not 1 <= ilosc_numerow <= 999:
                raise ValueError
        except ValueError:
            self.error_label.config(text="Błędna ilość numerów (1-999)")
            self.numery_listbox.delete(0, tk.END)
            return

        selected_option = self.przesylka_option_13znakow.get()
        if not selected_option:
            self.error_label.config(text="Proszę wybrać rodzaj przesyłki")
            self.numery_listbox.delete(0, tk.END)
            return

        poczatkowe_liczby_str = self.poczatkowe_liczby_entry.get()
        poczatkowe_liczby_str = ''.join(filter(str.isdigit, poczatkowe_liczby_str))

        if selected_option == 'numery pomocnicze APM':
            dlugosc_numeru = self.var_dlugosc.get()
            liczba_cyfr = 7 if dlugosc_numeru == 10 else 6
            poczatkowe_liczby_str = poczatkowe_liczby_str.zfill(liczba_cyfr)

            try:
                poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str]
            except ValueError:
                self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
                self.numery_listbox.delete(0, tk.END)
                return

            max_value = 10 ** liczba_cyfr - ilosc_numerow
            current_value = int(''.join(map(str, poczatkowe_liczby)))
            if current_value > max_value:
                self.error_label.config(text=f"Numer początkowy przekracza maksymalną wartość dla numeru pomocniczego APM ({max_value})")
                self.numery_listbox.delete(0, tk.END)
                return

        elif selected_option == 'PUH':
            poczatkowe_liczby_str = poczatkowe_liczby_str.lstrip('0')

            try:
                poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str]
            except ValueError:
                self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
                self.numery_listbox.delete(0, tk.END)
                return

            try:
                dlugosc_puh = int(self.dlugosc_puh_entry.get())
                if not 10 <= dlugosc_puh <= 20:
                    raise ValueError
            except ValueError:
                self.error_label.config(text="Długość numeru PUH musi być liczbą od 10 do 20")
                self.numery_listbox.delete(0, tk.END)
                return

            rodzaj_przesylki_puh = self.rodzaj_przesylki_puh_entry.get()

            num_zeros = dlugosc_puh - len(rodzaj_przesylki_puh) - len(poczatkowe_liczby_str) - 1  # -1 dla cyfry kontrolnej

            if num_zeros < 0:
                self.error_label.config(text="Numer początkowy jest zbyt długi dla wybranej długości")
                self.numery_listbox.delete(0, tk.END)
                return

            max_numer = int('9' * len(poczatkowe_liczby_str))
            current_value = int(poczatkowe_liczby_str) + ilosc_numerow - 1
            if current_value > max_numer:
                self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę dla PUH")
                self.numery_listbox.delete(0, tk.END)
                return

        else:
            # Dla pozostałych przesyłek
            poczatkowe_liczby_str = poczatkowe_liczby_str.zfill(8)

            try:
                poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str]
            except ValueError:
                self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
                self.numery_listbox.delete(0, tk.END)
                return

            max_value = 99999999 - ilosc_numerow + 1
            current_value = int(''.join(map(str, poczatkowe_liczby)))
            if current_value > max_value:
                self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
                self.numery_listbox.delete(0, tk.END)
                return

        self.numery_listbox.delete(0, tk.END)
        self.error_label.config(text="")

        for i in range(ilosc_numerow):
            if selected_option == 'numery pomocnicze APM':
                prefix = "APM"
                liczba_cyfr = 7 if self.var_dlugosc.get() == 10 else 6
                numer_cyfry = ''.join(map(str, poczatkowe_liczby)).zfill(liczba_cyfr)
                numer_nadawczy = f"{prefix}{numer_cyfry}"
            elif selected_option == 'PUH':
                numer_cyfry = ''.join(map(str, poczatkowe_liczby))
                num_zeros = dlugosc_puh - len(rodzaj_przesylki_puh) - len(numer_cyfry) - 1  # -1 dla cyfry kontrolnej

                numer_bez_k = rodzaj_przesylki_puh + ('0' * num_zeros) + numer_cyfry

                # Pobierz część numeryczną (bez prefixu)
                start_idx = len(rodzaj_przesylki_puh)
                digits_part = numer_bez_k[start_idx:]

                # Oblicz wagi w zależności od długości części numerycznej
                weights_length = len(digits_part)
                wagi = [8, 6, 4, 2, 3, 5, 9, 7] * ((weights_length // 8) + 1)
                wagi = wagi[:weights_length]

                # Oblicz sumę ważoną
                suma_wazona = 0
                for j in range(weights_length):
                    suma_wazona += int(digits_part[j]) * wagi[j]

                reszta = suma_wazona % 11
                if reszta == 0:
                    cyfra_kontrolna = '5'
                elif reszta == 1:
                    cyfra_kontrolna = '0'
                else:
                    cyfra_kontrolna = str(11 - reszta)
                numer_nadawczy = numer_bez_k + cyfra_kontrolna
            elif selected_option == 'Pocztex 2.0':
                prefix = "PX"
                trzeci_znak = self.trzeci_znak_entry.get().upper()
                if trzeci_znak.isalpha() and len(trzeci_znak) == 1:
                    prefix = prefix[:2] + trzeci_znak
                elif trzeci_znak != "":
                    self.error_label.config(text="Błędny trzeci znak prefixu (A-Z)")
                    self.numery_listbox.delete(0, tk.END)
                    return
                numer_cyfry = ''.join(map(str, poczatkowe_liczby))

                # Sprawdź, czy numer nie przekracza maksymalnej wartości
                max_value = 99999999 - ilosc_numerow + 1
                current_value = int(numer_cyfry)
                if current_value > max_value:
                    self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę dla Pocztex 2.0")
                    self.numery_listbox.delete(0, tk.END)
                    return

                # Oblicz cyfrę kontrolną
                suma_wazona = 0
                wagi = [8, 6, 4, 2, 3, 5, 9, 7]
                for j in range(8):
                    suma_wazona += int(numer_cyfry[j]) * wagi[j]
                reszta = suma_wazona % 11
                if reszta == 0:
                    cyfra_kontrolna = 5
                elif reszta == 1:
                    cyfra_kontrolna = 0
                else:
                    cyfra_kontrolna = 11 - reszta
                numer_nadawczy = f"{prefix}{numer_cyfry}{cyfra_kontrolna}"
            else:
                rodzaj_przesylki = self.rodzaj_przesylki_entry.get().upper()
                kraj_nadania = self.kraj_nadania_entry.get().upper()
                if len(rodzaj_przesylki) != 2 or not rodzaj_przesylki.isalpha():
                    self.error_label.config(text="Rodzaj przesyłki musi mieć 2 litery")
                    self.numery_listbox.delete(0, tk.END)
                    return
                if len(kraj_nadania) != 2 or not kraj_nadania.isalpha():
                    self.error_label.config(text="Kraj nadania musi mieć 2 litery")
                    self.numery_listbox.delete(0, tk.END)
                    return

                numer_cyfry = ''.join(map(str, poczatkowe_liczby))

                # Sprawdź, czy numer nie przekracza maksymalnej wartości
                max_value = 99999999 - ilosc_numerow + 1
                current_value = int(numer_cyfry)
                if current_value > max_value:
                    self.error_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
                    self.numery_listbox.delete(0, tk.END)
                    return

                # Oblicz cyfrę kontrolną
                suma_wazona = 0
                wagi = [8, 6, 4, 2, 3, 5, 9, 7]
                for j in range(8):
                    suma_wazona += int(numer_cyfry[j]) * wagi[j]
                reszta = suma_wazona % 11
                if reszta == 0:
                    cyfra_kontrolna = 5
                elif reszta == 1:
                    cyfra_kontrolna = 0
                else:
                    cyfra_kontrolna = 11 - reszta
                numer_nadawczy = f"{rodzaj_przesylki}{numer_cyfry}{cyfra_kontrolna}{kraj_nadania}"

            self.numery_listbox.insert(tk.END, numer_nadawczy)

            if i < ilosc_numerow - 1:
                # Zwiększ numer początkowy
                idx = len(poczatkowe_liczby) - 1
                while idx >= 0:
                    poczatkowe_liczby[idx] += 1
                    if poczatkowe_liczby[idx] < 10:
                        break
                    poczatkowe_liczby[idx] = 0
                    idx -= 1
                else:
                    self.error_label.config(text="Przekroczono maksymalną liczbę dla podanej długości")
                    break

    def on_listbox_select(self, event):
        """Wyświetla kod kreskowy wybranego numeru."""
        selected_indices = self.numery_listbox.curselection()
        if not selected_indices:
            self.barcode_label.config(image='')
            return
        selected_index = selected_indices[0]
        numer = self.numery_listbox.get(selected_index)
        # Generowanie obrazu kodu kreskowego
        barcode_image = self.generate_barcode_image(numer)
        if barcode_image:
            self.barcode_photo = ImageTk.PhotoImage(barcode_image)
            self.barcode_label.config(image=self.barcode_photo)
        else:
            self.barcode_label.config(image='')

    def generate_barcode_image(self, numer):
        """Generuje obraz kodu kreskowego dla danego numeru."""
        try:
            barcode = Code128(numer, writer=ImageWriter())
            barcode_image = barcode.render(writer_options={'module_height': 10.0, 'module_width': 0.2, 'quiet_zone': 1.0})
            return barcode_image
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie można wygenerować kodu kreskowego: {e}")
            return None

# --- Klasa dla zakładki "Numery 20 znaków" ---

class Numery20ZnakowTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.przesylki_data = {
            'przesyłka polecona': {
                'S1': ['1', '4', '0'],
                'IAC': {
                    '1': [str(i) for i in range(1, 9)],
                    '4': [str(i) for i in range(1, 10)],
                    '0': [str(i) for i in range(1, 5)] + ['9']
                }
            },
            'przesyłka listowa z zadeklarowaną wartością': {
                'S1': ['2'],
                'IAC': {
                    '2': [str(i) for i in range(1, 9)] + ['9']
                }
            },
            'paczka pocztowa (w tym z zadeklarowaną wartością)': {
                'S1': ['3'],
                'IAC': {
                    '3': [str(i) for i in range(1, 9)] + ['9']
                }
            },
            'przesyłka Paczka+': {
                'S1': ['5'],
                'IAC': {
                    '5': ['1', '8', '9']
                }
            },
            'paczka 24, paczka ekstra 24, paczka 48': {
                'S1': ['6'],
                'IAC': {
                    '6': [str(i) for i in range(3, 9)] + ['9']
                }
            },
            'przesyłka paletowa': {
                'S1': ['7'],
                'IAC': {
                    '7': [str(i) for i in range(4, 9)]
                }
            },
            'pocztex kurier 48, pocztex kurier 48 pobranie': {
                'S1': ['8'],
                'IAC': {
                    '8': ['1', '2', '3', '4', '5']
                }
            },
            'pozostałe usługi listowe lub reklamowe': {
                'S1': ['9'],
                'IAC': {
                    '9': ['1', '2', '3', '9']
                }
            },
            'przesyłka hybrydowa': {
                'S1': [],  # Brak wyboru S1
                'IAC': {}  # Brak wyboru IAC
            },
        }
        self.przesylka_option = tk.StringVar()
        self.s1_option = tk.StringVar()
        self.iac_option = tk.StringVar()
        self.init_ui()

    def init_ui(self):
        self.ilosc_przesylek_entry = create_entry_with_label(self, "Ilość numerów (1-999):", row=0)
        self.gs1_prefix_entry = create_entry_with_label(self, "Prefix GS1 (domyślnie 590):", "590", row=1)
        self.numer_jednostki_entry = create_entry_with_label(self, "Numer jednostki kodującej (domyślnie 0773):", "0773", row=2)
        self.poczatkowe_liczby_przesylek_entry = create_entry_with_label(self, "Pierwsze liczby (opcjonalnie):", row=3)
        self.prefix_przesylki_entry = create_entry_with_label(self, "Prefix (do wysyłania):", "@", row=4)
        self.suffix_przesylki_entry = create_entry_with_label(self, "Suffix (do wysyłania):", "@", row=5)

        przesylka_label = ttk.Label(self, text="Wybierz rodzaj przesyłki:")
        przesylka_label.grid(row=6, column=0, sticky=tk.W, padx=10, pady=2)
        przesylka_frame = ttk.Frame(self)
        przesylka_frame.grid(row=7, column=0, columnspan=2, sticky=tk.W, padx=10)

        for idx, przesylka in enumerate(self.przesylki_data.keys()):
            rb = ttk.Radiobutton(przesylka_frame, text=przesylka, variable=self.przesylka_option, value=przesylka, command=self.on_przesylka_selected)
            rb.grid(row=idx // 2, column=idx % 2, sticky=tk.W, padx=5, pady=2)

        # Frame dla S1
        self.s1_label = ttk.Label(self, text="Wybierz S1:")
        self.s1_label.grid(row=8, column=0, sticky=tk.W, padx=10, pady=2)
        self.s1_frame = ttk.Frame(self)
        self.s1_frame.grid(row=9, column=0, columnspan=2, sticky=tk.W, padx=10)

        # Frame dla IAC
        self.iac_label = ttk.Label(self, text="Wybierz IAC:")
        self.iac_label.grid(row=10, column=0, sticky=tk.W, padx=10, pady=2)
        self.iac_frame = ttk.Frame(self)
        self.iac_frame.grid(row=11, column=0, columnspan=2, sticky=tk.W, padx=10)

        # Przycisk generowania
        generuj_button = ttk.Button(self, text="Generuj", command=self.generuj_numery_przesylek)
        generuj_button.grid(row=12, column=0, padx=10, pady=10, sticky=tk.W)

        # Przycisk wysyłania numeru
        wyslij_button = ttk.Button(self, text="Wyślij", command=self.wyslij_numer_przesylki)
        wyslij_button.grid(row=12, column=1, padx=10, pady=10, sticky=tk.E)

        # Etykieta błędu
        self.error_przesylki_label = ttk.Label(self, text="", foreground="red")
        self.error_przesylki_label.grid(row=13, column=0, columnspan=2)

        # Listbox z numerami
        self.numery_przesylek_listbox = tk.Listbox(self, selectmode=tk.EXTENDED, width=70)
        self.numery_przesylek_listbox.grid(row=14, column=0, columnspan=2, pady=10)
        self.numery_przesylek_listbox.bind("<Double-Button-1>", kopiuj_do_schowka)
        # Dodane bindowanie zdarzenia wyboru
        self.numery_przesylek_listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        # Label do wyświetlania kodu kreskowego
        self.barcode_label = ttk.Label(self, width=70)
        self.barcode_label.grid(row=15, column=0, columnspan=2, pady=10)

        # Przyciski eksportu i drukowania
        eksport_csv_button = ttk.Button(self, text="Eksport do CSV", command=lambda: eksport_do_csv(self.numery_przesylek_listbox))
        eksport_csv_button.grid(row=13, column=0, pady=5, padx=10, sticky=tk.W)

        drukuj_kody_button = ttk.Button(self, text="Drukuj kody kreskowe", command=lambda: drukuj_kody_kreskowe(self.numery_przesylek_listbox, numer_formatowany=True))
        drukuj_kody_button.grid(row=13, column=1, pady=5, padx=10, sticky=tk.E)

    def on_przesylka_selected(self):
        """Aktualizuje dostępne opcje S1 i IAC po wybraniu przesyłki."""
        selected = self.przesylka_option.get()
        self.s1_option.set('')
        self.iac_option.set('')
        # Usunięcie istniejących radiobuttonów S1 i IAC
        for widget in self.s1_frame.winfo_children():
            widget.destroy()
        for widget in self.iac_frame.winfo_children():
            widget.destroy()

        if selected == 'przesyłka hybrydowa':
            # Ukryj ramki S1 i IAC
            self.s1_label.grid_remove()
            self.s1_frame.grid_remove()
            self.iac_label.grid_remove()
            self.iac_frame.grid_remove()
        else:
            # Przywróć ramki S1 i IAC
            self.s1_label.grid()
            self.s1_frame.grid()
            self.iac_label.grid()
            self.iac_frame.grid()
            s1_values = self.przesylki_data[selected]['S1']
            s1_label_list = [f"S1 = {s1}" for s1 in s1_values]
            for idx, s1_label_text in enumerate(s1_label_list):
                rb = ttk.Radiobutton(
                    self.s1_frame, text=s1_label_text, variable=self.s1_option,
                    value=s1_values[idx], command=self.on_s1_selected)
                rb.pack(side=tk.LEFT, padx=5)

    def on_s1_selected(self):
        """Aktualizuje dostępne opcje IAC po wybraniu S1."""
        selected = self.przesylka_option.get()
        selected_s1 = self.s1_option.get()
        iac_values = self.przesylki_data[selected]['IAC'][selected_s1]
        self.iac_option.set('')
        # Usunięcie istniejących radiobuttonów IAC
        for widget in self.iac_frame.winfo_children():
            widget.destroy()
        # Stworzenie nowych radiobuttonów IAC
        for idx, iac in enumerate(iac_values):
            label_text = f"IAC = {iac} (EN)" if iac == '9' else f"IAC = {iac}"
            rb = ttk.Radiobutton(
                self.iac_frame, text=label_text, variable=self.iac_option, value=iac)
            rb.pack(side=tk.LEFT, padx=5)

    def wyslij_numer_przesylki(self):
        """Wysyła zaznaczony numer do odpowiedniej aplikacji."""
        selected_indices = self.numery_przesylek_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Błąd", "Proszę wybrać numer z listy.")
            return

        numer = self.numery_przesylek_listbox.get(selected_indices[0])
        prefix = self.prefix_przesylki_entry.get() or "@"
        suffix = self.suffix_przesylki_entry.get() or "@"
        data_to_send = f"{prefix}{numer}{suffix}"
        activate_and_send(data_to_send)
        
    def generuj_numery_przesylek(self):
        """Generuje numery przesyłek zgodne z GS1-128."""
        try:
            ilosc_numerow = int(self.ilosc_przesylek_entry.get())
            if not 1 <= ilosc_numerow <= 999:
                raise ValueError
        except ValueError:
            self.error_przesylki_label.config(text="Błędna ilość numerów (1-999)")
            self.numery_przesylek_listbox.delete(0, tk.END)
            return

        gs1_prefix = self.gs1_prefix_entry.get().strip()
        if not gs1_prefix.isdigit() or len(gs1_prefix) != 3:
            self.error_przesylki_label.config(text="Błędny Prefix GS1 (3 cyfry)")
            self.numery_przesylek_listbox.delete(0, tk.END)
            return

        numer_jednostki = self.numer_jednostki_entry.get().strip()
        if not numer_jednostki.isdigit() or len(numer_jednostki) != 4:
            self.error_przesylki_label.config(
                text="Błędny Numer jednostki kodującej (4 cyfry)")
            self.numery_przesylek_listbox.delete(0, tk.END)
            return

        poczatkowe_liczby_str = self.poczatkowe_liczby_przesylek_entry.get()
        poczatkowe_liczby_str = poczatkowe_liczby_str.zfill(8)

        try:
            poczatkowe_liczby = [int(x) for x in poczatkowe_liczby_str if x.isdigit()]
            if len(poczatkowe_liczby) > 8:
                raise ValueError
        except ValueError:
            self.error_przesylki_label.config(text="Przekroczono maksymalną długość numeru, zmniejsz początkową liczbę")
            self.numery_przesylek_listbox.delete(0, tk.END)
            return

        poczatkowe_liczby = [0] * (8 - len(poczatkowe_liczby)) + poczatkowe_liczby

        selected_option = self.przesylka_option.get()
        if not selected_option:
            self.error_przesylki_label.config(text="Proszę wybrać rodzaj przesyłki")
            self.numery_przesylek_listbox.delete(0, tk.END)
            return

        if selected_option == 'przesyłka hybrydowa':
            selected_S1 = '0'
            selected_IAC = '1'
        else:
            selected_S1 = self.s1_option.get()
            selected_IAC = self.iac_option.get()
            if not selected_S1 or not selected_IAC:
                self.error_przesylki_label.config(text="Proszę wybrać S1 i IAC")
                self.numery_przesylek_listbox.delete(0, tk.END)
                return

        self.numery_przesylek_listbox.delete(0, tk.END)
        self.error_przesylki_label.config(text="")

        for i in range(ilosc_numerow):
            numer_bez_k = f"{selected_IAC}{gs1_prefix}{numer_jednostki}" \
                          f"{selected_S1}{''.join(map(str, poczatkowe_liczby))}"
            # Sprawdzenie zakresów dla przesyłki hybrydowej
            if selected_option == 'przesyłka hybrydowa':
                numer_int = int(''.join(map(str, poczatkowe_liczby)))
                if not (1 <= numer_int <= 10000000):
                    self.error_przesylki_label.config(
                        text="Numer poza zakresem dla przesyłki hybrydowej (1-10000000)")
                    return
            # Sprawdzenie zakresu dla przesyłki poleconej z IAC=1 i S1=0
            elif selected_option == 'przesyłka polecona' and selected_IAC == '1' and selected_S1 == '0':
                numer_int = int(''.join(map(str, poczatkowe_liczby)))
                if 1 <= numer_int <= 10000000:
                    self.error_przesylki_label.config(
                        text="Zakres numerów zarezerwowany dla przesyłki hybrydowej")
                    return

            cyfra_kontrolna = self.oblicz_cyfre_kontrolna_gs1(numer_bez_k)
            pelny_numer = f"(00){numer_bez_k}{cyfra_kontrolna}"
            self.numery_przesylek_listbox.insert(tk.END, pelny_numer)

            if i < ilosc_numerow - 1:
                for j in range(7, -1, -1):
                    poczatkowe_liczby[j] += 1
                    if poczatkowe_liczby[j] < 10:
                        break
                    poczatkowe_liczby[j] = 0

    def oblicz_cyfre_kontrolna_gs1(self, numer_bez_k):
        """Oblicza cyfrę kontrolną według algorytmu GS1."""
        suma = 0
        odwrocony_numer = numer_bez_k[::-1]
        for i, cyfra in enumerate(odwrocony_numer):
            cyfra = int(cyfra)
            if i % 2 == 0:
                suma += cyfra * 3
            else:
                suma += cyfra * 1
        kontrolna = (10 - (suma % 10)) % 10
        return str(kontrolna)

    def on_listbox_select(self, event):
        """Wyświetla kod kreskowy wybranego numeru."""
        selected_indices = self.numery_przesylek_listbox.curselection()
        if not selected_indices:
            self.barcode_label.config(image='')
            return
        selected_index = selected_indices[0]
        numer = self.numery_przesylek_listbox.get(selected_index)
        # Usuwamy (00) z numeru dla kodu kreskowego
        numer_kod = numer[4:] if numer.startswith('(00)') else numer
        # Generowanie obrazu kodu kreskowego
        barcode_image = self.generate_barcode_image(numer_kod)
        if barcode_image:
            self.barcode_photo = ImageTk.PhotoImage(barcode_image)
            self.barcode_label.config(image=self.barcode_photo)
        else:
            self.barcode_label.config(image='')

    def generate_barcode_image(self, numer):
        """Generuje obraz kodu kreskowego dla danego numeru."""
        try:
            barcode = Code128(numer, writer=ImageWriter())
            barcode_image = barcode.render(writer_options={'module_height': 10.0, 'module_width': 0.2, 'quiet_zone': 1.0})
            return barcode_image
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie można wygenerować kodu kreskowego: {e}")
            return None

def create_entry_with_label(parent, label_text, default_value="", row=None, column=0, columnspan=1):
    """Tworzy etykietę i pole wprowadzania."""
    label = ttk.Label(parent, text=label_text)
    entry = ttk.Entry(parent)
    entry.insert(0, default_value)
    if row is not None:
        label.grid(row=row, column=column, sticky=tk.W, padx=10, pady=2)
        entry.grid(row=row, column=column+1, sticky=tk.EW, padx=10, pady=2, columnspan=columnspan)
    return entry

# --- Klasa dla zakładki "Inne" ---

class InneTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.typ_numeru = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        # Wybór typu numeru
        typ_label = ttk.Label(self, text="Wybierz typ numeru:")
        typ_label.grid(row=0, column=0, sticky=tk.W, padx=10, pady=2)
        typ_frame = ttk.Frame(self)
        typ_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=10)

        typy_numerow = ['PESEL', 'NIP', 'REGON', 'NRB']
        for idx, typ in enumerate(typy_numerow):
            rb = ttk.Radiobutton(
                typ_frame, text=typ, variable=self.typ_numeru, value=typ, command=self.on_typ_selected)
            rb.grid(row=0, column=idx, sticky=tk.W, padx=5, pady=2)

        # Pola wspólne
        self.ilosc_entry = create_entry_with_label(self, "Ilość numerów (1-999):", row=2)
        self.prefix_entry = create_entry_with_label(self, "Prefix (do wysyłania):", "@", row=3)
        self.suffix_entry = create_entry_with_label(self, "Suffix (do wysyłania):", "@", row=4)

        # Pola specyficzne
        self.data_urodzenia_label = ttk.Label(self, text="Data urodzenia (RRRRMMDD):")
        self.data_urodzenia_entry = ttk.Entry(self)

        self.nip_kod_urzedu_label = ttk.Label(self, text="Kod urzędu skarbowego (3 cyfry):")
        self.nip_kod_urzedu_entry = ttk.Entry(self)

        self.regon_poczatkowe_label = ttk.Label(self, text="Początkowe cyfry REGON (7 lub 8 cyfr):")
        self.regon_poczatkowe_entry = ttk.Entry(self)

        self.nrb_bank_label = ttk.Label(self, text="Numer banku (8 cyfr):")
        self.nrb_bank_entry = ttk.Entry(self)

        self.nrb_rachunek_label = ttk.Label(self, text="Numer rachunku klienta (16 cyfr):")
        self.nrb_rachunek_entry = ttk.Entry(self)

        # Przycisk generowania
        generuj_button = ttk.Button(self, text="Generuj", command=self.generuj_numery)
        generuj_button.grid(row=10, column=0, padx=10, pady=10, sticky=tk.W)

        # Przycisk wysyłania numeru
        wyslij_button = ttk.Button(self, text="Wyślij", command=self.wyslij_numer)
        wyslij_button.grid(row=10, column=1, padx=10, pady=10, sticky=tk.E)

        # Etykieta błędu
        self.error_label = ttk.Label(self, text="", foreground="red")
        self.error_label.grid(row=11, column=0, columnspan=2)

        # Listbox z numerami
        self.numery_listbox = tk.Listbox(self, selectmode=tk.EXTENDED, width=70)
        self.numery_listbox.grid(row=12, column=0, columnspan=2, pady=10)
        self.numery_listbox.bind("<Double-Button-1>", kopiuj_do_schowka)
        self.numery_listbox.bind('<<ListboxSelect>>', self.on_listbox_select)

        # Label do wyświetlania kodu kreskowego
        self.barcode_label = ttk.Label(self)
        self.barcode_label.grid(row=13, column=0, columnspan=2, pady=10)

        # Przyciski eksportu i drukowania
        eksport_csv_button = ttk.Button(self, text="Eksport do CSV",
                                        command=lambda: eksport_do_csv(self.numery_listbox))
        eksport_csv_button.grid(row=14, column=0, pady=5, padx=10, sticky=tk.W)

        drukuj_kody_button = ttk.Button(self, text="Drukuj kody kreskowe",
                                        command=lambda: drukuj_kody_kreskowe(self.numery_listbox))
        drukuj_kody_button.grid(row=14, column=1, pady=5, padx=10, sticky=tk.E)

    def on_typ_selected(self):
        """Aktualizuje widok po wybraniu typu numeru."""
        selected = self.typ_numeru.get()
        # Ukryj wszystkie pola specyficzne
        self.data_urodzenia_label.grid_remove()
        self.data_urodzenia_entry.grid_remove()
        self.nip_kod_urzedu_label.grid_remove()
        self.nip_kod_urzedu_entry.grid_remove()
        self.regon_poczatkowe_label.grid_remove()
        self.regon_poczatkowe_entry.grid_remove()
        self.nrb_bank_label.grid_remove()
        self.nrb_bank_entry.grid_remove()
        self.nrb_rachunek_label.grid_remove()
        self.nrb_rachunek_entry.grid_remove()

        if selected == 'PESEL':
            self.data_urodzenia_label.grid(row=5, column=0, sticky=tk.W, padx=10, pady=2)
            self.data_urodzenia_entry.grid(row=5, column=1, sticky=tk.EW, padx=10, pady=2)
        elif selected == 'NIP':
            self.nip_kod_urzedu_label.grid(row=5, column=0, sticky=tk.W, padx=10, pady=2)
            self.nip_kod_urzedu_entry.grid(row=5, column=1, sticky=tk.EW, padx=10, pady=2)
        elif selected == 'REGON':
            self.regon_poczatkowe_label.grid(row=5, column=0, sticky=tk.W, padx=10, pady=2)
            self.regon_poczatkowe_entry.grid(row=5, column=1, sticky=tk.EW, padx=10, pady=2)
        elif selected == 'NRB':
            self.nrb_bank_label.grid(row=5, column=0, sticky=tk.W, padx=10, pady=2)
            self.nrb_bank_entry.grid(row=5, column=1, sticky=tk.EW, padx=10, pady=2)
            self.nrb_rachunek_label.grid(row=6, column=0, sticky=tk.W, padx=10, pady=2)
            self.nrb_rachunek_entry.grid(row=6, column=1, sticky=tk.EW, padx=10, pady=2)

    def wyslij_numer(self):
        """Wysyła zaznaczony numer z listy do odpowiedniej aplikacji."""
        selected_indices = self.numery_listbox.curselection()
        if not selected_indices:
            messagebox.showerror("Błąd", "Proszę wybrać numer z listy.")
            return
        numer = self.numery_listbox.get(selected_indices[0])
        prefix = self.prefix_entry.get() or "@"
        suffix = self.suffix_entry.get() or "@"
        data_to_send = f"{prefix}{numer}{suffix}"
        activate_and_send(data_to_send)

    def generuj_numery(self):
        """Generuje numery zgodnie z wybranym typem."""
        try:
            ilosc_numerow = int(self.ilosc_entry.get())
            if not 1 <= ilosc_numerow <= 999:
                raise ValueError
        except ValueError:
            self.error_label.config(text="Błędna ilość numerów (1-999)")
            self.numery_listbox.delete(0, tk.END)
            return

        selected_typ = self.typ_numeru.get()
        if not selected_typ:
            self.error_label.config(text="Proszę wybrać typ numeru")
            self.numery_listbox.delete(0, tk.END)
            return

        self.numery_listbox.delete(0, tk.END)
        self.error_label.config(text="")

        if selected_typ == 'PESEL':
            self.generuj_pesel(ilosc_numerow)
        elif selected_typ == 'NIP':
            self.generuj_nip(ilosc_numerow)
        elif selected_typ == 'REGON':
            self.generuj_regon(ilosc_numerow)
        elif selected_typ == 'NRB':
            self.generuj_nrb(ilosc_numerow)

    def generuj_pesel(self, ilosc):
        """Generuje numery PESEL."""
        data_urodzenia_input = self.data_urodzenia_entry.get()
        if data_urodzenia_input:
            if not data_urodzenia_input.isdigit() or len(data_urodzenia_input) != 8:
                self.error_label.config(text="Błędna data urodzenia (RRRRMMDD)")
                return

            rok = int(data_urodzenia_input[:4])
            miesiac = int(data_urodzenia_input[4:6])
            dzien = int(data_urodzenia_input[6:8])

            if not (1900 <= rok <= 2299) or not (1 <= miesiac <= 12) or not (1 <= dzien <= 31):
                self.error_label.config(text="Nieprawidłowa data urodzenia")
                return
        else:
            # Losowa data urodzenia
            start_date = datetime.date(1900, 1, 1)
            end_date = datetime.date(2299, 12, 31)
            delta_days = (end_date - start_date).days
            random_days = random.randint(0, delta_days)
            random_date = start_date + datetime.timedelta(days=random_days)
            rok = random_date.year
            miesiac = random_date.month
            dzien = random_date.day

        # Konwersja roku i miesiąca zgodnie z zasadami PESEL
        if 1900 <= rok <= 1999:
            miesiac_code = miesiac
        elif 2000 <= rok <= 2099:
            miesiac_code = miesiac + 20
        elif 2100 <= rok <= 2199:
            miesiac_code = miesiac + 40
        elif 2200 <= rok <= 2299:
            miesiac_code = miesiac + 60
        else:
            self.error_label.config(text="Rok poza zakresem dla numeru PESEL")
            return

        data_pesel = f"{str(rok)[-2:]}{miesiac_code:02d}{dzien:02d}"

        # Generowanie numerów
        for _ in range(ilosc):
            numer_porzadkowy = random.randint(0, 9999)
            numer_porzadkowy_str = f"{numer_porzadkowy:04d}"
            numer_bez_k = data_pesel + numer_porzadkowy_str
            wagi = [9, 7, 3, 1, 9, 7, 3, 1, 9, 7]
            suma = sum(int(numer_bez_k[i]) * wagi[i] for i in range(10))
            cyfra_kontrolna = suma % 10
            pesel = numer_bez_k + str(cyfra_kontrolna)
            self.numery_listbox.insert(tk.END, pesel)

    def generuj_nip(self, ilosc):
        """Generuje numery NIP."""
        kod_urzedu = self.nip_kod_urzedu_entry.get()
        if kod_urzedu:
            if not kod_urzedu.isdigit() or len(kod_urzedu) != 3:
                self.error_label.config(text="Błędny kod urzędu skarbowego (3 cyfry)")
                return
        else:
            kod_urzedu = f"{random.randint(101, 999):03d}"

        for _ in range(ilosc):
            numer_porzadkowy = random.randint(0, 999999)
            numer_bez_k = kod_urzedu + f"{numer_porzadkowy:06d}"
            wagi = [6, 5, 7, 2, 3, 4, 5, 6, 7]
            suma = sum(int(numer_bez_k[i]) * wagi[i] for i in range(9))
            reszta = suma % 11
            if reszta == 10:
                cyfra_kontrolna = 0
            else:
                cyfra_kontrolna = reszta
            nip = numer_bez_k + str(cyfra_kontrolna)
            self.numery_listbox.insert(tk.END, nip)

    def generuj_regon(self, ilosc):
        """Generuje numery REGON."""
        poczatkowe_cyfry = self.regon_poczatkowe_entry.get()
        if poczatkowe_cyfry:
            if not poczatkowe_cyfry.isdigit() or len(poczatkowe_cyfry) not in [7, 8]:
                self.error_label.config(text="Błędne początkowe cyfry REGON (7 lub 8 cyfr)")
                return
            dlugosc_regon = len(poczatkowe_cyfry) + 1  # Dodamy cyfrę kontrolną
        else:
            dlugosc_regon = 9
            poczatkowe_cyfry = ''.join([str(random.randint(0, 9)) for _ in range(dlugosc_regon - 1)])

        for _ in range(ilosc):
            numer_bez_k = poczatkowe_cyfry
            if len(numer_bez_k) == 8:
                wagi = [8, 9, 2, 3, 4, 5, 6, 7]
            else:
                self.error_label.config(text="Obsługiwane są tylko numery 9-cyfrowe REGON")
                return
            suma = sum(int(numer_bez_k[i]) * wagi[i] for i in range(8))
            reszta = suma % 11
            if reszta == 10:
                cyfra_kontrolna = 0
            else:
                cyfra_kontrolna = reszta
            regon = numer_bez_k + str(cyfra_kontrolna)
            self.numery_listbox.insert(tk.END, regon)
            # Inkrementuj początkowe cyfry dla kolejnych numerów
            poczatkowe_cyfry = str(int(poczatkowe_cyfry) + 1).zfill(len(poczatkowe_cyfry))

    def generuj_nrb(self, ilosc):
        """Generuje numery NRB."""
        numer_banku = self.nrb_bank_entry.get()
        numer_rachunku = self.nrb_rachunek_entry.get()

        if numer_banku:
            if not numer_banku.isdigit() or len(numer_banku) != 8:
                self.error_label.config(text="Błędny numer banku (8 cyfr)")
                return
        else:
            numer_banku = ''.join([str(random.randint(0, 9)) for _ in range(8)])

        if numer_rachunku:
            if not numer_rachunku.isdigit() or len(numer_rachunku) != 16:
                self.error_label.config(text="Błędny numer rachunku klienta (16 cyfr)")
                return
        else:
            numer_rachunku = ''.join([str(random.randint(0, 9)) for _ in range(16)])

        for _ in range(ilosc):
            # Oblicz sumę kontrolną
            numer_bez_k = numer_banku + numer_rachunku
            numer_do_obliczen = numer_bez_k + '252100'  # 'PL00' jako liczby
            numer_do_obliczen = numer_do_obliczen[2:] + numer_do_obliczen[:2]
            liczba = int(numer_do_obliczen)
            reszta = liczba % 97
            suma_kontrolna = (98 - reszta) % 97
            nrb = f"{suma_kontrolna:02d}{numer_bez_k}"
            self.numery_listbox.insert(tk.END, nrb)
            # Inkrementuj numer rachunku
            numer_rachunku = str(int(numer_rachunku) + 1).zfill(16)

    def on_listbox_select(self, event):
        """Wyświetla kod kreskowy wybranego numeru."""
        selected_indices = self.numery_listbox.curselection()
        if not selected_indices:
            self.barcode_label.config(image='')
            return
        selected_index = selected_indices[0]
        numer = self.numery_listbox.get(selected_index)
        # Generowanie obrazu kodu kreskowego
        barcode_image = self.generate_barcode_image(numer)
        if barcode_image:
            self.barcode_photo = ImageTk.PhotoImage(barcode_image)
            self.barcode_label.config(image=self.barcode_photo)
        else:
            self.barcode_label.config(image='')

    def generate_barcode_image(self, numer):
        """Generuje obraz kodu kreskowego dla danego numeru."""
        try:
            barcode = Code128(numer, writer=ImageWriter())
            barcode_image = barcode.render(writer_options={'module_height': 10.0, 'module_width': 0.2, 'quiet_zone': 1.0})
            return barcode_image
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie można wygenerować kodu kreskowego: {e}")
            return None

# --- Inicjalizacja GUI ---

root = tk.Tk()
root.title("Generator Numerów Nadawczych")
root.geometry("750x900")
root.resizable(True, True)

# Górny frame dla przycisku "Pomoc"
top_frame = ttk.Frame(root)
top_frame.pack(side=tk.TOP, fill=tk.X)

# Przycisk "?"
pomoc_button = ttk.Button(top_frame, text="?", command=pomoc)
pomoc_button.pack(side=tk.RIGHT, padx=10, pady=0)

# Notebook z zakładkami
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both", padx=10, pady=0)

# Zakładka "Numery 13 znaków"
# (Pozostawiamy istniejące zakładki bez zmian)
tab1 = Numery13ZnakowTab(notebook)
notebook.add(tab1, text="Numery 13 znaków")

# Zakładka "Numery 20 znaków"
tab2 = Numery20ZnakowTab(notebook)
notebook.add(tab2, text="Numery 20 znaków")

# Nowa zakładka "Inne"
inne_tab = InneTab(notebook)
notebook.add(inne_tab, text="Inne")

# Uruchomienie głównej pętli
root.mainloop()
