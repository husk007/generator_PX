# Changelog

## [4.0.1] - 2026-11-02

### Poprawki / Fixed
- **NRB:** poprawiono algorytm cyfr kontrolnych (mod 97) – wcześniej niemal wszystkie wygenerowane numery były niepoprawne. / Fixed the mod-97 check digits; previously almost every generated NRB was invalid.
- **NIP:** numery, dla których reszta z dzielenia wynosi 10, są losowane ponownie zamiast dostawać błędną cyfrę kontrolną `0`. / NIPs whose remainder is 10 are now re-drawn instead of getting an invalid check digit.
- **Numery 13-znakowe (S10):** numer początkowy dłuższy niż 8 cyfr jest odrzucany (wcześniej powstawały numery 14-znakowe z błędną cyfrą kontrolną). / Starting numbers longer than 8 digits are rejected.
- **PUH:** puste pole „Pierwsze liczby" nie powoduje już cichego błędu; licznik startuje od zera. / An empty starting number no longer fails silently; the counter starts at zero.
- **Numery 20-znakowe (GS1):** przepełnienie licznika (`99999999`) zatrzymuje generowanie zamiast zawijać numerację. / Counter overflow now stops generation instead of wrapping around.
- **PESEL:** walidacja kalendarzowa daty urodzenia (np. 31 lutego jest odrzucane). / Birth dates are validated against the calendar.
- **REGON:** akceptowane są wyłącznie 8 cyfr początkowych (wcześniej 7 było przyjmowane, a potem odrzucane). / Only 8 starting digits are accepted.
- **Wyślij:** komunikat i opóźnienie nie działają już w osobnym wątku (tkinter nie jest wątkowo bezpieczny). / The send fallback no longer touches tkinter from a worker thread.
- **Kody kreskowe:** używana jest czcionka dołączona do `python-barcode`; usunięto `arial.ttf` (czcionka komercyjna) z repozytorium. Podgląd działa też w zbudowanym pliku EXE. / Barcode preview uses the font bundled with `python-barcode`; `arial.ttf` was removed from the repository.
- **requirements.txt:** `pygetwindow==0.0.10` nie istnieje w PyPI – zmieniono na `0.0.9`; zluzowano przypięcia (m.in. `Pillow>=10.2.0`), `pywin32` tylko na Windows. / `pygetwindow==0.0.10` does not exist on PyPI; pins relaxed.
- **generator_PX.spec:** dołączane są dane pakietu `barcode` oraz hook z katalogu `hooks`. / The spec now bundles `barcode` data and uses the `hooks` directory.

### Zmiany techniczne / Maintenance
- GUI uruchamia się tylko przy starcie skryptu (`if __name__ == "__main__"`), moduł można importować w testach. / The GUI starts only when run as a script.
- Dodano `.gitignore`; usunięto z repozytorium `__pycache__`. / Added `.gitignore`; removed tracked `__pycache__`.
- README: informacja o wymaganiu Windows i Pythona 3.9+. / README documents the Windows and Python 3.9+ requirement.

> Plik `generator_v4.exe` z wydania v4.0.0 nie zawiera tych poprawek – do v4.0.1 dołączany jest kod źródłowy; EXE należy zbudować na Windows (`pyinstaller generator_PX.spec`).
