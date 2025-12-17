Projekt matematyka finansowa - analiza portfela Markowitza

# Jak uruchomić skrypt?
Wymagany jest interperter Python'a oraz system kontroli wersji - git. 

Interpreter można pobrać ze strony https://www.python.org/downloads/, natomiast git ze strony https://git-scm.com/.

Podczas instalacji interpretera ważne jest, aby zaznaczyć opcję automatycznego dodania ścieżki pythona do zmiennej środowiskowej PATH (kliknąć kwadracik podczas instalacji).

Aby pobrać kod należy w terminalu użyć komendy:
git clone https://github.com/uciekl/projekt_matematyka.git

Skrypt można uruchomić tylko wtedy, gdy w terminalu (cmd) znajdujemy się w katalogu zawierającym skrypt, dlatego należy sprawdzić ścieżkę projektu i skierować się do niego za pomocą komendy cd - change directory. Domyślnie powinien zostać pobrany w folderze User, zatem należy skorzystać z komendy: cd C:\Users\(User - nazwa użytkownika komputera)\projekt_matematyka\src - w katalogu src znajduje się plik z rozszerzeniem .py zawierający kod.

Jeżeli znajdujemy się w katalogu src, to do uruchomienia skryptu należy użyć komendy: python projektmf.py
Pokaże się sposób użycia kodu - ticker1 ticker2 ticker3 to wymagane symbole spółek, start_date to data, od której pobierane są dane (niewymagane, domyślnie 01.01.2025) i end_date to data, do której pobierany jest kod (niewymagane, domyślnie dzień dzisiejszy). 

Przykładowe użycie: python projektmf.py GOOG NVDA PYPL

Wyniki są zapisywane do nowo utworzonego folderu na pulpicie - wykresy, output z terminala, plik z rozszerzeniem .xlsx
