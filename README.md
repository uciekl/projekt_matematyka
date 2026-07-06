# Optymalizacja portfela inwestycyjnego metodą Markowitza (MPT)

## Opis projektu

Analiza portfelowa Markowitza (MPT - Modern Portfolio Theory) to matematyczny model zarządzania portfelem inwestycyjnym. Założeniem teorii jest obliczenie odpowiedniej dywersyfikacji aktywów porfela w celu minimalizacji ryzyka i maksymalizacji zwrotów. Wynikiem analizy jest wyznaczenie tzw. granicy efektywnej (Efficient Frontier), czyli zestawu portfeli, które oferują najwyższy zwrot przy danym poziomie ryzyka. 

## Jak działa kod?

Skrypt automatycznie pobiera historyczne ceny aktywów (Open, Close) dla trzech wybranych spółek z serwisu Yahoo Finance za pomocą biblioteki yfinance. Następnie stosuje algorytm okna przesuwnego (rolling window) o rozmiarze 25 dni, aby obliczyć stopy zwrotu i zidentyfikować podokres o najwyższej, skumulowanej stopie zwrotu dla całego koszyka akcji.

Główny część analizy polega na optymalizacji trzech portfeli dwuskładnikowych (pary AB, AC, BC dla akcji A, B, C). Dla każdej pary skrypt kalkuluje macierz korelacji, stopy zwrotu oraz ryzyko portfela (implementowane w formie odchylenia standardowego) przy różnych proporcjach wagowych aktywów. Na tej podstawie wyznaczana jest granica efektywna dla każdej pary, pozwalająca wskazać strukturę portfela o optymalnym stosunku zysku do ryzyka.

Wyniki działania skryptu są automatycznie zapisywane do dedykowanego katalogu na pulpicie i obejmują:

- Wizualizacje: 4 wykresy przedstawiające granice efektywne (trzy wykresy indywidualne dla każdej pary oraz jeden wykres zbiorczy, porównujący wszystkie pary na jednej przestrzeni).
- Raport tekstowy: plik zawierający tekstową analizę.
- Dane w formacie tabelarycznym: zapisane w pliku z rozszerzeniem .xlsx (Excel) zawierają wyniki obliczeń do dalszej pracy z danymi.

## Obsługa skryptu

Skrypt przyjmuje 5 argumentów. Pierwsze 3 dotyczą tickerów aktywów - wymagane. Pozostałe określają zakres czasowy, dla którego pobierane są dane giełdowe - niewymagane. Domyślnie za początek okresu skrypt uznaje 01.01.2025, a za koniec dzień dzisiejszy.

Dla tickerów notowanych na warszawskiej giełdzie należy użyć suffixu .WA (np. JSW.WA). 

**Wzór komendy:** 

```bash
$ python3 MarkowitzMPT.py ticker1 ticker2 ticker3 start_date(opcjonalne, domyślnie 2025-01-01) end_date(opcjonalne, domyślnie dzień dzisiejszy)
```

**Podstawowe użycie:**
```bash
$ python3 MarkowitzMPT.py WMT NEM GLD
```

**Wykorzystanie indywidualnych dat:**
```bash
$ python3 MarkowitzMPT.py WMT NEM GLD "2022-03-01" "2023-05-05"
```
