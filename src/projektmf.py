import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import xlsxwriter
import os
import sys

#zmienne globalne dla plików logów
LOG_FILE = None
RESULTS_DIR = None

#funkcja printująca wynik w terminalu i pliku tekstowym jednocześnie
def terminal_and_file_print(message, end="\n"):
    print(message, end=end)
    if LOG_FILE:
        LOG_FILE.write(message + end)


#uzyskanie ścieżki do pulpitu, na którym ma zostać zapisany folder z wynikami (ścieżki się różnią w zależności od systemu operacyjnego)
def get_desktop_path():
    #Windows
    if sys.platform.startswith("win"):
        return os.path.join(os.environ.get("USERPROFILE"), "Desktop")

    #Linux lub macos
    elif sys.platform.startswith("linux") or sys.platform.startswith("darwin"):
        return os.path.join(os.path.expanduser("~"), "Desktop")
    else:
        return os.path.expanduser("~")

#pobranie danych, wybranie najlepszego 25 dniowego okresu
def data_download(tickers, start_date="2025-01-01", end_date=None, workbook=None, worksheet=None, formats=None):
    if end_date is None:
        end_date = datetime.today().strftime("%Y-%m-%d")

    tickers_list = tickers.split() if not isinstance(tickers, list) else tickers

    try:
        df = yf.download(tickers_list, start=start_date, end=end_date)["Close"]
    except Exception as e:
        terminal_and_file_print(f"Błąd pobierania danych z yfinance: {e}")
        return None

    df = df.ffill().bfill()

    terminal_and_file_print(f"\nPobrano dane dla: {', '.join(tickers_list)}")
    terminal_and_file_print(f"Zakres dat: {df.index[0].strftime('%Y-%m-%d')} do {df.index[-1].strftime('%Y-%m-%d')}")
    terminal_and_file_print(f"Liczba dni: {len(df)}\n")

    period = 25 #25-dniowe okno (24 zwroty)
    results = [] #lista wyników dla każdego okna

    #przesuwanie okna po 1 dniu, zapis wyników i ich porównanie
    for start_idx in range(0, len(df) - period + 1):
        end_idx = start_idx + period
        window_df = df.iloc[start_idx:end_idx]

        if len(window_df) == period:
            daily_returns_df = window_df.pct_change().dropna()

            if len(daily_returns_df) != period - 1:
                continue

            returns_dict = {}
            avg_returns_dict = {}
            for ticker in df.columns:
                daily_returns = daily_returns_df[ticker]
                period_return = daily_returns.sum()
                avg_return = daily_returns.mean()
                returns_dict[f"return_sum_{ticker}"] = period_return
                avg_returns_dict[f"avg_return_{ticker}"] = avg_return

            total_return = sum(returns_dict.values())
            total_avg_return = sum(avg_returns_dict.values())
            mean_avg_return = total_avg_return / len(df.columns)

            result_entry = {
                "start_date": window_df.index[0],
                "end_date": window_df.index[-1],
                "total_return": total_return,
                "total_avg_return": total_avg_return,
                "mean_avg_return": mean_avg_return,
                "days_in_window": len(daily_returns_df),
                "start_index": start_idx,
                "end_index": end_idx
            }
            result_entry.update(returns_dict)
            result_entry.update(avg_returns_dict)
            results.append(result_entry)

    if not results:
        terminal_and_file_print("Brak wystarczającej liczby danych, aby utworzyć 25-dniowe okno")
        return None

    results_df = pd.DataFrame(results)
    best_period = results_df.loc[results_df["total_return"].idxmax()]
    best_start_idx = int(best_period["start_index"])
    best_end_idx = int(best_period["end_index"])
    df_best = df.iloc[best_start_idx:best_end_idx].copy()
    daily_returns_best_df = df_best.pct_change().dropna()

    terminal_and_file_print("=" * 80)
    terminal_and_file_print(" NAJLEPSZY 25-DNIOWY OKRES Z NAJWYŻSZĄ ZSUMOWANĄ STOPĄ ZWROTU")
    terminal_and_file_print("=" * 80)
    terminal_and_file_print(f" Start: {best_period['start_date'].strftime('%Y-%m-%d')}")
    terminal_and_file_print(f" Koniec: {best_period['end_date'].strftime('%Y-%m-%d')}")
    terminal_and_file_print(f" Dni w oknie (zwrotów): {int(best_period['days_in_window'])}")
    terminal_and_file_print(f"\n Stopy zwrotu z okresu (liczone jako suma dziennych zmian):")
    for ticker in df.columns:
        ret = best_period[f"return_sum_{ticker}"]
        avg_ret = best_period[f"avg_return_{ticker}"]
        terminal_and_file_print(f" - {ticker:10s}: {ret:7.4f} ({ret * 100:6.2f}%) | średnia: {avg_ret:7.6f} ({avg_ret * 100:7.4f}%)")
    terminal_and_file_print(f"\n ZSUMOWANA STOPA ZWROTU DLA 3 TICKERÓW: {best_period['total_return']:.4f} ({best_period['total_return'] * 100:.2f}%)")
    terminal_and_file_print(f" ZSUMOWANA ŚREDNIA STOPA ZWROTU DLA 3 TICKERÓW: {best_period['total_avg_return']:.6f} ({best_period['total_avg_return'] * 100:.4f}%)")
    terminal_and_file_print(f" ŚREDNIA STOPA ZWROTU DLA 3 TICKERÓW: {best_period['mean_avg_return']:.6f} ({best_period['mean_avg_return'] * 100:.4f}%)")
    terminal_and_file_print("=" * 80)
    terminal_and_file_print("\n--- DF z najlepszym okresem ---")
    terminal_and_file_print(df_best.to_string())
    terminal_and_file_print("\n--- DF z dziennymi zwrotami ---")
    terminal_and_file_print(daily_returns_best_df.to_string())
    terminal_and_file_print("\n--- 5 najlepszych okresów ---")
    top5 = results_df.nlargest(5, "total_return")[["start_date", "end_date", "total_return"]]
    top5["total_return_%"] = top5["total_return"] * 100
    terminal_and_file_print(top5.to_string(index=False))

    #zapisanie do pliku .xlsx zakresu dat, cen i zwrotów  dziennych
    if worksheet and formats:
        bold = formats["bold"]
        float_format = formats["float_format"]
        percent_format_6 = formats["percent_format_6"]

        #A1: zakres dat
        start_date_str = best_period["start_date"].strftime("%Y-%m-%d")
        end_date_str = best_period["end_date"].strftime("%Y-%m-%d")
        date_range = f"{start_date_str} do {end_date_str}"
        worksheet.write("A1", date_range, bold)

        #A2: "Data"
        worksheet.write("A2", "Data", bold)

        #A3-A27: daty
        dates_to_write = df_best.index.strftime("%Y-%m-%d")
        worksheet.write_column("A3", dates_to_write)

        #B2-D2: nazwy tickerów
        worksheet.write("B2", tickers_list[0], bold)
        worksheet.write("C2", tickers_list[1], bold)
        worksheet.write("D2", tickers_list[2], bold)

        #B3-D27: ceny
        worksheet.write_column("B3", df_best[tickers_list[0]], float_format)
        worksheet.write_column("C3", df_best[tickers_list[1]], float_format)
        worksheet.write_column("D3", df_best[tickers_list[2]], float_format)

        #E4-G27: dzienne zwroty
        worksheet.write_column("E4", daily_returns_best_df[tickers_list[0]], percent_format_6)
        worksheet.write_column("F4", daily_returns_best_df[tickers_list[1]], percent_format_6)
        worksheet.write_column("G4", daily_returns_best_df[tickers_list[2]], percent_format_6)

        #E28-G28: suma zwrotów
        worksheet.write("E28", best_period[f"return_sum_{tickers_list[0]}"], percent_format_6)
        worksheet.write("F28", best_period[f"return_sum_{tickers_list[1]}"], percent_format_6)
        worksheet.write("G28", best_period[f"return_sum_{tickers_list[2]}"], percent_format_6)

    return {
        "all_results": results_df,
        "best_period": best_period,
        "best_dataframe": df_best,
        "daily_returns_best_df": daily_returns_best_df,
        "original_data": df,
        "tickers_list": tickers_list
    }

def markowitz_analysis(df, tickers_list, dir_path, workbook=None, worksheet=None, formats=None):

    '''
    Oznaczenia poszczególnych składników obliczeń
    - r - stopa zwrotu dla całego okresu jednego tickera
    - w - wagi
    - sigma - odchylenie standardowe tickera, więc jest to jego ryzyko
    - jakaś_zmienna_p - skumulowany element dla A oraz B
    - cv - współczynnik zmienności (coefficient of variation)
    - rho - współczynnik korelacji (rho jako nazwa symbolu)
    '''

    daily_returns = df.pct_change().dropna()
    period_days = len(daily_returns)

    terminal_and_file_print("\n" + "=" * 80)
    terminal_and_file_print("WYNIKI POJEDYNCZYCH AKCJI")
    terminal_and_file_print("=" * 80)
    terminal_and_file_print(f" Okres analizy: {period_days} dni (stóp zwrotu)")

    sum_returns_dict = {}
    mean_daily_returns_dict = {}
    risk_daily_std_dict = {}
    cv_daily_dict = {}

    risk_dict_period = {}
    mean_return_dict_period = {}

    for ticker in tickers_list:
        daily_rets = daily_returns[ticker]

        #zsumowana stopa dziennych zwrotóœ
        total_return_period = daily_rets.sum()

        #średnia dzienna stopa dziennych zwrotów
        mean_daily_ret = daily_rets.mean()

        #ryzyko - odchylenie standardowe
        std_daily_ret = daily_rets.std()

        #współczynnik zmienności
        cv = std_daily_ret / abs(mean_daily_ret) if mean_daily_ret != 0 else np.inf

        cumulative_returns = (1 + daily_rets).cumprod() - 1
        risk_period = cumulative_returns.std()

        sum_returns_dict[ticker] = total_return_period
        mean_daily_returns_dict[ticker] = mean_daily_ret
        risk_daily_std_dict[ticker] = std_daily_ret
        cv_daily_dict[ticker] = cv

        risk_dict_period[ticker] = risk_period
        mean_return_dict_period[ticker] = total_return_period

        terminal_and_file_print(f"\n{ticker}:")
        terminal_and_file_print(f" Suma dziennych zwrotów: {total_return_period:.6f} ({total_return_period * 100:.4f}%)")
        terminal_and_file_print(f" Ryzyko: {risk_period:.6f} ({risk_period * 100:.4f}%)")
        terminal_and_file_print(f" Współczynnik zmienności: {risk_period / abs(total_return_period):.4f}" if total_return_period != 0 else "inf")

    #zapisanie do pliku .xlsx: R, Rśr, S(n-1), V
    if worksheet and formats:
        bold = formats["bold"]
        float_format = formats["float_format"]
        percent_format_6 = formats["percent_format_6"]

        for i, ticker in enumerate(tickers_list):
            col = chr(ord("K") + i)
            worksheet.write(f"{col}2", ticker, bold)

        #suma dziennych zwrotów
        worksheet.write("J3", "R", bold)
        for i, ticker in enumerate(tickers_list):
            col = chr(ord("K") + i)
            worksheet.write(f"{col}3", sum_returns_dict[ticker], percent_format_6)

        #średnie dzienne zwroty
        worksheet.write("J4", "Rśr", bold)
        for i, ticker in enumerate(tickers_list):
            col = chr(ord("K") + i)
            worksheet.write(f"{col}4", mean_daily_returns_dict[ticker], percent_format_6)

        #ryzyko
        worksheet.write("J5", "S(n-1)", bold)
        for i, ticker in enumerate(tickers_list):
            col = chr(ord("K") + i)
            worksheet.write(f"{col}5", risk_dict_period[ticker], float_format)

        #współczynnik zmienności
        worksheet.write("J6", "V", bold)
        for i, ticker in enumerate(tickers_list):
            col = chr(ord("K") + i)
            cv_markowitz = risk_dict_period[ticker] / abs(sum_returns_dict[ticker]) if sum_returns_dict[ticker] != 0 else np.inf
            worksheet.write(f"{col}6", cv_markowitz, float_format)

    #macierz korelacji
    terminal_and_file_print("\n" + "=" * 80)
    terminal_and_file_print(" WSPÓŁCZYNNIKI KORELACJI")
    terminal_and_file_print("=" * 80)
    correlation_matrix = daily_returns.corr()
    terminal_and_file_print("\nMacierz korelacji:")
    terminal_and_file_print(correlation_matrix.to_string())

    pairs = []
    for i in range(len(tickers_list)):
        for j in range(i + 1, len(tickers_list)):
            ticker_a = tickers_list[i]
            ticker_b = tickers_list[j]
            corr = correlation_matrix.loc[ticker_a, ticker_b]
            pairs.append((ticker_a, ticker_b))
            terminal_and_file_print(f"\nKorelacja {ticker_a} - {ticker_b}: {corr:.4f}")

    #zapisanie do pliku .xlsx: macierz korelacji
    if worksheet and formats:
        worksheet.write("J10", "Macierz Korelacji", bold)

        for i, ticker in enumerate(tickers_list):
            worksheet.write(10, 9 + i + 1, ticker, bold)
            worksheet.write(10 + i + 1, 9, ticker, bold)

        for i, ticker1 in enumerate(tickers_list):
            for j, ticker2 in enumerate(tickers_list):
                worksheet.write(10 + i + 1, 9 + j + 1, correlation_matrix.loc[ticker1, ticker2], float_format)

    #stosunki ryzyka
    terminal_and_file_print("\n" + "=" * 80)
    terminal_and_file_print(" STOSUNEK RYZYKA")
    terminal_and_file_print("=" * 80)
    risk_ratios = {}
    for ticker_a, ticker_b in pairs:
        risk_ratio = risk_dict_period[ticker_a] / risk_dict_period[ticker_b] if risk_dict_period[ticker_b] != 0 else np.inf
        risk_ratios[f"{ticker_a} / {ticker_b}"] = risk_ratio
        terminal_and_file_print(f"\nStosunek ryzyka {ticker_a} / {ticker_b}: {risk_ratio:.4f}")

    #zapisanie do pliku:stosunek ryzyka
    if worksheet and formats:
        worksheet.write("J15", "Stosunek Ryzyka", bold)
        row = 16
        for pair_name, ratio in risk_ratios.items():
            worksheet.write(f"J{row}", pair_name, bold)
            worksheet.write(f"K{row}", ratio, float_format)
            row += 1

    #analiza portfeli dwuskładnikowych i krzywe efektywności
    terminal_and_file_print("\n" + "=" * 80)
    terminal_and_file_print(" ANALIZA PORTFELI DWUSKŁADNIKOWYCH")
    terminal_and_file_print("=" * 80)
    portfolio_results = {}
    for ticker_a, ticker_b in pairs:
        terminal_and_file_print(f"\n{'=' * 40}")
        terminal_and_file_print(f" PORTFEL: {ticker_a} + {ticker_b}")
        terminal_and_file_print(f"{'=' * 40}")

        #parametry akcji
        r_a = sum_returns_dict[ticker_a]
        r_b = sum_returns_dict[ticker_b]
        sigma_a = risk_dict_period[ticker_a]
        sigma_b = risk_dict_period[ticker_b]
        rho = correlation_matrix.loc[ticker_a, ticker_b]

        terminal_and_file_print(f"\nParametry dla obliczeń portfela:")
        terminal_and_file_print(f" {ticker_a}: R={r_a:.6f} ({r_a*100:.2f}%), S={sigma_a:.6f} ({sigma_a*100:.2f}%)")
        terminal_and_file_print(f" {ticker_b}: R={r_b:.6f} ({r_b*100:.2f}%), S={sigma_b:.6f} ({sigma_b*100:.2f}%)")
        terminal_and_file_print(f" Korelacja: {rho:.4f}")

        #wagi portfela i obliczenia punktów na krzywej (jest ich 100 dla dokładniejszej analizy)
        weights_a_all = np.linspace(0, 1, 101)
        weights_b_all = 1 - weights_a_all
        weights_a_labeled = np.linspace(0, 1, 11)
        weights_b_labeled = 1 - weights_a_labeled
        portfolio_returns_all = []
        portfolio_risks_all = []
        portfolio_returns_labeled = []
        portfolio_risks_labeled = []
        for w_a, w_b in zip(weights_a_all, weights_b_all):
            r_p = w_a * r_a + w_b * r_b
            variance_p = (w_a**2 * sigma_a**2 + w_b**2 * sigma_b**2 + 2 * w_a * w_b * sigma_a * sigma_b * rho)
            if variance_p < 0: variance_p = 0
            sigma_p = np.sqrt(variance_p)
            portfolio_returns_all.append(r_p)
            portfolio_risks_all.append(sigma_p)
        for w_a, w_b in zip(weights_a_labeled, weights_b_labeled):
            r_p = w_a * r_a + w_b * r_b
            variance_p = (w_a**2 * sigma_a**2 + w_b**2 * sigma_b**2 + 2 * w_a * w_b * sigma_a * sigma_b * rho)
            if variance_p < 0: variance_p = 0
            sigma_p = np.sqrt(variance_p)
            portfolio_returns_labeled.append(r_p)
            portfolio_risks_labeled.append(sigma_p)

        #znalezienie portfela o minimalnym ryzyku
        min_risk_idx = np.argmin(portfolio_risks_all)
        min_risk_weight_a = weights_a_all[min_risk_idx]
        min_risk_return = portfolio_returns_all[min_risk_idx]
        min_risk_value = portfolio_risks_all[min_risk_idx]

        #znalezienie portfela o maksymalnej stopie zwrotu
        max_return_idx = np.argmax(portfolio_returns_all)
        max_return_weight_a = weights_a_all[max_return_idx]
        max_return_value = portfolio_returns_all[max_return_idx]
        max_return_risk = portfolio_risks_all[max_return_idx]

        terminal_and_file_print(f"\nPortfel o minimalnym ryzyku:")
        terminal_and_file_print(f" Udział {ticker_a}: {min_risk_weight_a * 100:.1f}%")
        terminal_and_file_print(f" Udział {ticker_b}: {(1 - min_risk_weight_a) * 100:.1f}%")
        terminal_and_file_print(f" Ryzyko: {min_risk_value:.6f} ({min_risk_value * 100:.4f}%)")
        terminal_and_file_print(f" Stopa zwrotu: {min_risk_return:.6f} ({min_risk_return * 100:.4f}%)")

        terminal_and_file_print(f"\nPortfel o maksymalnej stopie zwrotu:")
        terminal_and_file_print(f" Udział {ticker_a}: {max_return_weight_a * 100:.1f}%")
        terminal_and_file_print(f" Udział {ticker_b}: {(1 - max_return_weight_a) * 100:.1f}%")
        terminal_and_file_print(f" Ryzyko: {max_return_risk:.6f} ({max_return_risk * 100:.4f}%)")
        terminal_and_file_print(f" Stopa zwrotu: {max_return_value:.6f} ({max_return_value * 100:.4f}%)")

        portfolio_results[(ticker_a, ticker_b)] = {
            "weights_a_all": weights_a_all,
            "returns_all": portfolio_returns_all,
            "risks_all": portfolio_risks_all,
            "weights_a_labeled": weights_a_labeled,
            "returns_labeled": portfolio_returns_labeled,
            "risks_labeled": portfolio_risks_labeled,
            "min_risk_idx": min_risk_idx,
            "max_return_idx": max_return_idx
        }

    #rysowanie wykresów krzywych efektywności - osobne figure dla każdej pary
    for idx, (ticker_a, ticker_b) in enumerate(pairs):
        data = portfolio_results[(ticker_a, ticker_b)]

        risks_all_pct = np.array(data["risks_all"]) * 100
        returns_all_pct = np.array(data["returns_all"]) * 100

        risks_labeled_pct = np.array(data["risks_labeled"]) * 100
        returns_labeled_pct = np.array(data["returns_labeled"]) * 100
        weights_a_labeled = data["weights_a_labeled"]

        fig, ax = plt.subplots(figsize=(14, 10))

        ax.plot(risks_all_pct, returns_all_pct, "b-", linewidth=3, label="Krzywa efektywności", alpha=0.8, zorder=3)

        for i in range(len(weights_a_labeled)):
            w_a = weights_a_labeled[i]
            w_b = 1 - w_a
            risk = risks_labeled_pct[i]
            ret = returns_labeled_pct[i]

            ax.scatter(risk, ret, color="darkblue", s=120, zorder=5, alpha=0.9, edgecolors="white", linewidths=2)

            label_text = f"{w_a:.0%}-{w_b:.0%}\nS:{risk:.2f}%\nR:{ret:.2f}%"

            if w_a < 0.4:
                xytext = (15, 0); ha = "left"
            elif w_a > 0.6:
                xytext = (-15, 0); ha = "right"
            else:
                xytext = (0, 15 if i % 2 == 0 else -15); ha = "center"

            ax.annotate(label_text,
                        xy=(risk, ret),
                        xytext=xytext,
                        textcoords="offset points",
                        fontsize=9,
                        bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", alpha=0.9, edgecolor="darkblue", linewidth=1.5),
                        ha=ha, va="center", zorder=10,
                        arrowprops=dict(arrowstyle="->", color="gray", lw=1, alpha=0.6))

        #punkt minimalnego ryzyka
        min_risk_idx = data["min_risk_idx"]
        min_risk = risks_all_pct[min_risk_idx]
        min_ret = returns_all_pct[min_risk_idx]
        min_w_a = data["weights_a_all"][min_risk_idx]
        min_w_b = 1 - min_w_a
        ax.scatter(min_risk, min_ret, color="red", s=150, zorder=6, alpha=1.0, edgecolors="black", linewidths=2)
        min_label_text = f"{min_w_a:.0%}-{min_w_b:.0%}\nS:{min_risk:.2f}%\nR:{min_ret:.2f}%"
        ax.annotate(min_label_text,
                    xy=(min_risk, min_ret),
                    xytext=(0, 40),
                    textcoords="offset points",
                    fontsize=9,
                    bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", alpha=0.9, edgecolor="red", linewidth=1.5),
                    ha="center", va="bottom", zorder=10,
                    arrowprops=dict(arrowstyle="->", color="gray", lw=0.5, alpha=0.6))

        ax.set_xlabel("Ryzyko portfela dla 25. dniowego okresu (%)", fontsize=13, fontweight="bold")
        ax.set_ylabel("Stopa zwrotu portfela dla 25. dniowego okresu (%)", fontsize=13, fontweight="bold")
        ax.set_title(f"Krzywa Efektywności Markowitza\nPortfel: {ticker_a} + {ticker_b}",
                     fontsize=15, fontweight="bold", pad=20)
        ax.legend(fontsize=11, loc="best")
        ax.grid(True, alpha=0.3, linestyle="--", linewidth=0.8)
        plt.tight_layout()

        #zapis wykresu do folderu
        filename = f"markowitz_{ticker_a}_{ticker_b}.png"
        filepath = os.path.join(dir_path, filename)
        plt.savefig(filepath, dpi=300, bbox_inches="tight")

    fig, ax = plt.subplots(figsize=(14, 10))
    colors = ["b", "g", "r"] # Kolory dla 3 par
    for idx, (ticker_a, ticker_b) in enumerate(pairs):
        data = portfolio_results[(ticker_a, ticker_b)]
        risks_all_pct = np.array(data["risks_all"]) * 100
        returns_all_pct = np.array(data["returns_all"]) * 100
        ax.plot(risks_all_pct, returns_all_pct, "-", color=colors[idx], linewidth=3, label=f"{ticker_a} + {ticker_b}", alpha=0.8, zorder=3)

    ax.set_xlabel("Ryzyko portfela 25. dniowego okresu (%)", fontsize=13, fontweight="bold")
    ax.set_ylabel("Stopa zwrotu portfela 25. dniowego okres (%)", fontsize=13, fontweight="bold")
    ax.set_title("Merge krzywych efektywności Markowitza dla wszystkich par", fontsize=15, fontweight="bold", pad=20)
    ax.legend(fontsize=11, loc="best")
    ax.grid(True, alpha=0.3, linestyle="--", linewidth=0.8)
    plt.tight_layout()

    #zapis wykresu połączonego do folderu
    filename = "markowitz_combined.png"
    filepath = os.path.join(dir_path, filename)
    plt.savefig(filepath, dpi=300, bbox_inches="tight")
    return portfolio_results

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Obsługa skryptu: python \"nazwa_pliku\".py ticker1 ticker2 ticker3 start_date(opcjonalne, domyślnie 2025-01-01) end_date(opcjonalne, domyślnie dzień dzisiejszy)")
        sys.exit(1)

    ticker1 = sys.argv[1]
    ticker2 = sys.argv[2]
    ticker3 = sys.argv[3]
    tickers = f"{ticker1} {ticker2} {ticker3}"
    start_date = sys.argv[4] if len(sys.argv) > 4 else "2025-01-01"
    end_date = sys.argv[5] if len(sys.argv) > 5 else None

    #utworzenie ścieżek
    desktop_path = get_desktop_path()
    tickers_list = [ticker1, ticker2, ticker3]
    tickers_string = "_".join([t.replace(".WA", "") for t in tickers_list])

    #utworzenie folderu na pulpicie
    RESULTS_DIR = os.path.join(desktop_path, f"analiza_{tickers_string}")
    os.makedirs(RESULTS_DIR, exist_ok=True)

    #ścieżka do pliku logów (.txt)
    log_filename = os.path.join(RESULTS_DIR, f"{tickers_string}_analiza.txt")

    #ścieżka do pliku .xlsx w folderze
    xlsx_filename = os.path.join(RESULTS_DIR, f"{tickers_string}_analiza.xlsx")

    #otwarcie pliku logów (.txt)
    try:
        LOG_FILE = open(log_filename, "w", encoding="utf-8")

        workbook = xlsxwriter.Workbook(xlsx_filename)
        worksheet = workbook.add_worksheet()

        formats = {
            "bold": workbook.add_format({"bold": True}),
            "float_format": workbook.add_format({"num_format": "0.000000"}),
            "percent_format_6": workbook.add_format({"num_format": "0.000000%"}),
        }

        worksheet.set_column("A:A", 15)
        worksheet.set_column("B:I", 12)
        worksheet.set_column("J:J", 8)
        worksheet.set_column("K:M", 12)

        result = data_download(tickers, start_date=start_date, end_date=end_date,
                               workbook=workbook, worksheet=worksheet, formats=formats)

        if result:
            df_best = result["best_dataframe"]
            portfolio_results = markowitz_analysis(df_best, tickers_list, dir_path=RESULTS_DIR,
                                                   workbook=workbook, worksheet=worksheet, formats=formats)
    except Exception as e:
        terminal_and_file_print(f"\nBłąd: {e}")
    finally:
        #zamknięcie pliku po zakończonej analizie
        if "workbook" in locals() and workbook:
            workbook.close()
        if LOG_FILE:
            LOG_FILE.close()