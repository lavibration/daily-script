import yfinance as yf
import pandas as pd
import os

# Paramètres
ema_periods = list(range(60, 321, 10))
rolling_window = 60
tolerance = 0.01
input_file = "Euronext_Tickers.xlsx"
output_html = "docs/index.html"
long_term_ema_min_period = 220
volume_threshold = 5000

# Assurez-vous que le répertoire "docs" existe
os.makedirs("docs", exist_ok=True)

# Charger les tickers et leurs noms depuis le fichier Excel
try:
    tickers_df = pd.read_excel(input_file)
    if 'Ticker' not in tickers_df.columns or 'Name' not in tickers_df.columns:
        raise ValueError("Le fichier Excel doit contenir les colonnes 'Ticker' et 'Name'.")
    tickers_with_names = tickers_df[['Ticker', 'Name']].dropna()
    tickers_dict = tickers_with_names.set_index('Ticker')['Name'].to_dict()
    tickers = list(tickers_dict.keys())
except FileNotFoundError:
    print(f"Fichier introuvable : {input_file}")
    exit(1)
except Exception as e:
    print(f"Erreur lors du chargement des tickers : {e}")
    exit(1)

# Fonction de calcul du Z-Score
def calculate_z_score(data, ema):
    rolling_std = (data['Close'] - ema).rolling(window=rolling_window).std()
    if rolling_std.isnull().all():
        raise ValueError("Impossible de calculer l'écart-type pour le Z-Score. Vérifiez les données.")
    z_scores = (data['Close'] - ema) / rolling_std
    return z_scores

# Fonction principale pour l'analyse EMA
def find_ema_with_max_contacts_and_long_term(data, ema_periods, rolling_window=60, tolerance=0.01, long_term_min_period=220):
    max_contacts = 0
    best_ema = None
    best_ema_series = None
    best_z_score = None

    best_long_term_ema = None
    best_long_term_ema_series = None
    max_contacts_long_term = 0

    for period in ema_periods:
        ema = data['Close'].ewm(span=period, adjust=False).mean()
        data[f'EMA_{period}'] = ema

        contacts = ((data['Close'] - ema).abs() / ema <= tolerance).sum()

        if contacts > max_contacts:
            max_contacts = contacts
            best_ema = period
            best_ema_series = ema

        if period >= long_term_min_period and contacts > max_contacts_long_term:
            max_contacts_long_term = contacts
            best_long_term_ema = period
            best_long_term_ema_series = ema

    z_score_max_contacts = calculate_z_score(data, best_ema_series).iloc[-1] if best_ema_series is not None else None
    z_score_long_term = calculate_z_score(data, best_long_term_ema_series).iloc[-1] if best_long_term_ema_series is not None else None

    return best_ema, max_contacts, best_ema_series.iloc[-1], z_score_max_contacts, \
           best_long_term_ema, max_contacts_long_term, best_long_term_ema_series.iloc[-1], z_score_long_term

# Liste pour les résultats
results = []

for ticker in tickers:
    name = tickers_dict[ticker]
    print(f"Analyse de {ticker} ({name})...")
    try:
        data = yf.download(ticker, period="5y", interval="1wk")
        if data.empty:
            raise ValueError(f"Les données pour {ticker} sont vides. Vérifiez l'API yfinance ou les paramètres.")
        
        volume_data = yf.download(ticker, period="1mo", interval="1d")

        if data.empty or 'Close' not in data.columns:
            print(f"Données insuffisantes ou 'Close' manquant pour {ticker}.")
            continue
        if volume_data.empty or 'Volume' not in volume_data.columns:
            print(f"Données insuffisantes ou 'Volume' manquant pour {ticker}.")
            continue

        avg_daily_volume = volume_data['Volume'].mean()

        if avg_daily_volume < volume_threshold:
            print(f"{ticker} exclu pour volume moyen quotidien < {volume_threshold}.")
            continue

        best_ema, max_contacts, last_ema_value, z_score_max_contacts, \
        best_long_term_ema, max_contacts_long_term, last_long_term_value, z_score_long_term = \
            find_ema_with_max_contacts_and_long_term(
                data, ema_periods, rolling_window, tolerance, long_term_ema_min_period
            )

        last_price = data['Close'].iloc[-1]
        distance_to_best_ema = ((last_price - last_ema_value) / last_ema_value) * 100 if last_ema_value else None
        distance_to_long_term_ema = ((last_price - last_long_term_value) / last_long_term_value) * 100 if last_long_term_value else None

        trend_max_contacts = "Trend" if last_price > last_ema_value else ""
        trend_long_term = "Trend" if last_price > last_long_term_value else ""

        results.append({
            'Ticker': ticker,
            'Name': name,
            'Last Price': round(last_price, 2),
            'Période EMA Max Contacts': best_ema,
            'EMA Max Contacts': round(last_ema_value, 2) if last_ema_value else None,
            'Trend Max Contacts': trend_max_contacts,
            'Distance % Max Contacts': round(distance_to_best_ema, 2) if distance_to_best_ema else None,
            'Z-Score Max Contacts': round(z_score_max_contacts, 2) if z_score_max_contacts else None,
            'Période EMA Long Term': best_long_term_ema,
            'EMA Long Term': round(last_long_term_value, 2) if last_long_term_value else None,
            'Trend Long Term': trend_long_term,
            'Distance % Long Term': round(distance_to_long_term_ema, 2) if distance_to_long_term_ema else None,
            'Z-Score Long Term': round(z_score_long_term, 2) if z_score_long_term else None
        })
    except Exception as e:
        print(f"Erreur pour {ticker} ({name}): {e}")

# Générer le tableau avec DataTables
try:
    results_df = pd.DataFrame(results)
    if results_df.empty:
        print("Aucun résultat à exporter.")
    else:
         # Convertir le DataFrame en HTML et ajouter des options DataTables
        table_html = results_df.to_html(index=False, classes="display")
        
        # Ajouter les ressources CSS et JS de DataTables
        html_template = f"""
        <!DOCTYPE html>
        <html lang="fr">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Résultats EMA Analysis</title>
            <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
            <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
            <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
        </head>
        <body>
            <h1>Analyse des EMA</h1>
            <table id="resultsTable" class="display">
                {table_html}
            </table>
            <script>
                $(document).ready(function() {{
                    $('#resultsTable').DataTable({{
                        paging: true,
                        searching: true,
                        info: true,
                        ordering: true
                    }});
                }});
            </script>
        </body>
        </html>
        """

        # Enregistrer le fichier HTML dans le répertoire docs
        with open(output_html, "w", encoding="utf-8") as f:
            f.write(html_template)
        print(f"Résultats enregistrés dans {output_html}.")
except Exception as e:
    print(f"Erreur lors de l'exportation des résultats : {e}")
