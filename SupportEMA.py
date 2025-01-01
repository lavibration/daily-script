import yfinance as yf
import pandas as pd
import numpy as np

# Paramètres
ema_periods = list(range(60, 321, 10))  # EMAs à tester
rolling_window = 60  # Fenêtre glissante pour le Z-Score
tolerance = 0.01  # Tolérance pour les contacts avec l'EMA
input_file = "Euronext_Tickers.xlsx"  # Fichier source
output_file = "EMA_Analysis_Results.xlsx"  # Fichier de sortie
long_term_ema_min_period = 220  # Période minimum pour l'EMA long terme
volume_threshold = 5000  # Volume quotidien moyen minimum

# Charger les tickers et leurs noms depuis le fichier Excel
tickers_df = pd.read_excel(input_file)
tickers_with_names = tickers_df[['Ticker', 'Name']].dropna()
tickers_dict = tickers_with_names.set_index('Ticker')['Name'].to_dict()  # Dictionnaire {Ticker: Name}
tickers = list(tickers_dict.keys())  # Liste des tickers

# Fonction de calcul du Z-Score
def calculate_z_score(data, ema):
    rolling_std = (data['Close'] - ema).rolling(window=rolling_window).std()
    z_scores = (data['Close'] - ema) / rolling_std
    return z_scores

# Fonction principale pour l'analyse EMA avec le plus de contacts et l'EMA long terme
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

        # Compter les contacts
        contacts = (abs(data['Close'] - ema) / ema <= tolerance).sum()

        # Vérifier si cette EMA a le plus de contacts
        if contacts > max_contacts:
            max_contacts = contacts
            best_ema = period
            best_ema_series = ema

        # Vérifier si cette EMA peut être l'EMA long terme (supérieure ou égale à 220 jours)
        if period >= long_term_min_period and contacts > max_contacts_long_term:
            max_contacts_long_term = contacts
            best_long_term_ema = period
            best_long_term_ema_series = ema

    # Calculer les Z-scores pour les deux EMA
    z_score_max_contacts = calculate_z_score(data, best_ema_series).iloc[-1] if best_ema_series is not None else None
    z_score_long_term = calculate_z_score(data, best_long_term_ema_series).iloc[-1] if best_long_term_ema_series is not None else None

    # Retourner les résultats
    return best_ema, max_contacts, best_ema_series.iloc[-1], z_score_max_contacts, \
           best_long_term_ema, max_contacts_long_term, best_long_term_ema_series.iloc[-1], z_score_long_term

# Liste pour les résultats
results = []

# Parcourir chaque ticker
for ticker in tickers:
    name = tickers_dict[ticker]  # Récupérer le nom associé au ticker
    print(f"Analyse de {ticker} ({name})...")
    try:
        # Récupérer les données historiques
        data = yf.download(ticker, period="5y", interval="1wk")
        volume_data = yf.download(ticker, period="1mo", interval="1d")

        if data.empty or volume_data.empty:
            print(f"Aucune donnée pour {ticker}.")
            continue

        # Calculer le volume moyen quotidien du mois dernier
        avg_daily_volume = volume_data['Volume'].mean()

        # Filtrer les actions avec un volume moyen quotidien inférieur au seuil
        if avg_daily_volume < volume_threshold:
            print(f"{ticker} exclu pour volume moyen quotidien < {volume_threshold}.")
            continue

        # Effectuer l'analyse EMA
        best_ema, max_contacts, last_ema_value, z_score_max_contacts, \
        best_long_term_ema, max_contacts_long_term, last_long_term_value, z_score_long_term = \
            find_ema_with_max_contacts_and_long_term(
                data, ema_periods, rolling_window, tolerance, long_term_ema_min_period
            )

        # Calculer les distances en pourcentage
        last_price = data['Close'].iloc[-1]
        distance_to_best_ema = ((last_price - last_ema_value) / last_ema_value) * 100 if last_ema_value else None
        distance_to_long_term_ema = ((last_price - last_long_term_value) / last_long_term_value) * 100 if last_long_term_value else None

        # Déterminer les tendances ("Trend")
        trend_max_contacts = "Trend" if last_price > last_ema_value else ""
        trend_long_term = "Trend" if last_price > last_long_term_value else ""

        # Ajouter les résultats dans la liste
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

# Créer un DataFrame avec les résultats
results_df = pd.DataFrame(results)

# Réorganisation des colonnes
column_order = [
    'Ticker', 'Name', 'Last Price',
    'Période EMA Max Contacts', 'EMA Max Contacts', 'Trend Max Contacts', 'Distance % Max Contacts', 'Z-Score Max Contacts',
    'Période EMA Long Term', 'EMA Long Term', 'Trend Long Term', 'Distance % Long Term', 'Z-Score Long Term'
]
results_df = results_df[column_order]

# Exporter les résultats dans un fichier Excel
results_df.to_excel(output_file, index=False)
print(f"Analyse terminée. Résultats enregistrés dans {output_file}.")
