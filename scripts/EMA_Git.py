import os
import pandas as pd
import yfinance as yf

# Chemins des fichiers et répertoires
input_file = os.path.join(os.getcwd(), "Euronext_Tickers.xlsx")
output_dir = os.path.join(os.getcwd(), "output")
output_html = os.path.join(output_dir, "results.html")

# Création du répertoire output
try:
    os.makedirs(output_dir, exist_ok=True)
    print(f"Répertoire de sortie créé ou déjà existant : {output_dir}")
except Exception as e:
    print(f"Erreur lors de la création du répertoire : {e}")
    raise

# Chargement des tickers depuis le fichier Excel
try:
    tickers_df = pd.read_excel(input_file)
    print("Fichier Excel chargé avec succès.")
except FileNotFoundError:
    print(f"Fichier introuvable : {input_file}. Assurez-vous qu'il est présent dans le répertoire.")
    raise
except Exception as e:
    print(f"Erreur lors du chargement du fichier Excel : {e}")
    raise

# Vérification de la colonne contenant les tickers
if 'Ticker' not in tickers_df.columns:
    raise ValueError("La colonne 'Ticker' est absente du fichier Excel.")

tickers = tickers_df['Ticker'].dropna().tolist()
if not tickers:
    raise ValueError("Aucun ticker valide trouvé dans le fichier Excel.")

# Initialisation du DataFrame pour les résultats
results = []

# Téléchargement des données pour chaque ticker
for ticker in tickers:
    print(f"Traitement du ticker : {ticker}")
    try:
        data = yf.download(ticker, period="1mo", interval="1d")
        if data.empty:
            print(f"Pas de données disponibles pour {ticker}.")
            continue

        # Calcul de l'EMA sur 20 jours
        data['EMA_20'] = data['Close'].ewm(span=20).mean()
        last_close = data['Close'].iloc[-1]
        last_ema = data['EMA_20'].iloc[-1]
        results.append({
            'Ticker': ticker,
            'Dernier cours': last_close,
            'EMA 20 jours': last_ema,
            'Signal': 'Achat' if last_close > last_ema else 'Vente'
        })
        print(f"Données traitées pour {ticker} : Dernier cours={last_close}, EMA={last_ema}")
    except Exception as e:
        print(f"Erreur lors du traitement de {ticker} : {e}")
        continue

# Création d'un DataFrame à partir des résultats
results_df = pd.DataFrame(results)
if results_df.empty:
    print("Aucun résultat à exporter. Vérifiez vos tickers et les données disponibles.")
else:
    try:
        # Exportation en HTML
        results_df.to_html(output_html, index=False)
        print(f"Résultats exportés avec succès vers {output_html}")
    except Exception as e:
        print(f"Erreur lors de l'exportation des résultats : {e}")
        raise

# Résumé des résultats
print("\nRésumé des résultats :")
print(results_df)
