from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# Charger le fichier Excel depuis Google Drive
#fichier_excel = '/content/drive/My Drive/MonDossier/votre_fichier.xlsx'
fichier_excel = './Depouillement_saint_louis_4.xlsx'

df = pd.read_excel(fichier_excel)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Récupérer l'ID du formulaire
        id_recherche = request.form['id_recherche']
        
        # Appeler la fonction pour rechercher l'ID et sauvegarder le résultat
        message = rechercher_par_id_et_ajouter(id_recherche)
        
        # Retourner la réponse avec le message et le lien de téléchargement
        return render_template('index.html', message=message)
    
    return render_template('index.html')

def rechercher_par_id_et_ajouter(id_recherche):
    # Obtenir la date et l'heure actuelles
    maintenant = datetime.now()
    date_recherche = maintenant.strftime('%Y-%m-%d')  # Extraire la date
    heure_recherche = maintenant.strftime('%H:%M:%S')  # Extraire l'heure
    print(id_recherche)
    # Filtrer le DataFrame selon l'ID
    #resultat = df[df['ID'] == id_recherche]
    resultat = df[df['ID'].astype(str).str.strip() == str(id_recherche).strip()]
    print(resultat)
    
    if not resultat.empty:
        # Ajouter les colonnes Date et Heure de recherche au résultat
        resultat['Date de recherche'] = date_recherche
        resultat['Heure de recherche'] = heure_recherche
        
        # Définir le nom du fichier de sortie
        # fichier_resultat = '/content/drive/My Drive/MonDossier/resultats_recherche.xlsx'
        fichier_resultat = './resultats_recherche.xlsx'
        
        # Vérifier si le fichier existe déjà
        if os.path.exists(fichier_resultat):
            # Si le fichier existe, charger les données existantes
            df_exist = pd.read_excel(fichier_resultat)
            # Ajouter le nouveau résultat
            df_concat = pd.concat([df_exist, resultat], ignore_index=True)
        else:
            # Si le fichier n'existe pas, utiliser directement le résultat
            df_concat = resultat
        
        # Sauvegarder le fichier avec les données mises à jour
        df_concat.to_excel(fichier_resultat, index=False)
        
        return f"PIGEON CONSTATE {fichier_resultat}"
    else:
        return f"CODE ERRONE {id_recherche} le {date_recherche} à {heure_recherche}"


@app.route('/download')
def download_file():
    file_path = './resultats_recherche.xlsx'  # Remplacez par le chemin du fichier sur le serveur
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print(df.head())  # Affiche les premières lignes
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return "Fichier introuvable", 404
        

if __name__ == '__main__':
    app.run(debug=True)
