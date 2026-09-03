import requests
import pandas as pd

def scraper_donner():
    url = "https://fr.donnermusic.com/products.json"
    liste_produits = []
    page = 1

    print("Récupération des produits...")

    while True:
        reponse = requests.get(url, params={'limit': 250, 'page': page})
        if reponse.status_code != 200:
            break
            
        donnees = reponse.json()
        produits = donnees.get('products', [])
        if not produits:
            break
            
        for p in produits:
            titre = p.get('title')
            for v in p.get('variants', []):
                nom_variante = v.get('title')
                designation = f"{titre} - {nom_variante}" if nom_variante and nom_variante != "Default Title" else titre
                prix_ttc = float(v.get('price', 0))

                liste_produits.append({
                    "Désignation du produit": designation,
                    "Prix TTC (€)": prix_ttc
                })
        page += 1

    df = pd.DataFrame(liste_produits)
    fichier_excel = "Prix_Donner_TTC.xlsx"
    
    with pd.ExcelWriter(fichier_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Prix")
        worksheet = writer.sheets["Prix"]
        worksheet.column_dimensions['A'].width = 60
        worksheet.column_dimensions['B'].width = 18

    print(f"Terminé ! {len(liste_produits)} produits enregistrés.")

if __name__ == "__main__":
    scraper_donner()
