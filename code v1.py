import pandas as pd
import time
import random
import re
import json
import logging
import os
from datetime import datetime
from DrissionPage import ChromiumPage, ChromiumOptions
from openai import OpenAI
from html.parser import HTMLParser

# ==========================================
# CONFIGURATION ET SETUP
# ==========================================
GROK_API_KEY = os.environ.get("GROK_API_KEY")
FICHIER_ENTREE = "Exemple Scrap KP (1).xlsx"
DOSSIER_SORTIE = "Resultats_Extraction"

client = OpenAI(api_key=GROK_API_KEY, base_url="https://api.x.ai/v1")

os.makedirs(DOSSIER_SORTIE, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Structure exacte de ton fichier d'exemple
COLONNES_CIBLES = [
    'URL', 'Nom Entreprise', 'Rôles', 'Description', 'Code Postal', 'Ville', 
    'Pays', 'Téléphone', 'Fax', 'Email', 'Site Web', 'SIREN', 'SIRET', 'TVA', 
    'Capital', 'Forme Juridique', 'Année Création', 'Effectif Adresse', 
    'Effectif Entreprise', 'Activités Principales', 'Activités Secondaires', 
    'Autres Classifications', 'Parc Machine / Outil de Production', 'Erreur'
]

# ==========================================
# EXTRACTEUR DE TEXTE (Nettoyage HTML)
# ==========================================
class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.text_blocks = []
        self.current_block = []
        self.skip_depth = 0

    def handle_starttag(self, tag, attrs):
        if tag.lower() in {"script", "style", "noscript", "svg", "canvas", "nav", "footer", "header"}:
            self.skip_depth += 1
        elif tag.lower() in {"div", "p", "br", "li", "tr", "h1", "h2", "h3", "td"}:
            if self.current_block:
                self.text_blocks.append(" ".join(self.current_block))
                self.current_block = []

    def handle_endtag(self, tag):
        if tag.lower() in {"script", "style", "noscript", "svg", "canvas", "nav", "footer", "header"}:
            self.skip_depth = max(0, self.skip_depth - 1)
        elif tag.lower() in {"div", "p", "br", "li", "tr", "h1", "h2", "h3", "td"}:
            if self.current_block:
                self.text_blocks.append(" ".join(self.current_block))
                self.current_block = []

    def handle_data(self, data):
        if self.skip_depth == 0:
            txt = data.replace('\xa0', ' ').strip()
            if txt:
                self.current_block.append(txt)

# ==========================================
# MODULE 1 : EXTRACTION SCRAPING (DOM + REGEX)
# ==========================================
def extraction_scraping(url, page):
    logging.info(f"[SCRAPING] Début de l'analyse pour : {url}")
    donnees = {col: "" for col in COLONNES_CIBLES}
    donnees['URL'] = url
    
    try:
        page.get(url)
        page.wait.load_start()
        time.sleep(1.5)
        
        # 1. Tenter d'ouvrir l'onglet Produits (souvent là où se cache le parc machine)
        try:
            onglet = page.ele('xpath://a[contains(@href, "produits-et-services")]', timeout=1.5)
            if onglet:
                onglet.click(by_js=True)
                time.sleep(2)
        except: pass

        # 2. Scroll pour forcer le chargement de la page complète
        page.scroll.to_bottom()
        time.sleep(1)
        page.scroll.to_top()
        
        # 3. Ouvrir tous les accordéons
        try:
            page.run_js("document.querySelectorAll('a, button, span').forEach(e => { let t = e.innerText.toLowerCase(); if(t.includes('afficher') || t.includes('voir plus')) e.click(); });")
            time.sleep(1)
        except: pass

        html_brut = page.html
        
        # 4. JSON-LD (Données structurées SEO)
        json_blocks = re.findall(r'<script type="application/ld\+json">(.*?)</script>', html_brut, re.DOTALL)
        for block in json_blocks:
            try:
                jdata = json.loads(block)
                if 'telephone' in jdata: donnees['Téléphone'] = jdata.get('telephone')
                if 'name' in jdata: donnees['Nom Entreprise'] = jdata.get('name')
                if 'url' in jdata: donnees['Site Web'] = jdata.get('url')
                if 'address' in jdata:
                    donnees['Code Postal'] = jdata['address'].get('postalCode', '')
                    donnees['Ville'] = jdata['address'].get('addressLocality', '')
                    donnees['Pays'] = jdata['address'].get('addressCountry', 'France')
            except: pass

        # 5. Extraction via le DOM Tableaux (Très fiable pour le légal)
        for tr in page.eles('tag:tr'):
            txt = tr.text.upper()
            if "SIREN" in txt: donnees['SIREN'] = txt.replace("SIREN", "").replace("\n", "").strip()
            elif "SIRET" in txt: donnees['SIRET'] = txt.replace("SIRET", "").replace("\n", "").strip()
            elif "TVA" in txt: donnees['TVA'] = txt.replace("NUMÉRO DE TVA", "").replace("TVA", "").replace("\n", "").strip()
            elif "CAPITAL" in txt: donnees['Capital'] = txt.replace("CAPITAL", "").replace("\n", "").strip()
            elif "FORME JURIDIQUE" in txt: donnees['Forme Juridique'] = txt.replace("FORME JURIDIQUE", "").replace("\n", "").strip()
            elif "CRÉATION" in txt or "CREATION" in txt: donnees['Année Création'] = txt.replace("ANNÉE DE CRÉATION", "").replace("\n", "").strip()

        # 6. Extraction NAF et Activités DOM
        if naf := page.ele('xpath://*[contains(text(), "NAF Rev")]', timeout=0.5):
            donnees['Autres Classifications'] = naf.parent().text.replace('\n', ' | ').strip()

        if act := page.ele('xpath://*[@id="activities-tree"]', timeout=0.5):
            donnees['Activités Principales'] = act.text.replace('\n', ' | ').strip()

        if role := page.ele('xpath://*[contains(@class, "companyRoles")]', timeout=0.5):
            donnees['Rôles'] = role.text.replace('\n', ' ').strip()

        # Parc Machine DOM
        machine_xpath = 'xpath://*[contains(translate(text(), "PARC MACHINEoutil de production", "parc machineoutil de production"), "parc machine") or contains(translate(text(), "PARC MACHINEoutil de production", "parc machineoutil de production"), "outil de production")]/ancestor::div[contains(@class, "block")][1]'
        if bloc_machine := page.ele(machine_xpath, timeout=0.5):
            donnees['Parc Machine / Outil de Production'] = bloc_machine.text.replace('\n', ' | ').strip()

        # 7. Préparation du texte brut pour l'IA
        parser = TextExtractor()
        parser.feed(html_brut)
        if parser.current_block: parser.text_blocks.append(" ".join(parser.current_block))
        blocks_propres = [re.sub(r'\s+', ' ', b).strip() for b in parser.text_blocks if len(b.strip()) > 1]
        
        # On sauvegarde le texte brut dans une variable temporaire non exportée
        donnees['_texte_brut_pour_ia'] = " | ".join(blocks_propres)[:6000]

    except Exception as e:
        logging.error(f"[SCRAPING] Erreur sur {url} : {e}")
        donnees['Erreur'] = str(e)
        donnees['_texte_brut_pour_ia'] = ""

    return donnees

# ==========================================
# MODULE 2 : ENRICHISSEMENT IA (GROK)
# ==========================================
def extraction_ia(nom_entreprise, texte_brut):
    logging.info(f"[IA] Appel Grok pour {nom_entreprise}...")
    prompt = f"""
    Voici le texte extrait du profil B2B de l'entreprise : {nom_entreprise}.
    Texte brut :
    ---
    {texte_brut}
    ---

    Ta tâche est de restructurer ce texte pour combler les informations d'une base de données.
    DÉDUIS les informations si elles ne sont pas explicites, particulièrement pour le parc machine basé sur leur activité industrielle.
    
    Réponds EXCLUSIVEMENT avec un objet JSON contenant ces clés :
    "Description": Résumé de ce que fait l'entreprise.
    "Rôles": (Producteur, Distributeur, Prestataire).
    "Activités Principales": Les activités (séparées par |).
    "Activités Secondaires": Les activités (séparées par |).
    "Parc Machine / Outil de Production": Liste des machines et équipements. SI NON PRÉCISÉ, DÉDUIS l'équipement probable selon leur métier (ex: "Découpe laser, plieuses, postes à souder").
    "SIREN": Numéro à 9 chiffres.
    "SIRET": Numéro à 14 chiffres.
    "TVA": Numéro de TVA.
    "Capital": Montant du capital.
    "Forme Juridique": SAS, SARL...
    "Année Création": Année.
    "Fax": Numéro.
    "Email": Contact email.

    Si une info est introuvable et indéductible, mets "". Ne mets pas de markdown.
    """
    try:
        rep = client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Tu es un extracteur de données B2B. Tu réponds uniquement en JSON valide."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2 
        )
        content = rep.choices[0].message.content.strip()
        if content.startswith("```json"): content = content[7:-3].strip()
        return json.loads(content)
    except Exception as e:
        logging.error(f"[IA] Erreur API pour {nom_entreprise}: {e}")
        return {}

# ==========================================
# PIPELINE PRINCIPAL
# ==========================================
def executer_pipeline(limite=2):
    debut = time.perf_counter()
    logging.info("Démarrage du pipeline hybride (Scraping + IA)")
    
    df_entree = pd.read_excel(FICHIER_ENTREE)
    liens_a_traiter = df_entree.iloc[:, 0].dropna().head(limite).tolist()
    resultats_finaux = []
    
    co = ChromiumOptions().auto_port()
    co.headless() # INDISPENSABLE SUR LE CLOUD
    co.set_argument('--no-sandbox') # Nécessaire pour les serveurs Linux
    co.set_argument('--start-maximized')
    co.set_argument('--disable-blink-features=AutomationControlled')
    
    for url in liens_a_traiter:
        # 1. Récupération des données brutes
        donnees = extraction_scraping(url, navigateur)
        texte_brut = donnees.pop('_texte_brut_pour_ia', '')
        
        # 2. Vérification des manques critiques
        champs_critiques = ['SIREN', 'SIRET', 'Parc Machine / Outil de Production', 'Activités Principales']
        besoin_ia = False
        for champ in champs_critiques:
            if not donnees.get(champ):
                besoin_ia = True
                break
                
        # 3. Appel de l'IA si nécessaire
        if besoin_ia and texte_brut:
            nom = donnees.get('Nom Entreprise') or "Entreprise"
            donnees_ia = extraction_ia(nom, texte_brut)
            
            # Fusion : L'IA comble les trous, mais n'écrase pas ce que le scraper a trouvé de manière certaine
            for cle in COLONNES_CIBLES:
                if cle in donnees_ia and donnees_ia[cle]:
                    if not donnees.get(cle):
                        donnees[cle] = str(donnees_ia[cle])
                        
        resultats_finaux.append(donnees)
        time.sleep(random.uniform(3.0, 5.0))

    navigateur.quit()

    # Génération du fichier Excel
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom_fichier = os.path.join(DOSSIER_SORTIE, f"Kompass_Extrait_{horodatage}.xlsx")
    
    df_sortie = pd.DataFrame(resultats_finaux, columns=COLONNES_CIBLES)
    df_sortie.to_excel(nom_fichier, index=False)
    
    logging.info(f"Pipeline terminé en {time.perf_counter() - debut:.2f} secondes.")
    logging.info(f"Fichier de résultat : {nom_fichier}")

if __name__ == "__main__":
    executer_pipeline(limite=2)
