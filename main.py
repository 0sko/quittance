import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from num2words import num2words
import json
import os
import textwrap
import re

CONFIG_FILE = "config_quittance.json"
DOSSIER_PDF = os.getcwd()

# ---------------------------
# Chargement / sauvegarde config
# ---------------------------
def charger_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def sauvegarder_config(data):
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=4)

config = charger_config()
SIGNATURE_PATH = config.get("signature", "")
DOSSIER_PDF = config.get("dossier_pdf", DOSSIER_PDF)

# ---------------------------
# Choisir fichier signature
# ---------------------------
def choisir_signature():
    global SIGNATURE_PATH
    fichier = filedialog.askopenfilename(
        title="Choisir l'image de signature",
        filetypes=[("Images", "*.png *.jpg *.jpeg")]
    )
    if fichier:
        SIGNATURE_PATH = fichier
        label_signature.config(text=os.path.basename(fichier))
        # Mise à jour config immédiatement
        config["signature"] = SIGNATURE_PATH
        sauvegarder_config(config)

# ---------------------------
# Choisir dossier PDF
# ---------------------------
def choisir_dossier():
    global DOSSIER_PDF
    dossier = filedialog.askdirectory()
    if dossier:
        DOSSIER_PDF = dossier
        label_dossier.config(text=dossier)
        # Mise à jour config immédiatement
        config["dossier_pdf"] = DOSSIER_PDF
        sauvegarder_config(config)

# ---------------------------
# Génération PDF
# ---------------------------
def generer_quittance():
    # --- Collecte des données ---
    nom_proprietaire = entry_nom_proprietaire.get().strip()
    prenom_proprietaire = entry_prenom_proprietaire.get().strip()
    adresse_proprietaire = entry_adresse_proprietaire.get().strip()
    cp_proprietaire = entry_cp_proprietaire.get().strip()
    ville_proprietaire = entry_ville_proprietaire.get().strip()

    civilite = civilite_var.get()
    nom_locataire = entry_nom_locataire.get().strip()
    prenom_locataire = entry_prenom_locataire.get().strip()
    adresse_location = entry_adresse_location.get().strip()
    cp_location = entry_cp_location.get().strip()
    ville_location = entry_ville_location.get().strip()

    montant_str = entry_montant.get().strip()
    mois = mois_var.get()
    annee = entry_annee.get().strip()

    # --- Validation complète ---
    if not all([nom_proprietaire, prenom_proprietaire, adresse_proprietaire, cp_proprietaire, ville_proprietaire,
                nom_locataire, prenom_locataire, adresse_location, cp_location, ville_location,
                montant_str, mois, annee]):
        messagebox.showerror("Erreur", "Tous les champs doivent être remplis")
        return

    # --- Montant ---
    try:
        montant = float(montant_str.replace(",", "."))
    except ValueError:
        messagebox.showerror("Erreur", "Montant invalide")
        return

    montant_lettres = num2words(montant, lang="fr")

    # --- Nettoyage nom pour fichier ---
    nom_locataire_clean = re.sub(r'[\\/*?:"<>|]', "_", nom_locataire)
    fichier = os.path.join(DOSSIER_PDF, f"Quittance_{mois}_{annee}_{nom_locataire_clean}.pdf")

    # --- Sauvegarde complète dans config ---
    config_to_save = {
        "proprietaire": {
            "nom": nom_proprietaire,
            "prenom": prenom_proprietaire,
            "adresse": adresse_proprietaire,
            "cp": cp_proprietaire,
            "ville": ville_proprietaire
        },
        "locataire": {
            "civilite": civilite,
            "nom": nom_locataire,
            "prenom": prenom_locataire,
            "adresse": adresse_location,
            "cp": cp_location,
            "ville": ville_location
        },
        "loyer": {
            "montant": montant,
            "mois": mois,
            "annee": annee
        },
        "signature": SIGNATURE_PATH
    }
    sauvegarder_config(config_to_save)

    # --- Création PDF ---
    c = canvas.Canvas(fichier, pagesize=A4)
    width, height = A4
    y = height - 80

    # --- Métadonnées ---
    c.setTitle(f"Quittance de loyer - {mois} {annee} - {prenom_locataire} {nom_locataire}")
    c.setAuthor(f"{prenom_proprietaire} {nom_proprietaire}")
    c.setSubject("Quittance de loyer")
    c.setKeywords(f"quittance, loyer, {mois}, {annee}, {prenom_locataire}, {nom_locataire}")

    # fonts
    # Chemin vers le TTF
    LIB_SERIF_PATH = "LiberationSerif-Regular.ttf"
    LIB_SERIF_BOLD_PATH = "LiberationSerif-Bold.ttf"

    # Nom interne pour ReportLab
    police = "LiberationSerif"
    police_bold = "LiberationSerif-Bold"

    # Police par défaut
    police_default = "Helvetica"
    police_bold_default = "Helvetica-Bold"

    try:
        pdfmetrics.registerFont(TTFont(police, LIB_SERIF_PATH))
        pdfmetrics.registerFont(TTFont(police_bold, LIB_SERIF_BOLD_PATH))
    except Exception as e:
        print(f"Impossible de charger {LIB_SERIF_PATH} : {e}")
        police = police_default
        police_bold = police_bold_default

    c.setFont(police_bold, 12)
    c.drawCentredString(width/2, y, f"Quittance de loyer du mois de {mois} {annee}")

    y -= 50
    c.setFont(police, 12)

    # Propriétaire
    c.drawString(60, y, f"{prenom_proprietaire} {nom_proprietaire}")
    y -= 15
    c.drawString(60, y, adresse_proprietaire)
    y -= 15
    c.drawString(60, y, f"{cp_proprietaire} {ville_proprietaire}")
    y -= 15
    c.drawString(60, y, "France")

    # Locataire
    y -= 15
    c.drawString(350, y, f"{prenom_locataire} {nom_locataire}")
    y -= 15
    c.drawString(350, y, adresse_location)
    y -= 15
    c.drawString(350, y, f"{cp_location} {ville_location}")
    y -= 15
    c.drawString(350, y, "France")

    y -= 40
    date_du_jour = datetime.today().strftime("%d/%m/%Y")
    c.drawRightString(width-60, y, f"Fait le {date_du_jour}")

    y -= 40
    c.drawString(60, y, "Adresse de la location :")
    y -= 30
    c.drawString(200, y, f"{adresse_location} {cp_location} {ville_location} France")

    y -= 40
    texte = (
        f"Je soussigné {prenom_proprietaire} {nom_proprietaire}, propriétaire du logement désigné ci-dessus, "
        f"déclare avoir reçu de {civilite} {prenom_locataire} {nom_locataire} la somme de {montant} € "
        f"({montant_lettres} euros) pour la période du mois de {mois} {annee} "
        f"et lui en donne quittance, sous réserve de tous mes droits."
    )

    text = c.beginText(60, y)
    text.setLeading(16)
    max_chars = int((width-20) / 6)  # approx 6 points per character
    for line in textwrap.wrap(texte, max_chars):
        text.textLine(line)
    c.drawText(text)

    # Signature
    y -= 80
    c.drawString(60, y, f"{prenom_proprietaire} {nom_proprietaire}")
    y -= 150

    if SIGNATURE_PATH and os.path.exists(SIGNATURE_PATH):
        img = ImageReader(SIGNATURE_PATH)
        iw, ih = img.getSize()
        max_width = 200
        max_height = 150
        aspect = ih / iw
        aspect1 = iw/ih
        c.drawImage(img, 60, y, width=max_height*aspect1, height=max_height, mask='auto')

    y -= 60
    texte = (
        f"Cette quittance annule tous les reçus qui auraient pu être établis précédemment en cas de paiement "
        f"partiel du montant du présent terme. Elle est à conserver pendant trois ans par le locataire (loi n° 89-"
        f"462 du 6 juillet 1989 : art. 7-1). \nTexte de référence : loi du 6.7.89 : art. 21"
    )

    text = c.beginText(60, y)
    text.setLeading(16)
    max_chars = 100
    for line in textwrap.wrap(texte, max_chars):
        text.textLine(line)
    c.drawText(text)

    c.save()

    # Nom du fichier texte
    fichier_txt = os.path.join(DOSSIER_PDF, f"mail_quittance_{mois}_{annee}_{nom_locataire_clean}.txt")

    # Contenu du mail
    contenu_mail = f"""
-------------------------------------------------------------------------------------
############################### Version informelle ##################################
-------------------------------------------------------------------------------------

Bonjour {prenom_locataire},
\n
Tu peux trouver ci-joint la quittance de loyer de {mois} {annee}.
\n
Bonne journée,
\n
{prenom_proprietaire} {nom_proprietaire}

-------------------------------------------------------------------------------------
################################# Version formelle ##################################
-------------------------------------------------------------------------------------

Bonjour {prenom_locataire} {nom_locataire},
\n
Veuillez trouver ci-joint la quittance de loyer de {mois} {annee}.
\n
Cordialement,
\n
{prenom_proprietaire} {nom_proprietaire}
    """

    # Écriture dans le fichier
    with open(fichier_txt, "w", encoding="utf-8") as f:
        f.write(contenu_mail)

    messagebox.showinfo("Succès", f"PDF généré : {fichier}")

# ---------------------------
# Interface graphique
# ---------------------------
root = tk.Tk()
root.title("Générateur de quittance de loyer")
root.geometry("425x700")

# -- Frame Propriétaire
frame_prop = tk.LabelFrame(root, text="Propriétaire", padx=10, pady=10)
frame_prop.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

tk.Label(frame_prop, text="Nom").grid(row=0, column=0)
entry_nom_proprietaire = tk.Entry(frame_prop)
entry_nom_proprietaire.grid(row=0, column=1)

tk.Label(frame_prop, text="Prénom").grid(row=1, column=0)
entry_prenom_proprietaire = tk.Entry(frame_prop)
entry_prenom_proprietaire.grid(row=1, column=1)

tk.Label(frame_prop, text="Adresse").grid(row=2, column=0)
entry_adresse_proprietaire = tk.Entry(frame_prop)
entry_adresse_proprietaire.grid(row=2, column=1)

tk.Label(frame_prop, text="Code postal").grid(row=3, column=0)
entry_cp_proprietaire = tk.Entry(frame_prop)
entry_cp_proprietaire.grid(row=3, column=1)

tk.Label(frame_prop, text="Ville").grid(row=4, column=0)
entry_ville_proprietaire = tk.Entry(frame_prop)
entry_ville_proprietaire.grid(row=4, column=1)

# -- Frame Locataire
frame_loc = tk.LabelFrame(root, text="Locataire", padx=10, pady=10)
frame_loc.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

civilite_var = tk.StringVar(value="Monsieur")
tk.Radiobutton(frame_loc, text="Monsieur", variable=civilite_var, value="Monsieur").grid(row=0,column=0)
tk.Radiobutton(frame_loc, text="Madame", variable=civilite_var, value="Madame").grid(row=0,column=1)

tk.Label(frame_loc,text="Nom").grid(row=1,column=0)
entry_nom_locataire = tk.Entry(frame_loc)
entry_nom_locataire.grid(row=1,column=1)

tk.Label(frame_loc,text="Prénom").grid(row=2,column=0)
entry_prenom_locataire = tk.Entry(frame_loc)
entry_prenom_locataire.grid(row=2,column=1)

tk.Label(frame_loc,text="Adresse logement").grid(row=3,column=0)
entry_adresse_location = tk.Entry(frame_loc)
entry_adresse_location.grid(row=3,column=1)

tk.Label(frame_loc,text="Code postal").grid(row=4,column=0)
entry_cp_location = tk.Entry(frame_loc)
entry_cp_location.grid(row=4,column=1)

tk.Label(frame_loc,text="Ville").grid(row=5,column=0)
entry_ville_location = tk.Entry(frame_loc)
entry_ville_location.grid(row=5,column=1)

# -- Frame Quittance
frame_quittance = tk.LabelFrame(root, text="Informations quittance", padx=10, pady=10)
frame_quittance.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

tk.Label(frame_quittance,text="Montant (€)").grid(row=0,column=0)
entry_montant = tk.Entry(frame_quittance)
entry_montant.grid(row=0,column=1)

tk.Label(frame_quittance,text="Mois").grid(row=1,column=0)
mois_var = tk.StringVar()
ttk.Combobox(
    frame_quittance,
    textvariable=mois_var,
    values=["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
).grid(row=1,column=1)

tk.Label(frame_quittance,text="Année").grid(row=2,column=0)
entry_annee = tk.Entry(frame_quittance)
entry_annee.grid(row=2,column=1)

# -- Choisir dossier PDF
frame_dossier = tk.Frame(root)
frame_dossier.grid(row=3,column=0,pady=10)
tk.Button(frame_dossier,text="Choisir dossier PDF",command=choisir_dossier).grid(row=0,column=0)
label_dossier = tk.Label(frame_dossier,text=DOSSIER_PDF)
label_dossier.grid(row=4,column=0,padx=10)

# -- Choisir signature
frame_signature = tk.Frame(root)
frame_signature.grid(row=5, column=0, pady=10)
tk.Button(frame_signature,text="Choisir signature",command=choisir_signature).grid(row=0, column=0)
label_signature = tk.Label(frame_signature,text=os.path.basename(SIGNATURE_PATH) if SIGNATURE_PATH else "Aucune signature")
label_signature.grid(row=6, column=0, padx=10)

# -- Bouton génération
tk.Button(root,text="Générer la quittance",command=generer_quittance,width=25).grid(pady=20)

# --- Remplissage automatique des champs à l'ouverture ---
if "proprietaire" in config:
    entry_nom_proprietaire.insert(0, config["proprietaire"].get("nom", ""))
    entry_prenom_proprietaire.insert(0, config["proprietaire"].get("prenom", ""))
    entry_adresse_proprietaire.insert(0, config["proprietaire"].get("adresse", ""))
    entry_cp_proprietaire.insert(0, config["proprietaire"].get("cp", ""))
    entry_ville_proprietaire.insert(0, config["proprietaire"].get("ville", ""))

if "locataire" in config:
    civilite_var.set(config["locataire"].get("civilite", "Monsieur"))
    entry_nom_locataire.insert(0, config["locataire"].get("nom", ""))
    entry_prenom_locataire.insert(0, config["locataire"].get("prenom", ""))
    entry_adresse_location.insert(0, config["locataire"].get("adresse", ""))
    entry_cp_location.insert(0, config["locataire"].get("cp", ""))
    entry_ville_location.insert(0, config["locataire"].get("ville", ""))

if "loyer" in config:
    entry_montant.insert(0, str(config["loyer"].get("montant", "")))
    mois_var.set(config["loyer"].get("mois", ""))
    entry_annee.delete(0, tk.END)
    entry_annee.insert(0, config["loyer"].get("annee", datetime.now().year))

root.mainloop()