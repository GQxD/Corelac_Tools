import os

# Dossier local contenant les fichiers Excel de plaques (Plaque_XXX.xlsx)
dossier = r"A_REMPLACER_PAR_CHEMIN_DOSSIER"

# Scanner QR code ou coller le texte
nom_plaque = input("Collez le texte du QR code (ex: Plaque_001.xlsx) : ").strip()

chemin_fichier = os.path.join(dossier, nom_plaque)
if os.path.exists(chemin_fichier):
    os.startfile(chemin_fichier)  # ouvre directement avec Excel
    print(f"Ouverture de {chemin_fichier}")
else:
    print("âŒ Fichier introuvable ! VÃ©rifie le nom et le dossier.")

