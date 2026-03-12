import os

# Dossier des plaques
dossier = r"C:\IE\Etudes\ET_Corélac\CORELAC_300_plaques_24_femelles_groupées\plaques_complètes(dernière version)"

# Scanner QR code ou coller le texte
nom_plaque = input("Collez le texte du QR code (ex: Plaque_001.xlsx) : ").strip()

chemin_fichier = os.path.join(dossier, nom_plaque)
if os.path.exists(chemin_fichier):
    os.startfile(chemin_fichier)  # ouvre directement avec Excel
    print(f"Ouverture de {chemin_fichier}")
else:
    print("❌ Fichier introuvable ! Vérifie le nom et le dossier.")
