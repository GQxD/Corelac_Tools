import os
import requests
import xml.etree.ElementTree as ET
import qrcode
import csv

# ---------------- CONFIG ----------------
# ⚠️ Remplace ce lien par le lien public de du dossier Nextcloud
nextcloud_public_link = "https://nextcloud.inrae.fr/s/L6xzaMQsRysbqi3"

# Dossier local pour sauvegarder les QR codes
output_dir = r"C:\IE\Etudes\ET_Corélac\CORELAC_300_plaques\QR_codes"
os.makedirs(output_dir, exist_ok=True)

# Fichier CSV pour lister les liens
csv_file = os.path.join(output_dir, "liens_qr_codes.csv")
# ----------------------------------------

# Extraire le token du lien public (ex: https://nextcloud.inrae.fr/s/ABCD1234xyz)
public_token = nextcloud_public_link.rstrip("/").split("/")[-1]

# URL WebDAV publique
webdav_url = f"https://nextcloud.inrae.fr/public.php/webdav/"

# Récupérer la liste des fichiers via WebDAV
response = requests.request("PROPFIND", webdav_url, headers={"Depth": "1"}, auth=(public_token, ""))
if response.status_code != 207:
    raise RuntimeError(f"Impossible d'accéder à la liste WebDAV, code {response.status_code}")

# Parser XML pour extraire les noms de fichiers
root = ET.fromstring(response.text)
namespaces = {"d": "DAV:"}
fichiers = []

for response_elem in root.findall("d:response", namespaces):
    href = response_elem.find("d:href", namespaces)
    if href is not None:
        filename = os.path.basename(href.text)
        if filename.lower().endswith(".xlsx"):
            fichiers.append(filename)

print(f"Fichiers détectés : {len(fichiers)}")

# Générer liens publics et QR codes
with open(csv_file, "w", newline='', encoding='utf-8') as csvf:
    writer = csv.writer(csvf)
    writer.writerow(["Nom_du_fichier", "Lien_public"])
    
    for f in fichiers:
        # Construire le lien public téléchargeable
        lien_public = f"{nextcloud_public_link}/download?path=%2F&files={f}"
        writer.writerow([f, lien_public])
        
        # Générer QR code
        qr = qrcode.make(lien_public)
        qr_file = os.path.join(output_dir, f"{os.path.splitext(f)[0]}.png")
        qr.save(qr_file)

print(f" QR codes générés dans : {output_dir}")
print(f" CSV des liens : {csv_file}")
