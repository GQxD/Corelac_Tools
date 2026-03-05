import os
import requests
import xml.etree.ElementTree as ET
import qrcode
import csv

# ---------------- CONFIG ----------------
# Lien public de partage Nextcloud (dossier partagé)
nextcloud_public_link = "A_REMPLACER_PAR_LIEN_PUBLIC_NEXTCLOUD"

# Dossier local de sortie pour les PNG de QR codes + le CSV des liens
output_dir = r"A_REMPLACER_PAR_CHEMIN_DOSSIER"
os.makedirs(output_dir, exist_ok=True)

# Fichier CSV pour lister les liens
csv_file = os.path.join(output_dir, "liens_qr_codes.csv")
# ----------------------------------------

# Extraire le token du lien public (ex: https://nextcloud.inrae.fr/s/ABCD1234xyz)
public_token = nextcloud_public_link.rstrip("/").split("/")[-1]

# URL WebDAV publique
webdav_url = f"https://nextcloud.inrae.fr/public.php/webdav/"

# RÃ©cupÃ©rer la liste des fichiers via WebDAV
response = requests.request("PROPFIND", webdav_url, headers={"Depth": "1"}, auth=(public_token, ""))
if response.status_code != 207:
    raise RuntimeError(f"Impossible d'accÃ©der Ã  la liste WebDAV, code {response.status_code}")

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

print(f"Fichiers dÃ©tectÃ©s : {len(fichiers)}")

# GÃ©nÃ©rer liens publics et QR codes
with open(csv_file, "w", newline='', encoding='utf-8') as csvf:
    writer = csv.writer(csvf)
    writer.writerow(["Nom_du_fichier", "Lien_public"])
    
    for f in fichiers:
        # Construire le lien public tÃ©lÃ©chargeable
        lien_public = f"{nextcloud_public_link}/download?path=%2F&files={f}"
        writer.writerow([f, lien_public])
        
        # GÃ©nÃ©rer QR code
        qr = qrcode.make(lien_public)
        qr_file = os.path.join(output_dir, f"{os.path.splitext(f)[0]}.png")
        qr.save(qr_file)

print(f" QR codes gÃ©nÃ©rÃ©s dans : {output_dir}")
print(f" CSV des liens : {csv_file}")

