# -*- coding: utf-8 -*-
"""
Automatisation Nextcloud -> QR codes OnlyOffice (API WebDAV).
- Utilise NEXTCLOUD_USERNAME / NEXTCLOUD_PASSWORD si definis
- Sinon, ouvre une popup pour les saisir
- Fallback terminal si popup indisponible
"""

import csv
import os
import getpass
import re
import xml.etree.ElementTree as ET
from urllib.parse import quote

import qrcode
import requests
from requests.auth import HTTPBasicAuth

# ---------------- CONFIG ----------------
NEXTCLOUD_URL = "https://nextcloud.inrae.fr"
FOLDER_PATH = "/carrtel-documents-collaboratifs/Corélac/TEst/plaques_modifiees"
OUTPUT_DIR = r"C:\IE\Etudes\ET_Corélac\QR_codes"
os.makedirs(OUTPUT_DIR, exist_ok=True)
CSV_FILE = os.path.join(OUTPUT_DIR, "liens_onlyoffice.csv")
PLAQUE_FILE_RE = re.compile(r"^Plaque_\d{3}\.xlsx$", re.IGNORECASE)


def resolve_credentials():
    """Resolve les identifiants Nextcloud.

    Priorite:
    1) Variables d'environnement
    2) Popup Tkinter
    3) Saisie terminal
    """
    username = os.getenv("NEXTCLOUD_USERNAME")
    password = os.getenv("NEXTCLOUD_PASSWORD")

    if username and password:
        return username.strip(), password.strip()

    # Tente une popup graphique
    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        if not username:
            username = simpledialog.askstring(
                "Nextcloud", "Identifiant Nextcloud:", parent=root
            )

        if not password:
            password = simpledialog.askstring(
                "Nextcloud", "Token / mot de passe d'application:", show="*", parent=root
            )

        root.destroy()
    except Exception:
        pass

    # Fallback terminal
    if not username:
        username = input("Identifiant Nextcloud: ").strip()
    if not password:
        password = getpass.getpass("Token Nextcloud: ").strip()

    if not username or not password:
        raise ValueError("Identifiants Nextcloud manquants.")

    return username, password


def is_valid_plaque_file(filename):
    """Garde uniquement les fichiers plaques attendus."""
    if not filename:
        return False

    name = filename.strip()
    if name.startswith("~$") or name.lower().startswith("~temp"):
        return False
    if name.lower() == "plaque_000.xlsx":
        return False

    return PLAQUE_FILE_RE.match(name) is not None


def get_all_files_webdav(base_url, username, password, folder_path):
    """Recupere les fichiers d'un dossier via WebDAV."""
    webdav_url = f"{base_url}/remote.php/dav/files/{username}{folder_path}"
    headers = {"Depth": "1", "Content-Type": "application/xml"}

    body = """<?xml version="1.0"?>
    <d:propfind xmlns:d="DAV:" xmlns:oc="http://owncloud.org/ns" xmlns:nc="http://nextcloud.org/ns">
        <d:prop>
            <d:displayname />
            <oc:fileid />
            <d:resourcetype />
        </d:prop>
    </d:propfind>"""

    response = requests.request(
        "PROPFIND",
        webdav_url,
        headers=headers,
        data=body,
        auth=HTTPBasicAuth(username, password),
        timeout=60,
    )

    if response.status_code != 207:
        print(f"Erreur API: {response.status_code}")
        print(response.text)
        return []

    root = ET.fromstring(response.content)
    ns = {"d": "DAV:", "oc": "http://owncloud.org/ns"}

    fichiers = []
    for response_elem in root.findall(".//d:response", ns):
        resourcetype = response_elem.find(".//d:resourcetype", ns)
        is_collection = resourcetype is not None and resourcetype.find("d:collection", ns) is not None

        if not is_collection:
            name_elem = response_elem.find(".//d:displayname", ns)
            fileid_elem = response_elem.find(".//oc:fileid", ns)
            if name_elem is not None and fileid_elem is not None:
                name = (name_elem.text or "").strip()
                fileid = (fileid_elem.text or "").strip()
                if name and fileid and is_valid_plaque_file(name):
                    fichiers.append((name, fileid))

    return fichiers


def main():
    username, password = resolve_credentials()

    print("Recuperation de la liste des fichiers via API WebDAV...")
    fichiers = get_all_files_webdav(NEXTCLOUD_URL, username, password, FOLDER_PATH)

    if not fichiers:
        print("Aucun fichier trouve ou erreur d'authentification.")
        return

    print(f"{len(fichiers)} fichiers detectes.")
    print("Generation des QR codes...")

    encoded_folder = quote(FOLDER_PATH, safe="")

    with open(CSV_FILE, "w", newline="", encoding="utf-8") as csvf:
        writer = csv.writer(csvf)
        writer.writerow(["Nom_du_fichier", "Lien_onlyoffice"])

        for name, fileid in fichiers:
            encoded_name = quote(name, safe="")
            link = f"{NEXTCLOUD_URL}/apps/onlyoffice/{fileid}?filePath={encoded_folder}%2F{encoded_name}"
            writer.writerow([name, link])

            qr = qrcode.QRCode(
                version=2,
                error_correction=qrcode.constants.ERROR_CORRECT_H,
                box_size=6,
                border=2,
            )
            qr.add_data(link)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")

            safe_name = name.replace("/", "_").replace("\\", "_").replace(":", "_")
            output_path = os.path.join(OUTPUT_DIR, f"{safe_name}.png")
            img.save(output_path)
            print(f"OK: {name}")

    print(f"QR codes generes dans: {OUTPUT_DIR}")
    print(f"CSV genere: {CSV_FILE}")


if __name__ == "__main__":
    main()
