import os
import xlsxwriter

# --- CONFIGURATION ---
input_csv = r"C:\IE\Etudes\ET_Corégone\ET_Corélac\CORELAC_300_plaques_Aléa\12_grilles_5x5_CORELAC_LB-MATRICE.csv"
output_dir = r"C:\IE\Etudes\ET_Corélac\CORELAC_300_plaques_24_femelles_groupées"
os.makedirs(output_dir, exist_ok=True)

# --- LECTURE DU CSV ET EXTRACTION DES CROISEMENTS PAR GROUPE ---
groupes = []  # Liste de groupes, chaque groupe = liste de 25 croisements
groupes_noms = []  # Pour garder les noms des groupes

with open(input_csv, "r", encoding="utf-8") as f:
    lines = f.readlines()
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Identifier les lignes de titre de groupe (ex: B_x_B_G1)
        if line and '_G' in line and not ',' in line:
            nom_groupe = line
            print(f"📋 Groupe trouvé : {nom_groupe}")
            groupe_croisements = []
            i += 1
            
            # Ligne suivante = en-tête femelles (ignorer)
            if i < len(lines):
                i += 1
            
            # Lignes suivantes = croisements (5 lignes de mâles × 5 femelles = 25)
            for _ in range(5):
                if i < len(lines):
                    croi_line = lines[i].strip()
                    if croi_line and ',' in croi_line:
                        parts = croi_line.split(',')
                        # parts[0] = mâle, parts[1:6] = croisements
                        for croi in parts[1:6]:
                            croi = croi.strip()
                            if croi and 'x' in croi:  # Vérifier que c'est bien un croisement
                                groupe_croisements.append(croi)
                    i += 1
            
            if len(groupe_croisements) == 25:
                groupes.append(groupe_croisements)
                groupes_noms.append(nom_groupe)
        else:
            i += 1

print(f"✅ Nombre de groupes extraits : {len(groupes)}")
print(f"✅ Total croisements : {sum(len(g) for g in groupes)}")

# --- ORGANISATION DES GROUPES ---
# Groupes 0-2 : B×B G1-G3
# Groupes 3-5 : L×L G1-G3
# Groupes 6-8 : L♂×B♀ G1-G3 (femelles B)
# Groupes 9-11 : B♂×L♀ G1-G3 (femelles L)

# --- STRUCTURE DES PLAQUES ---
nb_plaques = 200
lignes = ['A', 'B', 'C', 'D']
colonnes = [1, 2, 3, 4, 5, 6]

# Plaques 1-100 → 5°C
# Plaques 101-200 → 9°C

# --- COULEURS ---
color_header = '#D9D9D9'
color_oeuf = '#00FF00'

# --- INITIALISATION DES PLAQUES ---
plaques_data = {}
for i in range(1, nb_plaques + 1):
    plaque_name = f"Plaque_{i:03d}"
    temp = 5 if i <= 100 else 9
    plaques_data[plaque_name] = {
        'temp': temp,
        'wells': {}  # {position: croisement_code}
    }

# --- RÉPARTITION AVEC REGROUPEMENT PAR FEMELLE ---
# Nouvelle logique : pour chaque position, on groupe les croisements par femelle
# Position 1 → F1, F6, F11 avec leurs 2 types de croisements (B×B + L×B ou L×L + B×L)

plaque_5C_counter = 1
plaque_9C_counter = 101

for pos in range(25):  # 25 positions dans chaque groupe (5×5)
    
    # --- SÉRIE 1 : Femelles B (B×B + L♂×B♀) ---
    # 3 femelles B × 2 types de croisements = 6 colonnes
    serie1_plaques = [
        plaque_5C_counter,      # Plaque N à 5°C
        plaque_5C_counter + 1,  # Plaque N+1 à 5°C
        plaque_9C_counter,      # Plaque N à 9°C
        plaque_9C_counter + 1   # Plaque N+1 à 9°C
    ]
    
    col_idx = 0
    # Pour chaque groupe B×B (groupes 0-2)
    for g_idx in range(3):
        if g_idx < len(groupes) and pos < len(groupes[g_idx]):
            # Croisement B×B
            croi_BxB = groupes[g_idx][pos]
            # Croisement L♂×B♀ correspondant (groupe 6-8)
            croi_LxB = groupes[6 + g_idx][pos] if (6 + g_idx) < len(groupes) else None
            
            # Placer B×B en colonne col_idx
            for plaque_num in serie1_plaques:
                plaque_name = f"Plaque_{plaque_num:03d}"
                for ligne in lignes:
                    position = f"{ligne}{col_idx + 1}"
                    plaques_data[plaque_name]['wells'][position] = croi_BxB
            
            # Placer L♂×B♀ en colonne col_idx+1
            if croi_LxB:
                for plaque_num in serie1_plaques:
                    plaque_name = f"Plaque_{plaque_num:03d}"
                    for ligne in lignes:
                        position = f"{ligne}{col_idx + 2}"
                        plaques_data[plaque_name]['wells'][position] = croi_LxB
            
            col_idx += 2  # Passer à la paire de colonnes suivante
    
    plaque_5C_counter += 2
    plaque_9C_counter += 2
    
    # --- SÉRIE 2 : Femelles L (L×L + B♂×L♀) ---
    # 3 femelles L × 2 types de croisements = 6 colonnes
    serie2_plaques = [
        plaque_5C_counter,      # Plaque N+2 à 5°C
        plaque_5C_counter + 1,  # Plaque N+3 à 5°C
        plaque_9C_counter,      # Plaque N+2 à 9°C
        plaque_9C_counter + 1   # Plaque N+3 à 9°C
    ]
    
    col_idx = 0
    # Pour chaque groupe L×L (groupes 3-5)
    for g_idx in range(3, 6):
        if g_idx < len(groupes) and pos < len(groupes[g_idx]):
            # Croisement L×L
            croi_LxL = groupes[g_idx][pos]
            # Croisement B♂×L♀ correspondant (groupe 9-11)
            croi_BxL = groupes[6 + g_idx][pos] if (6 + g_idx) < len(groupes) else None
            
            # Placer L×L en colonne col_idx
            for plaque_num in serie2_plaques:
                plaque_name = f"Plaque_{plaque_num:03d}"
                for ligne in lignes:
                    position = f"{ligne}{col_idx + 1}"
                    plaques_data[plaque_name]['wells'][position] = croi_LxL
            
            # Placer B♂×L♀ en colonne col_idx+1
            if croi_BxL:
                for plaque_num in serie2_plaques:
                    plaque_name = f"Plaque_{plaque_num:03d}"
                    for ligne in lignes:
                        position = f"{ligne}{col_idx + 2}"
                        plaques_data[plaque_name]['wells'][position] = croi_BxL
            
            col_idx += 2  # Passer à la paire de colonnes suivante
    
    plaque_5C_counter += 2
    plaque_9C_counter += 2

print(f"✅ Répartition terminée sur {plaque_5C_counter - 1} plaques à 5°C et {plaque_9C_counter - 101} plaques à 9°C")

# --- GÉNÉRATION DES FICHIERS EXCEL ---
for plaque_name, data in plaques_data.items():
    file_path = os.path.join(output_dir, f"{plaque_name}.xlsx")
    workbook = xlsxwriter.Workbook(file_path)
    
    # === FEUILLE 1 : DISPOSITION ===
    ws = workbook.add_worksheet("Disposition")
    
    fmt_header = workbook.add_format({'bg_color': color_header, 'bold': True, 'align': 'center'})
    fmt_oeuf = workbook.add_format({'bg_color': color_oeuf, 'align': 'center'})
    
    # En-têtes colonnes
    for idx, col in enumerate(colonnes):
        ws.write(0, idx + 1, col, fmt_header)
    
    # En-têtes lignes + contenu
    for r_idx, ligne in enumerate(lignes):
        ws.write(r_idx + 1, 0, ligne, fmt_header)
        
        for c_idx, col in enumerate(colonnes):
            pos = f"{ligne}{col}"
            if pos in data['wells']:
                ws.write(r_idx + 1, c_idx + 1, data['wells'][pos], fmt_oeuf)
            else:
                ws.write(r_idx + 1, c_idx + 1, "", workbook.add_format({'bg_color': '#FFFFFF'}))
    
    ws.set_column(0, 6, 20)
    
    # === FEUILLE 2 : SUIVI ===
    suivi_ws = workbook.add_worksheet("Suivi")
    
    headers = [
        "Well ID",
        "Row", 
        "Column",
        "Cross",
        "Temperature (°C)",
        "Fertilization Date",
        "Eyespot Stage Date",
        "Hatching Date",
        "Status",
        "Death Date",
        "Notes"
    ]
    
    fmt_suivi_header = workbook.add_format({
        'bg_color': '#4472C4',
        'font_color': 'white',
        'bold': True,
        'align': 'center',
        'border': 1
    })
    
    for col_idx, header in enumerate(headers):
        suivi_ws.write(0, col_idx, header, fmt_suivi_header)
    
    # Remplir les données de suivi
    row_idx = 1
    for ligne in lignes:
        for col in colonnes:
            pos = f"{ligne}{col}"
            croi = data['wells'].get(pos, "")
            
            suivi_ws.write(row_idx, 0, pos)
            suivi_ws.write(row_idx, 1, ligne)
            suivi_ws.write(row_idx, 2, col)
            suivi_ws.write(row_idx, 3, croi)
            suivi_ws.write(row_idx, 4, data['temp'])
            # Colonnes dates et notes vides
            for i in range(5, 11):
                suivi_ws.write(row_idx, i, "")
            
            row_idx += 1
    
    suivi_ws.set_column(0, 0, 10)  # Well ID
    suivi_ws.set_column(1, 2, 8)   # Row, Column
    suivi_ws.set_column(3, 3, 25)  # Cross
    suivi_ws.set_column(4, 4, 15)  # Temperature
    suivi_ws.set_column(5, 9, 18)  # Dates et Status
    suivi_ws.set_column(10, 10, 30) # Notes
    
    # === FEUILLE 3 : INFOS ===
    info_ws = workbook.add_worksheet("Infos")
    
    info_ws.write(0, 0, "Plate ID:", workbook.add_format({'bold': True}))
    info_ws.write(0, 1, plaque_name)
    
    info_ws.write(1, 0, "Temperature:", workbook.add_format({'bold': True}))
    info_ws.write(1, 1, f"{data['temp']}°C")
    
    info_ws.write(3, 0, "Crosses in this plate:", workbook.add_format({'bold': True}))
    
    crosses_in_plate = sorted(set(data['wells'].values()))
    for idx, croi in enumerate(crosses_in_plate):
        info_ws.write(4 + idx, 0, croi)
    
    # Ajouter info sur les femelles présentes
    femelles_in_plate = set()
    for croi in crosses_in_plate:
        if 'x' in croi:
            femelle = croi.split('x')[1]
            femelles_in_plate.add(femelle)
    
    info_ws.write(4 + len(crosses_in_plate) + 1, 0, "Females in this plate:", workbook.add_format({'bold': True}))
    for idx, femelle in enumerate(sorted(femelles_in_plate)):
        info_ws.write(5 + len(crosses_in_plate) + 1 + idx, 0, femelle)
    
    info_ws.set_column(0, 0, 25)
    info_ws.set_column(1, 1, 20)
    
    workbook.close()

print(f"✅ Génération terminée : {nb_plaques} plaques créées dans '{output_dir}'")
print(f"   - Plaques 001-100 : 5°C")
print(f"   - Plaques 101-200 : 9°C")
print(f"   - Organisation : croisements regroupés par femelle")
print(f"   - Exemple Plaque_001 colonnes 1-2 : B_F1 (B×B + L×B)")
print(f"   - Exemple Plaque_001 colonnes 3-4 : B_F6 (B×B + L×B)")
print(f"   - Exemple Plaque_001 colonnes 5-6 : B_F11 (B×B + L×B)")