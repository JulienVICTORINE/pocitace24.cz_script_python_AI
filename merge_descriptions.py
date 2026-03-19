import argparse
import csv
import html
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.request


# Fonction qui permet de nettoyer le HTML
def strip_html(text: str) -> str:
    if not text:
        return ""
    return re.sub(r"<[^>]+>", " ", text).strip()

# Fonction qui permet de supprimer les accents
def deaccent(text: str) -> str:
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(c for c in normalized if not unicodedata.combining(c))

# Fonction qui va charger le fichier .env manuellement pour ajouter des variables d'environnement
def load_env_file(path: str):
    if not os.path.isfile(path):
        return
    with open(path, "r", encoding="utf-8") as f:

        # Parcourt chaque ligne du fichier
        for raw_line in f:
            # Supprime les espaces et les retours à la ligne
            line = raw_line.strip()
            # Ignore les lignes vides ou les commentaires (commençant par #)
            if not line or line.startswith("#"):
                continue
            # Ignore les lignes qui ne contiennent pas de clé=valeur
            if "=" not in line:
                continue
            # Sépare la clé et la valeur au niveau du premier
            key, value = line.split("=", 1)
            # Nettoie les espaces autour de la clé
            key = key.strip()
            # Nettoie les espaces et supprime les guillemets éventuels
            value = value.strip().strip('"').strip("'")
            # Ajoute la variable dans les variables d'environnements uniquement si elle n'existe déjà pas
            if key and key not in os.environ:
                os.environ[key] = value


# Fonction qui va extraire le délai de retry envoyé par une API Google
def parse_retry_delay_seconds(body_text: str) -> float:
    try:
        # Convertit la chaîne JSON en dictionnaire Python
        payload = json.loads(body_text)
        # Récupère la liste "details" dans payload["error"]
        # Si "error" ou "details" n'existe pas, une liste vide est utilisée par défaut
        details = payload.get("error", {}).get("details", [])
        # Parcourt chaque élément de la liste details
        for item in details:
            # Vérifie si l'élément correspond au type RetryInfo de l'API Google
            # Cela indique qu'un délai avant nouvelle tentative est fourni
            if item.get("@type") == "type.googleapis.com/google.rpc.RetryInfo":
                # Récupère la valeur du délai de retry
                delay = item.get("retryDelay", "")
                # Vérifie que la valeur se termine par "s" (secondes)
                if delay.endswith("s"):
                    # Supprime le "s" et convertit la valeur en float
                    return float(delay[:-1])
    # Si une erreur se produit (JSON invalide, structure inattendue, etc.), il retourne 0.0 pour indiquer qu'aucun délai valide n'a été trouvé
    except Exception:
        return 0.0
    # Si aucun RetryInfo n'est trouvé dans les détails, retourne 0.0
    return 0.0


# Fonction qui va lire un fichier CSV en testant plusieurs encodages (utile pour les fichiers Excel)
def read_csv_with_fallback(path: str, delimiter: str = ";"):
    encodings = ["utf-8-sig", "cp1250", "latin-1"]
    last_error = None
    for enc in encodings:
        try:
            # Ouvre le fichier CSV avec l'encodage courant
            with open(path, "r", encoding=enc, newline="") as f:
                # Crée un lecteur CSV qui retourne chaque ligne sous forme de dictionnaire
                # Les clés du dictionnaire correspondent aux noms de colonnes
                reader = csv.DictReader(f, delimiter=delimiter)
                # Convertit toutes les lignes en liste
                rows = list(reader)
                # Il retourne les lignes, les noms des colonnes, l'encodage utilisé avec succès
                return rows, reader.fieldnames or [], enc
        # Si l'encodage ne fonctionne pas, on capture l'erreur
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    # Si aucun encodage ne fonctionne, on relance la dernière erreur
    raise last_error or UnicodeDecodeError("utf-8", b"", 0, 1, "Unknown decode error")

# Fonction qui va écrire des données dans un CSV
def write_csv(path: str, rows, fieldnames, delimiter: str = ";", encoding: str = "utf-8"):
    # Ouvre le fichier CSV en mode écriture
    with open(path, "w", encoding=encoding, newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
        writer.writeheader()
        # Parcourt chaque ligne et l'écrit dans le fichier
        for row in rows:
            writer.writerow(row)

# Fonction qui va extraire des caractéristiques produit (label, valeur) depuis les colonnes
def extract_specs(row, max_specs: int = 18):
    # Liste qui contiendra les spécifications extraites
    specs = []
    # Parcourt toutes les colonnes de la ligne
    for key, value in row.items():
        if not value:
            continue
        # Vérifie si la colonne correspond à une proriété de filtrage
        if key.startswith("filteringProperty:"):
            # Il extrait le nom du label après "filteringProperty:"
            label = key.split(":", 1)[1].strip()
            # Ajoute la paire (label, valeur)
            specs.append((label, value))
    # Parcourt de nouveau les colonnes pour traiter les textProperty
    for key, value in row.items():
        if not value:
            continue
        # Exemple de colonne : textProperty1 = "Couleur;Rouge"
        if key.startswith("textProperty"):
            # Sépare le label et la valeur
            parts = value.split(";", 1)
            # Vérifie que les deux parties existent
            if len(parts) == 2 and parts[0].strip() and parts[1].strip():
                # Ajoute (label, valeur) après nettoyage
                specs.append((parts[0].strip(), parts[1].strip()))
    # Keep deterministic ordering
    # Limite le nombre de spécifications retournées
    # Cela garantit un ordre déterministe et évite les listes trop longues
    return specs[:max_specs]

# Fonction qui va rechercher une propriété spécifique dans les colonnes textProperty
def get_text_property(row, label_contains: str):
    # Convertit le texte recherché en minuscule pour une comparaison insensible à la casse
    target = label_contains.lower()
    # Parcourt toutes les colonnes de la ligne
    for key, value in row.items():
        if not value:
            continue
        # Recherche dans les colonnes textProperty
        if key.startswith("textProperty"):
            # Exemple : "Couleur;Rouge"
            parts = value.split(";", 1)
            if len(parts) == 2:
                # Récupère le label
                label = parts[0].strip().lower()
                # Vérifie si le label contient le texte recherché
                if target in label:
                    # Retourne la valeur correspondante
                    return parts[1].strip()
    # Si aucune propriété correspondante n'est trouvée
    return ""


def get_text_property_any(row, label_contains_list):
    for label in label_contains_list:
        value = get_text_property(row, label)
        if value:
            return value
    return ""


def render_inline_icon(icon_name: str, color: str) -> str:
    icons = {
        "bolt": (
            '<svg viewBox="0 0 24 24" aria-hidden="true">'
            '<path d="M13 2L5 14h5l-1 8 8-12h-5l1-8z"></path>'
            "</svg>"
        ),
        "tasks": (
            '<svg viewBox="0 0 24 24" aria-hidden="true">'
            '<path d="M9 6h11"></path>'
            '<path d="M9 12h11"></path>'
            '<path d="M9 18h11"></path>'
            '<path d="M4 6h.01"></path>'
            '<path d="M4 12h.01"></path>'
            '<path d="M4 18h.01"></path>'
            "</svg>"
        ),
        "leaf": (
            '<svg viewBox="0 0 24 24" aria-hidden="true">'
            '<path d="M5 21c6-1 12-7 14-14-7 2-13 8-14 14z"></path>'
            '<path d="M9 15c1.5-.5 3.5-2.5 5-5"></path>'
            "</svg>"
        ),
    }
    svg = icons.get(icon_name, icons["leaf"])
    return f'<span class="bullet-glyph" style="color:{color}">{svg}</span>'

# Fonction qui transforme une ligne CSV produit en spécifications structurées prêtes à être affichées dans une fiche produit
def build_spec_fields(row):
    # ---------------------------------------------------
    # 1. Construction du champ PROCESSEUR
    # ---------------------------------------------------
    # On récupère le modèle exact du CPU depuis les propriétés textuelles
    cpu_model = get_text_property(row, "model cpu")
    # On récupère la valeur générique du processeur depuis les propriétés de filtrage du CSV
    processor = row.get("filteringProperty:Procesor") or ""
    # Si aucun modèle précis n'a été trouvé, on tente de l'extraire depuis la description courte du produit (shortDescription)
    if not cpu_model:
        short_desc = strip_html(row.get("shortDescription") or "")
        # Extract CPU model from short description like "Intel® Core™ i5-9400 Processor (9M Cache, up to 4.10 GHz)"
        # Expression régulière pour extraire un modèle de CPU
        m = re.search(r"(Intel®?\s+Core™?\s+[A-Za-z0-9-]+)\s+Processor\s*\(([^)]+)\)", short_desc)
        if m:
            cpu_model = f"{m.group(1)} ({m.group(2)})"
    # Si on a à la fois un modèle détaillé et un nom générique, on tente de les combiner proprement sans duplication
    if cpu_model and processor:
        cpu_lower = cpu_model.lower()
        proc_lower = processor.lower()
        # Si le processeur générique est déjà inclus dans le modèle, on garde simplement le modèle complet
        if proc_lower in cpu_lower:
            spec_processor = cpu_model
        else:
            # Cas particulier pour les processeurs Intel Core afin de reconstruire un nom cohérent
            m = re.match(r"(Intel®?\s+Core™?\s+)(i[3579])", processor, re.IGNORECASE)
            if m:
                prefix = m.group(1)
                series = m.group(2)
                if cpu_model.lower().startswith(series.lower() + "-"):
                    spec_processor = f"{prefix}{cpu_model}"
                else:
                    spec_processor = f"{processor} {cpu_model}"
            else:
                spec_processor = f"{processor} {cpu_model}"
    else:
        # Si on a qu'une seule information, on l'utilise telle quelle
        spec_processor = cpu_model or processor

    # ---------------------------------------------------
    # 2. Construction du champ MÉMOIRE RAM
    # ---------------------------------------------------
    # On récupère la taille et loe type de la RAM
    ram_size = row.get("filteringProperty:Velikost operační paměti RAM") or ""
    ram_type = row.get("filteringProperty:Typ operační paměti") or ""

    # On combine taille + type (ex : "16 GB DDR4")
    spec_ram = " ".join(x for x in [ram_size, ram_type] if x).strip()

    # Certaines fiches produit contiennent des options d’upgrade RAM
    ram_upgrades = row.get("surchargeParameter:Navýšení operační paměti RAM") or ""

    # Si des options existent, on cherche la RAM maximale possible
    if ram_size and ram_upgrades:
        # Find max RAM size in upgrade options, e.g. "na 64 GB"
        matches = re.findall(r"\b(\d+)\s*GB\b", ram_upgrades)
        if matches:
            try:
                current = int(re.findall(r"\d+", ram_size)[0])
                max_ram = max(int(m) for m in matches)

                # Si la RAM maximale est supérieure à celle installée, on génère une indication du type "16 GB až 64 GB"
                if max_ram > current:
                    spec_ram = f"{current} GB až {max_ram} GB"
                    if ram_type:
                        spec_ram = f"{spec_ram} {ram_type}"
            except Exception:
                pass

    # ---------------------------------------------------
    # 3. Construction du champ STOCKAGE
    # ---------------------------------------------------
    storage_size = row.get("filteringProperty:Kapacita úložiště") or ""
    storage_type = row.get("filteringProperty:Typ úložiště") or ""

    # Combinaison simple taille + type (ex : "512 GB SSD")
    spec_storage = " ".join(x for x in [storage_size, storage_type] if x).strip()

    # Certaines fiches possèdent une configuration complète (ex : "256 GB SSD + 1 TB HDD")
    storage_config = get_text_property(row, "konfigurace úložiště")
    if storage_config:
        spec_storage = storage_config
    # Add a short note for NVMe SSDs
    # Ajout d'une petite description si le stockage est NVMe SSD
    if "SSD" in storage_type and "NVMe" in storage_type and "vysoká rychlost" not in spec_storage:
        if storage_type and storage_size:
            spec_storage = f"{storage_size} {storage_type} (vysoká rychlost čtení i zápisu)"
        else:
            spec_storage = f"{storage_type} (vysoká rychlost čtení i zápisu)"

    # ---------------------------------------------------
    # 4. Construction du champ CARTE GRAPHIQUE
    # ---------------------------------------------------
    # On tente plusieurs sources pour trouver le GPU
    gpu = get_text_property(row, "grafická karta") or ""
    if not gpu:
        gpu = row.get("filteringProperty:Grafická karta") or ""
    if not gpu:
        gpu = row.get("filteringProperty:Typ grafické karty") or ""
    # If we only have a generic value like "integrovaná", try to extract a real model from shortDescription.
    # Si la valeur est trop générique ("integrated"), on tente d'extraire le vrai modèle depuis la description
    if not gpu or gpu.strip().lower() in ("integrovaná", "integrovana", "integrated"):
        short_desc = strip_html(row.get("shortDescription") or "")
        # Try to extract GPU from short description like "... Intel® UHD Graphics 630, ..."
        # Extraction possible de GPU Intel, AMD ou NVIDIA
        m = re.search(r"(Intel®?\s+[A-Za-z0-9\s]+Graphics\s*\d*|AMD\s+Radeon\s+[A-Za-z0-9\s]+|NVIDIA\s+GeForce\s+[A-Za-z0-9\s]+)", short_desc, re.IGNORECASE)
        if m:
            gpu = m.group(1).strip()
    spec_gpu = gpu

    # ---------------------------------------------------
    # 5. Dimensions et système d’exploitation
    # ---------------------------------------------------
    spec_dimensions = get_text_property(row, "rozměry") or get_text_property(row, "rozm") or ""
    case_type = row.get("filteringProperty:Typ skříně") or ""

    # Ajoute le type de boîtier aux dimensions si ce n'est pas déjà présent
    if spec_dimensions and case_type and case_type not in spec_dimensions:
        case_label = case_type
        m = re.match(r"([A-Z]+)\s*\(([^)]+)\)", case_type)
        if m:
            case_label = m.group(2).strip()
        spec_dimensions = f"{spec_dimensions} ({case_label})"

    # Récupération du système d'exploitation
    spec_os = get_text_property(row, "operační systém") or get_text_property(row, "opera") or ""

    # Feature labels for hero icons
    # ---------------------------------------------------
    # 6. Création des labels pour les icônes "features"
    # ---------------------------------------------------
    feature_cpu = row.get("filteringProperty:Procesor") or spec_processor

    # Normalisation du label RAM
    feature_ram = ram_type or "RAM"
    if "DDR" in feature_ram:
        feature_ram = f"{feature_ram} RAM"
    # Normalisation du label stockage
    feature_storage = storage_type or "Úložiště"
    if "NVMe" in feature_storage:
        feature_storage = "NVMe SSD"
    elif "SSD" in feature_storage:
        feature_storage = "SSD"
    elif "HDD" in feature_storage:
        feature_storage = "HDD"
    # Construction du label design du boitier
    feature_case = case_type
    if feature_case:
        m_case = re.match(r"([A-Z]+)\s*\(([^)]+)\)", feature_case)
        feature_case = m_case.group(1) + " Design" if m_case else f"{feature_case} Design"
    else:
        feature_case = "Kompaktní"

    # ---------------------------------------------------
    # 7. Retour des champs de spécifications structurées
    # ---------------------------------------------------
    return {
        "spec_processor": spec_processor or "-",
        "spec_ram": spec_ram or "-",
        "spec_storage": spec_storage or "-",
        "spec_gpu": spec_gpu or "-",
        "spec_dimensions": spec_dimensions or "-",
        "spec_os": spec_os or "-",
        "feature_cpu": feature_cpu,
        "feature_ram": feature_ram,
        "feature_storage": feature_storage,
        "feature_case": feature_case,
    }


def normalize_boolish(value: str) -> bool:
    if value is None:
        return False
    val = str(value).strip().lower()
    if not val:
        return False
    if val in ("ne", "no", "false", "0", "bez", "n/a", "na"):
        return False
    return True


# Fonction qui génère une liste de badges indiquant les ports disponibles sur un ordinateur
def build_port_badges(row, max_items: int = 6):
    # ---------------------------------------------------
    # Fonction interne : formater l'affichage d'un port
    # ---------------------------------------------------
    # Cette fonction vérifie si un port est présent et, si possible, extrait le nombre de ports pour afficher un badge du type  "USB 3.2 × 4"
    def count_label(label, raw_value):
        # Vérifie si la valeur correspond à quelque chose de "vrai" via la fonction normalize_boolish
        if not normalize_boolish(raw_value):
            return ""
        # Il va rechercher un nombre dans la valeur brute
        m = re.search(r"\d+", str(raw_value))
        # Si un nombre est trouvé, on indique la quantité
        if m:
            return f"{label} × {m.group(0)}"
        # Sinon on retrourne le label du port
        return label

    # ---------------------------------------------------
    # Liste des ports possibles à détecter dans les données
    # ---------------------------------------------------
    # Chaque élément associe :
    # - le nom affiché du port
    # - la colonne correspondante dans le CSV
    port_map = [
        ("USB 3.2", row.get("filteringProperty:USB-3.2")),
        ("USB 3.1", row.get("filteringProperty:USB-3.1")),
        ("USB 3.0", row.get("filteringProperty:USB-3.0")),
        ("USB 2.0", row.get("filteringProperty:USB-2.0")),
        ("USB-C", row.get("filteringProperty:USB Type-C™")),
        ("HDMI", row.get("filteringProperty:HDMI")),
        ("DisplayPort", row.get("filteringProperty:DisplayPort")),
        ("Mini DP", row.get("filteringProperty:Mini DisplayPort")),
        ("VGA", row.get("filteringProperty:VGA")),
        ("DVI", row.get("filteringProperty:DVI")),
        ("LAN", row.get("filteringProperty:LAN")),
        ("Audio", row.get("filteringProperty:Audio JACK")),
        ("Thunderbolt", row.get("filteringProperty:Thunderbolt™")),
    ]

    # ---------------------------------------------------
    # Construction de la liste finale des badges
    # ---------------------------------------------------
    # Cette liste contiendra les ports détectés sous forme de texte prêt à être affiché
    badges = []
    for label, raw in port_map:
        # Formate le texte du badge
        text = count_label(label, raw)
        # Si un port valide est détecté, on l'ajoute à la liste
        if text:
            badges.append(text)
        # On limite le nombre total de badges affichées (par défaut : 6) pour éviter une interface trop chargée
        if len(badges) >= max_items:
            break
    # Retourne la liste finale des badges
    return badges


# Fonction qui génère les champs de la section "features" d'une fiche produit à partir des spécifications extraites. Elle retourne un dictionnaire prêt à être injecté dans un template HTML
def build_feature_section_fields(row, spec_fields):
    spec_cpu = spec_fields.get("spec_processor", "")
    spec_ram = spec_fields.get("spec_ram", "")
    spec_storage = spec_fields.get("spec_storage", "")
    spec_gpu = spec_fields.get("spec_gpu", "")
    form_factor = guess_form_factor(row)

    # Feature 1: Performance
    # On construit une description de la performance qui met en avant les points forts du CPU, de la RAM et du stockage. Si aucune information n'est disponible, on génère une description générique
    perf_bits = []
    if spec_cpu and spec_cpu != "-":
        perf_bits.append(f"Základem výkonu je {spec_cpu}.")
    if spec_ram and spec_ram != "-":
        perf_bits.append(f"Operační paměť {spec_ram} pomáhá udržet plynulý multitasking.")
    if spec_storage and spec_storage != "-":
        if "SSD" in spec_storage:
            perf_bits.append("SSD zrychluje start systému i načítání aplikací.")
        else:
            perf_bits.append(f"Úložiště {spec_storage} nabízí dostatek místa na data.")
    if not perf_bits:
        perf_bits.append("Spolehlivý výkon pro každodenní práci, web a kancelářské aplikace.")
    feature1_text = " ".join(perf_bits)

    # On construit une liste de "bullets" pour les points forts à mettre en avant avec des icônes. Ces points sont choisis en fonction des spécifications disponibles : la présence d'un SSD, la quantité de RAM, etc. Chaque bullet contient une icône, un label et une couleur associée.
    bullets = []
    if spec_storage and "SSD" in spec_storage:
        bullets.append(("bolt", "Rychlý start systému", "#eab308"))
    if spec_ram and spec_ram != "-":
        bullets.append(("tasks", "Plynulý multitasking", "#3b82f6"))
    bullets.append(("leaf", "Svižná odezva aplikací", "#22c55e"))
    bullets = bullets[:3]

    feature1_bullets_html = "".join(
        f'<li class="bullet-item">{render_inline_icon(icon, color)} {label}</li>'
        for icon, label, color in bullets
    )
    
    feature1_image = get_text_property_any(
        row,
        ["feature image 1", "obrázek 1", "obrazek 1", "performance image", "výkon"],
    ) or "https://img.notebooksbilliger.de/images/products/340000/348141/Lenovo_ThinkCentre_M710e_SFF_2.jpg"

    # Feature 2: Ports
    # On génère une description de la connectique qui met en avant les types de ports disponibles. Si aucune information n'est disponible, on génère une description générique
    ports = build_port_badges(row)
    has_video = any(p in ("HDMI", "DisplayPort", "Mini DP", "VGA", "DVI") for p in ports)
    has_usb = any(p.startswith("USB") or p == "USB-C" for p in ports)
    has_lan = "LAN" in ports
    parts = ["Konektorová výbava se liší dle konkrétní konfigurace."]
    # Si des ports USB sont présents, on met en avant la facilité de connexion de périphériques externes (clavier, souris, imprimante, etc.)
    if has_usb:
        parts.append("USB porty umožní snadné připojení periferií.")
    # Si des ports vidéo sont présents, on met en avant la possibilité de connecter un ou plusieurs moniteurs externes
    if has_video:
        parts.append("Video výstupy slouží k připojení monitoru.")
    # Si un port LAN est présent, on met en avant la possibilité d'une connexion internet filaire stable et rapide, idéale pour le télétravail ou les visioconférences
    if has_lan:
        parts.append("LAN zajistí stabilní kabelové připojení.")
    feature2_text = " ".join(parts)

    # On génère une liste de badges HTML pour les ports détectés, qui seront affichés dans la section "features" de la fiche produit. Chaque badge est stylisé avec une classe CSS et contient le nom du port
    feature2_ports_html = "".join(
        f'<span class="port-badge">{html.escape(p)}</span>'
        for p in ports
    )
    # On tente de récupérer une image spécifique pour illustrer la section des ports. Si aucune image n'est fournie dans les données, on utilise une image générique par défaut
    feature2_image = get_text_property_any(
        row,
        ["feature image 2", "obrázek 2", "obrazek 2", "ports image", "porty"],
    ) or "https://img.notebooksbilliger.de/images/products/340000/348141/Lenovo_ThinkCentre_M710e_SFF_3.jpg"

    # Feature 3: Internal / Design
    # On génère une description du design et de l'intérieur du boîtier qui met en avant les avantages d'un format compact (mini, SFF) ou d'un design pratique pour l'entretien. Si aucune information n'est disponible, on génère une description générique
    if form_factor == "mini":
        internal_text = "Mini provedení je ideální tam, kde je málo místa a chcete čistý pracovní stůl."
    elif form_factor == "sff":
        internal_text = "SFF skříň nabízí dobrý kompromis mezi kompaktními rozměry a možnostmi rozšíření."
    else:
        internal_text = "Praktické provedení usnadňuje každodenní používání i údržbu."

    if spec_gpu and spec_gpu != "-" and is_dedicated_gpu(spec_gpu):
        internal_text += " Dedikovaná grafika přidává výkon pro náročnější vizuální úkoly."

    feature3_image = get_text_property_any(
        row,
        ["feature image 3", "obrázek 3", "obrazek 3", "internal image", "vnitřek"],
    ) or "https://img.notebooksbilliger.de/images/products/340000/348141/Lenovo_ThinkCentre_M710e_SFF_4.jpg"

    # Retourne un dictionnaire contenant tous les champs nécessaires pour remplir la section "features" d'une fiche produit, avec des titres, des textes descriptifs, des listes de points forts et des images associées à chaque feature
    return {
        "feature1_title": "Výkon, který Vás nezpomalí",
        "feature1_text": feature1_text,
        "feature1_bullets_html": feature1_bullets_html,
        "feature1_image": feature1_image,
        "feature1_image_alt": "Výkon",
        "feature2_title": "Vše, co potřebujete k připojení",
        "feature2_text": feature2_text,
        "feature2_ports_html": feature2_ports_html,
        "feature2_image": feature2_image,
        "feature2_image_alt": "Porty",
        "feature3_title": "Chytrý design a praktické provedení",
        "feature3_text": internal_text,
        "feature3_quote_html": "Praktické provedení pro každodenní práci.",
        "feature3_image": feature3_image,
        "feature3_image_alt": "Vnitřek",
    }


# Fonction qui génère les champs de la section "suitability" d'une fiche produit, qui liste les cas d'usage recommandés et déconseillés pour un ordinateur donné en fonction de ses spécifications
# Elle retourne un dictionnaire prêt à être injecté dans un template HTML
# Cette fonction transforme des données techniques en recommandations d'utilisation concrètes et compréhensibles pour les clients, en mettant en avant les points forts du produit tout en étant transparent sur ses limites
def build_suitability_fields(row, spec_fields):
    # On identifie la catégorie du produit et son format pour adapter les recommandations d'usage en fonction du type de machine
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    spec_gpu = spec_fields.get("spec_gpu", "")
    spec_ram = spec_fields.get("spec_ram", "")
    spec_storage = spec_fields.get("spec_storage", "")

    # En fonction de la catégorie et des spécifications, on construit deux listes : une pour les usages recommandés (good) et une pour les usages déconseillés (bad). Ces listes sont ensuite transformées en HTML avec des icônes pour être affichées dans la fiche produit
    good = []
    bad = []

    # On adapte les recommandations d'usage en fonction de la catégorie du produit (desktop, laptop, monitor) et de son format (mini, sff, etc.) pour mettre en avant les cas d'utilisation les plus pertinents
    if category in ("desktop", "computer"):
        good.extend(
            [
                "Kancelářské aplikace, e-maily a web",
                "Domácí práce a studium",
                "Online schůzky a videohovory",
            ]
        )
        if form_factor == "mini":
            good.append("Prostorově úsporné pracovní místo")
    elif category == "laptop":
        good.extend(
            [
                "Škola, domácí úkoly a e-maily",
                "Práce na cestách i doma",
                "Online výuka a videohovory",
            ]
        )
    elif category == "monitor":
        good.extend(
            [
                "Kancelář a domácí pracoviště",
                "Druhá obrazovka k notebooku",
                "Běžné multimediální použití",
            ]
        )
    
    # En fonction des spécifications techniques, on ajoute des recommandations d'usage spécifiques qui mettent en avant les points forts du produit 
    if spec_storage and "SSD" in spec_storage:
        good.append("Rychlé spuštění systému a aplikací")
    if spec_ram and spec_ram != "-":
        good.append("Plynulý multitasking v běžných aplikacích")

    # Si une carte graphique dédiée est présente, on recommande des usages plus exigeants en ressources graphiques, tout en restant réaliste sur les limites d'une machine d'entrée de gamme ou milieu de gamme
    # Si aucune carte dédiée n'est présente, on met en avant les limites pour les usages graphiques intensifs
    if spec_gpu and is_dedicated_gpu(spec_gpu):
        good.append("Graficky náročnější aplikace a lehčí hraní")
        bad.append("Profesionální 3D rendering a špičkové herní sestavy")
    else:
        # Si la machine ne possède pas de GPU dédié, on déconseille les usages graphiques intensifs
        bad.extend(
            [
                "Nejnovější 3D hry na vysoké detaily",
                "Profesionální střih 4K videa",
                "Náročné CAD/3D modelování",
            ]
        )
    
    # Si la machine possède une quantité de RAM limitée (ex : 4 GB ou moins), on déconseille les usages multitasking intensifs et les applications gourmandes en mémoire
    if category == "monitor":
        bad.append("Práce bez samostatného počítače")

    # Fonction interne pour formater une recommandation d'usage avec une icône "check" pour les usages recommandés et une icône "times" pour les usages déconseillés, en échappant le texte pour éviter les problèmes de sécurité
    def li_check(text):
        return f'<li><span class="icon-good">✔</span> {html.escape(text)}</li>'

    def li_times(text):
        return f'<li><span class="icon-bad">✖</span> {html.escape(text)}</li>'
    
    # On limite le nombre de recommandations affichées pour éviter une section trop longue, en gardant les plus pertinentes
    good = good[:4] if len(good) > 4 else good
    bad = bad[:3] if len(bad) > 3 else bad

    # On génère le HTML final pour les listes d'usages recommandés et déconseillés, qui seront affichées dans la section "suitability" de la fiche produit, avec des icônes visuelles pour renforcer la compréhension et l'impact des recommandations
    return {
        "suitability_good_html": "".join(li_check(t) for t in good),
        "suitability_bad_html": "".join(li_times(t) for t in bad),
    }

# Function for generate the FAQ SECTION
# Fonction qui génère les champs de la section "FAQ" d'une fiche produit, avec des questions et réponses basées sur les spécifications du produit et les questions fréquemment posées par les clients
def build_faq_fields(row, spec_fields):
    # On identifie la catégorie du produit pour adapter les questions et réponses en fonction du type de machine
    category = guess_category(row)
    # On détermine les labels à utiliser pour "ordinateur" en fonction de la catégorie (notebook pour les laptops, monitor pour les moniteurs, etc.) pour rendre les réponses plus naturelles et adaptées au type de produit
    is_laptop = category == "laptop"
    is_monitor = category == "monitor"
    device_label = "notebook" if is_laptop else "počítač"
    device_label_plural = "notebooky" if is_laptop else "počítače"
    # Cas particulier pour les moniteurs
    if is_monitor:
        device_label = "monitor"
        device_label_plural = "monitory"
    # On récupère le système d'exploitation pour adapter les réponses liées à Windows en fonction de la présence ou non d'un OS installé sur la machine, et éviter de mentionner des licences Windows pour les moniteurs qui n'en ont pas besoin
    spec_os = (spec_fields.get("spec_os") or "").lower()

    # Fonction interne pour formater une question et sa réponse dans un bloc HTML structuré, en échappant le texte de la question pour éviter les problèmes de sécurité, et en permettant d'inclure du HTML dans la réponse pour mettre en forme les informations importantes
    def details_block(question, answer):
        return (
            '<div class="faq-item">'
            f'<div class="faq-summary">{html.escape(question)}</div>'
            f'<div class="faq-answer">{answer}</div>'
            "</div>"
        )
    
    # Liste qui contiendra les blocs de questions/réponses à afficher dans la section FAQ de la fiche produit
    faq_items = []

    #-----------------------------------------------------------------------------
    # Question 1 : Garantie
    #-----------------------------------------------------------------------------
    faq_items.append(
        details_block(
            "Jaká je záruka na repasované počítače a notebooky?",
            "Na <strong>veškerou repasovanou výpočetní techniku</strong> poskytujeme záruku na skryté vady v délce "
            "<strong>24 měsíců</strong>. "
            "Pokud budete mít zájem, můžete si záruku prodloužit až na 5 let.",
        )
    )

    #-----------------------------------------------------------------------------
    # Question 2 : Baterie (uniquement pour les laptops)
    #-----------------------------------------------------------------------------
    if is_laptop:
        faq_items.append(
            details_block(
                "Jak dlouho vydrží baterie, vztahuje se záruka i na ni?",
                "Všechny baterie testujeme</strong> tak, aby notebook při prodeji vydržel pracovat alespoň 60 minut z baterie</strong>. "
                "Pokud baterie tento parametr nesplňuje, automaticky ji měníme za novou. "
                "<strong>U nových baterií poskytujeme záruku 24 měsíců na elektroniku baterie a "
                "12 měsíců na kapacitu článků</strong>."
            )
        )
        # Contenu de la question 3 pour les laptops : ce qui est inclus dans le pack
        pack_answer = (
            "Pokud objednáte <strong>repasovaný notebook</strong>, je součástí balení <strong>notebook</strong>, "
            "<strong>napájecí adaptér</strong> a přívodní kabel pro zapojení nabíječky do zásuvky."
        )
    else:
        # Contenu de la question 3 pour les desktops et les moniteurs : ce qui est inclus dans le pack
        pack_answer = (
            "Pokud objednáte <strong>repasovaný notebook</strong>, je součástí balení <strong>notebook</strong>, <strong>napájecí adaptér</strong> a přívodní kabel pro zapojení nabíječky do zásuvky."
            "<br><br>"
            "Pokud objednáte <strong>repasovaný počítač</strong>, je součástí balení <strong>počítač</strong> a <strong>napájecí kabel</strong>, případně napájecí adaptér a přívodní kabel pro zapojení nabíječky do zásuvky."
            "<br><br>"
            "Pokud objednáte <strong>repasovaný monitor</strong>, je součástí balení <strong>monitor</strong> vč. stojanu, <strong>napájecí kabel</strong> a <strong>video kabel</strong> k propojení monitoru s počítačem."
        )

    #-----------------------------------------------------------------------------
    # Question 3 : Contenu du pack
    #-----------------------------------------------------------------------------
    faq_items.append(details_block("Co je součástí balení?", pack_answer))

    #-----------------------------------------------------------------------------
    # Question 4 : Licence Windows
    #-----------------------------------------------------------------------------
    faq_items.append(
        details_block(
            "Kde najdu licenční klíč k Windows? Co je to elektronická licence?",
            (
                "U monitorů se licenční klíč Windows netýká, protože neobsahují operační systém. "
                "Monitor je samostatné zobrazovací zařízení, které se připojuje k počítači nebo notebooku."
                if is_monitor
                else
                "<strong>Licenční klíč</strong> k operačnímu systému Windows se obvykle nachází na nálepce, která je umístěna na počítači nebo notebooku. "
                "Nálepka se může nacházet na <strong>spodní straně zařízení nebo pod baterií</strong>, pokud je baterie odnímatelná."
                "<br><br>"
                "<strong>Elektronická licence je alternativou ke klasickému licenčnímu klíči</strong>, který se dnes již téměř nepoužívá. "
                "Elektronická licence je přidělena na základě hardwarových informací vašeho počítače a <strong>je uložena na serverech Microsoftu</strong>. "
                "Po případné reinstalaci zařízení a opětovném připojení počítače k internetu tak dojde k automatické aktivaci Windows pomocí digitálního klíče. "
                "<strong>Licence není přenositelná na jiné zařízení</strong>."
                + (
                    "<br><br>Všechny naše počítače a notebooky prodáváme s předinstalovaným operačním systémem Windows 10 nebo Windows 11. "
                    "Elektronická licence je tak již nainstalována a zařízení je aktivováno a připraveno k použití."
                    if "windows" in spec_os
                    else ""
                )
            ),
        )
    )

    #-----------------------------------------------------------------------------
    # Question 5 : Upgrade de la RAM ou du stockage
    #-----------------------------------------------------------------------------
    upgrade_answer = (
        # Pour les moniteurs, la question de l'upgrade de la RAM ou du stockage n'est pas pertinente, donc on indique que cela ne s'applique pas
        "U monitorů se operační paměť ani disk nevyměňuje, takže se tento bod netýká."
        if is_monitor
        else
        # Pour les desktops et les laptops, on indique que l'upgrade de la RAM ou du stockage est possible sans perdre la garantie, mais que la garantie ne couvre que les composants achetés chez nous
        f"Ano, můžete. <strong>Otevřením počítače a výměnou komponent nepřicházíte o záruku</strong>. I nadále však platí, že se záruka vztahuje jen na hardware zakoupený v našem obchodě. Pokud si například v notebooku vyměníte pevný disk za nový, který vám po chvíli přestane fungovat, nebudeme vám moct tento nový rozbitý disk vyreklamovat."
    )
    faq_items.append(
        details_block(
            f"Můžu v počítači nebo notebooku vyměnit operační paměť nebo přidat disk? Nepřijdu o záruku?",
            upgrade_answer
            + ("" if is_monitor else " I nadále však platí, že se záruka vztahuje jen na hardware zakoupený v našem obchodě."),
        )
    )

    # -----------------------------------------------------------------------------
    # Question 6 : définition de ce qu'est un ordinateur ou un notebook reconditionné, et d'où proviennent les machines destinées à être reconditionnées, pour rassurer les clients sur la qualité des produits et l'origine des machines
    #-----------------------------------------------------------------------------
    faq_items.append(
        details_block(
            "Co je to repas?",
            "Repas neboli <strong>repasovaná výpočetní technika je již dříve používané zařízení</strong>. V rámci repase jsou počítače a notebooky "
            "kompletně rozebrány za účelem důkladného čištění a testování všech komponentů. Pokud se při testech hardwaru ukážou některé "
            "komponenty jako nefunkční, jsou nahrazeny funkčními nebo novými tak, aby bylo možné zařízení i nadále používat a my na něj mohli "
            "poskytnout záruku. <strong>Repasovaný notebook nebo počítač je tak vhodnou volbou</strong> pro ty, <strong>kteří hledají levnější alternativu k novým "
            "počítačům, ale stále chtějí mít výkonné a spolehlivé zařízení</strong>.",
        )
    )

    # -----------------------------------------------------------------------------
    # Question 7 : d'où proviennent les machines destinées à être reconditionnées
    #-----------------------------------------------------------------------------
    faq_items.append(
        details_block(
            "Kde nakupujete počítače a notebooky určené k repasi?",
            "Původ těchto zařízení je ve <strong>velkých korporátních firmách</strong>, které tyto počítače prodávají za účelem obnovy firemního vybavení, "
            "které je vyměňováno za novější technologie. Počítače a notebooky určené k repasi vykupujeme z českých nebo zahraničních firem napřímo. "
            "Firemní počítače a notebooky jsou vyrobeny tak, aby měly co nejdelší životnost a co největší výkon, proto <strong>jsou vhodnou alternativou pro ty, "
            "kteří hledají levnější variantu k novým počítačům, ale stále chtějí mít výkonné a spolehlivé zařízení</strong>.",
        )
    )

    # Retourne un dictionnaire contenant le champ "faq_html" qui contient le HTML complet de la section FAQ, avec toutes les questions et réponses formatées et prêtes à être injectées dans la fiche produit
    return {"faq_html": "".join(faq_items)}


# Fonction qui va essayer de deviner la catégorie d'un produit informatique
def guess_category(row):
    # Récupère la valeur de la colonne "name", puis convertit le texte en minuscules
    name = (row.get("name") or "").lower()
    # Supprime les accents du nom pour faciliter la comparaison
    name_ascii = deaccent(name)
    # Vérifie si le nom contient des mots-clés indiquant un ordinateur de bureau
    if "thinkcentre" in name_ascii or "desktop" in name_ascii or "pc" in name_ascii:
        return "desktop"

    # Inspect spec keys for hints without hardcoding non-ascii labels
    # Analyse les colonnes de spécifications pour deviner la catégorie sans dépendre des accents ou des langues spécifiques
    has_display = False # Indique si un écran est mentionné
    has_cpu = False # Indique si un processeur est mentionné
    # Parcourt toutes les colonnes de la ligne
    for key, value in row.items():
        # Ignore les valeurs vides
        if not value:
            continue
        # Supprime les accents et convertit en minuscules pour rendre la détection plus robuste
        key_ascii = deaccent(key).lower()
        # Vérifie si la clé indique un écran
        # "displej" est l'équivalent en tchèque/slovaque
        if "displej" in key_ascii or "display" in key_ascii:
            has_display = True
        # Vérifie si la clé indique un processeur
        if "procesor" in key_ascii or "cpu" in key_ascii:
            has_cpu = True
    # Si on trouve un écran ET un processeur → probablement un laptop
    if has_display and has_cpu:
        return "laptop"
    # Si on trouve seulement un écran → probablement un moniteur
    if has_display and not has_cpu:
        return "monitor"
    # Si aucun indice clair n'est trouvé → catégorie générique
    return "computer"

# Fonction qui va essayer de deviner le format d'un ordinateur (mini, sff, etc.) en se basant sur le nom et les spécifications, pour adapter les descriptions et recommandations d'usage en fonction du type de machine
def guess_form_factor(row):
    # On récupère le nom du produit et la valeur de la colonne "Typ skříně" (type de boîtier), puis on supprime les accents et on convertit en minuscules pour faciliter la détection de mots-clés indiquant un format compact
    name = deaccent((row.get("name") or "")).lower()
    case_type = deaccent((row.get("filteringProperty:Typ skříně") or "")).lower()
    # On combine le nom et le type de boîtier pour avoir plus de chances de détecter des indices sur le format de la machine
    hay = " ".join([name, case_type])
    # On vérifie la présence de mots-clés indiquant un format mini ou très compact, puis de mots-clés indiquant un format SFF (Small Form Factor), en respectant l'ordre de priorité pour éviter les confusions
    if any(token in hay for token in ["mini", "tiny", "micro", "mff", "usff"]):
        return "mini"
    # Si aucun format mini n'est détecté, on vérifie la présence de mots-clés indiquant un format SFF, qui est plus grand que le mini mais toujours compact
    if any(token in hay for token in ["sff"]):
        return "sff"
    # Si aucun indice de format compact n'est trouvé, on retourne une chaîne vide pour indiquer un format standard
    return ""


# Fonction qui extrait le nom court du modèle à partir du nom complet, en retirant la partie configuration (ex: " - 8 Go - 500 Go...") pour éviter de répéter les spécifications dans le corps du texte et respecter la règle de ne pas utiliser le nom complet de la variante dans les descriptions
def _extract_short_model_name(row) -> str:
    """
    Extrait le nom court du modèle (ex: "Lenovo ThinkCentre M710e")
    sans la partie configuration (" - 8 Go - 500 Go...").
    Utilisé dans les prompts pour respecter la règle : ne pas répéter
    le nom complet de la variante dans le corps du texte.
    """
    # Récupère le nom complet du produit à partir de la colonne "name", ou une chaîne vide si la colonne est absente ou vide
    full_name = row.get("name") or ""
    # On utilise un regex pour couper au premier tiret suivi d'un chiffre ou "Go" pour retirer la config
    m = re.match(r"^(.*?)\s*[-–]\s*\d", full_name)
    if m:
        return m.group(1).strip()
    # Si aucun tiret suivi d'une config n'est trouvé, on retourne le nom complet tel quel, en supprimant les espaces superflus
    return full_name.strip()


# ============================================================
# PROMPTS ADAPTÉS SELON LA DOCUMENTATION DU PROJET
# ============================================================

# Fonction qui construit un prompt destiné à un modèle d'IA afin de générer une description produit HTML complète pour une fiche e-commerce
def build_prompt(row, specs, language: str) -> str:
    """
    Prompt principal : génère la description longue structurée en 7 sections
    telles que définies dans la documentation du projet.
    """
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    manufacturer = row.get("manufacturer") or ""
    condition = row.get("filteringProperty:Stav") or ""
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    short_desc = strip_html(row.get("shortDescription") or "")

    specs_lines = "\n".join(f"- {label}: {value}" for label, value in specs)

    prompt = f"""You are a professional copywriter for an online store selling refurbished computers.
Write a long product description in {language} following EXACTLY the 7-section structure below.

PRODUCT DATA:
- Full product name (for <h3> title only): {name}
- Short model name (use in body text): {short_model}
- Manufacturer: {manufacturer}
- Condition: {condition}
- Category: {category}
- Form factor: {form_factor}
- Short description: {short_desc}
Key specs:
{specs_lines}

SOURCE RULE:
- Use ONLY the data provided above from the CSV row.
- Do NOT invent any specification, feature, usage, compatibility, accessory, or benefit not supported by the data.
- If a piece of information is missing in the data, omit it.

OUTPUT FORMAT:
Return ONLY an HTML fragment. Use <h3>, <p>, <strong>. No <ul>, no <li>, no bullet points, no semicolons (;).
Write in coherent sentences and short paragraphs. Professional, calm, understandable tone for non-technical users.
Do NOT use marketing filler phrases. Do NOT explain what "refurbished" means.

MANDATORY 7-SECTION STRUCTURE:

<h3>{name}</h3>

Section 1 - Introduction (1-2 sentences):
What is this product and what is its main benefit for the user?

Section 2 - {manufacturer} as a brand (2-4 sentences):
General presentation of the brand. Focus on reliability of professional PCs.
Do NOT mention the full variant name here.

Section 3 - Product line / series (2-4 sentences):
What does this product line mean in practice: stability, office use, quiet operation, compact size.
Use only "{short_model}" or "this configuration", never the full variant name with specs.
Do NOT write about what is "in the name" or "part of the name".

Section 4 - Model and configuration (2-4 sentences):
Link the specific hardware to real-world usage. Mention processor name exactly as given.
Mention RAM size in GB. Mention storage size in GB and storage type exactly as provided by the data.
If the storage type explicitly says NVMe, you may mention the fast read/write benefit. Do NOT use semicolons.

Section 5 - Who it suits (1-2 sentences):
Start with "Convient pour...". Name specific use cases: Office, web, email, accounting, studies, video.
This must be a use case description, not a repetition of specs.

Section 6 - Who it does NOT suit (1-2 sentences):
Start with "Ne convient pas pour..." or "N'est pas destiné à...".
Do NOT use "not recommended for". Be factual, not judgmental.

Section 7 - Ready to use (1 sentence):
Mention: tested, Windows 11 Pro, 24-month warranty.

STRICT RULES:
- Outside the <h3> title, NEVER use the full variant name "{name}" (with config like "- 8 Go - 500 Go").
  In the body text, use only "{short_model}", "ThinkCentre M710e" when appropriate, or "cette configuration".
- No bullet points, no numbered lists, no semicolons.
- No phrases like "as shown above", "listed parameters", "series name".
- Do NOT mention price, stock, delivery.
- Do NOT mention that this text was generated by AI.
- Each section must bring NEW information, not repeat specs as a list.
- In suitability paragraphs, use the expressions "Convient pour..." and "Ne convient pas pour...".
"""
    return prompt.strip()


# Fonction qui va créer un prompt plus marketing et plus long, avec une mise en page enrichie et des images
def build_marketing_prompt(row, specs, language: str) -> str:
    """
    Prompt pour la description marketing longue.
    Suit la même structure en 7 sections mais avec un ton légèrement
    plus engageant, tout en restant factuel et sans puces.
    """
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    manufacturer = row.get("manufacturer") or ""
    condition = row.get("filteringProperty:Stav") or ""
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    short_desc = strip_html(row.get("shortDescription") or "")

    specs_lines = "\n".join(f"- {label}: {value}" for label, value in specs)

    prompt = f"""You are a professional e-commerce copywriter for a store selling refurbished computers.
Write a long product description in {language} following EXACTLY the 7-section structure below.

PRODUCT DATA:
- Full product name (for <h3> title only): {name}
- Short model name (use in body text): {short_model}
- Manufacturer: {manufacturer}
- Condition: {condition}
- Category: {category}
- Form factor: {form_factor}
- Short description: {short_desc}
Key specs:
{specs_lines}

SOURCE RULE:
- Use ONLY the data provided above from the CSV row.
- Do NOT invent any specification, feature, usage, compatibility, accessory, or benefit not supported by the data.
- If a piece of information is missing in the data, omit it.

OUTPUT FORMAT:
Return ONLY an HTML fragment. Use <h3>, <p>, <strong>. No <ul>, no <li>, no bullet points, no semicolons (;).
Write in coherent sentences and short paragraphs. Professional, calm, understandable tone for non-technical users.
Be commercial but factual. No marketing filler.
Do NOT explain what "refurbished" means.

MANDATORY 7-SECTION STRUCTURE:

<h3>{name}</h3>

Section 1 - Introduction (1-2 sentences):
What is this product and what concrete advantage does it bring to the user?

Section 2 - {manufacturer} brand (2-4 sentences):
Present the brand. Emphasize reliability and longevity of professional-grade hardware.
Do NOT use the full variant name here.

Section 3 - Product line (2-4 sentences):
Explain what this product line offers in everyday use: stability, quiet operation, compact design, office focus.
Use only "{short_model}" or "this configuration", never the full name with specs.
Do NOT write about what is "in the name" or "part of the name".

Section 4 - Configuration details (2-4 sentences):
Connect the hardware to real usage scenarios. Name the exact processor. State RAM in GB.
State storage in GB and storage type exactly as provided by the data. If the storage type explicitly says NVMe, mention the speed benefit.
No semicolons. No lists.

Section 5 - Ideal users (1-2 sentences):
Begin with "Convient pour..." List specific use cases: Office apps, web browsing, email,
accounting software, studies, video calls.

Section 6 - Not suitable for (1-2 sentences):
Begin with "Ne convient pas pour..." or "N'est pas destiné à..."
Never use "not recommended for". State clearly and factually.

Section 7 - Ready to use (1 sentence):
Mention that the device is tested, comes with Windows 11 Pro, and includes a 24-month warranty.

STRICT RULES:
- Outside the <h3> title, NEVER write the full variant name "{name}" (the one containing config like "- 8 Go").
  Always use "{short_model}" or "cette configuration".
- No bullet points or numbered lists anywhere.
- No semicolons (;) — always separate items with a period or comma.
- No empty phrases like "base uniforme", "équipement sélectionné", "prêt à l'emploi".
  Replace them with a concrete benefit.
- Do NOT mention price, stock, or delivery.
- Do NOT mention that this text was generated by AI.
- Do NOT write about what is "in the name" or "part of the name".
- In suitability paragraphs, use the expressions "Convient pour..." and "Ne convient pas pour...".
"""
    return prompt.strip()


# Fonction qui va créer un prompt pour générer une description très courte (teaser)
def build_short_prompt(row, specs, language: str) -> str:
    """
    Prompt pour le teaser court (1-2 phrases, 30-40 mots max).
    Texte brut sans HTML.
    """
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    manufacturer = row.get("manufacturer") or ""
    condition = row.get("filteringProperty:Stav") or ""
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    short_desc = strip_html(row.get("shortDescription") or "")

    specs_lines = "\n".join(f"- {label}: {value}" for label, value in specs)

    prompt = f"""You are a copywriter for a refurbished computer store.
Write a short product teaser in {language}.

PRODUCT DATA:
- Product name: {name}
- Short model name: {short_model}
- Manufacturer: {manufacturer}
- Condition: {condition}
- Category: {category}
- Form factor: {form_factor}
- Short description: {short_desc}
Key specs:
{specs_lines}

OUTPUT FORMAT:
Return ONLY plain text (no HTML, no bullet points).
Length: 1-2 sentences, maximum 30-40 words.

CONTENT RULES:
- Use ONLY the data provided above from the CSV row.
- Do NOT invent any specification, feature, usage, compatibility, accessory, or benefit not supported by the data.
- Summarize the main use case and 1-2 key benefits.
- Do NOT list full specs. Mention at most one component if it adds value.
- Use only "{short_model}" in the text, never the full variant name with configuration.
- Do NOT use semicolons (;).
- Do NOT mention price, stock, delivery, or refurbished process.
- Do NOT mention AI generation.
"""
    return prompt.strip()


# ============================================================
# DESCRIPTIONS OFFLINE (sans API)
# ============================================================

# Fonction qui va générer une description HTML simple sans utiliser IA
# Sert de fallback si l'API n'est pas disponible
def generate_offline_description(row, specs, language: str):
    """Génère une fiche produit simple sans IA (fallback)."""
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    short_desc = strip_html(row.get("shortDescription") or "")
    condition = row.get("filteringProperty:Stav") or ""
    category = guess_category(row)

    category_map = {
        "desktop": "stolní počítač",
        "laptop": "notebook",
        "monitor": "monitor",
        "computer": "počítač",
    }
    category_label = category_map.get(category, "počítač")

    specs_text = ". ".join(f"{label}: {value}" for label, value in specs[:6])

    parts = [
        f"<h3>{html.escape(name)}</h3>",
        f"<p>{html.escape(short_model)} je {category_label} určený pro každodenní kancelářskou práci, studium a domácí použití.</p>",
    ]
    if short_desc:
        parts.append(f"<p>{html.escape(short_desc)}</p>")
    if specs_text:
        parts.append(f"<p><strong>Konfigurace:</strong> {html.escape(specs_text)}.</p>")
    if condition:
        parts.append(f"<p>Stav zařízení: {html.escape(condition)}.</p>")
    parts.append("<p>Zařízení je otestováno a dodáváno se zárukou <strong>24 měsíců</strong>.</p>")
    return "".join(parts)


# Fonction qui va déterminer si la carte graphique est dédiée (NVIDIA / AMD) ou intégrée
def is_dedicated_gpu(gpu: str) -> bool:
    if not gpu:
        return False
    gpu_l = gpu.lower()
    # La fonction va rechercher ds mots clés dans le nom du GPU
    return any(token in gpu_l for token in ["nvidia", "geforce", "rtx", "gtx", "radeon"])


# Fonction qui va créer une description longue sans IA, mais structurée automatiquement
def generate_offline_long_description(row, specs_fields, language: str):
    """Génère une description longue structurée sans IA."""
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    spec_cpu = specs_fields.get("spec_processor", "")
    spec_ram = specs_fields.get("spec_ram", "")
    spec_storage = specs_fields.get("spec_storage", "")
    spec_gpu = specs_fields.get("spec_gpu", "")
    spec_dimensions = specs_fields.get("spec_dimensions", "")
    case_type = row.get("filteringProperty:Typ skříně") or ""

    category_map = {
        "desktop": "stolní počítač",
        "laptop": "notebook",
        "monitor": "monitor",
        "computer": "počítač",
    }
    category_label = category_map.get(category, "počítač")

    parts = []

    # Section 1 : Introduction
    parts.append(
        f"<p>{html.escape(short_model)} je {category_label} vhodný pro každodenní kancelářskou práci, "
        "domácí použití a studium. Spolehlivě zvládne e-maily, webové aplikace, práci s dokumenty i online schůzky.</p>"
    )

    # Section 2 : Marque
    parts.append(
        "<p><strong>Lenovo</strong> patří mezi přední světové výrobce výpočetní techniky. "
        "Jejich kancelářské počítače jsou navrženy pro dlouhodobé a spolehlivé používání v náročném firemním prostředí.</p>"
    )

    # Section 3 : Gamme produit
    if "thinkcentre" in short_model.lower():
        parts.append(
            f"<p>Řada <strong>ThinkCentre</strong> je určena pro profesionální použití v kanceláři i doma. "
            "Vyznačuje se tichým chodem, kompaktními rozměry a důrazem na stabilitu a dlouhou životnost.</p>"
        )

    # Section 4 : Configuration
    perf_bits = []
    if spec_cpu:
        perf_bits.append(f"Tato konfigurace je osazena procesorem <strong>{html.escape(spec_cpu)}</strong>.")
    if spec_ram:
        perf_bits.append(f"Operační paměť <strong>{html.escape(spec_ram)}</strong> zajišťuje plynulý multitasking.")
    if spec_storage:
        if "NVMe" in spec_storage:
            perf_bits.append(f"Úložiště <strong>{html.escape(spec_storage)}</strong> zajišťuje velmi rychlé načítání systému i aplikací díky technologii NVMe.")
        elif "SSD" in spec_storage:
            perf_bits.append(f"Úložiště <strong>{html.escape(spec_storage)}</strong> urychluje start systému i otevírání souborů.")
        else:
            perf_bits.append(f"Kapacita úložiště <strong>{html.escape(spec_storage)}</strong> poskytuje dostatek místa pro data a dokumenty.")
    if perf_bits:
        parts.append(f"<p>{' '.join(perf_bits)}</p>")

    # Section 5 : Hodí se pro (Convient pour)
    parts.append(
        "<p>Hodí se pro každodenní kancelářské aplikace, prohlížení webu, e-maily, videokonference a studium. "
        "Je vhodný pro domácí i firemní použití.</p>"
    )
 
    # Section 6 : Nehodí se pro (Ne convient pas pour)
    if spec_gpu and is_dedicated_gpu(spec_gpu):
        parts.append(
            "<p>Nehodí se pro profesionální 3D rendering ani pro nejnáročnější počítačové hry.</p>"
        )
    else:
        parts.append(
            "<p>Nehodí se pro nejnovější 3D hry ve vysoké kvalitě, profesionální střih 4K videa ani náročné 3D modelování.</p>"
        )
 
    # Section 7 : Připraveno k použití (Prêt à l'emploi)
    parts.append(
        "<p>Zařízení je otestováno, dodáváno s Windows 11 Pro a kryto <strong>zárukou 24 měsíců</strong>.</p>"
    )
 
    return "".join(parts)


# Fonction qui va générer un teaser court sans IA, en se basant sur les spécifications clés pour mettre en avant les avantages principaux de la machine
def generate_offline_short_description(row, specs_fields, language: str):
    """Génère un teaser court sans IA."""
    name = row.get("name") or ""
    short_model = _extract_short_model_name(row)
    category = guess_category(row)
    form_factor = guess_form_factor(row)
    spec_storage = specs_fields.get("spec_storage", "")
    spec_ram = specs_fields.get("spec_ram", "")

    category_map = {
        "desktop": "stolní počítač",
        "laptop": "notebook",
        "monitor": "monitor",
        "computer": "počítač",
    }
    category_label = category_map.get(category, "počítač")

    parts = []
    if form_factor == "mini":
        parts.append(f"Kompaktní {category_label} {html.escape(short_model)} je ideální pro kancelář, domácí práci i studium.")
    else:
        parts.append(f"{html.escape(short_model)} je spolehlivý {category_label} pro kancelář, domácí práci i studium.")

    if spec_storage and "SSD" in spec_storage:
        parts.append("Rychlé SSD zkracuje načítání systému i aplikací.")
    elif spec_ram:
        parts.append("Dostatek operační paměti usnadní multitasking.")

    return " ".join(parts).strip()


# ============================================================
# APPELS API
# ============================================================

# Fonction qui va faire un appel API celui de Chat GPT de Openai
def call_openai(prompt: str, api_key: str, model: str, max_output_tokens: int, retries: int = 2, retry_delay: float = 2.0):
    url = "https://api.openai.com/v1/responses"
    body = {
        "model": model,
        "input": prompt,
        "max_output_tokens": max_output_tokens,
        "temperature": 0.7,
    }
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST")
    req.add_header("Authorization", f"Bearer {api_key}")
    req.add_header("Content-Type", "application/json")

    attempt = 0
    while True:
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as exc:
            body_text = exc.read().decode("utf-8", "ignore")
            if exc.code in (429, 503) and "insufficient_quota" not in body_text and attempt < retries:
                time.sleep(retry_delay * (2 ** attempt))
                attempt += 1
                continue
            raise RuntimeError(f"OpenAI HTTP error: {exc.code} {exc.reason}: {body_text}") from exc

    if isinstance(payload, dict) and payload.get("output_text"):
        return payload["output_text"]

    parts = []
    for item in payload.get("output", []):
        if item.get("type") == "message":
            for content in item.get("content", []):
                if content.get("type") in ("output_text", "text"):
                    text = content.get("text")
                    if text:
                        parts.append(text)
    return "\n".join(parts).strip()


# Fonction qui va faire un appel API à Gemini de Google
def call_gemini(prompt: str, api_key: str, model: str, max_output_tokens: int, retries: int = 2, retry_delay: float = 2.0):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    body = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": max_output_tokens, "temperature": 0.7},
    }
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST")
    req.add_header("Content-Type", "application/json")

    attempt = 0
    while True:
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            break
        except urllib.error.HTTPError as exc:
            body_text = exc.read().decode("utf-8", "ignore")
            if exc.code in (429, 503) and attempt < retries:
                retry_after = parse_retry_delay_seconds(body_text)
                if retry_after > 0:
                    time.sleep(retry_after)
                else:
                    time.sleep(retry_delay * (2 ** attempt))
                attempt += 1
                continue
            raise RuntimeError(f"Gemini HTTP error: {exc.code} {exc.reason}: {body_text}") from exc

    candidates = payload.get("candidates", [])
    if not candidates:
        return ""
    content = candidates[0].get("content", {})
    parts = content.get("parts", [])
    texts = [p.get("text", "") for p in parts if p.get("text")]
    return "\n".join(texts).strip()


# ============================================================
# FIN APPELS API
# ============================================================


# Fonction qui va vérifier si le texte généré est un fragment HTML complet (toutes les balises ouvertes sont fermées), pour éviter les coupures brutales dans les descriptions et améliorer la qualité du rendu final
def is_complete_html_fragment(text: str) -> bool:
    if not text or not text.strip():
        return False

    stripped = text.strip()
    if stripped.endswith("<") or stripped.endswith("</") or stripped.count("<") != stripped.count(">"):
        return False

    tag_pattern = re.compile(r"<(/?)([a-zA-Z0-9]+)(?:\s[^>]*)?>")
    void_tags = {"br", "hr", "img", "input", "meta", "link"}
    stack = []

    for match in tag_pattern.finditer(stripped):
        is_closing = match.group(1) == "/"
        tag_name = match.group(2).lower()
        full_tag = match.group(0)

        if tag_name in void_tags or full_tag.endswith("/>"):
            continue
        if not is_closing:
            stack.append(tag_name)
            continue
        if not stack or stack[-1] != tag_name:
            return False
        stack.pop()

    return not stack


# Fonction qui va générer le texte de la description en appelant l'API du provider choisi, avec une logique de fallback et de validation pour s'assurer d'obtenir un fragment HTML complet et de bonne qualité, même en cas de coupure ou d'erreur temporaire
def generate_text_with_provider(
    prompt: str,
    provider: str,
    openai_key: str,
    openai_model: str,
    gemini_key: str,
    gemini_model: str,
    max_output_tokens: int,
    retries: int,
    retry_delay: float,
    expect_html: bool = False,
    validation_retries: int = 1,
):
    current_tokens = max_output_tokens

    for validation_attempt in range(validation_retries + 1):
        if provider == "openai":
            try:
                text = call_openai(
                    prompt, openai_key, openai_model, current_tokens, retries, retry_delay
                )
            except RuntimeError as exc:
                msg = str(exc)
                if "insufficient_quota" in msg or "quota" in msg:
                    if not gemini_key:
                        raise
                    text = call_gemini(
                        prompt, gemini_key, gemini_model, current_tokens, retries, retry_delay
                    )
                else:
                    raise
        else:
            text = call_gemini(
                prompt, gemini_key, gemini_model, current_tokens, retries, retry_delay
            )

        if not expect_html or is_complete_html_fragment(text):
            return text

        if validation_attempt < validation_retries:
            next_tokens = max(current_tokens * 2, current_tokens + 300)
            print(
                f"Warning: incomplete HTML fragment detected. Retrying with max_output_tokens={next_tokens}."
            )
            current_tokens = next_tokens

    return text


# Fonction qui va appliquer les valeurs d'un dictionnaire dans un template HTML en remplaçant les placeholders du type {{key}} par les valeurs correspondantes, pour générer les pages HTML finales à partir du template et des données produits
def apply_template(template_text: str, values: dict) -> str:
    result = template_text
    for key, value in values.items():
        result = result.replace(f"{{{{{key}}}}}", value)
    return result

# Fonction qui va aplatir un texte HTML en supprimant les retours à la ligne et les espaces superflus, pour faciliter la validation de la complétude du fragment HTML généré par l'IA
def flatten_html(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "")

# Fonction qui va extraire les paragraphes d'un texte HTML en utilisant une expression régulière pour trouver les balises <p>, et retourner une liste de paragraphes nettoyés, pour construire l'extrait "hero" à partir des 1-2 premiers paragraphes de la description longue
def extract_paragraphs_from_html(text: str):
    if not text:
        return []
    return [p.strip() for p in re.findall(r"<p\b[^>]*>.*?</p>", text, flags=re.IGNORECASE | re.DOTALL) if p.strip()]

# Fonction qui va construire un extrait "hero" à partir des 1-2 premiers paragraphes d'une description longue HTML, pour mettre en avant les avantages principaux de la machine dans les zones de mise en avant (bannières, listes de produits, etc.) et améliorer le taux de clic
def build_hero_excerpt_html(long_description_html: str) -> str:
    paragraphs = extract_paragraphs_from_html(long_description_html)
    if not paragraphs:
        return ""
    return "\n".join(paragraphs[:2]).strip()


def build_html_block_fragment(name: str, long_description_html: str, description_html: str) -> str:
    title_html = f"<h3><strong>{html.escape((name or '').strip())}</strong></h3>"
    source_html = (long_description_html or "").strip() or (description_html or "").strip()
    paragraphs = extract_paragraphs_from_html(source_html)
    body_html = "\n".join(paragraphs) if paragraphs else source_html
    return f"{title_html}\n{body_html}".strip()

# Fonction qui va normaliser les espaces dans un texte en remplaçant les séquences d'espaces, de tabulations et de retours à la ligne par un seul espace, et en supprimant les espaces en début et fin de texte, pour améliorer la qualité du rendu final et éviter les problèmes de mise en page liés aux espaces superflus
def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()

# Fonction qui va construire la description courte à partir du teaser généré par l'IA, avec une logique de validation pour éviter les descriptions trop courtes ou génériques qui n'apporteraient pas de valeur ajoutée par rapport au titre et aux spécifications, et en utilisant une description offline en fallback si nécessaire
def build_display_short_description(row, specs_fields, short_description_ai: str) -> str:
    candidate = normalize_whitespace(strip_html(short_description_ai or ""))
    fallback = normalize_whitespace(generate_offline_short_description(row, specs_fields, ""))

    if not candidate:
        return fallback

    too_short = len(candidate) < 35
    generic_values = {
        normalize_whitespace(strip_html(row.get("manufacturer") or "")),
        normalize_whitespace(_extract_short_model_name(row)),
    }
    is_generic = candidate in generic_values or len(candidate.split()) <= 2

    if too_short or is_generic:
        return fallback
    return candidate

# Fonction principale qui va orchestrer le processus de lecture du CSV, génération des descriptions, application du template et écriture des fichiers de sortie, avec une logique de reprise en cas d'interruption et de validation pour garantir la qualité des descriptions générées
def main():
    parser = argparse.ArgumentParser(description="Generate product descriptions via AI and build HTML pages.")
    parser.add_argument("--input", default="products.csv", help="Input CSV file.")
    parser.add_argument("--output-csv", default="products_with_descriptions.csv", help="Output CSV file.")
    parser.add_argument("--template", default="template_product.html", help="HTML template file.")
    parser.add_argument("--output-dir", default="out_html", help="Directory for per-product HTML files.")
    parser.add_argument("--html-blocks-out", default="produkt-ntb.html", help="HTML file with CODE blocks.")
    parser.add_argument("--csv-import-out", default="products_import.csv", help="CSV file for Shoptet import.")
    parser.add_argument("--provider", choices=["openai", "gemini"], default="openai", help="AI provider.")
    parser.add_argument("--model", default="", help="Model name. If empty, a sensible default is used.")
    parser.add_argument("--max-output-tokens", type=int, default=500, help="Max output tokens.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of products processed.")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing description field.")
    parser.add_argument("--language", default="Czech", help="Output language, e.g. Czech, French.")
    parser.add_argument("--delay", type=float, default=0.5, help="Delay between API calls in seconds.")
    parser.add_argument("--retries", type=int, default=2, help="Retry count for transient API errors.")
    parser.add_argument("--retry-delay", type=float, default=2.0, help="Base delay for retries in seconds.")
    parser.add_argument("--checkpoint-every", type=int, default=5, help="Write partial outputs every N generated items.")
    parser.add_argument("--resume", action="store_true", help="Resume from existing output CSV if available.")
    parser.add_argument("--env-file", default=".env.local", help="Optional env file with API keys.")
    parser.add_argument("--offline", action="store_true", help="Generate descriptions without API.")
    args = parser.parse_args()

    load_env_file(args.env_file)

    openai_key = os.getenv("OPENAI_API_KEY", "")
    gemini_key = os.getenv("GEMINI_API_KEY", "") or os.getenv("GOOGLE_API_KEY", "")
    openai_model = args.model or os.getenv("OPENAI_MODEL", "") or "gpt-4o-mini"
    gemini_model = args.model or os.getenv("GEMINI_MODEL", "") or "gemini-2.5-flash"

    if not args.offline:
        if args.provider == "openai":
            if not openai_key:
                raise SystemExit("Missing OPENAI_API_KEY environment variable.")
        else:
            if not gemini_key:
                raise SystemExit("Missing GEMINI_API_KEY (or GOOGLE_API_KEY) environment variable.")

    rows, fieldnames, encoding = read_csv_with_fallback(args.input, delimiter=";")
    if "description" not in fieldnames:
        fieldnames.append("description")
    if "long_description" not in fieldnames:
        fieldnames.append("long_description")
    if "short_description_ai" not in fieldnames:
        fieldnames.append("short_description_ai")

    if not os.path.isfile(args.template):
        raise SystemExit(f"Template file not found: {args.template}")
    with open(args.template, "r", encoding="utf-8") as f:
        template_text = f.read()

    os.makedirs(args.output_dir, exist_ok=True)

    processed = 0
    blocks = []
    resume_codes = set()
    if args.resume and os.path.isfile(args.output_csv):
        try:
            resume_rows, _, _ = read_csv_with_fallback(args.output_csv, delimiter=";")
            for r in resume_rows:
                if (r.get("description") or "").strip():
                    resume_codes.add((r.get("code") or "").strip())
        except Exception:
            resume_codes = set()
    for row in rows:
        if args.limit and processed >= args.limit:
            break

        code = (row.get("code") or "").strip()
        existing_desc = (row.get("description") or "").strip()
        if (code in resume_codes) and not args.overwrite:
            continue
        if existing_desc and not args.overwrite:
            continue

        # Extraction des valeurs du CSV
        specs = extract_specs(row)
        spec_fields = build_spec_fields(row)

        # Génération de la description avec IA ou en mode offline
        if args.offline:
            description_html = generate_offline_description(row, specs, args.language)
            long_description_html = generate_offline_long_description(row, spec_fields, args.language)
            short_description_ai = generate_offline_short_description(row, spec_fields, args.language)
        else:
            # Le code va construire les prompts et faire les appels API pour générer les descriptions
            prompt = build_prompt(row, specs, args.language)
            description_html = generate_text_with_provider(
                prompt,
                args.provider,
                openai_key,
                openai_model,
                gemini_key,
                gemini_model,
                args.max_output_tokens,
                args.retries,
                args.retry_delay,
                expect_html=True,
            )
            long_prompt = build_marketing_prompt(row, specs, args.language)
            short_prompt = build_short_prompt(row, specs, args.language)
            long_description_html = generate_text_with_provider(
                long_prompt,
                args.provider,
                openai_key,
                openai_model,
                gemini_key,
                gemini_model,
                args.max_output_tokens,
                args.retries,
                args.retry_delay,
                expect_html=True,
            )
            short_description_ai = generate_text_with_provider(
                short_prompt,
                args.provider,
                openai_key,
                openai_model,
                gemini_key,
                gemini_model,
                120,
                args.retries,
                args.retry_delay,
                expect_html=False,
            )

        if not description_html:
            continue

        row["description"] = description_html
        row["long_description"] = long_description_html or ""
        row["short_description_ai"] = strip_html((short_description_ai or "").strip())
        hero_excerpt_html = build_hero_excerpt_html(row["long_description"])
        display_short_description = build_display_short_description(
            row, spec_fields, row["short_description_ai"]
        )

        # Le code va ensuite appliquer les descriptions générées dans le template HTML et écrire les fichiers de sortie (Préparer les valeurs pour le template, en échappant les champs nécessaires)
        values = {
            "name": html.escape(row.get("name") or ""),
            "short_description": html.escape(display_short_description),
            "description_html": description_html,
            "long_description_html": row.get("long_description") or "",
            "hero_excerpt_html": hero_excerpt_html,
            "feature_cpu": spec_fields.get("feature_cpu", ""),
            "feature_ram": spec_fields.get("feature_ram", ""),
            "feature_storage": spec_fields.get("feature_storage", ""),
            "feature_case": spec_fields.get("feature_case", ""),
        }
        values.update(spec_fields)
        values.update(build_feature_section_fields(row, spec_fields))
        values.update(build_suitability_fields(row, spec_fields))
        values.update(build_faq_fields(row, spec_fields))

        # Générer le contenu HTML final en appliquant les valeurs au template
        html_out = apply_template(template_text, values)

        code = row.get("code") or row.get("externalId") or f"product_{processed+1}"
        name = row.get("name") or ""
        blocks.append(
            "<!-- CODE: {code} | NAME: {name} -->\n{block}\n".format(
                code=code.strip(),
                name=name.strip(),
                block=build_html_block_fragment(
                    name=name,
                    long_description_html=row.get("long_description") or "",
                    description_html=description_html,
                ),
            )
        )

        # Enregistre le contenu HTML dans un fichier individuel pour chaque produit
        safe_code = re.sub(r"[^a-zA-Z0-9_-]+", "_", code)
        out_path = os.path.join(args.output_dir, f"{safe_code}.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html_out)

        processed += 1
        time.sleep(max(0.0, args.delay))

        if args.checkpoint_every > 0 and processed % args.checkpoint_every == 0:
            write_csv(args.output_csv, rows, fieldnames, delimiter=";", encoding="utf-8")
            with open(args.html_blocks_out, "w", encoding="utf-8") as f:
                f.write("\n".join(blocks).strip() + "\n")
            rows_import = []
            for row_cp in rows:
                row_copy = dict(row_cp)
                row_copy["description"] = flatten_html(row_copy.get("description") or "")
                rows_import.append(row_copy)
            write_csv(args.csv_import_out, rows_import, fieldnames, delimiter=";", encoding="utf-8")

    write_csv(args.output_csv, rows, fieldnames, delimiter=";", encoding="utf-8")

    # Write HTML blocks file compatible with boss script format
    with open(args.html_blocks_out, "w", encoding="utf-8") as f:
        f.write("\n".join(blocks).strip() + "\n")

    # Write flattened HTML descriptions for Shoptet import
    rows_import = []
    for row in rows:
        row_copy = dict(row)
        row_copy["description"] = flatten_html(row_copy.get("description") or "")
        rows_import.append(row_copy)
    write_csv(args.csv_import_out, rows_import, fieldnames, delimiter=";", encoding="utf-8")

    print(
        f"Done. Generated {processed} descriptions. "
        f"HTML blocks: {args.html_blocks_out}. "
        f"CSV import: {args.csv_import_out}."
    )

# Point d'entrée du script
if __name__ == "__main__":
    main()
