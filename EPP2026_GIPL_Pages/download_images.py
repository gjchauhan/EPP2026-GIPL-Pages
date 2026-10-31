# -*- coding: utf-8 -*-
"""
Download all available exhibit images from GIPL server.
Saves to EPP2026_GIPL_Pages/images/ with naming:
  001_The_Kaaleen_Corridor.jpg
  056_Installation_Of_Vande_Mataram.png
  etc.
IDs 87-153: not available on server — listed in MISSING_IMAGES.txt
"""

import sys, io, os, re, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import urllib.request as ul
except ImportError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "images")
os.makedirs(IMG_DIR, exist_ok=True)

BASE_URL = "https://rashtriyaektadiwas.com/assets/img/EktaPrakashParvDetails/"

# (sr_no, gipl_id, title_en) — active only
EXHIBITS = [
    (1,1,"THE KAALEEN CORRIDOR"),(2,2,"AARI EMBROIDERY"),
    (3,3,"KASHMIRI LATTICE WORK"),(4,4,"THE BOAT HOUSE"),
    (5,5,"SNOW LEOPARD"),(6,6,"MANI CHAKRA"),
    (7,7,"KULLU SHAWL"),(8,8,"KINNAURI NATI DANCE"),
    (9,9,"PHULKARI"),(10,10,"GIDDHA DANCE"),
    (11,11,"LOTUS TEMPLE"),(12,12,"BAISAKHI"),
    (13,13,"BHANGRA"),(14,14,"ASHOK STAMBH"),
    (15,15,"GANGA AARTI"),(16,16,"RAM MANDIR"),
    (17,17,"BRIJ HOLI"),(18,18,"PICHWAI ART"),
    (19,19,"JHANSI FORT"),(20,20,"DIWALI"),
    (21,21,"KAVAD ART"),(22,22,"KALBELIYA"),
    (23,23,"GHOOMAR"),(24,24,"CANE CRAFT"),
    (25,25,"BIHU DANCE"),(26,26,"GAYAN BAYAN"),
    (27,27,"MUKHA SILPA"),(28,28,"TEA PLANTATION"),
    (29,29,"WARRIOR'S CREST"),(30,30,"TREE OF LIFE"),
    (31,31,"TRIBAL HOUSE"),(32,32,"FOLK DANCES OF EASTERN INDIA"),
    (33,33,"PATTACHITRA"),(34,34,"KONARK ARCHITECTURE"),
    (35,35,"MANDALA CHAKRA"),(36,36,"RATH YATRA"),
    (37,37,"HOWRAH BRIDGE"),(38,38,"GARBA"),
    (39,39,"LIPPAN ART"),(40,40,"SABARMATI ASHRAM"),
    (41,41,"FABRIC OF WEST"),(42,42,"SHANIWAR WADA"),
    (43,43,"STRAWBERRY FARMING"),(44,44,"ORANGE FARMING"),
    (45,45,"MULBERRY FARMING"),(46,46,"DABBAWALA"),
    (47,47,"SHIVNERI FORT"),(48,48,"LOCAL TRAIN"),
    (49,49,"MEENAKSHI TEMPLE"),(50,50,"THIRUVALLUVAR STATUE"),
    (51,51,"SWAMI VIVEKANANDA MEMORIAL"),(52,52,"SPICES INSTALLATION"),
    (53,53,"SILHOUETTE INSTALLATION"),(54,54,"HAMPI STONE"),
    (55,55,"KERALA BOAT RACE"),
    # IDs 56-86: PNG format
    (56,56,"INSTALLATION OF VANDE MATARAM"),(57,57,"PARADE"),
    (58,58,"ENVIRONMENT"),(59,59,"FLAMES OF TIME"),
    (60,60,"PANCH TATTVA"),(61,61,"GIFT CITY"),
    (62,62,"SHADOW INSTALLATION"),(63,63,"RIVER OF TIME"),
    (64,64,"THE TREE OF GRACE"),(65,65,"NARI UDAY"),
    (66,66,"WINGS OF DREAMS"),(67,67,"STAIRS TO KNOWLEDGE"),
    (68,68,"PETALORA OF LIGHT"),(69,69,"SHE LEADS"),
    (70,70,"WOMEN OF TODAY"),(71,71,"NURTURE"),
    (72,72,"ITBP"),(73,73,"CRPF"),
    (74,74,"CAPF"),(75,75,"POLICE ON DUTY"),
    (76,76,"BSF"),(77,77,"BSF"),
    (78,78,"CISF"),(79,79,"COAST GUARD"),
    (80,80,"AIR FORCE"),(81,81,"HANDS OF PROGRESS"),
    (82,82,"TECH BLOOM"),(83,83,"MICROCHIP REVOLUTION"),
    (84,84,"BLUEPRINT OF INTELLIGENCE"),(85,85,"INFINITE CONNECTION"),
    (86,86,"THE RISING HORIZON"),
    # IDs 87-153: NOT available on server (listed in MISSING_IMAGES.txt)
    (87,87,"KITE FESTIVAL"),(88,88,"SURFING STATION"),
    (89,89,"SACRED LOTUS"),(90,90,"BLOSSOM HORSE"),
    (91,91,"PUNJAB DI SHAAN"),(92,92,"WELL"),
    (93,93,"NALANDA UNIVERSITY"),(94,94,"CAMELS"),
    (95,95,"WALL OF KATHPUTLI"),(96,96,"HAWA MAHAL"),
    (97,97,"KAZIRANGA'S RHINOCEROS"),(98,98,"LIVING ROOT BRIDGE"),
    (99,99,"BENGALI TOUCH"),(100,100,"RANI KI VAV"),
    (101,101,"FIRE DANCER"),(102,102,"VIBE OF GOA"),
    (103,103,"NASHIK DHOL"),(104,104,"MARATHAS"),
    (105,105,"CAMEL CART"),(106,106,"MARINE DRIVE"),
    (107,107,"RAM SETU"),(108,108,"ELEPHANT OF SOUTH"),
    (109,109,"GOLDEN BIRD"),(110,110,"SAMAY KA CHAKRA"),
    (111,111,"SEVEN HORSES"),(112,114,"ELEPHANT"),
    (113,115,"DEEP DHARA"),(115,116,"FLICKER AND FLAME"),
    (116,117,"WAVE OF LIGHT"),(117,118,"TUNNEL OF HOT AIR BALLOON"),
    (118,119,"TUNNEL OF LAMPS"),(119,120,"TUNNEL OF THREADS"),
    (120,121,"FIGHTER JET"),(121,122,"INDIA IN SPACE"),
    (122,123,"SPACE AND TECHNOLOGY"),(123,124,"MAKE IN INDIA"),
    (124,125,"YOGA AASAN"),(125,126,"SURYA NAMASKAR"),
    (126,127,"MATSYASANA"),(127,128,"SIRSHASANA"),
    (128,129,"SARVANGASANA"),(129,130,"PASCHIMOTTANASANA"),
    (130,131,"HALASANA"),(131,132,"BHUJANGASANA"),
    (132,133,"ARDHA MATSYENDRASANA"),(133,134,"PADA HASTASANA"),
    (134,135,"KAKASANA"),(135,136,"DHANURASANA"),
    (136,137,"SALABHASANA"),(137,138,"TRIKONASANA"),
    (138,139,"SPIDER WEB"),(139,140,"ENCHANTED FOREST"),
    (140,141,"TUNNEL OF DEEP WATER"),(141,142,"AQUATIC TUNNEL"),
    (142,143,"LVM3 ROCKET"),(143,144,"PSLV ROCKET"),
    (144,145,"GAGANYAAN ASTRONAUT"),(145,146,"MARS ORBITER MISSION"),
    (146,147,"ADITYA-L1 SOLAR MISSION"),(147,148,"SPACE PORTAL"),
    (148,149,"VIKRAM LANDER"),(149,150,"PRAGYAN ROVER"),
    (150,151,"SATELLITE COMMUNICATION DISH"),(151,152,"ARTIFICIAL SATELLITE"),
    (152,153,"INDIA IN SPACE TITLE PANEL"),
]

def slug(title):
    return re.sub(r"[^A-Z0-9]+", "_", title.upper()).strip("_").title().replace("_", "_")

def title_slug(title):
    return re.sub(r"[^A-Za-z0-9]+", "_", title.title()).strip("_")

downloaded = []
skipped = []
missing = []

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
}

print(f"Downloading images to: {IMG_DIR}\n")

for sr, gid, ten in EXHIBITS:
    ext = 'jpg' if gid <= 55 else ('png' if gid <= 86 else None)
    name_slug = title_slug(ten)
    fname = f"{sr:03d}_{name_slug}.{ext}" if ext else None

    if ext is None:
        missing.append((sr, gid, ten))
        continue

    url = f"{BASE_URL}{gid}.{ext}"
    fpath = os.path.join(IMG_DIR, fname)

    if os.path.exists(fpath):
        print(f"  SKIP (exists) {fname}")
        skipped.append(fname)
        continue

    try:
        req = ul.Request(url, headers=headers)
        with ul.urlopen(req, timeout=15) as resp:
            data = resp.read()
        with open(fpath, 'wb') as f:
            f.write(data)
        kb = len(data) // 1024
        print(f"  OK  {fname}  ({kb} KB)")
        downloaded.append(fname)
        time.sleep(0.15)
    except Exception as e:
        print(f"  ERR {fname}: {e}")
        missing.append((sr, gid, ten))

# Write missing list
missing_path = os.path.join(IMG_DIR, "MISSING_IMAGES.txt")
with open(missing_path, 'w', encoding='utf-8') as f:
    f.write("Images not available on GIPL server (IDs 87-153)\n")
    f.write("Please provide these images and share Google Drive link.\n")
    f.write("Naming convention: {sr:03d}_{Title_Slug}.jpg\n\n")
    for sr, gid, ten in missing:
        name_slug = title_slug(ten)
        f.write(f"Sr {sr:03d} | GIPL_Id {gid:03d} | {ten} | filename: {sr:03d}_{name_slug}.jpg\n")

print(f"\n{'='*60}")
print(f"Downloaded : {len(downloaded)}")
print(f"Skipped    : {len(skipped)} (already existed)")
print(f"Missing    : {len(missing)} (not on server)")
print(f"Missing list: {missing_path}")
print(f"\nImages folder: {IMG_DIR}")
