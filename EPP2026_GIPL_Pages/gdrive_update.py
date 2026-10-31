# -*- coding: utf-8 -*-
"""
gdrive_update.py
- Stores GDrive file IDs for each Sr exhibit
- Converts sharing links to direct image URLs
- Updates HTML pages to use GDrive URLs
- Run again after adding more Sr entries to GDRIVE_IDS below
"""

import sys, io, os, re, csv
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
HTML_DIR  = os.path.join(BASE_DIR, "html_pages")
CSV_PATH  = os.path.join(BASE_DIR, "gdrive_links.csv")

# ============================================================
# GDRIVE FILE IDs  — key = Sr_No, value = Google Drive File ID
# Add more entries here as the user provides them.
# Sharing URL format: https://drive.google.com/file/d/{ID}/view
# ============================================================
GDRIVE_IDS = {
    1:   "1iSHXgFSj5AjBZGudw90X560kskZwJpa-",
    2:   "1vs7HZ9ZLYPYokSLc88SX6hKJJYrrpdig",
    3:   "1auKngYM7A8Jja_j-SwyUdI3YfQYodKSx",
    4:   "1Rkk2B4Aj7LU2uSWop215HBU6tVgxqcoj",
    5:   "10cftNVc6UNaSQIWj3g89KaRLBqoeCepA",
    6:   "1poZsbPYp6aWUEXq-zTgLv0jfBFy0Nv_O",
    7:   "1efWLaYCDykRsD_B-zYRyMIhoUdS1HyKQ",
    8:   "1_NCxbfLmgR7yoQCGM9G903daBavI4JBo",
    9:   "14LAyFjJCvJIxIi1h3wcT2r-IRN7Uq9H0",
    10:  "1Z-lUNmdwRtY_aGEfURoSi7MNF8BQedfn",
    11:  "1mGgMkoulkvY4cldfNCdWCrucgoCvnx3I",
    12:  "1Y_kn0ZAPu_DCZeMIdPdoNzvA7FFtvh78",
    13:  "1lltO1cQTEBZQ4SADoAMTlu9uxOxfI4sn",
    14:  "1m6MKHLLr9-nTiiKTaentZq1XfwBfQbou",
    15:  "1UlBTzT2RzVLPsy-6EQnk26ypsr0gmXXw",
    16:  "1zmqLClzG6qZ4BRy7kGYZYvs4Vk9ZsUn_",
    17:  "1FifmforIO3tmA-WTDiR7XPbp_KMAQUc_",
    18:  "1jWb-9fTJnq-nHwzsCpYt_fYFp_gogmkd",
    19:  "1Ii8sPwP5v8RQjvNiOb0tSZv0e2_PBxNI",
    20:  "116NqBTDMW0wqDOBZjrW7q6iJAxNmhu0f",
    21:  "1KDDAF1wuLOO37zCaZGPwDyUj_ymqQICx",
    22:  "1d4cqm3U4h51DPbHgMep-qPpCmWTg1Nlz",
    23:  "1XEe_0VVo6qTxbEJHeCCWuNa2VY1lo8Wn",
    24:  "1okOrV_07y7tueS-0jkEejXiYn2K-5gsp",
    25:  "1GDnPFeqDogPLcFF_b8Gx_H2KF5XeDpDl",
    26:  "12wJ-jvNCcDqtpOaeRXZ9MoemxGmxT07s",
    27:  "1mE3QahwMcuD_9kTg4PGcwflo7H_BBPGb",
    28:  "141bNf_5IHZyGt5_ISKW-OBGF6mKOn2K0",
    29:  "1J2ZHSisDr1hqotChRUFMzTUshbYu8fxM",
    30:  "1EoETuYM4PEsuMe9yYOUcInk0MclR4u38",
    31:  "1fo3Oyo4eIKcq78m_R1ezijzk7DDXAUxp",
    32:  "182mMtxOLfZE4kfZCzpNltpbC1Fg0Gcwd",
    33:  "1kmT4UavVDmtAk7oKA6ZAhBbdSMB6K6k5",
    34:  "1AVXqDV7NObxZTQ1Ae_jOCsQeKjuFxIJ-",
    35:  "1VcN4UKbk9YhlQ9-ZULQrnGgj8Rd2zJsj",
    36:  "1sdEgxAtL8JrJdcGV4rDS9XIGxGvTSdqH",
    37:  "16mbAjaRUfMfIMADOwx6BdLafV0aLjCra",
    38:  "1B8il0ASy9JfFCpdh1TBxMLqmdcdApHQe",
    39:  "1n3RmCD5mV3Haq5ZcNfp4Eh_cuOGcJCGk",
    40:  "1ELn2CS7UG6X2oCX9_N7gX0o10A6I7V3i",
    41:  "1j6klPS2RHa5IRJIdlD7aN_MiZpTWDTcU",
    42:  "1P3xBDrjlVHD4RUsF_TcRdYnHG8ICNzhs",
    43:  "1ATZtkDScpIIt7KEtAjWsUFMLXAUVFTaL",
    44:  "1NAXyILHt-ROwW9DCcmoea3cDlWBewlm_",
    45:  "1bfwKYB3rw7rlK4lCAPHKI0nSpAyv2HEB",
    46:  "1Eik2IPmWsV-7UIBSAY47wtcCSItgcqpX",
    47:  "1M0qL-zdga6obfX5lR0HaHqh65ZRbBpjV",
    48:  "1RMoYd29pGrarnHysBrOnWbJNS4nkIPD0",
    49:  "1kC6XlUuWABvRaRdnH2AKNm4novAuGY8x",
    50:  "1uF6ZPO7GgCEM2D2-tkKlmjHqghwzaBdg",
    51:  "1UVlNpuVYXd5ksyxuna3vEA_1GpPANeBk",  # SWAMI VIVEKANANDA MEMORIAL
    52:  "1UVlNpuVYXd5ksyxuna3vEA_1GpPANeBk",  # SPICES INSTALLATION — WARNING: same ID as Sr51, likely wrong
    53:  "17CsfuJfwAaCe1LY5o4U88Ub-uLWfcV-D",
    54:  "1hL4TkUFw9r1h5qsDYdR0y2vdCYODC-UO",
    55:  "1ztd8-c48vKOEVTPZwPMtJRzQT59IyL9Z",
    # 56: INSTALLATION OF VANDE MATARAM — link not provided, PENDING
    57:  "1yHPOXqasGQ3WMRp4U5BGRyOu1cMAp10E",
    58:  "1MbbTqwVij69GKzN_UszDiUWjtC-djJqW",
    59:  "13fjdQMr6P4RKRdPRGl8rB1K3EcM_BBlH",
    60:  "1HQT4eytWmbk8VZ8NFN1RZ6xg0r-Rx0XO",
    61:  "1XMBOGoydzr7wTzBcspzHc5-Kzl_hS92Y",
    62:  "1-xlODEAwmKPt6oalmhURujnFSDl01HKx",
    63:  "1niJg8s3ninhzQquLDVviWLdNFI3NckIA",
    64:  "1T8UvHIlpgCE01noxfYUvfP4s1iziISny",
    65:  "1fM-YeanwVYbsNNdMRDcFrZ7o6ffpOOY0",
    66:  "1bkIJ70TjOtvipxzaAR9HrqWeZC8EoSfr",
    67:  "1QJ4XyQF79QRPgCr1brepyxW4GTnScWZ6",
    68:  "1HTR7kLwEuW5Ll779wyU7EukFHOL8nF9V",
    69:  "1y5YPd4UxcC8i2yBiqCYRTBT9o7UTt3mA",
    70:  "1O0THAHc6pPXuTeR9nqvu9jtT_N9whK1r",
    71:  "17fEAFXGXRgm10rffbJ5A-YFq---qIruj",
    72:  "1goglWMzCORSa3CLUH5qx2sXRjVpdodXv",
    73:  "1k3to38G9CVRj85BTYcBa5j7CQf9qFQXu",
    74:  "1kAjRhD3yhvlmpiB0B9XR8F_ozD0d97gr",
    75:  "1VQUHHNZlRoGzT6U1EQBhdY62VYHU8yJc",
    # --- Add Sr 76-152 below as links arrive ---
}

# Sr -> (gipl_id, title_en) — needed for HTML filename lookup
SR_INFO = {
    1:(1,"THE KAALEEN CORRIDOR"),2:(2,"AARI EMBROIDERY"),
    3:(3,"KASHMIRI LATTICE WORK"),4:(4,"THE BOAT HOUSE"),
    5:(5,"SNOW LEOPARD"),6:(6,"MANI CHAKRA"),
    7:(7,"KULLU SHAWL"),8:(8,"KINNAURI NATI DANCE"),
    9:(9,"PHULKARI"),10:(10,"GIDDHA DANCE"),
    11:(11,"LOTUS TEMPLE"),12:(12,"BAISAKHI"),
    13:(13,"BHANGRA"),14:(14,"ASHOK STAMBH"),
    15:(15,"GANGA AARTI"),16:(16,"RAM MANDIR"),
    17:(17,"BRIJ HOLI"),18:(18,"PICHWAI ART"),
    19:(19,"JHANSI FORT"),20:(20,"DIWALI"),
    21:(21,"KAVAD ART"),22:(22,"KALBELIYA"),
    23:(23,"GHOOMAR"),24:(24,"CANE CRAFT"),
    25:(25,"BIHU DANCE"),26:(26,"GAYAN BAYAN"),
    27:(27,"MUKHA SILPA"),28:(28,"TEA PLANTATION"),
    29:(29,"WARRIOR'S CREST"),30:(30,"TREE OF LIFE"),
    31:(31,"TRIBAL HOUSE"),32:(32,"FOLK DANCES OF EASTERN INDIA"),
    33:(33,"PATTACHITRA"),34:(34,"KONARK ARCHITECTURE"),
    35:(35,"MANDALA CHAKRA"),36:(36,"RATH YATRA"),
    37:(37,"HOWRAH BRIDGE"),38:(38,"GARBA"),
    39:(39,"LIPPAN ART"),40:(40,"SABARMATI ASHRAM"),
    41:(41,"FABRIC OF WEST"),42:(42,"SHANIWAR WADA"),
    43:(43,"STRAWBERRY FARMING"),44:(44,"ORANGE FARMING"),
    45:(45,"MULBERRY FARMING"),46:(46,"DABBAWALA"),
    47:(47,"SHIVNERI FORT"),48:(48,"LOCAL TRAIN"),
    49:(49,"MEENAKSHI TEMPLE"),50:(50,"THIRUVALLUVAR STATUE"),
    51:(51,"SWAMI VIVEKANANDA MEMORIAL"),52:(52,"SPICES INSTALLATION"),
    53:(53,"SILHOUETTE INSTALLATION"),54:(54,"HAMPI STONE"),
    55:(55,"KERALA BOAT RACE"),56:(56,"INSTALLATION OF VANDE MATARAM"),
    57:(57,"PARADE"),58:(58,"ENVIRONMENT"),
    59:(59,"FLAMES OF TIME"),60:(60,"PANCH TATTVA"),
    61:(61,"GIFT CITY"),62:(62,"SHADOW INSTALLATION"),
    63:(63,"RIVER OF TIME"),64:(64,"THE TREE OF GRACE"),
    65:(65,"NARI UDAY"),66:(66,"WINGS OF DREAMS"),
    67:(67,"STAIRS TO KNOWLEDGE"),68:(68,"PETALORA OF LIGHT"),
    69:(69,"SHE LEADS"),70:(70,"WOMEN OF TODAY"),
    71:(71,"NURTURE"),72:(72,"ITBP"),
    73:(73,"CRPF"),74:(74,"CAPF"),
    75:(75,"POLICE ON DUTY"),76:(76,"BSF"),
    77:(77,"BSF"),78:(78,"CISF"),
    79:(79,"COAST GUARD"),80:(80,"AIR FORCE"),
    81:(81,"HANDS OF PROGRESS"),82:(82,"TECH BLOOM"),
    83:(83,"MICROCHIP REVOLUTION"),84:(84,"BLUEPRINT OF INTELLIGENCE"),
    85:(85,"INFINITE CONNECTION"),86:(86,"THE RISING HORIZON"),
    87:(87,"KITE FESTIVAL"),88:(88,"SURFING STATION"),
    89:(89,"SACRED LOTUS"),90:(90,"BLOSSOM HORSE"),
    91:(91,"PUNJAB DI SHAAN"),92:(92,"WELL"),
    93:(93,"NALANDA UNIVERSITY"),94:(94,"CAMELS"),
    95:(95,"WALL OF KATHPUTLI"),96:(96,"HAWA MAHAL"),
    97:(97,"KAZIRANGA'S RHINOCEROS"),98:(98,"LIVING ROOT BRIDGE"),
    99:(99,"BENGALI TOUCH"),100:(100,"RANI KI VAV"),
    101:(101,"FIRE DANCER"),102:(102,"VIBE OF GOA"),
    103:(103,"NASHIK DHOL"),104:(104,"MARATHAS"),
    105:(105,"CAMEL CART"),106:(106,"MARINE DRIVE"),
    107:(107,"RAM SETU"),108:(108,"ELEPHANT OF SOUTH"),
    109:(109,"GOLDEN BIRD"),110:(110,"SAMAY KA CHAKRA"),
    111:(111,"SEVEN HORSES"),112:(114,"ELEPHANT"),
    113:(115,"DEEP DHARA"),115:(116,"FLICKER AND FLAME"),
    116:(117,"WAVE OF LIGHT"),117:(118,"TUNNEL OF HOT AIR BALLOON"),
    118:(119,"TUNNEL OF LAMPS"),119:(120,"TUNNEL OF THREADS"),
    120:(121,"FIGHTER JET"),121:(122,"INDIA IN SPACE"),
    122:(123,"SPACE AND TECHNOLOGY"),123:(124,"MAKE IN INDIA"),
    124:(125,"YOGA AASAN"),125:(126,"SURYA NAMASKAR"),
    126:(127,"MATSYASANA"),127:(128,"SIRSHASANA"),
    128:(129,"SARVANGASANA"),129:(130,"PASCHIMOTTANASANA"),
    130:(131,"HALASANA"),131:(132,"BHUJANGASANA"),
    132:(133,"ARDHA MATSYENDRASANA"),133:(134,"PADA HASTASANA"),
    134:(135,"KAKASANA"),135:(136,"DHANURASANA"),
    136:(137,"SALABHASANA"),137:(138,"TRIKONASANA"),
    138:(139,"SPIDER WEB"),139:(140,"ENCHANTED FOREST"),
    140:(141,"TUNNEL OF DEEP WATER"),141:(142,"AQUATIC TUNNEL"),
    142:(143,"LVM3 ROCKET"),143:(144,"PSLV ROCKET"),
    144:(145,"GAGANYAAN ASTRONAUT"),145:(146,"MARS ORBITER MISSION"),
    146:(147,"ADITYA-L1 SOLAR MISSION"),147:(148,"SPACE PORTAL"),
    148:(149,"VIKRAM LANDER"),149:(150,"PRAGYAN ROVER"),
    150:(151,"SATELLITE COMMUNICATION DISH"),151:(152,"ARTIFICIAL SATELLITE"),
    152:(153,"INDIA IN SPACE TITLE PANEL"),
}

def gdrive_direct(file_id):
    return f"https://lh3.googleusercontent.com/d/{file_id}"

def gdrive_view(file_id):
    return f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

def html_fname(sr, title):
    slug = re.sub(r"[^A-Z0-9]+", "_", title.upper()).strip("_")
    return f"EPP2026_{sr:03d}_{slug}.html"

# ============================================================
# 1. Save / update gdrive_links.csv
# ============================================================
rows = []
for sr in sorted(SR_INFO.keys()):
    gid_val, title = SR_INFO[sr]
    fid = GDRIVE_IDS.get(sr, "")
    rows.append({
        "Sr_No":      f"{sr:03d}",
        "GIPL_Id":    gid_val,
        "Title_EN":   title,
        "GDrive_ID":  fid,
        "View_URL":   gdrive_view(fid)  if fid else "",
        "Direct_URL": gdrive_direct(fid) if fid else "",
        "Status":     "LINKED" if fid else "PENDING",
    })

with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=["Sr_No","GIPL_Id","Title_EN","GDrive_ID","View_URL","Direct_URL","Status"])
    w.writeheader()
    w.writerows(rows)

linked = sum(1 for r in rows if r["Status"] == "LINKED")
print(f"Saved gdrive_links.csv  ({linked} linked, {len(rows)-linked} pending)")

# ============================================================
# 2. Update HTML pages for every Sr that has a GDrive ID
# ============================================================
updated = skipped = 0
for sr, fid in GDRIVE_IDS.items():
    info = SR_INFO.get(sr)
    if not info:
        print(f"  WARN: Sr {sr} not in SR_INFO")
        continue
    _, title = info
    fname = html_fname(sr, title)
    fpath = os.path.join(HTML_DIR, fname)
    if not os.path.exists(fpath):
        print(f"  MISS: {fname}")
        continue

    direct_url = gdrive_direct(fid)
    with open(fpath, encoding="utf-8") as f:
        html = f.read()

    # Replace GIPL server URLs (original) and old GDrive uc?export=view URLs
    new_html = re.sub(
        r'(src=")https://rashtriyaektadiwas\.com/assets/img/EktaPrakashParvDetails/[^"]+(")',
        rf'\g<1>{direct_url}\g<2>',
        html
    )
    new_html = re.sub(
        r'(src=")https://drive\.google\.com/uc\?export=view&id=[^"]+(")',
        rf'\g<1>{direct_url}\g<2>',
        new_html
    )

    if new_html == html:
        print(f"  SKIP (no GIPL img src found): {fname}")
        skipped += 1
        continue

    with open(fpath, "w", encoding="utf-8") as f:
        f.write(new_html)
    print(f"  OK  {fname}")
    updated += 1

print(f"\nHTML pages updated : {updated}")
print(f"Skipped (no change): {skipped}")
print(f"\nGDrive CSV saved to: {CSV_PATH}")
print(f"\nNext: add more Sr entries to GDRIVE_IDS in this script and re-run.")
