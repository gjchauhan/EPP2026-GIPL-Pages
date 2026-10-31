# -*- coding: utf-8 -*-
"""
EPP2026 GIPL Immersive Pages Generator
Generates 152 HTML exhibit pages for Ekta Prakash Parv 2026
Run: python generate_pages.py
"""

import sys, io, os, re, csv
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import openpyxl
except ImportError:
    print("Installing openpyxl..."); os.system("pip install openpyxl"); import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(BASE_DIR), "Data")
HTML_DIR = os.path.join(BASE_DIR, "html_pages")
os.makedirs(HTML_DIR, exist_ok=True)

# =============================================================
# MASTER EXHIBIT LIST: (sr_no, gipl_id, title_en, active)
# NOTE: sr_no skips 114 (inactive DEEP DHARA at Id=112)
# =============================================================
EXHIBITS_RAW = [
    (1,1,"THE KAALEEN CORRIDOR",True),(2,2,"AARI EMBROIDERY",True),
    (3,3,"KASHMIRI LATTICE WORK",True),(4,4,"THE BOAT HOUSE",True),
    (5,5,"SNOW LEOPARD",True),(6,6,"MANI CHAKRA",True),
    (7,7,"KULLU SHAWL",True),(8,8,"KINNAURI NATI DANCE",True),
    (9,9,"PHULKARI",True),(10,10,"GIDDHA DANCE",True),
    (11,11,"LOTUS TEMPLE",True),(12,12,"BAISAKHI",True),
    (13,13,"BHANGRA",True),(14,14,"ASHOK STAMBH",True),
    (15,15,"GANGA AARTI",True),(16,16,"RAM MANDIR",True),
    (17,17,"BRIJ HOLI",True),(18,18,"PICHWAI ART",True),
    (19,19,"JHANSI FORT",True),(20,20,"DIWALI",True),
    (21,21,"KAVAD ART",True),(22,22,"KALBELIYA",True),
    (23,23,"GHOOMAR",True),(24,24,"CANE CRAFT",True),
    (25,25,"BIHU DANCE",True),(26,26,"GAYAN BAYAN",True),
    (27,27,"MUKHA SILPA",True),(28,28,"TEA PLANTATION",True),
    (29,29,"WARRIOR'S CREST",True),(30,30,"TREE OF LIFE",True),
    (31,31,"TRIBAL HOUSE",True),(32,32,"FOLK DANCES OF EASTERN INDIA",True),
    (33,33,"PATTACHITRA",True),(34,34,"KONARK ARCHITECTURE",True),
    (35,35,"MANDALA CHAKRA",True),(36,36,"RATH YATRA",True),
    (37,37,"HOWRAH BRIDGE",True),(38,38,"GARBA",True),
    (39,39,"LIPPAN ART",True),(40,40,"SABARMATI ASHRAM",True),
    (41,41,"FABRIC OF WEST",True),(42,42,"SHANIWAR WADA",True),
    (43,43,"STRAWBERRY FARMING",True),(44,44,"ORANGE FARMING",True),
    (45,45,"MULBERRY FARMING",True),(46,46,"DABBAWALA",True),
    (47,47,"SHIVNERI FORT",True),(48,48,"LOCAL TRAIN",True),
    (49,49,"MEENAKSHI TEMPLE",True),(50,50,"THIRUVALLUVAR STATUE",True),
    (51,51,"SWAMI VIVEKANANDA MEMORIAL",True),(52,52,"SPICES INSTALLATION",True),
    (53,53,"SILHOUETTE INSTALLATION",True),(54,54,"HAMPI STONE",True),
    (55,55,"KERALA BOAT RACE",True),(56,56,"INSTALLATION OF VANDE MATARAM",True),
    (57,57,"PARADE",True),(58,58,"ENVIRONMENT",True),
    (59,59,"FLAMES OF TIME",True),(60,60,"PANCH TATTVA",True),
    (61,61,"GIFT CITY",True),(62,62,"SHADOW INSTALLATION",True),
    (63,63,"RIVER OF TIME",True),(64,64,"THE TREE OF GRACE",True),
    (65,65,"NARI UDAY",True),(66,66,"WINGS OF DREAMS",True),
    (67,67,"STAIRS TO KNOWLEDGE",True),(68,68,"PETALORA OF LIGHT",True),
    (69,69,"SHE LEADS",True),(70,70,"WOMEN OF TODAY",True),
    (71,71,"NURTURE",True),(72,72,"ITBP",True),
    (73,73,"CRPF",True),(74,74,"CAPF",True),
    (75,75,"POLICE ON DUTY",True),(76,76,"BSF",True),
    (77,77,"BSF",True),(78,78,"CISF",True),
    (79,79,"COAST GUARD",True),(80,80,"AIR FORCE",True),
    (81,81,"HANDS OF PROGRESS",True),(82,82,"TECH BLOOM",True),
    (83,83,"MICROCHIP REVOLUTION",True),(84,84,"BLUEPRINT OF INTELLIGENCE",True),
    (85,85,"INFINITE CONNECTION",True),(86,86,"THE RISING HORIZON",True),
    (87,87,"KITE FESTIVAL",True),(88,88,"SURFING STATION",True),
    (89,89,"SACRED LOTUS",True),(90,90,"BLOSSOM HORSE",True),
    (91,91,"PUNJAB DI SHAAN",True),(92,92,"WELL",True),
    (93,93,"NALANDA UNIVERSITY",True),(94,94,"CAMELS",True),
    (95,95,"WALL OF KATHPUTLI",True),(96,96,"HAWA MAHAL",True),
    (97,97,"KAZIRANGA'S RHINOCEROS",True),(98,98,"LIVING ROOT BRIDGE",True),
    (99,99,"BENGALI TOUCH",True),(100,100,"RANI KI VAV",True),
    (101,101,"FIRE DANCER",True),(102,102,"VIBE OF GOA",True),
    (103,103,"NASHIK DHOL",True),(104,104,"MARATHAS",True),
    (105,105,"CAMEL CART",True),(106,106,"MARINE DRIVE",True),
    (107,107,"RAM SETU",True),(108,108,"ELEPHANT OF SOUTH",True),
    (109,109,"GOLDEN BIRD",True),(110,110,"SAMAY KA CHAKRA",True),
    (111,111,"SEVEN HORSES",True),(112,114,"ELEPHANT",True),
    (113,115,"DEEP DHARA",True),(114,112,"DEEP DHARA",False),
    (115,116,"FLICKER AND FLAME",True),(116,117,"WAVE OF LIGHT",True),
    (117,118,"TUNNEL OF HOT AIR BALLOON",True),(118,119,"TUNNEL OF LAMPS",True),
    (119,120,"TUNNEL OF THREADS",True),(120,121,"FIGHTER JET",True),
    (121,122,"INDIA IN SPACE",True),(122,123,"SPACE AND TECHNOLOGY",True),
    (123,124,"MAKE IN INDIA",True),(124,125,"YOGA AASAN",True),
    (125,126,"SURYA NAMASKAR",True),(126,127,"MATSYASANA",True),
    (127,128,"SIRSHASANA",True),(128,129,"SARVANGASANA",True),
    (129,130,"PASCHIMOTTANASANA",True),(130,131,"HALASANA",True),
    (131,132,"BHUJANGASANA",True),(132,133,"ARDHA MATSYENDRASANA",True),
    (133,134,"PADA HASTASANA",True),(134,135,"KAKASANA",True),
    (135,136,"DHANURASANA",True),(136,137,"SALABHASANA",True),
    (137,138,"TRIKONASANA",True),(138,139,"SPIDER WEB",True),
    (139,140,"ENCHANTED FOREST",True),(140,141,"TUNNEL OF DEEP WATER",True),
    (141,142,"AQUATIC TUNNEL",True),(142,143,"LVM3 ROCKET",True),
    (143,144,"PSLV ROCKET",True),(144,145,"GAGANYAAN ASTRONAUT",True),
    (145,146,"MARS ORBITER MISSION",True),(146,147,"ADITYA-L1 SOLAR MISSION",True),
    (147,148,"SPACE PORTAL",True),(148,149,"VIKRAM LANDER",True),
    (149,150,"PRAGYAN ROVER",True),(150,151,"SATELLITE COMMUNICATION DISH",True),
    (151,152,"ARTIFICIAL SATELLITE",True),(152,153,"INDIA IN SPACE TITLE PANEL",True),
]
ACTIVE = [(sr,gid,ten,act) for sr,gid,ten,act in EXHIBITS_RAW if act]

# Build quick lookups
BY_SR  = {sr:(sr,gid,ten) for sr,gid,ten,_ in EXHIBITS_RAW if _}
BY_GID = {gid:(sr,gid,ten) for sr,gid,ten,_ in EXHIBITS_RAW if _}
ACTIVE_SRS = [sr for sr,*_ in ACTIVE]

# =============================================================
# HINDI NAMES
# =============================================================
HINDI_NAME = {
    1:"कालीन गलियारा",2:"आरी कढ़ाई",3:"कश्मीर की जालीदार कला",
    4:"शिकारा",5:"स्नो लेपर्ड",6:"मणि चक्र",7:"कुल्लू शॉल",
    8:"किन्नौरी नाटी नृत्य",9:"फुलकारी",10:"गिद्धा नृत्य",
    11:"कमल मंदिर",12:"बैसाखी",13:"भांगड़ा",14:"अशोक स्तंभ",
    15:"गंगा आरती",16:"राम मंदिर",17:"ब्रज की होली",18:"पिछवाई कला",
    19:"झाँसी का किला",20:"दीवाली",21:"कावड़ कला",22:"कालबेलिया नृत्य",
    23:"घूमर",24:"बेंत शिल्प",25:"बिहू नृत्य",26:"गायन-बायन",
    27:"मुखा शिल्प",28:"चाय के बागान",29:"योद्धा का मुकुट",30:"जीवन वृक्ष",
    31:"जनजातीय घर",32:"पूर्वी भारत के लोकनृत्य",33:"पट्टचित्र",
    34:"कोणार्क की वास्तुकला",35:"मंडल चक्र",36:"रथ यात्रा",37:"हावड़ा पुल",
    38:"गरबा",39:"लिप्पन कला",40:"साबरमती आश्रम",41:"पश्चिम भारत के वस्त्र",
    42:"शनिवार वाड़ा",43:"स्ट्रॉबेरी की खेती",44:"संतरे की खेती",
    45:"शहतूत की खेती",46:"डब्बावाला",47:"शिवनेरी किला",48:"लोकल ट्रेन",
    49:"मीनाक्षी मंदिर",50:"तिरुवल्लुवर प्रतिमा",51:"स्वामी विवेकानंद स्मारक",
    52:"मसालों की प्रदर्शनी",53:"छाया आकृति स्थापना",54:"हम्पी के पत्थर",
    55:"केरल की नौका दौड़",56:"वंदे मातरम् की स्थापना",57:"परेड की झाँकी",
    58:"पर्यावरण",59:"समय की ज्वालाएँ",60:"पंचतत्त्व",61:"गिफ्ट सिटी",
    62:"छाया स्थापना",63:"समय की धारा",64:"अनुग्रह का वृक्ष",65:"नारी उदय",
    66:"सपनों के पंख",67:"ज्ञान की सीढ़ियाँ",68:"प्रकाश की पंखुड़ियाँ",
    69:"नारी नेतृत्व",70:"आज की नारी",71:"पोषण और विकास",72:"आईटीबीपी",
    73:"केंद्रीय रिजर्व पुलिस बल",74:"केंद्रीय सशस्त्र पुलिस बल",
    75:"कर्तव्य पर पुलिस",76:"सीमा सुरक्षा बल",77:"सीमा सुरक्षा बल",
    78:"केंद्रीय औद्योगिक सुरक्षा बल",79:"भारतीय तटरक्षक बल",
    80:"भारतीय वायु सेना",81:"प्रगति के हाथ",82:"तकनीकी प्रस्फुटन",
    83:"माइक्रोचिप क्रांति",84:"बुद्धिमत्ता की रूपरेखा",85:"अनंत संपर्क",
    86:"उभरता क्षितिज",87:"उत्तरायण",88:"सर्फिंग स्टेशन",89:"पवित्र कमल",
    90:"ब्लॉसम हॉर्स",91:"पंजाब का गौरव",92:"कुआँ",93:"नालंदा विश्वविद्यालय",
    94:"ऊँट",95:"कठपुतली की दीवार",96:"हवा महल",97:"काजीरंगा के गैंडे",
    98:"जीवित जड़ पुल",99:"बंगाली टच",100:"रानी की वाव",101:"अग्नि नर्तक",
    102:"गोवा का माहौल",103:"नासिक ढोल",104:"मराठा",105:"ऊँट गाड़ी",
    106:"मरीन ड्राइव",107:"राम सेतु",108:"दक्षिण का हाथी",109:"सोने की चिड़िया",
    110:"समय का चक्र",111:"सात घोड़े",114:"हाथी",115:"दीप धारा",
    116:"टिमटिमाहट",117:"प्रकाश की लहर",118:"हॉट एयर बैलून सुरंग",
    119:"टनल ऑफ लैम्प्स",120:"टनल ऑफ थ्रेड्स",121:"फ़ाइटर जेट",
    122:"अंतरिक्ष में भारत",123:"अंतरिक्ष और प्रौद्योगिकी",124:"मेक इन इंडिया",
    125:"योग आसन",126:"सूर्य नमस्कार",127:"मत्स्यासन",128:"शीर्षासन",
    129:"सर्वांगासन",130:"पश्चिमोत्तानासन",131:"हलासन",132:"भुजंगासन",
    133:"अर्ध मत्स्येन्द्रासन",134:"पादहस्तासन",135:"काकासन",136:"धनुरासन",
    137:"शलभासन",138:"त्रिकोणासन",139:"मकड़ी का जाला",140:"एनचांटेड फॉरेस्ट",
    141:"टनल ऑफ डीप वाटर",142:"एक्वेटिक टनल",143:"एलवीएम3 रॉकेट",
    144:"पीएसएलवी रॉकेट",145:"गगनयान अंतरिक्ष यात्री",146:"मंगलयान मिशन",
    147:"आदित्य-एल1 सौर मिशन",148:"अंतरिक्ष द्वार",149:"विक्रम लैंडर",
    150:"प्रज्ञान रोवर",151:"उपग्रह संचार एंटीना",152:"कृत्रिम उपग्रह",
    153:"अंतरिक्ष में भारत शीर्षक पट्ट",
}

# =============================================================
# ZONE & PLACE & EFFECT
# =============================================================
ZONE = {}
for gid in range(1,24):   ZONE[gid]='North'
for gid in range(24,38):  ZONE[gid]='East'
for gid in range(38,49):  ZONE[gid]='West'
for gid in range(49,56):  ZONE[gid]='South'
for gid in range(56,89):  ZONE[gid]='Outdoor'
for gid in [89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108]: ZONE[gid]='Indoor'
for gid in [109,110,111,114,115,116,117,118,119,120,121,122,123,124]: ZONE[gid]='Outdoor'
for gid in range(125,154): ZONE[gid]='Outdoor'

PLACE_EN = {
    1:"Kashmir",2:"Kashmir",3:"Kashmir",4:"Dal Lake, Kashmir",5:"Trans-Himalayan Region",
    6:"Ladakh",7:"Himachal Pradesh",8:"Kinnaur, Himachal Pradesh",9:"Punjab",10:"Punjab",
    11:"New Delhi",12:"Punjab",13:"Punjab",14:"Sarnath, Uttar Pradesh",
    15:"Varanasi, Uttar Pradesh",16:"Ayodhya, Uttar Pradesh",17:"Braj, Uttar Pradesh",
    18:"Nathdwara, Rajasthan",19:"Jhansi, Uttar Pradesh",20:"North India",
    21:"Rajasthan",22:"Rajasthan",23:"Rajasthan",24:"Assam",25:"Assam",
    26:"Majuli, Assam",27:"Majuli, Assam",28:"Assam",29:"Northeast India",
    30:"Northeast India",31:"Northeast India",32:"Eastern India",33:"Odisha",
    34:"Konark, Odisha",35:"Odisha",36:"Puri, Odisha",37:"Kolkata, West Bengal",
    38:"Gujarat",39:"Gujarat",40:"Ahmedabad, Gujarat",41:"West India",
    42:"Pune, Maharashtra",43:"Maharashtra",44:"Maharashtra",45:"Maharashtra",
    46:"Mumbai, Maharashtra",47:"Junnar, Maharashtra",48:"Mumbai, Maharashtra",
    49:"Madurai, Tamil Nadu",50:"Kanyakumari, Tamil Nadu",51:"Kanyakumari, Tamil Nadu",
    52:"South India",53:"South India",54:"Hampi, Karnataka",55:"Kerala",
}
PLACE_HI = {
    1:"कश्मीर",2:"कश्मीर",3:"कश्मीर",4:"डल झील, कश्मीर",5:"ट्रांस-हिमालयी क्षेत्र",
    6:"लद्दाख",7:"हिमाचल प्रदेश",8:"किन्नौर, हिमाचल प्रदेश",9:"पंजाब",10:"पंजाब",
    11:"नई दिल्ली",12:"पंजाब",13:"पंजाब",14:"सारनाथ, उत्तर प्रदेश",
    15:"वाराणसी, उत्तर प्रदेश",16:"अयोध्या, उत्तर प्रदेश",17:"ब्रज, उत्तर प्रदेश",
    18:"नाथद्वारा, राजस्थान",19:"झाँसी, उत्तर प्रदेश",20:"उत्तर भारत",
    21:"राजस्थान",22:"राजस्थान",23:"राजस्थान",24:"असम",25:"असम",
    26:"माजुली, असम",27:"माजुली, असम",28:"असम",29:"पूर्वोत्तर भारत",
    30:"पूर्वोत्तर भारत",31:"पूर्वोत्तर भारत",32:"पूर्वी भारत",33:"ओडिशा",
    34:"कोणार्क, ओडिशा",35:"ओडिशा",36:"पुरी, ओडिशा",37:"कोलकाता, पश्चिम बंगाल",
    38:"गुजरात",39:"गुजरात",40:"अहमदाबाद, गुजरात",41:"पश्चिम भारत",
    42:"पुणे, महाराष्ट्र",43:"महाराष्ट्र",44:"महाराष्ट्र",45:"महाराष्ट्र",
    46:"मुंबई, महाराष्ट्र",47:"जुन्नर, महाराष्ट्र",48:"मुंबई, महाराष्ट्र",
    49:"मदुरई, तमिलनाडु",50:"कन्याकुमारी, तमिलनाडु",51:"कन्याकुमारी, तमिलनाडु",
    52:"दक्षिण भारत",53:"दक्षिण भारत",54:"हम्पी, कर्नाटक",55:"केरल",
}
EPP_PLACE_EN = "Ekta Prakash Parv 2026, Ekta Nagar"
EPP_PLACE_HI = "एकता प्रकाश पर्व 2026, एकता नगर"

# Additional place overrides for 56+ exhibits
PLACE_EN.update({
    61:"Gandhinagar, Gujarat",87:"Gujarat",88:"Goa, India",
    97:"Kaziranga, Assam",98:"Meghalaya",99:"West Bengal",
    100:"Patan, Gujarat",102:"Goa",103:"Nashik, Maharashtra",
    104:"Maharashtra",106:"Mumbai, Maharashtra",107:"Tamil Nadu",
    108:"South India",111:"India",121:"India",122:"India",
    123:"India",124:"India",125:"India",126:"India",
    143:"India",144:"India",145:"India",146:"India",
    147:"India",148:"India",149:"India",150:"India",
    151:"India",152:"India",153:"India",
})
PLACE_HI.update({
    61:"गांधीनगर, गुजरात",87:"गुजरात",88:"गोवा, भारत",
    97:"काजीरंगा, असम",98:"मेघालय",99:"पश्चिम बंगाल",
    100:"पाटन, गुजरात",102:"गोवा",103:"नासिक, महाराष्ट्र",
    104:"महाराष्ट्र",106:"मुंबई, महाराष्ट्र",107:"तमिलनाडु",
    108:"दक्षिण भारत",111:"भारत",121:"भारत",122:"भारत",
    123:"भारत",124:"भारत",125:"भारत",126:"भारत",
    143:"भारत",144:"भारत",145:"भारत",146:"भारत",
    147:"भारत",148:"भारत",149:"भारत",150:"भारत",
    151:"भारत",152:"भारत",153:"भारत",
})

EFFECT = {}
# snow for Kashmir/Himalayan
for gid in [1,2,3,4,5,6,7,8]: EFFECT[gid]='snow'
# petals for flower/dance/celebration themes
for gid in [9,11,12,17,23,38,58,64,65,68,71,89,90]: EFFECT[gid]='petals'
# sparks for rest (default)

# =============================================================
# REAL IMAGE URLs scraped from GIPL pages (IDs 87-153)
# Pattern: /assets/img/EktaPrakashParvDetails/{date}/{TITLE}_{uuid}.ext
# =============================================================
_B = "https://rashtriyaektadiwas.com"
GIPL_IMAGE = {
    87:  _B+"/assets/img/EktaPrakashParvDetails/17102026/KITE FESTIVAL_d11599fb-2fa8-401c-9d33-7053c4c33c26.png",
    88:  _B+"/assets/img/EktaPrakashParvDetails/17102026/SURFING STATION_018c626b-21c5-4773-9dc8-00d2a711aa6d.png",
    89:  _B+"/assets/img/EktaPrakashParvDetails/28102026/SACRED LOTUS_79372139-9c96-4620-8a71-04fa2d6fdf1b.jpg",
    90:  _B+"/assets/img/EktaPrakashParvDetails/28102026/BLOSSOM HORSE_e1b41ca9-438d-4d1a-8b36-b0f70854b9d9.png",
    91:  _B+"/assets/img/EktaPrakashParvDetails/28102026/PUNJAB DI SHAAN_6c4e3daf-bc08-4fca-b860-3559fdfa5285.png",
    92:  _B+"/assets/img/EktaPrakashParvDetails/28102026/WELL_2acde5d1-e775-4a62-8f14-d6cbda2dfe27.jpg",
    93:  _B+"/assets/img/EktaPrakashParvDetails/28102026/NALANDA UNIVERSITY_9a1f89eb-831b-4590-baf5-43888be93890.png",
    94:  _B+"/assets/img/EktaPrakashParvDetails/28102026/CAMELS_5b33d4e3-1287-4268-a72c-d3e5c1222ad3.png",
    95:  _B+"/assets/img/EktaPrakashParvDetails/28102026/WALL OF KATHPUTLI _7de4fb4d-7253-444f-96de-47446705df15.jpg",
    96:  _B+"/assets/img/EktaPrakashParvDetails/28102026/HAWA MAHAL_0c94a679-f153-4434-987f-c61a9ae2066b.jpg",
    97:  _B+"/assets/img/EktaPrakashParvDetails/28102026/KAZIRANGA'S RHINOCEROS_65c92954-b6ae-49a6-a686-f1e13f368079.png",
    98:  _B+"/assets/img/EktaPrakashParvDetails/28102026/LIVING ROOT BRIDGE_12ce350e-ad4b-41ef-90f7-526b838959bc.png",
    99:  _B+"/assets/img/EktaPrakashParvDetails/28102026/BENGALI TOUCH_66da6c8f-705f-4475-943c-4e4422408fad.jpg",
    100: _B+"/assets/img/EktaPrakashParvDetails/28102026/RANI KI VAV_a5014d97-63cf-4013-905e-de51755190bd.jpg",
    101: _B+"/assets/img/EktaPrakashParvDetails/28102026/FIRE DANCER_59240197-bfd6-4546-97d8-cfcb9e75d249.jpg",
    102: _B+"/assets/img/EktaPrakashParvDetails/28102026/VIBE OF GOA_d077149e-4668-473d-87d7-ef7d780b451b.png",
    103: _B+"/assets/img/EktaPrakashParvDetails/28102026/NASHIK DHOL_6dc660fd-84fd-447e-9a62-0c17865bee64.jpg",
    104: _B+"/assets/img/EktaPrakashParvDetails/28102026/MARATHAS_35058057-311f-46be-8fe2-dac098a53b02.png",
    105: _B+"/assets/img/EktaPrakashParvDetails/28102026/CAMEL CART_c239b90e-2fe8-4550-8c01-050f3e64b778.jpg",
    106: _B+"/assets/img/EktaPrakashParvDetails/28102026/MARINE DRIVE_423ced6c-6517-485f-8fa5-ad8be7f808c5.png",
    107: _B+"/assets/img/EktaPrakashParvDetails/28102026/RAM SETU_0e87a4ab-3ee1-47d9-a0fe-319f6b8832aa.jpg",
    108: _B+"/assets/img/EktaPrakashParvDetails/28102026/ELEPHANT OF SOUTH_780fddb5-da45-4433-ada3-9a1c5242b8c3.png",
    109: _B+"/assets/img/EktaPrakashParvDetails/28102026/GOLDEN BIRD_b4b80cb9-379a-4009-9a2e-366831fbca6b.jpg",
    110: _B+"/assets/img/EktaPrakashParvDetails/28102026/SAMAY KA CHAKRA_880bdbde-fd11-4545-b149-518710df030f.jpg",
    111: _B+"/assets/img/EktaPrakashParvDetails/28102026/SEVEN HORSES_3c19d415-7de8-4ddc-838b-b5979a90fe1d.jpg",
    114: _B+"/assets/img/EktaPrakashParvDetails/28102026/ELEPHANT_fb6e4e2c-5aad-425e-82bb-2a7565991195.jpg",
    115: _B+"/assets/img/EktaPrakashParvDetails/28102026/DEEP DHARA_ae0003e7-e56e-4ff9-86de-e02a7e1c7d0d.jpg",
    116: _B+"/assets/img/EktaPrakashParvDetails/28102026/FLICKER AND FLAME_441df884-ff49-4214-9ca4-fa92e5c556ed.jpg",
    117: _B+"/assets/img/EktaPrakashParvDetails/28102026/WAVE OF LIGHT_f9b93cf5-52a2-4be5-b689-9ff67080bb22.jpg",
    118: _B+"/assets/img/EktaPrakashParvDetails/28102026/TUNNEL OF HOT AIR BALLOON_6332b566-3da2-46f6-9c6d-28522180b797.jpg",
    119: _B+"/assets/img/EktaPrakashParvDetails/28102026/TUNNEL OF LAMPS_a96cddd6-cfda-44fc-87ce-69bd1e8836fd.jpg",
    120: _B+"/assets/img/EktaPrakashParvDetails/28102026/TUNNEL OF THREADS_23630e1f-a458-4bdb-8d18-ce95017d0858.jpg",
    121: _B+"/assets/img/EktaPrakashParvDetails/28102026/FIGHTER JET_1367a5ea-3669-44ac-9a19-5064696ee97d.jpg",
    122: _B+"/assets/img/EktaPrakashParvDetails/28102026/INDIA IN SPACE_6d9953a1-6d5d-41e4-bba3-6c2cc4ab6c99.png",
    123: _B+"/assets/img/EktaPrakashParvDetails/28102026/SPACE AND TECHNOLOGY_7789de00-8f7c-4c07-a8c3-fd77c231c1a2.jpg",
    124: _B+"/assets/img/EktaPrakashParvDetails/28102026/MAKE IN INDIA_fa15c21e-a8f1-49d6-a054-a59334246f69.jpg",
    125: _B+"/assets/img/EktaPrakashParvDetails/28102026/YOGA AASAN_d8724ff8-cc36-44b5-b9bf-812cc45b4e3f.jpg",
    126: _B+"/assets/img/EktaPrakashParvDetails/28102026/SURYA NAMASKAR_eb90e044-b65a-4e4e-ba40-56e2b19cbe7c.jpg",
    127: _B+"/assets/img/EktaPrakashParvDetails/28102026/MATSYASANA_593e4a37-cdc8-402b-a8e7-b477fbbd808a.png",
    128: _B+"/assets/img/EktaPrakashParvDetails/28102026/SIRSHASANA_2ca50cf4-ad2e-41c9-a36f-5f9adc5b0361.png",
    129: _B+"/assets/img/EktaPrakashParvDetails/28102026/SARVANGASANA_2f154809-0ccb-4e09-842f-20a71d9b75e4.png",
    130: _B+"/assets/img/EktaPrakashParvDetails/28102026/PASCHIMOTTANASANA_2768f686-fcd5-44b3-909a-c856ac826362.png",
    131: _B+"/assets/img/EktaPrakashParvDetails/28102026/HALASANA_4afd3dc4-16a8-4369-9eed-02b321433d80.png",
    132: _B+"/assets/img/EktaPrakashParvDetails/28102026/BHUJANGASANA_8217ec12-b042-4538-9528-a1a71162fc7a.png",
    133: _B+"/assets/img/EktaPrakashParvDetails/28102026/ARDHA MATSYENDRASANA_b35b227f-ecd7-4e16-a4e0-b5bbef5d300b.png",
    134: _B+"/assets/img/EktaPrakashParvDetails/28102026/PADA HASTASANA_4d84a4ba-7654-4a76-959c-3b3e3833a197.png",
    135: _B+"/assets/img/EktaPrakashParvDetails/28102026/KAKASANA_6c86c52b-9e5f-4b0c-a0ec-9ec42af0b99c.png",
    136: _B+"/assets/img/EktaPrakashParvDetails/28102026/DHANURASANA_db9e1e32-19fc-4384-bab3-1843ff10f70c.png",
    137: _B+"/assets/img/EktaPrakashParvDetails/28102026/SALABHASANA_73dcf857-6283-4cb2-96fd-ca0ada8599f7.png",
    138: _B+"/assets/img/EktaPrakashParvDetails/28102026/TRIKONASANA_f90b5893-cfad-45be-9616-30b4c41f1684.png",
    139: _B+"/assets/img/EktaPrakashParvDetails/28102026/SPIDER WEB_cae71cf7-a069-4e0e-bccc-7b43ecfa222c.jpg",
    140: _B+"/assets/img/EktaPrakashParvDetails/28102026/ENCHANTED FOREST_626f336b-0dfe-4411-bfa5-01e3253377ef.png",
    141: _B+"/assets/img/EktaPrakashParvDetails/28102026/TUNNEL OF DEEP WATER_7b2e2483-5bea-491f-bc51-13143f5528c6.jpg",
    142: _B+"/assets/img/EktaPrakashParvDetails/28102026/AQUATIC TUNNEL_71b1b506-637f-401d-9fa1-b7d810179239.jpg",
    143: _B+"/assets/img/EktaPrakashParvDetails/28102026/LVM3 ROCKET_f2c228e2-7f35-4452-bc93-60daec5eef11.png",
    144: _B+"/assets/img/EktaPrakashParvDetails/28102026/PSLV ROCKET_e10ca07d-9726-44d6-a4b3-1ce8cb50cf69.png",
    145: _B+"/assets/img/EktaPrakashParvDetails/28102026/GAGANYAAN ASTRONAUT_65bdd9bc-fafb-4e11-8cbc-902cfc2a5a86.png",
    146: _B+"/assets/img/EktaPrakashParvDetails/28102026/MARS ORBITER MISSION_17287563-9c01-438d-804d-035ec74d41e7.png",
    147: _B+"/assets/img/EktaPrakashParvDetails/28102026/ADITYA-L1 SOLAR MISSION_b3b3a06e-1b82-4b52-8df9-553bf5c7d4c6.png",
    148: _B+"/assets/img/EktaPrakashParvDetails/28102026/SPACE PORTAL_8ecf3ac3-8759-43a2-885f-00c317b018a5.png",
    149: _B+"/assets/img/EktaPrakashParvDetails/28102026/VIKRAM LANDER_2e3e7739-5f01-4994-b9ad-191b7c265f3a.png",
    150: _B+"/assets/img/EktaPrakashParvDetails/28102026/PRAGYAN ROVER_54cc3ccf-d98a-497d-a629-783e1a0bc236.png",
    151: _B+"/assets/img/EktaPrakashParvDetails/28102026/SATELLITE COMMUNICATION DISH_dffcae78-a147-4624-ad6c-fa9c041f0025.png",
    152: _B+"/assets/img/EktaPrakashParvDetails/28102026/ARTIFICIAL SATELLITE_a42d1c3c-0ba5-454e-87d8-7a7eebd9af8f.png",
    153: _B+"/assets/img/EktaPrakashParvDetails/28102026/INDIA IN SPACE TITLE PANEL_6778772e-17a1-453c-9b07-f0e358e64eed.png",
}

# =============================================================
# HIGHLIGHTS BY CATEGORY (4 items: icon_key, en, hi)
# =============================================================
# icon keys: craft, dance, temple, nature, heritage, military, space, yoga, art, festival
HIGHLIGHTS = {
    1: [("craft","Hand-knotted","हाथ से गाँठ"),("craft","Wool & Silk","ऊन और रेशम"),("art","Chinar Motifs","चिनार रूपांकन"),("craft","Handloom Heritage","हथकरघा परंपरा")],
    2: [("craft","Hook Needle","हुक सुई"),("craft","Chain Stitch","चेन स्टिच"),("art","Floral Design","पुष्प आकृति"),("heritage","Kashmir Craft","कश्मीरी शिल्प")],
    3: [("craft","Wood Carving","लकड़ी नक्काशी"),("art","Geometric Lattice","ज्यामितीय जाली"),("heritage","Kashmir Heritage","कश्मीरी विरासत"),("craft","Pinjrakari Art","पिंजराकारी कला")],
    4: [("nature","Dal Lake","डल झील"),("craft","Carved Woodwork","नक्काशीदार लकड़ी"),("nature","Floating Homes","तैरते घर"),("heritage","Kashmir Life","कश्मीरी जीवन")],
    5: [("nature","High Altitude","ऊँचाई पर रहने वाला"),("nature","Rocky Mountains","पथरीले पहाड़"),("nature","Endangered Species","संकटग्रस्त प्रजाति"),("nature","Snow Habitat","बर्फीला आवास")],
    6: [("heritage","Prayer Wheel","प्रार्थना चक्र"),("heritage","Buddhist Tradition","बौद्ध परंपरा"),("heritage","Om Mani Padme Hum","ॐ मणि पद्मे हूँ"),("heritage","Ladakh Culture","लद्दाख संस्कृति")],
    7: [("craft","Wool Weaving","ऊन बुनाई"),("craft","Geometric Borders","ज्यामितीय किनारे"),("heritage","Himachal Craft","हिमाचली शिल्प"),("craft","Mountain Warmth","पर्वतीय गर्माहट")],
    8: [("dance","Folk Dance","लोकनृत्य"),("dance","Coordinated Steps","समन्वित कदम"),("heritage","Kinnaur Culture","किन्नौर संस्कृति"),("craft","Woollen Attire","ऊनी परिधान")],
    9: [("craft","Flower Work","फूलों का काम"),("craft","Silk Thread","रेशमी धागा"),("heritage","Punjab Craft","पंजाब शिल्प"),("craft","Bridal Art","विवाह कला")],
    10:[("dance","Women's Dance","महिला नृत्य"),("dance","Boliyan Songs","बोलियाँ"),("heritage","Punjab Culture","पंजाब संस्कृति"),("dance","Festive Dance","उत्सवी नृत्य")],
    11:[("temple","Bahá'í Temple","बहाई मंदिर"),("art","Lotus Architecture","कमल वास्तुकला"),("heritage","New Delhi Icon","दिल्ली की पहचान"),("heritage","27 Marble Petals","27 संगमरमर पंखुड़ियाँ")],
    12:[("festival","Wheat Harvest","गेहूँ की कटाई"),("festival","Sikh Heritage","सिख विरासत"),("festival","Khalsa Foundation","खालसा स्थापना"),("festival","April Festival","अप्रैल उत्सव")],
    13:[("dance","Energetic Dance","ऊर्जावान नृत्य"),("dance","Dhol Beat","ढोल की थाप"),("heritage","Punjab Pride","पंजाब की शान"),("dance","Harvest Dance","कटाई नृत्य")],
    14:[("heritage","Lion Capital","सिंह शीर्ष"),("heritage","State Emblem","राजकीय प्रतीक"),("heritage","Satyameva Jayate","सत्यमेव जयते"),("heritage","Emperor Ashoka","सम्राट अशोक")],
    15:[("festival","Sacred River","पवित्र नदी"),("festival","Lamp Ceremony","दीप समारोह"),("heritage","Varanasi Ghat","वाराणसी घाट"),("festival","Devotional Music","भक्तिगीत")],
    16:[("temple","Ram Janmabhoomi","राम जन्मभूमि"),("temple","Ayodhya Mandir","अयोध्या मंदिर"),("heritage","Traditional Design","पारंपरिक डिज़ाइन"),("temple","Sacred Heritage","पवित्र विरासत")],
    17:[("festival","Lathmar Holi","लट्ठमार होली"),("festival","Colours & Joy","रंग और उल्लास"),("heritage","Radha Krishna","राधा कृष्ण"),("festival","Barsana Festival","बरसाना उत्सव")],
    18:[("art","Temple Paintings","मंदिर चित्रकारी"),("art","Nathdwara Art","नाथद्वारा कला"),("art","Krishna Themes","कृष्ण विषय"),("art","Cloth Painting","कपड़े पर चित्र")],
    19:[("heritage","Jhansi Fort","झाँसी किला"),("heritage","Rani Lakshmibai","रानी लक्ष्मीबाई"),("heritage","1857 Uprising","1857 विद्रोह"),("heritage","Rocky Hill Fort","पथरीले किले")],
    20:[("festival","Festival of Lights","प्रकाश पर्व"),("festival","Oil Lamps Diyas","मिट्टी के दीये"),("festival","Family Celebration","पारिवारिक उत्सव"),("festival","Lakshmi Puja","लक्ष्मी पूजा")],
    21:[("art","Storytelling Shrine","कथा मंदिर"),("art","Painted Panels","चित्रित पट्ट"),("heritage","Rajasthan Craft","राजस्थान शिल्प"),("craft","Bassi Woodwork","बस्सी लकड़ी")],
    22:[("dance","Snake-Inspired","साँप से प्रेरित"),("dance","Swirling Skirts","घूमते घाघरे"),("heritage","UNESCO Heritage","यूनेस्को विरासत"),("dance","Kalbelia Community","कालबेलिया समुदाय")],
    23:[("dance","Circular Dance","गोल नृत्य"),("dance","Colourful Ghagra","रंगीन घाघरा"),("heritage","Rajasthan Dance","राजस्थान नृत्य"),("dance","Wedding Dance","विवाह नृत्य")],
    24:[("craft","Cane Weaving","बेंत बुनाई"),("nature","Renewable Craft","नवीकरणीय शिल्प"),("heritage","Northeast India","पूर्वोत्तर भारत"),("craft","Bamboo Artistry","बाँस कला")],
    25:[("dance","Spring Festival","वसंत उत्सव"),("dance","Assamese Dance","असमिया नृत्य"),("heritage","Rongali Bihu","रोंगाली बिहू"),("dance","Dhol Pepa Music","ढोल-पेपा संगीत")],
    26:[("heritage","Satra Tradition","सत्र परंपरा"),("heritage","Devotional Music","भक्ति संगीत"),("craft","Khol Drum","खोल वाद्य"),("heritage","Sankardev Legacy","शंकरदेव विरासत")],
    27:[("art","Mask Making","मुखौटा निर्माण"),("art","Bhaona Theatre","भाओना नाट्य"),("heritage","Majuli Satra","माजुली सत्र"),("craft","Bamboo Clay Mask","बाँस-मिट्टी मुखौटा")],
    28:[("nature","Tea Gardens","चाय बागान"),("nature","Camellia Sinensis","कैमेलिया साइनेंसिस"),("heritage","Assam Tea","असम की चाय"),("craft","Tea Processing","चाय प्रसंस्करण")],
    29:[("heritage","Ceremonial Headgear","औपचारिक शिरोभूषण"),("heritage","Northeast Tribes","पूर्वोत्तर जनजातियाँ"),("craft","Cane & Feather","बेंत और पंख"),("heritage","Community Identity","सामुदायिक पहचान")],
    30:[("nature","Tree Symbol","वृक्ष प्रतीक"),("heritage","Ancient Symbol","प्राचीन प्रतीक"),("nature","Nature Connection","प्रकृति से जुड़ाव"),("heritage","Life & Growth","जीवन और विकास")],
    31:[("heritage","Stilt Houses","खंभों पर घर"),("craft","Bamboo Architecture","बाँस वास्तुकला"),("heritage","Tribal Living","जनजातीय जीवन"),("nature","Northeast Homes","पूर्वोत्तर घर")],
    32:[("dance","Folk Traditions","लोक परंपराएँ"),("dance","Gotipua & Chhau","गोटीपुआ और छऊ"),("heritage","Eastern Heritage","पूर्वी विरासत"),("dance","Community Dances","सामुदायिक नृत्य")],
    33:[("art","Patta Painting","पट्ट चित्रकारी"),("art","Natural Pigments","प्राकृतिक रंग"),("heritage","Raghurajpur Village","रघुराजपुर गाँव"),("art","Mythological Art","पौराणिक कला")],
    34:[("temple","Sun Temple","सूर्य मंदिर"),("temple","Stone Chariot","पत्थर रथ"),("heritage","UNESCO World Heritage","यूनेस्को विश्व विरासत"),("art","Kalinga Architecture","कलिंग वास्तुकला")],
    35:[("art","Sacred Geometry","पवित्र ज्यामिति"),("heritage","Hindu Buddhist Jain","हिंदू बौद्ध जैन"),("art","Concentric Patterns","संकेंद्रित पैटर्न"),("heritage","Spiritual Symbol","आध्यात्मिक प्रतीक")],
    36:[("festival","Jagannath Temple","जगन्नाथ मंदिर"),("festival","Chariot Festival","रथ उत्सव"),("heritage","Puri Odisha","पुरी, ओडिशा"),("festival","Devotion & Unity","भक्ति और एकता")],
    37:[("heritage","Cantilever Bridge","कैंटिलीवर पुल"),("heritage","Rabindra Setu","रवींद्र सेतु"),("heritage","Kolkata Icon","कोलकाता की शान"),("heritage","Steel Rivets","इस्पात रिवेट")],
    38:[("dance","Navratri Dance","नवरात्रि नृत्य"),("dance","Circular Formation","गोल घेरा"),("heritage","UNESCO Heritage","यूनेस्को विरासत"),("heritage","Gujarat Culture","गुजरात संस्कृति")],
    39:[("art","Mirror Work","शीशे का काम"),("art","Clay & Mud Art","मिट्टी कला"),("heritage","Gujarat Craft","गुजरात शिल्प"),("craft","Geometric Patterns","ज्यामितीय आकृतियाँ")],
    40:[("heritage","Gandhi Ashram","गांधी आश्रम"),("heritage","Freedom Movement","स्वतंत्रता आंदोलन"),("heritage","Sabarmati River","साबरमती नदी"),("heritage","National Heritage","राष्ट्रीय विरासत")],
    41:[("craft","Bandhani Craft","बंधनी शिल्प"),("craft","Ikat Weaving","इकत बुनाई"),("heritage","Textile Heritage","वस्त्र विरासत"),("craft","West India Fabrics","पश्चिम भारत के वस्त्र")],
    42:[("heritage","Maratha Empire","मराठा साम्राज्य"),("heritage","Peshwa Seat","पेशवा का केंद्र"),("heritage","Historic Fort","ऐतिहासिक वाड़ा"),("heritage","Pune Heritage","पुणे विरासत")],
    43:[("nature","Hill Farming","पहाड़ी खेती"),("nature","Maharashtra Farms","महाराष्ट्र के खेत"),("nature","Red Strawberries","लाल स्ट्रॉबेरी"),("heritage","Agricultural Craft","कृषि शिल्प")],
    44:[("nature","Orange Groves","संतरे के बाग"),("nature","Nagpur Oranges","नागपुर संतरा"),("nature","Citrus Farming","खट्टे फल खेती"),("heritage","Maharashtra Fruit","महाराष्ट्र फल")],
    45:[("nature","Mulberry Leaves","शहतूत की पत्तियाँ"),("craft","Silk Farming","रेशम खेती"),("nature","Sericulture","रेशम उद्योग"),("heritage","Maharashtra Craft","महाराष्ट्र शिल्प")],
    46:[("heritage","Tiffin Service","टिफिन सेवा"),("heritage","Mumbai Icon","मुंबई की पहचान"),("heritage","Zero Error Delivery","शून्य त्रुटि वितरण"),("heritage","Community System","सामुदायिक व्यवस्था")],
    47:[("heritage","Shivaji Birthplace","शिवाजी जन्मस्थान"),("heritage","Maratha Fort","मराठा किला"),("heritage","Pune District","पुणे जिला"),("heritage","Historic Fortification","ऐतिहासिक किलेबंदी")],
    48:[("heritage","Mumbai Lifeline","मुंबई की जीवनरेखा"),("heritage","Daily Commute","दैनिक यात्रा"),("heritage","Urban Transport","शहरी परिवहन"),("heritage","Railway Network","रेलवे नेटवर्क")],
    49:[("temple","Meenakshi Temple","मीनाक्षी मंदिर"),("temple","Dravidian Architecture","द्रविड़ वास्तुकला"),("heritage","Madurai Heritage","मदुरई विरासत"),("temple","Goddess Meenakshi","देवी मीनाक्षी")],
    50:[("heritage","Thiruvalluvar","तिरुवल्लुवर"),("heritage","133 Feet Statue","133 फीट प्रतिमा"),("heritage","Tamil Literature","तमिल साहित्य"),("heritage","Kanyakumari Landmark","कन्याकुमारी स्थल")],
    51:[("heritage","Vivekananda Memorial","विवेकानंद स्मारक"),("heritage","1970 Founded","1970 में स्थापित"),("heritage","Kanyakumari Cape","कन्याकुमारी अंतरीप"),("heritage","Island Rock","द्वीप पत्थर")],
    52:[("nature","Spice Trade","मसाला व्यापार"),("nature","Kerala Spices","केरल के मसाले"),("heritage","Ancient Routes","प्राचीन व्यापार मार्ग"),("nature","Aromatic Herbs","सुगंधित जड़ी-बूटियाँ")],
    53:[("art","Shadow Art","छाया कला"),("art","Silhouette Installation","छाया स्थापना"),("art","Light & Shadow","प्रकाश और छाया"),("heritage","Indian Silhouettes","भारतीय आकृतियाँ")],
    54:[("heritage","Vijayanagara Empire","विजयनगर साम्राज्य"),("heritage","Stone Carvings","पत्थर नक्काशी"),("heritage","UNESCO Site","यूनेस्को स्थल"),("heritage","Karnataka Heritage","कर्नाटक विरासत")],
    55:[("festival","Vallam Kali","वल्लम काली"),("heritage","Kerala Backwaters","केरल बैकवाटर"),("festival","Snake Boat Race","साँप नाव दौड़"),("heritage","Onam Festival","ओणम पर्व")],
}
# Default highlights for exhibits without specific ones
DEFAULT_HL = {
    'North': [("heritage","North India Heritage","उत्तर भारतीय विरासत"),("art","Cultural Tradition","सांस्कृतिक परंपरा"),("craft","Artisan Skill","कारीगर कौशल"),("heritage","Living Legacy","जीवंत विरासत")],
    'East':  [("heritage","East India Heritage","पूर्वी भारतीय विरासत"),("art","Cultural Craft","सांस्कृतिक शिल्प"),("heritage","Community Art","सामुदायिक कला"),("heritage","Living Legacy","जीवंत विरासत")],
    'West':  [("heritage","West India Heritage","पश्चिम भारतीय विरासत"),("craft","Traditional Craft","पारंपरिक शिल्प"),("heritage","Cultural Pride","सांस्कृतिक गर्व"),("heritage","Living Legacy","जीवंत विरासत")],
    'South': [("heritage","South India Heritage","दक्षिण भारतीय विरासत"),("temple","Temple Culture","मंदिर संस्कृति"),("heritage","Dravidian Arts","द्रविड़ कलाएँ"),("heritage","Living Legacy","जीवंत विरासत")],
    'Outdoor':[("art","Light Installation","प्रकाश स्थापना"),("heritage","Cultural Display","सांस्कृतिक प्रदर्श"),("art","Ekta Nagar Art","एकता नगर कला"),("heritage","Unity in Diversity","विविधता में एकता")],
    'Indoor':[("heritage","Cultural Exhibit","सांस्कृतिक प्रदर्श"),("art","Artistic Heritage","कलात्मक विरासत"),("heritage","India's Pride","भारत का गर्व"),("heritage","Unity in Diversity","विविधता में एकता")],
}

# SVG icons pool (compact inline SVGs, reused by key)
ICONS = {
    'craft':  '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M9 9c4 0 6 3 7 7s3 7 7 7M23 9c-4 0-6 3-7 7s-3 7-7 7"/><circle cx="16" cy="16" r="2.4"/></svg>',
    'art':    '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M8 24s2-8 8-14c4 3 6 10 6 14"/><path d="M11 20c2-4 4-8 5-10"/><circle cx="22" cy="10" r="2"/></svg>',
    'dance':  '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="16" cy="7" r="3"/><path d="M16 10v6l-5 5M16 16l5 5M12 15l-4 2M20 15l4 2"/></svg>',
    'temple': '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M6 26h20M8 26v-8h16v8M4 18h24M16 4l-6 6h12l-6-6z"/><path d="M13 26v-5h6v5"/></svg>',
    'nature': '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M16 28V16"/><path d="M16 16c0-6 8-10 8-10s-2 8-8 10"/><path d="M16 16c0-6-8-10-8-10s2 8 8 10"/></svg>',
    'heritage':'<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M16 4l2.5 6 5-3-1.5 6 6 1-5 4 3 5-6-1-1 6-2-5-2 5-1-6-6 1 3-5-5-4 6-1-1.5-6 5 3Z"/></svg>',
    'festival':'<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M16 4l1.5 4.5L22 7l-3.5 3 1.5 4.5-4-2.5-4 2.5 1.5-4.5L10 7l4.5 1.5Z"/><path d="M8 22l16 0M10 26h12"/><path d="M16 12v10"/></svg>',
    'military':'<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4l10 6v8l-10 10L6 18V10z"/><path d="M16 10v8M12 14h8"/></svg>',
    'space':  '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M16 28V8"/><path d="M10 14l6-10 6 10"/><path d="M12 20h8"/><circle cx="16" cy="24" r="2"/></svg>',
    'yoga':   '<svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><circle cx="16" cy="6" r="3"/><path d="M16 9v7M8 12l8 4 8-4M8 22h16"/><path d="M10 22v4M22 22v4"/></svg>',
}

# =============================================================
# BUILD DESCRIPTIONS FROM EXCEL
# =============================================================
DESC = {}  # key=gipl_id, val={en, hi}

def norm(s): return re.sub(r'\s+',' ',str(s or '').replace('\n',' ').replace('\r','')).strip()

def load_excel_desc(path, sheet_idx=0):
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[sheet_idx]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            if not row or len(row) < 5: continue
            c1,c2,c3,c4,c5 = [norm(c) for c in row[:5]]
            if c1.isdigit() and '/' in c2 and c4:
                yield int(c1), c2, c4, c5
    except Exception as e:
        print(f"  WARN {path}: {e}")

# Excel SR to GIPL ID mapping (Excel SR number != GIPL ID)
# We build by matching English title
TITLE_EN_MAP = {norm(ten).upper(): gid for _,gid,ten,_ in EXHIBITS_RAW}

def add_desc_from_excel(path, sheet_idx=0):
    for xsr, xname, en, hi in load_excel_desc(path, sheet_idx):
        en_title = xname.split('/')[0].strip().upper()
        # Try direct match
        gid = TITLE_EN_MAP.get(en_title)
        if gid and en:
            DESC[gid] = {'en': en, 'hi': hi}

print("Loading Excel descriptions...")
add_desc_from_excel(os.path.join(DATA_DIR, "NORTH ZONE English + Hindi.xlsx"))
print(f"  North Zone: {len(DESC)} loaded")
add_desc_from_excel(os.path.join(DATA_DIR, "EAST ZONE English + Hindi (1).xlsx"))
print(f"  East Zone: {len(DESC)} loaded")
add_desc_from_excel(os.path.join(DATA_DIR, "WEST ZONE English + Hindi.xlsx"))
print(f"  West Zone: {len(DESC)} loaded")
add_desc_from_excel(os.path.join(DATA_DIR, "SOUTH ZONE English + Hindi.xlsx"))
print(f"  South Zone: {len(DESC)} loaded")
add_desc_from_excel(os.path.join(DATA_DIR, "OUTDOOR English + Hindi.xlsx"))
print(f"  Outdoor: {len(DESC)} loaded")
add_desc_from_excel(os.path.join(DATA_DIR, "Remaining QRs.xlsx"), 0)
add_desc_from_excel(os.path.join(DATA_DIR, "Remaining QRs.xlsx"), 1)
print(f"  Remaining: {len(DESC)} loaded")
# India in Space: different structure (cols: no, image, name+hi, en, hi)
try:
    wb = openpyxl.load_workbook(os.path.join(DATA_DIR, "India_in_Space_11_Bilingual_Name_Tags_With_Images.xlsx"), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    for row in ws.iter_rows(values_only=True):
        if not row or len(row) < 5: continue
        c1,c2,c3,c4,c5 = [norm(c) for c in row[:5]]
        if c1.isdigit() and c4:
            xno = int(c1)
            gid = xno + 2  # India in Space NO = GIPL_ID - 2
            if gid in BY_GID: DESC[gid] = {'en': c4, 'hi': c5}
    print(f"  India in Space: {len(DESC)} loaded")
except Exception as e:
    print(f"  WARN India in Space: {e}")

# Fix duplicate-name exhibits: OUTDOOR has two "BSF" rows (SR=33 → gid=76, SR=34 → gid=77)
# The title map only kept gid=77, so gid=76 falls back to generic text. Reload explicitly.
_bsf_gids = [76, 77]
_bsf_idx = 0
for _xsr, _xname, _en, _hi in load_excel_desc(os.path.join(DATA_DIR, "OUTDOOR English + Hindi.xlsx")):
    if _xname.split('/')[0].strip().upper() == 'BSF' and _bsf_idx < len(_bsf_gids):
        DESC[_bsf_gids[_bsf_idx]] = {'en': _en, 'hi': _hi}
        _bsf_idx += 1

print(f"\nTotal descriptions loaded: {len(DESC)}")

# =============================================================
# IMAGE INDEX CSV
# =============================================================
CSV_PATH = os.path.join(BASE_DIR, "images_index.csv")
with open(CSV_PATH, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['Sr_No','GIPL_Id','Title_EN','Title_HI','Zone','Ambient_Effect','Image_URL','GIPL_URL','Has_Description'])
    for sr,gid,ten,act in EXHIBITS_RAW:
        if not act: continue
        if gid >= 87:
            img_url = GIPL_IMAGE.get(gid, "").replace(' ', '%20')
        else:
            csv_ext = 'jpg' if gid <= 55 else 'png'
            img_url = f"https://rashtriyaektadiwas.com/assets/img/EktaPrakashParvDetails/{gid}.{csv_ext}"
        gipl_url = f"https://rashtriyaektadiwas.com/EktaPrakashParv/EktaPrakashParvDetails?Id={gid}&qid=0"
        eff = EFFECT.get(gid,'sparks')
        zone = ZONE.get(gid,'Outdoor')
        hi = HINDI_NAME.get(gid,'')
        has_desc = 'YES' if gid in DESC else 'NO'
        w.writerow([f'{sr:03d}',gid,ten,hi,zone,eff,img_url,gipl_url,has_desc])
print(f"CSV saved: {CSV_PATH}")

# =============================================================
# HTML GENERATOR
# =============================================================
def hl_html(hl_list):
    parts = []
    for key,en,hi in hl_list:
        svg = ICONS.get(key, ICONS['heritage'])
        parts.append(f'<div><i>{svg}</i><span class="en">{en}</span><span class="hi">{hi}</span></div>')
    return '\n'.join(parts)

def make_html(sr, gid, ten, act):
    hi = HINDI_NAME.get(gid, ten.title())
    eff = EFFECT.get(gid,'sparks')
    zone = ZONE.get(gid,'Outdoor')
    pe = PLACE_EN.get(gid, "Gujarat, India")
    ph = PLACE_HI.get(gid, "गुजरात, भारत")
    place_en = pe
    place_hi = ph

    d = DESC.get(gid, {})
    desc_en = d.get('en','') or f"{ten.title()} is a treasured cultural exhibit at Ekta Prakash Parv 2026, celebrating India's rich heritage, traditions and unity."
    desc_hi = d.get('hi','') or f"{hi} एकता प्रकाश पर्व 2026 में प्रदर्शित एक विशेष सांस्कृतिक प्रदर्श है, जो भारत की समृद्ध विरासत, परंपराओं और एकता का उत्सव मनाता है।"

    # Highlights
    hl_list = HIGHLIGHTS.get(gid, DEFAULT_HL.get(zone, DEFAULT_HL['Outdoor']))
    highlights = hl_html(hl_list)

    if gid >= 87:
        img_url = GIPL_IMAGE.get(gid, "").replace(' ', '%20')
    else:
        img_ext = 'jpg' if gid <= 55 else 'png'
        img_url = f"https://rashtriyaektadiwas.com/assets/img/EktaPrakashParvDetails/{gid}.{img_ext}"

    # Next exhibit
    cur_idx = ACTIVE_SRS.index(sr) if sr in ACTIVE_SRS else -1
    if cur_idx >= 0 and cur_idx < len(ACTIVE_SRS)-1:
        next_sr = ACTIVE_SRS[cur_idx+1]
        next_tuple = BY_SR.get(next_sr)
        next_gid = next_tuple[1] if next_tuple else gid
        next_en = next_tuple[2].title() if next_tuple else ""
        next_hi = HINDI_NAME.get(next_gid,'')
        next_url = f"https://rashtriyaektadiwas.com/EktaPrakashParv/EktaPrakashParvDetails?Id={next_gid}&qid=0"
        next_block = f'''  <a class="next rv" href="{next_url}">
    <div class="grow"><small class="en">Next exhibit · № {next_sr:02d}</small><small class="hi">अगला प्रदर्श · № {next_sr:02d}</small>
      <b class="en">{next_en}</b><b class="hi">{next_hi}</b></div>
    <span class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
  </a>'''
    else:
        next_block = ''

    ten_title = ten.title()
    # Replace apostrophes in JS-safe meta
    meta_desc = f"{ten_title} ({hi}) — Ekta Prakash Parv 2026, Rashtriya Ekta Diwas, Ekta Nagar."

    html = f'''<!DOCTYPE html>
<html lang="en" data-theme="day" data-lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{ten_title} · Ekta Prakash Parv 2026</title>
<meta name="description" content="{meta_desc}">
<meta name="theme-color" content="#1a0f0a">
<link rel="icon" href="https://rashtriyaektadiwas.com/assets/img/favicon.png">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@500;600;700&family=Poppins:wght@400;500;600;700&family=Noto+Sans+Devanagari:wght@400;500;600;700&display=swap" rel="stylesheet">
<script>(function(){{var r=document.documentElement;try{{if(localStorage.getItem('epp-theme')==='night')r.setAttribute('data-theme','night');
var l=localStorage.getItem('epp-lang');if(!l){{l=/^(hi|gu|mr)/.test((navigator.language||'').toLowerCase())?'hi':'en';}}
r.setAttribute('data-lang',l);r.lang=l;}}catch(e){{}}}})()</script>
<style>
:root{{
  --sf:#ff6a2b; --sf-2:#ffa53b; --gold:#e9b45a; --gold-hi:#ffd98a;
  --bg:#fff8ee; --bg-2:#fdeedb; --surface:#ffffff; --line:rgba(169,67,22,.16);
  --ink:#2a1a10; --ink-2:#5a4636; --muted:#8c7462; --title:#6e2a10;
  --hero-fade:#fff8ee; --shadow:0 18px 50px -20px rgba(110,42,16,.35);
  --en:"Poppins",system-ui,sans-serif; --display:"Cormorant Garamond",Georgia,serif;
  --hi:"Noto Sans Devanagari","Poppins",sans-serif; --hi-display:"Noto Sans Devanagari","Poppins",sans-serif;
}}
[data-theme="night"]{{
  --bg:#08070f; --bg-2:#120b18; --surface:#14101d; --line:rgba(255,190,110,.16);
  --ink:#f6ecdf; --ink-2:#d6c8b8; --muted:#a8998a; --title:#fff1dc;
  --hero-fade:#08070f; --shadow:0 18px 60px -18px rgba(0,0,0,.8);
}}
*{{box-sizing:border-box}}
html{{-webkit-text-size-adjust:100%;scroll-behavior:smooth;overflow-x:hidden}}
body{{margin:0;background:var(--bg);color:var(--ink);font-family:var(--en);overflow-x:hidden;transition:background .4s,color .4s}}
img{{display:block;max-width:100%}}
a{{color:inherit}}
button{{font:inherit}}
[data-lang="en"] .hi{{display:none!important}}
[data-lang="hi"] .en{{display:none!important}}
.hi{{font-family:var(--hi)}}
#fx{{position:fixed;inset:0;width:100%;height:100%;pointer-events:none;z-index:40}}
.top{{position:fixed;top:0;left:0;right:0;z-index:30;display:flex;align-items:center;justify-content:space-between;gap:10px;
  padding:10px 14px;padding-top:max(10px,env(safe-area-inset-top));transition:background .3s,box-shadow .3s}}
.top.solid{{background:color-mix(in srgb,var(--bg) 88%,transparent);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);box-shadow:0 1px 0 var(--line)}}
.red-logo{{background:#fff;border-radius:12px;padding:4px 8px;box-shadow:0 4px 14px rgba(0,0,0,.25);display:block}}
.red-logo img{{height:40px;width:auto}}
.tools{{display:flex;gap:8px;align-items:center}}
.seg{{display:flex;border-radius:999px;overflow:hidden;background:rgba(255,255,255,.92);border:1px solid rgba(255,255,255,.6);box-shadow:0 4px 14px rgba(0,0,0,.2)}}
.seg button{{border:0;background:none;padding:7px 13px;font-weight:600;font-size:12.5px;color:#5a4636;cursor:pointer}}
.seg button.h{{font-family:var(--hi);font-size:13px}}
.seg button[aria-pressed="true"]{{background:linear-gradient(135deg,var(--sf),var(--sf-2));color:#fff}}
.round{{width:36px;height:36px;border-radius:50%;border:1px solid rgba(255,255,255,.6);background:rgba(255,255,255,.92);color:#6e2a10;
  display:grid;place-items:center;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.2);padding:0}}
.round svg{{width:18px;height:18px}}
.round .moon{{display:none}}[data-theme="night"] .round .moon{{display:block}}[data-theme="night"] .round .sun{{display:none}}
[data-theme="night"] .seg,[data-theme="night"] .round{{background:rgba(20,16,29,.85);border-color:rgba(255,200,130,.25)}}
[data-theme="night"] .seg button{{color:#e8d8c4}}[data-theme="night"] .round{{color:var(--gold-hi)}}
.hero{{position:relative;min-height:100svh;display:flex;flex-direction:column;justify-content:flex-end;overflow:hidden;isolation:isolate;background:#0b0705}}
.hero-img{{position:absolute;inset:0;z-index:-2}}
.hero-img img{{width:100%;height:100%;object-fit:cover;object-position:center 30%;transform:scale(1.12);animation:kb 22s ease-in-out infinite alternate}}
@keyframes kb{{to{{transform:scale(1.02) translateY(-2%)}}}}
.hero::before{{content:"";position:absolute;inset:0;z-index:-1;
  background:linear-gradient(180deg,rgba(8,5,3,.55) 0%,rgba(8,5,3,.05) 20%,rgba(8,5,3,.18) 40%,rgba(14,7,4,.86) 66%,rgba(14,7,4,.94) 100%)}}
.hero::after{{content:"";position:absolute;left:0;right:0;bottom:0;height:46px;z-index:0;background:linear-gradient(180deg,rgba(14,7,4,0),var(--hero-fade))}}
.lights{{position:absolute;top:64px;left:-2%;width:104%;height:70px;z-index:1;pointer-events:none}}
.lights path{{fill:none;stroke:rgba(255,220,160,.35);stroke-width:1.2}}
.lights circle{{animation:bulb 1.8s ease-in-out infinite;filter:drop-shadow(0 0 5px currentColor)}}
@keyframes bulb{{0%,100%{{opacity:.35}}50%{{opacity:1}}}}
.hero-in{{position:relative;z-index:1;padding:0 20px 52px;max-width:760px;width:100%;margin:0 auto}}
.epp-badge{{display:flex;align-items:center;gap:10px;margin-bottom:16px}}
.epp-badge svg{{width:46px;height:46px;flex:none;filter:drop-shadow(0 0 10px rgba(255,150,60,.7))}}
.epp-badge .t{{line-height:1.15;color:#ffe9c9}}
.epp-badge b{{display:block;font-family:var(--display);font-weight:700;font-size:17px;letter-spacing:2.2px}}
.epp-badge span{{display:block;font-size:10.5px;letter-spacing:1.6px;text-transform:uppercase;opacity:.85}}
.epp-badge span.hi{{letter-spacing:.3px;text-transform:none;font-size:12px}}
.chips{{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}}
.chip{{padding:5px 11px;border-radius:999px;font-size:11.5px;font-weight:600;letter-spacing:.4px;color:#ffe6c4;
  background:rgba(255,255,255,.1);border:1px solid rgba(255,214,150,.35);backdrop-filter:blur(6px)}}
.chip.n{{background:linear-gradient(135deg,var(--sf),var(--sf-2));border-color:transparent;color:#fff}}
h1{{margin:0;font-family:var(--display);font-weight:700;color:#fff6e8;line-height:.95;letter-spacing:.5px;
  font-size:clamp(46px,14vw,92px);text-transform:uppercase;text-shadow:0 2px 30px rgba(0,0,0,.45)}}
h1 .glow{{background:linear-gradient(100deg,#fff6e8 0%,#fff6e8 40%,#ffd27a 50%,#fff6e8 60%,#fff6e8 100%);background-size:250% 100%;
  -webkit-background-clip:text;background-clip:text;color:transparent;animation:shine 5s linear infinite}}
@keyframes shine{{from{{background-position:120% 0}}to{{background-position:-120% 0}}}}
h1.hi{{font-family:var(--hi-display);font-weight:700;text-transform:none;line-height:1.3;font-size:clamp(40px,11.5vw,78px)}}
.alt{{margin:10px 0 0;font-size:clamp(20px,5.4vw,28px);color:#ffd9a8;font-family:var(--hi-display)}}
.alt.en{{font-family:var(--display);font-style:italic;font-weight:600;letter-spacing:.5px}}
.rule{{width:56px;height:3px;border-radius:3px;margin:16px 0;background:linear-gradient(90deg,var(--sf),var(--gold-hi));box-shadow:0 0 12px rgba(255,140,50,.8)}}
.place{{margin:0;color:#f3dcc0;font-size:15px}}
.cta-row{{display:flex;gap:10px;margin-top:22px;flex-wrap:wrap}}
.cta{{flex:1 1 auto;display:inline-flex;align-items:center;justify-content:center;gap:10px;padding:15px 22px;border-radius:999px;border:0;cursor:pointer;
  text-decoration:none;color:#fff;font-weight:600;font-size:14.5px;letter-spacing:1.4px;text-transform:uppercase;
  background:linear-gradient(135deg,#ff5a1f,#ff9a3c);position:relative;overflow:hidden;
  box-shadow:0 0 0 1px rgba(255,200,140,.4),0 10px 30px -6px rgba(255,100,30,.75);animation:pulse 2.8s ease-in-out infinite}}
.cta.hi{{letter-spacing:0;text-transform:none;font-size:15.5px}}
.cta::after{{content:"";position:absolute;inset:0;background:linear-gradient(110deg,transparent 30%,rgba(255,255,255,.45) 50%,transparent 70%);transform:translateX(-120%);animation:sweep 3.5s ease-in-out infinite}}
@keyframes sweep{{60%,100%{{transform:translateX(120%)}}}}
@keyframes pulse{{0%,100%{{box-shadow:0 0 0 1px rgba(255,200,140,.4),0 10px 30px -6px rgba(255,100,30,.75)}}50%{{box-shadow:0 0 0 1px rgba(255,220,160,.7),0 10px 44px -2px rgba(255,120,40,.95)}}}}
.scroll{{display:flex;flex-direction:column;align-items:center;gap:2px;margin-top:22px;color:#f3dcc0;font-size:12px;letter-spacing:1px;text-decoration:none}}
.scroll svg{{width:20px;height:20px;animation:bob 2s ease-in-out infinite}}
@keyframes bob{{50%{{transform:translateY(5px)}}}}
.wrap{{max-width:760px;margin:0 auto;padding:0 18px}}
.hl{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:26px 0 8px}}
.hl div{{text-align:center}}
.hl i{{display:grid;place-items:center;width:62px;height:62px;margin:0 auto 8px;border-radius:50%;color:var(--sf);
  border:1.5px solid color-mix(in srgb,var(--gold) 70%,transparent);background:var(--surface);
  box-shadow:0 0 0 4px color-mix(in srgb,var(--gold) 12%,transparent),0 8px 22px -10px rgba(255,120,40,.6)}}
[data-theme="night"] .hl i{{color:var(--gold-hi);box-shadow:0 0 0 4px rgba(255,200,120,.06),0 0 24px -4px rgba(255,140,50,.45)}}
.hl svg{{width:30px;height:30px}}
.hl span{{display:block;font-size:12.5px;line-height:1.35;color:var(--ink-2);font-weight:500}}
.story{{padding:30px 0 10px}}
.kicker{{display:flex;align-items:center;gap:12px;margin:0 0 14px;color:var(--sf);font-weight:700;font-size:12px;letter-spacing:2.4px;text-transform:uppercase}}
.kicker.hi{{letter-spacing:.4px;font-size:14px;text-transform:none}}
.kicker::after{{content:"";flex:1;height:1px;background:linear-gradient(90deg,var(--gold),transparent)}}
.text p{{margin:0;font-size:17px;line-height:1.95;color:var(--ink-2)}}
.text p.en::first-letter{{font-family:var(--display);font-weight:700;font-size:4.2em;line-height:.8;float:left;padding:8px 10px 0 0;color:var(--sf)}}
.text p.hi{{font-size:18px;line-height:2.05}}
.frame{{position:relative;margin:34px auto 10px;border-radius:22px;padding:3px;overflow:hidden;max-width:520px;
  box-shadow:var(--shadow)}}
.frame::before{{content:"";position:absolute;inset:-60%;background:conic-gradient(from 0deg,transparent 0 60%,#ffcf7a 72%,#ff6a2b 80%,transparent 90%);animation:spin 5s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
.frame-in{{position:relative;border-radius:19px;overflow:hidden;background:#000}}
.frame button{{all:unset;display:block;cursor:zoom-in;width:100%}}
.frame img{{width:100%;height:auto}}
.zoom{{position:absolute;right:12px;bottom:12px;display:flex;align-items:center;gap:6px;padding:6px 12px;border-radius:999px;background:rgba(0,0,0,.6);color:#fff;font-size:12px;font-weight:600}}
.zoom svg{{width:14px;height:14px}}
.cap{{text-align:center;margin:12px 0 0;color:var(--muted);font-size:13px}}
.next{{display:flex;align-items:center;gap:14px;margin:34px 0 0;padding:16px 18px;border-radius:20px;text-decoration:none;
  background:var(--surface);border:1px solid var(--line);box-shadow:var(--shadow)}}
.next .ic{{flex:none;width:46px;height:46px;border-radius:50%;display:grid;place-items:center;color:#fff;background:linear-gradient(135deg,var(--sf),var(--sf-2));box-shadow:0 0 18px rgba(255,120,40,.5)}}
.next .ic svg{{width:20px;height:20px}}
.next small{{display:block;font-size:11px;letter-spacing:1.8px;text-transform:uppercase;color:var(--muted);font-weight:600}}
.next small.hi{{letter-spacing:0;text-transform:none;font-size:12.5px}}
.next b{{display:block;font-family:var(--display);font-size:22px;color:var(--title);letter-spacing:.5px;line-height:1.2}}
.next b.hi{{font-family:var(--hi-display);font-weight:600;font-size:19px}}
.next .grow{{flex:1;min-width:0}}
.all{{display:block;text-align:center;margin:16px 0 0;color:var(--sf);font-weight:600;font-size:14px;text-decoration:none;padding:8px}}
footer{{margin-top:44px;padding:34px 18px calc(30px + env(safe-area-inset-bottom));text-align:center;position:relative;
  background:linear-gradient(180deg,transparent,color-mix(in srgb,var(--sf) 8%,transparent))}}
footer .mark{{width:64px;height:64px;margin:0 auto 10px;filter:drop-shadow(0 0 14px rgba(255,140,50,.6))}}
footer h3{{margin:0;font-family:var(--display);font-size:24px;letter-spacing:3px;color:var(--title)}}
footer h3.hi{{font-family:var(--hi-display);font-weight:600;letter-spacing:0;font-size:22px}}
footer p{{margin:6px 0 0;color:var(--muted);font-size:13px;line-height:1.7}}
footer .motto{{margin-top:12px;color:var(--sf);font-family:var(--hi-display);font-size:17px}}
dialog{{border:0;padding:0;background:transparent;width:100vw;height:100vh;max-width:100vw;max-height:100vh}}
dialog::backdrop{{background:rgba(5,3,2,.95)}}
dialog .w{{width:100%;height:100%;display:flex;align-items:center;justify-content:center;padding:14px}}
dialog img{{max-width:100%;max-height:100%;object-fit:contain;border-radius:8px}}
dialog button{{all:unset;position:absolute;top:16px;right:16px;width:42px;height:42px;border-radius:50%;display:grid;place-items:center;background:rgba(255,255,255,.16);color:#fff;font-size:26px;cursor:pointer}}
.rv{{opacity:0;transform:translateY(22px);transition:opacity .8s ease,transform .8s ease}}.rv.in{{opacity:1;transform:none}}
@media (min-width:900px){{.hero-in{{padding-bottom:60px}}.hl i{{width:72px;height:72px}}}}
@media (max-width:380px){{.red-logo img{{height:34px}}.seg button{{padding:6px 10px}}.hl i{{width:54px;height:54px}}.hl svg{{width:26px;height:26px}}.hl span{{font-size:11.5px}}}}
@media (prefers-reduced-motion:reduce){{*,*::before,*::after{{animation:none!important;transition:none!important}}.rv{{opacity:1;transform:none}}#fx{{display:none}}}}
</style>
</head>
<body data-effect="{eff}">
<svg width="0" height="0" style="position:absolute" aria-hidden="true">
  <defs>
    <radialGradient id="g-core" cx="50%" cy="60%" r="60%"><stop offset="0" stop-color="#fff7d6"/><stop offset=".45" stop-color="#ffc24a"/><stop offset="1" stop-color="#ff5a1f"/></radialGradient>
    <linearGradient id="g-ring" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#ffd98a"/><stop offset="1" stop-color="#ff6a2b"/></linearGradient>
    <symbol id="epp-mark" viewBox="0 0 100 100">
      <circle cx="50" cy="50" r="46" fill="none" stroke="url(#g-ring)" stroke-width="2.2"/>
      <circle cx="50" cy="50" r="39" fill="none" stroke="url(#g-ring)" stroke-width=".8" stroke-dasharray="1.5 4"/>
      <g fill="#ffd98a">
        <circle cx="50" cy="4" r="3"/><circle cx="82.5" cy="17.5" r="3"/><circle cx="96" cy="50" r="3"/><circle cx="82.5" cy="82.5" r="3"/>
        <circle cx="50" cy="96" r="3"/><circle cx="17.5" cy="82.5" r="3"/><circle cx="4" cy="50" r="3"/><circle cx="17.5" cy="17.5" r="3"/>
      </g>
      <path d="M50 18C58 30 64 38 64 48C64 57 57.5 63 50 63C42.5 63 36 57 36 48C36 38 42 30 50 18Z" fill="url(#g-core)"/>
      <path d="M50 34C54 40 56 44 56 49C56 53.5 53.3 56.5 50 56.5C46.7 56.5 44 53.5 44 49C44 44 46 40 50 34Z" fill="#fff8e0" opacity=".9"/>
      <path d="M26 66H74C72 77 62 83 50 83C38 83 28 77 26 66Z" fill="url(#g-ring)"/>
      <path d="M31 71C38 74 62 74 69 71" stroke="#7a2a0e" stroke-width="1.2" fill="none" opacity=".55"/>
    </symbol>
  </defs>
</svg>
<canvas id="fx" aria-hidden="true"></canvas>
<header class="top" id="top">
  <a class="red-logo" href="https://rashtriyaektadiwas.com/" aria-label="Rashtriya Ekta Diwas 2026"><img src="https://rashtriyaektadiwas.com/assets/img/logo.png" alt="Statue of Unity · Rashtriya Ekta Diwas 2026"></a>
  <div class="tools">
    <div class="seg" role="group" aria-label="Language">
      <button type="button" data-l="en" aria-pressed="true">English</button>
      <button type="button" class="h" data-l="hi" aria-pressed="false">हिंदी</button>
    </div>
    <button type="button" class="round" id="themeBtn" aria-label="Day / Night">
      <svg class="sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4.5"/><path d="M12 2v2.5M12 19.5V22M2 12h2.5M19.5 12H22M4.9 4.9l1.8 1.8M17.3 17.3l1.8 1.8M4.9 19.1l1.8-1.8M17.3 6.7l1.8-1.8"/></svg>
      <svg class="moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5Z"/></svg>
    </button>
  </div>
</header>
<section class="hero">
  <div class="hero-img"><img src="{img_url}" alt=""></div>
  <svg class="lights" id="lights" viewBox="0 0 400 70" preserveAspectRatio="none" aria-hidden="true"></svg>
  <div class="hero-in">
    <div class="epp-badge rv">
      <svg><use href="#epp-mark"/></svg>
      <div class="t"><b>EKTA PRAKASH PARV</b><span class="en">Rashtriya Ekta Diwas 2026</span><span class="hi">राष्ट्रीय एकता दिवस 2026</span></div>
    </div>
    <h1 class="en rv"><span class="glow">{ten_title}</span></h1>
    <h1 class="hi rv" lang="hi"><span class="glow">{hi}</span></h1>
    <p class="alt en rv" lang="hi" style="font-family:var(--hi-display);font-style:normal">{hi}</p>
    <p class="alt hi rv" style="font-family:var(--display);font-style:italic">{ten_title}</p>
    <div class="rule rv"></div>
    <p class="place rv"><span class="en">{place_en}</span><span class="hi">{place_hi}</span></p>
    <div class="cta-row rv">
      <a class="cta" href="#story"><span class="en">Explore the story</span><span class="hi">कहानी पढ़ें</span>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></a>
    </div>
    <a class="scroll" href="#story"><span class="en">Scroll to discover</span><span class="hi">आगे देखें</span>
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg></a>
  </div>
</section>
<main class="wrap">
  <div class="hl rv">
    {highlights}
  </div>
  <article class="story" id="story">
    <p class="kicker en rv">About the exhibit</p>
    <p class="kicker hi rv">प्रदर्श के बारे में</p>
    <div class="text rv" id="desc">
      <p class="en">{desc_en}</p>
      <p class="hi" lang="hi">{desc_hi}</p>
    </div>
    <figure class="frame rv" style="margin-left:auto;margin-right:auto">
      <div class="frame-in">
        <button type="button" id="zoomBtn" aria-label="Enlarge image / चित्र बड़ा करें">
          <img src="{img_url}" loading="lazy" alt="{ten_title} — {hi}">
          <span class="zoom"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4M11 8v6M8 11h6"/></svg><span class="en">Tap to zoom</span><span class="hi">बड़ा देखें</span></span>
        </button>
      </div>
    </figure>
    <p class="cap"><span class="en">The installation at Ekta Prakash Parv 2026</span><span class="hi">एकता प्रकाश पर्व 2026 में प्रदर्शित स्थापना</span></p>
  </article>
{next_block}
  <a class="all" href="https://rashtriyaektadiwas.com/Home/EktaPrakashParv2026"><span class="en">View all exhibits</span><span class="hi">सभी प्रदर्श देखें</span></a>
</main>
<footer>
  <svg class="mark"><use href="#epp-mark"/></svg>
  <h3 class="en">EKTA PRAKASH PARV 2026</h3><h3 class="hi">एकता प्रकाश पर्व 2026</h3>
  <p class="en">Rashtriya Ekta Diwas · Statue of Unity, Ekta Nagar, Gujarat<br>17 Oct 2026 – 15 Jan 2027 · 6:00 PM – 10:30 PM · Near Valley of Flowers</p>
  <p class="hi">राष्ट्रीय एकता दिवस · स्टैच्यू ऑफ यूनिटी, एकता नगर, गुजरात<br>17 अक्टूबर 2026 – 15 जनवरी 2027 · शाम 6:00 – रात 10:30 · वैली ऑफ फ्लावर्स के पास</p>
  <p class="motto">एक भारत, श्रेष्ठ भारत</p>
</footer>
<dialog id="lb"><div class="w"><img alt=""></div><button type="button" aria-label="Close">×</button></dialog>
<script>
(function(){{
  var R=document.documentElement, reduce=matchMedia('(prefers-reduced-motion: reduce)').matches;
  function put(k,v){{try{{localStorage.setItem(k,v);}}catch(e){{}}}}
  var lb=document.querySelectorAll('[data-l]');
  function setLang(l){{R.setAttribute('data-lang',l);R.lang=l;lb.forEach(function(b){{b.setAttribute('aria-pressed',b.dataset.l===l);}});put('epp-lang',l);}}
  lb.forEach(function(b){{b.onclick=function(){{setLang(b.dataset.l);}};}}); setLang(R.getAttribute('data-lang')||'en');
  document.getElementById('themeBtn').onclick=function(){{var t=R.getAttribute('data-theme')==='night'?'day':'night';R.setAttribute('data-theme',t);put('epp-theme',t);}};
  var top=document.getElementById('top');
  addEventListener('scroll',function(){{top.classList.toggle('solid',scrollY>innerHeight*.7);}},{{passive:true}});
  var L=document.getElementById('lights'),cols=['#ffd98a','#ff6a2b','#fff3d0','#ffb347','#ff8a3c'],h='<path d="M0 10 Q100 60 200 18 T400 12"/>';
  for(var i=0;i<24;i++){{var t=i/23,x=t*400,y=t<.5?10+(1-Math.pow((t-.25)/.25,2))*30+t*16:18+(1-Math.pow((t-.75)/.25,2))*22-(t-.5)*12;
    h+='<circle cx="'+x.toFixed(1)+'" cy="'+(y+4).toFixed(1)+'" r="3.2" fill="'+cols[i%5]+'" color="'+cols[i%5]+'" style="animation-delay:'+(i%6*.3).toFixed(1)+'s"/>';}}
  L.innerHTML=h;
  var c=document.getElementById('fx'),x=c.getContext('2d'),W,H,P=[],eff=document.body.dataset.effect||'sparks';
  function size(){{var d=Math.min(devicePixelRatio||1,2);W=c.width=innerWidth*d;H=c.height=innerHeight*d;x.setTransform(1,0,0,1,0,0);}}
  function make(){{var n=eff==='snow'?(innerWidth<600?70:120):(eff==='petals'?50:40);P=[];for(var i=0;i<n;i++)P.push(spawn(true));}}
  function spawn(any){{var d=Math.min(devicePixelRatio||1,2);return{{x:Math.random()*W,y:any?Math.random()*H:-10*d,r:(Math.random()*2.4+0.8)*d,vy:(Math.random()*.6+.35)*d,vx:(Math.random()-.5)*.3*d,ph:Math.random()*6.28,a:Math.random()*.5+.45}};}}
  function frame(){{x.clearRect(0,0,W,H);var night=R.getAttribute('data-theme')==='night';x.globalAlpha=scrollY>innerHeight*.6?.3:1;
    for(var i=0;i<P.length;i++){{var p=P[i];p.ph+=.012;p.y+=p.vy;p.x+=p.vx+Math.sin(p.ph)*.35;
      if(p.y>H+10||p.x<-20||p.x>W+20){{P[i]=spawn(false);continue;}}
      x.beginPath();x.arc(p.x,p.y,p.r,0,6.283);
      if(eff==='snow'){{x.fillStyle=night?'rgba(255,255,255,'+p.a+')':'rgba(255,255,255,'+(p.a*.95)+')';x.shadowColor=night?'rgba(200,225,255,.9)':'rgba(120,90,60,.35)';x.shadowBlur=night?8:3;}}
      else if(eff==='petals'){{x.fillStyle='rgba(255,'+(180+Math.random()*60|0)+','+(100+Math.random()*80|0)+','+p.a+')';x.shadowColor='rgba(255,100,180,.7)';x.shadowBlur=8;}}
      else{{x.fillStyle='rgba(255,'+(170+Math.random()*60|0)+',90,'+p.a+')';x.shadowColor='rgba(255,140,50,.9)';x.shadowBlur=10;}}
      x.fill();}}
    requestAnimationFrame(frame);}}
  if(!reduce){{size();make();frame();addEventListener('resize',function(){{size();make();}});}}
  var d=document.getElementById('lb'),di=d.querySelector('img'),zb=document.getElementById('zoomBtn');
  zb.onclick=function(){{di.src=zb.querySelector('img').src;if(d.showModal)d.showModal();}};
  d.onclick=function(e){{if(e.target!==di)d.close();}};
  var els=document.querySelectorAll('.rv');
  if('IntersectionObserver' in window&&!reduce){{var io=new IntersectionObserver(function(es){{es.forEach(function(e){{if(e.isIntersecting){{e.target.classList.add('in');io.unobserve(e.target);}}}});}},{{threshold:.08}});els.forEach(function(e){{io.observe(e);}});}}
  else els.forEach(function(e){{e.classList.add('in');}});
}})();
</script>
</body>
</html>'''
    return html

# =============================================================
# GENERATE ALL HTML FILES
# =============================================================
print("\nGenerating HTML pages...")
generated = 0
skipped = 0
for sr, gid, ten, act in EXHIBITS_RAW:
    if not act:
        skipped += 1
        continue
    slug = re.sub(r"[^A-Z0-9]+", "_", ten.upper()).strip("_")
    fname = f"EPP2026_{sr:03d}_{slug}.html"
    fpath = os.path.join(HTML_DIR, fname)
    html = make_html(sr, gid, ten, act)
    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(html)
    generated += 1
    if generated % 20 == 0:
        print(f"  Generated {generated} pages...")

print(f"\nDONE! Generated: {generated} | Skipped (inactive): {skipped}")
print(f"HTML pages: {HTML_DIR}")
print(f"CSV index:  {CSV_PATH}")
print(f"\nFolder ready for Google Drive upload:")
print(f"  {BASE_DIR}")
