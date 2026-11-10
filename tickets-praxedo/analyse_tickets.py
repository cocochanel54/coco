# Analyse des tickets Praxedo non planifiés : priorités, tickets proches, doublons et incohérences.
#
#   python3 analyse_tickets.py export_interventions.csv [dossier_sortie]
#
# Entrée : l'export CSV « Interventions » de Praxedo (séparateur « ; », UTF-8).
# Sortie : Analyse_tickets_AAAA-MM-JJ.xlsx et Carte_tickets_AAAA-MM-JJ.html dans le dossier de sortie.
# Les adresses (rue, code postal, ville, rien d'autre) sont placées sur la carte avec le service
# d'adresses officiel de l'IGN (data.geopf.fr). Le résultat est gardé dans geocodage_cache.json.
import csv, json, math, os, re, sys, time, unicodedata, urllib.parse, urllib.request
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SRC = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(SRC))
NOW = datetime.now(ZoneInfo("Europe/Paris")).replace(tzinfo=None, second=0, microsecond=0)
TODAY = datetime(NOW.year, NOW.month, NOW.day)

PROCHE_M = 300          # deux tickets à moins de 300 m : « tout proches »
TOURNEE_M = 800         # rayon d'une tournée de tickets à planifier ensemble
TOURNEE_MAX = 8         # tickets maxi par tournée (une journée)
CASER_M = 1000          # distance maxi pour caser un ticket avec une intervention déjà planifiée
AGE_VIEUX_J = 60        # ticket créé il y a plus de 60 jours et toujours pas planifié
PM_ALERTE = 3           # nombre de tickets ouverts sur un même PM qui fait penser à un défaut au PM
NOUVEAU_H = 24          # un ticket à planifier reste « nouveau » (violet sur la carte) pendant 24 h après sa création

# Délais de GTR comptés depuis la création du ticket. L'échéance « À faire avant » de Praxedo n'est pas fiable
# (souvent 3 mois) : c'est ce calcul qui sert pour les retards et les priorités.
#   [4H HO] / [8H HO] dans la description : 4 ou 8 heures ouvrées
#   « OTR 5J OUVRES » écrit dans la description : ce nombre de jours ouvrés
#   Niveau 1 : 4 heures (24 h/24) · Niveau 2 hors RDV client : 1 jour ouvré · Niveau 3 : 5 jours ouvrés · Niveau 4 : 10 jours ouvrés
#   Niveau 2 avec RDV client : pas de délai fixe, il dépend du RDV
GTR_NIVEAU = {1: ("heures", 4), 2: ("jours ouvrés", 1), 3: ("jours ouvrés", 5), 4: ("jours ouvrés", 10)}
HO_DEBUT, HO_FIN = 8, 18  # heures ouvrées : de 8 h à 18 h, du lundi au vendredi hors jours fériés

# Retard justifié / ticket en attente : une note Praxedo (ou la description) dit pourquoi on ne peut pas intervenir.
# Ces tickets passent en gris sur la carte et ne comptent plus comme « en retard ». Liste à compléter au besoin.
RE_JUSTIF = re.compile(r"""(
    GTR\s+justifi\w*[^\n.]{0,60}                                                   # « GTR justifiée car poteau »
  | justifi(?:é|e|ée|ee|cation)\b[^\n.]{0,60}                                      # « retard justifié … »
  | (?:poteau|pot\b\.?|appui)[^\n.]{0,25}?\b(?:à|a)\s+(?:remplacer|changer|signaler|poser|redresser)\b[^\n.]{0,40}
  | (?:en\s+)?attente\s+(?:d[eu]\s+|de\s+la\s+|d['’]\s*)?(?:orange|enedis|poteau|nacelle|autorisation|arr[êe]t[ée]|mairie
      |[ée]lagage|mat[ée]riel|accord|syndic|gestionnaire)\b[^\n.]{0,40}            # « en attente d'Orange », « attente nacelle »
  | signalisation[^\n.]{0,20}orange[^\n.]{0,30}                                    # « signalisation en cours vers Orange »
  | sig\s+faite[^\n.]{0,50}
)""", re.I | re.X)
# Colonnes de notes de l'export, reconnues à leur nom (ajoute la colonne des notes à ton filtre Praxedo)
RE_COL_NOTES = re.compile(r"note|comment|remarque|observ|justif", re.I)


# ---------- lecture ----------

def txt(s):
    return (s or "").replace("\r\n", "\n").replace("\n\r", "\n").replace("\r", "\n").strip()


def dt(s):
    s = (s or "").strip()
    for f in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().upper()
    s = re.sub(r"\b(RUE|R)\b", "RUE", s)
    s = re.sub(r"\b(AVENUE|AV|AVE)\b", "AVENUE", s)
    s = re.sub(r"\b(BOULEVARD|BD|BLD)\b", "BOULEVARD", s)
    s = re.sub(r"\b(ST)\b", "SAINT", s)
    s = re.sub(r"\b(STE)\b", "SAINTE", s)
    return re.sub(r"[^A-Z0-9]+", " ", s).strip()


rows = list(csv.DictReader(open(SRC, encoding="utf-8-sig", newline=""), delimiter=";"))
COLS = [c for c in (rows[0].keys() if rows else []) if c]
COLS_NOTES = [c for c in COLS if RE_COL_NOTES.search(c) and c != "Description"]

RE_CLS = re.compile(r"\b[A-Z]{2}\d{3}-[A-Z]{5}-\d{3}\b")
RE_PTO = re.compile(r"\b(?:TN-[A-Z0-9]{3,5}-[A-Z0-9]{4,5}|GE-\d{4}-[A-Z0-9]{4})\b")
RE_BOITE = re.compile(r"\b(?:GNY-(?:BPI|BPE)-[A-Z0-9]{4,6}(?:-\d{2})?|(?:PBO|BPE|BPI|PEC|PEO)-\d{2}-\d{3}-\d{3}-[A-Z0-9]{4}(?:-\d)?)\b", re.I)
RE_REF = re.compile(r"(?:BB-)?\b(20\d{6}T\d{4})\b", re.I)
RE_IMMEUBLE = re.compile(r"^(?:GNY-SITE-|IMB-|TN-BAT-)", re.I)
RE_PHONE = re.compile(r"(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}")


def feries(annee, departement):
    a, b, c = annee % 19, annee // 100, annee % 100
    d, e = b // 4, b % 4
    g = (8 * b + 13) // 25
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 19 * l) // 433
    mois = (h + l - 7 * m + 90) // 25
    paques = date(annee, mois, (h + l - 7 * m + 33 * mois + 19) % 32)
    j = {date(annee, 1, 1): "Jour de l'an", paques + timedelta(1): "Lundi de Pâques", date(annee, 5, 1): "1er mai",
         date(annee, 5, 8): "8 mai", paques + timedelta(39): "Ascension", paques + timedelta(50): "Lundi de Pentecôte",
         date(annee, 7, 14): "14 juillet", date(annee, 8, 15): "15 août", date(annee, 11, 1): "Toussaint",
         date(annee, 11, 11): "11 novembre", date(annee, 12, 25): "Noël"}
    if departement in ("57", "67", "68"):
        j[paques - timedelta(2)] = "Vendredi saint (Alsace-Moselle)"
        j[date(annee, 12, 26)] = "Saint-Étienne (Alsace-Moselle)"
    return j


def ouvre(d, dept):
    return d.weekday() < 5 and d.date() not in feries(d.year, dept)


def jours_ouvres(depart, n, dept=""):
    d = depart
    while n > 0:
        d += timedelta(1)
        if ouvre(d, dept):
            n -= 1
    return d


def heures_ouvrees(depart, n, dept=""):
    d, reste = depart, timedelta(hours=n)
    while True:
        if not ouvre(d, dept) or d.hour >= HO_FIN:
            d = (d + timedelta(1)).replace(hour=HO_DEBUT, minute=0)
            continue
        if d.hour < HO_DEBUT:
            d = d.replace(hour=HO_DEBUT, minute=0)
        fin = d.replace(hour=HO_FIN, minute=0)
        if d + reste <= fin:
            return d + reste
        reste -= fin - d
        d = (d + timedelta(1)).replace(hour=HO_DEBUT, minute=0)


def regle_gtr(t):
    """(libellé de la règle, échéance calculée ou None)"""
    c = t["creee"]
    for tag in t["tags"]:
        m = re.fullmatch(r"(\d+)\s*H\s*HO", tag)
        if m:
            n = int(m.group(1))
            return f"[{tag}] : {n} h ouvrées", heures_ouvrees(c, n, t["dept"]) if c else None
    m = re.search(r"OTR\s*(\d+)\s*J\s*OUVR", t["desc"], re.I)
    if m:
        n = int(m.group(1))
        return f"GTR écrite « OTR {n}J » : {n} j ouvrés", jours_ouvres(c, n, t["dept"]) if c else None
    if t["niveau"] == 2 and t["rdv"]:
        return "Niveau 2 avec RDV client : selon le RDV", None
    if t["niveau"] in GTR_NIVEAU:
        unite, n = GTR_NIVEAU[t["niveau"]]
        hors = " hors RDV" if t["niveau"] == 2 else ""
        if unite == "heures":
            return f"Niveau {t['niveau']} : {n} h", c + timedelta(hours=n) if c else None
        return f"Niveau {t['niveau']}{hors} : {n} j ouvré{'s' if n > 1 else ''}", jours_ouvres(c, n, t["dept"]) if c else None
    return "Sans niveau : échéance Praxedo", t["avant"]


def fretard(delta):
    h = int(delta.total_seconds() // 3600)
    return f"{h} h" if h < 24 else f"{h // 24} j"


T = []
for r in rows:
    num = (r.get("N°") or "").strip()
    if not num:
        continue
    desc = txt(r.get("Description"))
    tete = desc[:250]
    tags = [t.strip().upper() for t in re.findall(r"\[([^\]]+)\]", tete)]
    niv = re.search(r"niveau\s*(\d)", tete, re.I)
    contact = (r.get("Contact") or "").strip()
    equip = (r.get("Equipement") or "").strip()
    nom = " ".join(x for x in ((r.get("Nom technicien") or "").strip(), (r.get("Prénom technicien") or "").strip()) if x)
    base = num.split("#")[0].upper()
    info = re.search(r"Information du client\s*:\s*([A-Z0-9-]+)\s*/\s*([^\n]*)", desc, re.I)
    notes = " · ".join(txt(r.get(c)) for c in COLS_NOTES if txt(r.get(c)))
    mj = RE_JUSTIF.search(notes) or RE_JUSTIF.search(desc)
    justif = (("Note : " if RE_JUSTIF.search(notes) else "Description : ") + "« " + " ".join(mj.group(0).split()).rstrip(" ,;:-") + " »") if mj else ""
    T.append(dict(
        num=num, base=base, iteration=int(num.split("#")[1]) if "#" in num and num.split("#")[1].isdigit() else 0,
        statut=(r.get("Statut") or "").strip(), type=(r.get("Type inter.") or "").strip(),
        drapeaux=(r.get("Drapeaux") or "").strip(), planif=dt(r.get("Planifiée")), tech=nom,
        client=(r.get("Client") or "").strip(), site=(r.get("Site") or "").strip(), equip=equip,
        creee=dt(r.get("Créée")), avant=dt(r.get("À faire avant")),
        adresse=(r.get("Adresse") or "").strip(), cp=(r.get("CP") or "").strip(), ville=(r.get("Ville site") or "").strip(),
        contact=contact, immeuble=contact.upper() if RE_IMMEUBLE.match(contact) else "",
        desc=desc, tags=tags, niveau=int(niv.group(1)) if niv else None,
        rdv="RDVCLI" in tags, nacelle=bool(re.search(r"nacelle", desc, re.I)),
        cls=sorted(set(RE_CLS.findall(desc + " " + (r.get("Client") or "")))),
        pto=sorted(set(RE_PTO.findall(desc))),
        boites=sorted({b.upper() for b in RE_BOITE.findall(desc + " " + equip)}),
        refs=sorted({"BB-" + x.upper() for x in RE_REF.findall(desc)} - {base}),
        info_site=info.group(1).upper() if info else "", info_adr=info.group(2).strip() if info else "",
        notes=notes, justif=justif,
    ))

for t in T:
    t["ouvert_non_planifie"] = t["statut"] == "Qualifiée"
    t["a_planifier"] = t["statut"] in ("Qualifiée", "Pré-planifiée")
    t["deplanifie"] = t["statut"] == "Qualifiée" and t["planif"] is not None
    t["dept"] = t["cp"][:2] if t["cp"] else (t["site"].split("-")[1] if re.match(r"^(SRO|PMU|NRO)-\d{2}", t["site"]) else "54")
    t["regle"], t["echeance"] = regle_gtr(t)
    t["ech_gtr"] = not t["regle"].startswith("Sans niveau")


# ---------- géocodage (rue + CP + ville uniquement) ----------

cache_path = os.path.join(OUT, "geocodage_cache.json")
cache = json.load(open(cache_path, encoding="utf-8")) if os.path.exists(cache_path) else {}


def geocoder(params):
    url = "https://data.geopf.fr/geocodage/search?" + urllib.parse.urlencode(params)
    for essai in range(4):
        try:
            with urllib.request.urlopen(url, timeout=20) as resp:
                return json.load(resp).get("features") or []
        except Exception as e:
            erreur = e
            time.sleep(2 ** essai)
    raise erreur


echecs = 0
for t in T:
    q = " ".join(x for x in (t["adresse"], t["cp"], t["ville"]) if x)
    t["geo"] = None
    if not q:
        continue
    if q not in cache and echecs < 3:
        try:
            f = geocoder({"q": q, "limit": 1, **({"postcode": t["cp"]} if t["cp"] else {})})
            if not f and t["cp"]:  # code postal faux : on réessaie sans
                f = geocoder({"q": q, "limit": 1})
            cache[q] = {"lon": f[0]["geometry"]["coordinates"][0], "lat": f[0]["geometry"]["coordinates"][1],
                        **{k: f[0]["properties"].get(k) for k in ("label", "score", "type", "city", "postcode")}} if f else None
            echecs = 0
            time.sleep(0.03)
        except Exception as e:  # service injoignable : l'analyse continue, ces adresses ne sont pas sur la carte
            print("Géocodage impossible pour", q, ":", e)
            echecs += 1
    t["geo"] = cache.get(q, False) or None
    t["geo_inconnu"] = q not in cache
json.dump(cache, open(cache_path, "w", encoding="utf-8"), ensure_ascii=False, indent=0)


def dist(a, b):
    if not a["geo"] or not b["geo"]:
        return None
    la1, lo1, la2, lo2 = map(math.radians, (a["geo"]["lat"], a["geo"]["lon"], b["geo"]["lat"], b["geo"]["lon"]))
    h = math.sin((la2 - la1) / 2) ** 2 + math.cos(la1) * math.cos(la2) * math.sin((lo2 - lo1) / 2) ** 2
    return 6371000 * 2 * math.asin(math.sqrt(h))


def fdist(m):
    return f"{m:.0f} m" if m < 1000 else f"{m / 1000:.1f} km".replace(".", ",")


def fd(d, heure=True):
    return d.strftime("%d/%m/%Y %H:%M" if heure else "%d/%m/%Y") if d else ""


# ---------- incohérences ----------

INC = []  # (gravité, ticket, problème, détail)


def inc(g, t, probleme, detail=""):
    INC.append((g, t, probleme, detail))
    t.setdefault("alertes", []).append(probleme)


for t in T:
    t.setdefault("alertes", [])
    ouvert = t["statut"] in ("Qualifiée", "Pré-planifiée", "Planifiée", "Synchronisée")
    e = t["echeance"]
    if ouvert and e and e < NOW and t["justif"]:
        inc("4 - Info", t, f"Retard justifié ({fretard(NOW - e)})", f"{t['justif']} · échéance GTR {fd(e)}")
    elif ouvert and t["justif"]:
        t["alertes"].append("Justifié / en attente")
    elif ouvert and e and e < NOW:
        g = "1 - Urgent" if t["statut"] == "Qualifiée" else "2 - Important"
        inc(g, t, f"Échéance dépassée de {fretard(NOW - e)}",
            f"{t['regle']} · à faire avant le {fd(e)} · statut {t['statut']}" + (f" · planifiée le {fd(t['planif'])}" if t["planif"] and not t["deplanifie"] else ""))
    elif t["a_planifier"] and e and e < NOW + timedelta(days=2):
        inc("1 - Urgent", t, "Échéance dans moins de 48 h et pas planifié", f"{t['regle']} · à faire avant le {fd(e)}")
    if t["ech_gtr"] and e and t["avant"] and abs((t["avant"] - e).total_seconds()) > 2 * 86400:
        inc("3 - À vérifier", t, "Échéance Praxedo différente de la GTR",
            f"{t['regle']} → {fd(e)} · Praxedo : {fd(t['avant'])}")
    if ouvert and not e and not t["rdv"]:
        inc("3 - À vérifier", t, "Pas d'échéance (ni niveau, ni « À faire avant »)")
    if t["deplanifie"]:
        inc("2 - Important", t, "Déplanifié : revenu en « Qualifiée »",
            f"Était planifié le {fd(t['planif'])}, n'a plus de technicien" + (" · déjà en retard à cette date" if t["echeance"] and t["planif"] > t["echeance"] else ""))
    if t["statut"] in ("Planifiée", "Synchronisée") and t["planif"]:
        if t["planif"] < TODAY:
            inc("2 - Important", t, "Date planifiée passée, intervention pas clôturée",
                f"Prévue le {fd(t['planif'])} ({t['tech']}), toujours « {t['statut']} »")
        if t["echeance"] and t["planif"] > t["echeance"]:
            inc("2 - Important", t, "Planifiée après l'échéance",
                f"Prévue le {fd(t['planif'])}, à faire avant le {fd(t['echeance'])} ({t['regle']})")
    if t["statut"] == "Pré-planifiée" and t["planif"] and t["echeance"] and t["planif"] > t["echeance"]:
        inc("2 - Important", t, "Pré-planifiée après l'échéance", f"Prévue le {fd(t['planif'])}, à faire avant le {fd(t['echeance'])} ({t['regle']})")
    if t["planif"] and not t["deplanifie"]:
        jf = feries(t["planif"].year, t["dept"]).get(t["planif"].date())
        if jf:
            inc("3 - À vérifier", t, f"Planifiée un jour férié ({jf})", f"{fd(t['planif'])} · {t['tech']}")
        elif t["planif"].weekday() >= 5:
            inc("3 - À vérifier", t, "Planifiée un week-end", f"{fd(t['planif'])} · {t['tech']}")
    if t["statut"] in ("Planifiée", "Synchronisée", "Pré-planifiée") and not t["tech"]:
        inc("2 - Important", t, "Planifiée sans technicien")
    if t["a_planifier"] and t["creee"] and (NOW - t["creee"]).days > AGE_VIEUX_J:
        inc("2 - Important", t, f"Vieux ticket : créé il y a {(NOW - t['creee']).days} j", f"Créé le {fd(t['creee'], False)}")
    if not t["cp"]:
        inc("3 - À vérifier", t, "Code postal vide", "Le code postal est écrit dans la colonne Adresse" if re.search(r"\b\d{5}\b", t["adresse"]) else "")
    if t["geo"] is None and not t.get("geo_inconnu") and (t["adresse"] or t["ville"]):
        inc("3 - À vérifier", t, "Adresse introuvable", f"{t['adresse']} {t['cp']} {t['ville']}")
    elif t["geo"]:
        g = t["geo"]
        if norm(g.get("city")) != norm(t["ville"]) and norm(t["ville"]) not in norm(g.get("city")) and norm(g.get("city")) not in norm(t["ville"]):
            inc("3 - À vérifier", t, "Adresse trouvée dans une autre commune", f"Praxedo : {t['adresse']} {t['cp']} {t['ville']} · trouvée : {g.get('label')}")
        elif t["cp"] and g.get("postcode") and g["postcode"] != t["cp"]:
            inc("3 - À vérifier", t, "Code postal qui ne correspond pas à la ville", f"Praxedo : {t['cp']} {t['ville']} · officiel : {g['postcode']}")
        elif (g.get("score") or 0) < 0.5 or g.get("type") == "municipality":
            inc("3 - À vérifier", t, "Adresse imprécise (placée au centre de la commune)", f"{t['adresse']} {t['cp']} {t['ville']}")
    if re.search(r"L'incident concerne\s*:\s*undefined|Localisation la plus proche communiquée\s*:\s*0\b", t["desc"], re.I):
        inc("3 - À vérifier", t, "Localisation vague dans la description", "« undefined » ou numéro « 0 » : repérer le lieu exact avant d'envoyer quelqu'un")
    if t["info_site"] and t["immeuble"] and t["info_site"] != t["immeuble"] and t["info_site"] != t["site"].upper():
        inc("3 - À vérifier", t, "Site de la description ≠ champ Contact", f"Description : {t['info_site']} · Contact : {t['immeuble']}")
    if t["client"] and RE_CLS.fullmatch(t["client"]):
        inc("4 - Info", t, "Champ Client rempli avec la référence client au lieu de la commune", t["client"])
    if t["iteration"] >= 2:
        inc("4 - Info", t, f"Ticket rouvert {t['iteration']} fois (#{t['iteration']})")
    if t["rdv"] and t["a_planifier"]:
        t["alertes"].append("RDV client à prendre")
    if t["nacelle"]:
        t["alertes"].append("Nacelle")

# Même technicien, même créneau, ou deux interventions trop éloignées pour l'écart de temps
par_tech = defaultdict(list)
for t in T:
    if t["tech"] and t["planif"] and not t["deplanifie"]:
        par_tech[t["tech"]].append(t)
for tech, L in par_tech.items():
    L.sort(key=lambda x: x["planif"])
    for a, b in zip(L, L[1:]):
        ecart = (b["planif"] - a["planif"]).total_seconds() / 60
        if ecart == 0:
            inc("2 - Important", b, "Même technicien, même créneau", f"{tech} : {a['num']} et {b['num']} le {fd(a['planif'])}")
        else:
            d = dist(a, b)
            if d and a["planif"].date() == b["planif"].date() and ecart <= 60 and d > 20000:
                inc("3 - À vérifier", b, "Trajet serré entre deux interventions",
                    f"{tech} : {a['num']} ({a['ville']}) à {a['planif']:%H:%M} puis {b['num']} ({b['ville']}) à {b['planif']:%H:%M}, {fdist(d)} à vol d'oiseau")


# ---------- tickets liés (même immeuble, même boîte, suite de...) ----------

parent = {t["num"]: t["num"] for t in T}


def find(x):
    while parent[x] != x:
        parent[x] = parent[parent[x]]
        x = parent[x]
    return x


LIENS = defaultdict(set)  # paire -> raisons


def lier(a, b, raison):
    if a is b:
        return
    LIENS[tuple(sorted((a["num"], b["num"])))].add(raison)
    parent[find(a["num"])] = find(b["num"])


for cle, raison in ((lambda t: [t["immeuble"]] if t["immeuble"] else [], "même immeuble"),
                    (lambda t: [norm(t["adresse"]) + "|" + norm(t["ville"])] if re.match(r"^\s*[1-9]", t["adresse"]) else [], "même adresse"),
                    (lambda t: t["cls"], "même référence client"),
                    (lambda t: t["pto"], "même PTO"),
                    (lambda t: t["boites"], "même boîte")):
    idx = defaultdict(list)
    for t in T:
        for k in cle(t):
            idx[k].append(t)
    for k, L in idx.items():
        for b in L[1:]:
            lier(L[0], b, raison)
par_base = {t["base"]: t for t in T}
for t in T:
    for ref in t["refs"]:
        if ref in par_base:
            lier(t, par_base[ref], f"suite de {ref}")
for t in T:  # tout proches à vol d'oiseau
    for u in T:
        if t["num"] < u["num"] and (t["a_planifier"] or u["a_planifier"]):
            d = dist(t, u)
            if d is not None and d < 30 and (t["immeuble"] != u["immeuble"] or not t["immeuble"]):
                lier(t, u, "même endroit (< 30 m)")

GROUPES = defaultdict(list)
for t in T:
    GROUPES[find(t["num"])].append(t)
GROUPES = [g for g in GROUPES.values() if len(g) > 1 and any(x["a_planifier"] for x in g)]
GROUPES.sort(key=lambda g: (-sum(x["a_planifier"] for x in g), min(x["creee"] or NOW for x in g)))
for i, g in enumerate(GROUPES, 1):
    for t in g:
        t["groupe"] = i
    raisons = set()
    for a in g:
        for b in g:
            raisons |= LIENS.get((a["num"], b["num"]), set())
    for t in g:
        t["groupe_raisons"] = ", ".join(sorted(raisons))
    nums = {x["num"] for x in g}
    for t in g:
        if t["a_planifier"]:
            autres = sorted(nums - {t["num"]})
            deja = [x for x in g if x is not t and x["statut"] in ("Planifiée", "Synchronisée") and x["planif"]]
            if deja:
                x = deja[0]
                inc("2 - Important", t, "Lié à une intervention déjà planifiée",
                    f"{', '.join(sorted(raisons))} : {x['num']} prévue le {fd(x['planif'])} avec {x['tech']}")
            elif any(x["a_planifier"] for x in g if x is not t) and ("même référence client" in raisons or "même PTO" in raisons):
                inc("2 - Important", t, "Doublon probable", f"{', '.join(sorted(raisons))} avec {', '.join(autres)}")

# Plusieurs tickets ouverts sur un même PM
PM = defaultdict(list)
for t in T:
    if t["statut"] in ("Qualifiée", "Pré-planifiée", "Planifiée", "Synchronisée") and t["site"]:
        PM[t["site"].upper()].append(t)
PM_ALERTES = {k: v for k, v in PM.items() if len(v) >= PM_ALERTE}


# ---------- proximité : tournées et interventions où caser un ticket ----------

planifies = [t for t in T if t["statut"] in ("Planifiée", "Synchronisée") and t["planif"] and t["planif"] >= TODAY and t["geo"]]
for t in T:
    if not t["a_planifier"] or not t["geo"]:
        continue
    cands = [(dist(t, p), p) for p in planifies if p is not t]
    cands = sorted([c for c in cands if c[0] is not None and c[0] <= CASER_M], key=lambda c: (c[1]["planif"], c[0]))
    if cands:
        d, p = min(cands, key=lambda c: c[0])
        t["caser"] = f"{p['num']} · {fd(p['planif'])} · {p['tech'] or 'sans technicien'} · à {fdist(d)}"
        t["caser_p"], t["caser_d"] = p, d
    voisins = sorted([(dist(t, u), u) for u in T if u is not t and u["a_planifier"] and dist(t, u) is not None and dist(t, u) <= PROCHE_M], key=lambda c: c[0])
    t["voisins"] = ", ".join(f"{u['num']} ({fdist(d)})" for d, u in voisins)


def priorite(t):
    # 1 : en retard ou sur le point de l'être · 2 : client coupé, déplanifié, niveau 3 · 3 : le reste · 4 : justifié / en attente
    a = t.get("alertes", [])
    if t["justif"]:
        return 4
    if any(x.startswith(("Échéance dépassée", "Échéance dans moins")) for x in a):
        return 1
    if t["niveau"] in (2, 3) and t["equip"] in ("Service Interrompu", "Service Dégradé") or t["deplanifie"] or t["niveau"] == 3:
        return 2
    return 3


A_PLANIFIER = [t for t in T if t["a_planifier"]]
for t in A_PLANIFIER:
    t["prio"] = priorite(t)
    t["nouveau"] = bool(t["creee"] and NOW - t["creee"] < timedelta(hours=NOUVEAU_H))
NOUVEAUX = sorted([t for t in A_PLANIFIER if t["nouveau"] and not t["justif"]], key=lambda t: t["creee"], reverse=True)
JUSTIFIES = [t for t in A_PLANIFIER if t["justif"]]
# par échéance GTR, la plus proche (ou la plus dépassée) d'abord, puis les plus anciens
A_PLANIFIER.sort(key=lambda t: (t["prio"], t["echeance"] or datetime.max, t["creee"] or NOW))

# Tournées : tickets à planifier regroupés par zone, en partant des plus prioritaires
reste = [t for t in A_PLANIFIER if t["geo"]]
TOURNEES = []
while reste:
    graine = reste[0]
    zone = sorted([u for u in reste if dist(graine, u) <= TOURNEE_M], key=lambda u: dist(graine, u))[:TOURNEE_MAX]
    if len(zone) > 1:
        TOURNEES.append(zone)
    for u in zone:
        if len(zone) > 1:
            u["tournee"] = len(TOURNEES)
        reste.remove(u)


# ---------- Excel ----------

wb = openpyxl.Workbook()
H = Font(bold=True, color="FFFFFF")
HF = PatternFill("solid", fgColor="1F4E78")
WRAP = Alignment(wrap_text=True, vertical="top")
COUL = {"nouveau": "E4D7F5", "justifie": "DDE1E6", 4: "DDE1E6", 1: "F8CBAD", 2: "FFE699", 3: "E2EFDA", "1 - Urgent": "F8CBAD", "2 - Important": "FFE699", "3 - À vérifier": "DDEBF7", "4 - Info": "EDEDED"}


def feuille(ws, entetes, lignes, largeurs, couleur=None):
    ws.append(entetes)
    for c in ws[1]:
        c.font, c.fill, c.alignment = H, HF, Alignment(wrap_text=True, vertical="center")
    for ligne in lignes:
        ws.append(ligne)
        if couleur:
            k = couleur(ligne)
            if k in COUL:
                for c in ws[ws.max_row]:
                    c.fill = PatternFill("solid", fgColor=COUL[k])
    for i, w in enumerate(largeurs, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = WRAP
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions


def court(s, n=300):
    s = RE_PHONE.sub("[tél.]", re.sub(r"\s+", " ", s or "")).strip()
    return s if len(s) <= n else s[:n] + "…"


# Résumé
ws = wb.active
ws.title = "Résumé"
nb = Counter(t["statut"] for t in T)
qual = [t for t in T if t["statut"] == "Qualifiée"]
lignes = [
    ("Analyse de l'export Praxedo", os.path.basename(SRC)),
    ("Faite le", fd(NOW)),
    ("Interventions dans l'export", len(T)),
    ("", ""),
    (f"NOUVEAUX (créés depuis moins de {NOUVEAU_H} h)", len(NOUVEAUX)),
    (f"dont à moins de {fdist(CASER_M)} d'une intervention déjà planifiée : à programmer dans la foulée", sum(1 for t in NOUVEAUX if t.get("caser"))),
    ("", ""),
    ("JUSTIFIÉS / EN ATTENTE (GTR justifiée, poteau à remplacer, attente Orange…)", len(JUSTIFIES)),
    ("dont échéance GTR dépassée (retard justifié)", sum(1 for t in JUSTIFIES if t["echeance"] and t["echeance"] < NOW)),
    ("Colonnes de notes lues dans l'export", ", ".join(COLS_NOTES) or "aucune : seule la description est lue"),
    ("", ""),
    ("TICKETS À PLANIFIER", ""),
    ("Qualifiée (jamais planifiés)", sum(1 for t in qual if not t["deplanifie"])),
    ("Qualifiée mais déjà planifiés une fois (déplanifiés)", sum(1 for t in qual if t["deplanifie"])),
    ("Pré-planifiée (à confirmer)", nb.get("Pré-planifiée", 0)),
    ("dont priorité 1 (échéance GTR dépassée ou dans moins de 48 h)", sum(1 for t in A_PLANIFIER if t["prio"] == 1)),
    ("dont priorité 2 (client coupé, déjà déplanifié, niveau 3)", sum(1 for t in A_PLANIFIER if t["prio"] == 2)),
    ("dont justifiés / en attente (priorité 4)", sum(1 for t in A_PLANIFIER if t["prio"] == 4)),
    ("dont échéance GTR dépassée sans justification", sum(1 for t in A_PLANIFIER if t["echeance"] and t["echeance"] < NOW and not t["justif"])),
    ("dont créés il y a plus de 60 jours", sum(1 for t in A_PLANIFIER if t["creee"] and (NOW - t["creee"]).days > AGE_VIEUX_J)),
    ("dont RDV client à prendre", sum(1 for t in A_PLANIFIER if t["rdv"])),
    ("dont avec nacelle", sum(1 for t in A_PLANIFIER if t["nacelle"])),
    ("", ""),
    ("TICKETS PROCHES", ""),
    ("Groupes de tickets liés (même immeuble, boîte, référence, suite de…)", len(GROUPES)),
    (f"Tickets à planifier à moins de {fdist(CASER_M)} d'une intervention déjà planifiée", sum(1 for t in A_PLANIFIER if t.get("caser"))),
    (f"Tournées possibles (tickets à planifier à moins de {fdist(TOURNEE_M)} les uns des autres)", len(TOURNEES)),
    (f"PM avec au moins {PM_ALERTE} tickets ouverts", len(PM_ALERTES)),
    ("", ""),
    ("INCOHÉRENCES", len(INC)),
] + [(f"  {g}", n) for g, n in sorted(Counter(i[0] for i in INC).items())] + [
    ("", ""),
    ("Planifiées / synchronisées", f"{nb.get('Planifiée', 0)} / {nb.get('Synchronisée', 0)}"),
    ("Adresses placées sur la carte", f"{sum(1 for t in T if t['geo'])} sur {len(T)}"),
]
for l in lignes:
    ws.append(l)
    if l[0].isupper() and l[0]:
        ws.cell(ws.max_row, 1).font = Font(bold=True, color="1F4E78")
ws.cell(1, 1).font = Font(bold=True, size=14)
ws.column_dimensions["A"].width = 80
ws.column_dimensions["B"].width = 30

# Nouveaux : à programmer dans la foulée s'ils sont à côté d'une intervention déjà planifiée
ws = wb.create_sheet(f"Nouveaux ({NOUVEAU_H} h)")
feuille(ws, ["N°", "Créé le", "Prio", "Niveau", "Règle GTR", "Échéance GTR", "Ville", "Adresse", "PM (Site)",
             "À programmer avec (intervention déjà planifiée)", "Tout proches (< 300 m)", "Alertes", "Description"],
        [[t["num"], fd(t["creee"]), t["prio"], t["niveau"] or "", t["regle"], fd(t["echeance"]), t["ville"], t["adresse"], t["site"],
          t.get("caser", "") or "Aucune à moins de " + fdist(CASER_M), t.get("voisins", ""),
          " · ".join(dict.fromkeys(t.get("alertes", []))), court(t["desc"], 200)] for t in NOUVEAUX],
        [22, 16, 6, 8, 26, 16, 22, 30, 18, 45, 30, 40, 60], couleur=lambda l: "nouveau")

# Justifiés / en attente
ws = wb.create_sheet("Justifiés")
feuille(ws, ["N°", "Statut", "Niveau", "Échéance GTR", "Retard", "Justification", "Ville", "Adresse", "PM (Site)", "Description"],
        [[t["num"], t["statut"], t["niveau"] or "", fd(t["echeance"]),
          fretard(NOW - t["echeance"]) if t["echeance"] and t["echeance"] < NOW else "", t["justif"], t["ville"], t["adresse"],
          t["site"], court(t["desc"], 200)] for t in JUSTIFIES],
        [22, 12, 8, 16, 9, 50, 22, 30, 18, 70], couleur=lambda l: "justifie")

# À planifier
ws = wb.create_sheet("À planifier")
feuille(ws, ["Prio", "N°", "Nouveau", "Statut", "Niveau", "Échéance GTR", "Retard", "Règle GTR", "Échéance Praxedo", "Âge (j)", "Ville", "Adresse", "PM (Site)", "Immeuble",
             "Boîte(s)", "Équipement", "Ancienne date planifiée", "Justification", "Alertes", "Tickets liés (groupe)", "Tout proches (< 300 m)",
             "À caser avec (intervention déjà planifiée)", "Tournée", "Description"],
        [[t["prio"], t["num"], "oui" if t["nouveau"] else "", t["statut"], t["niveau"] or "", fd(t["echeance"]),
          fretard(NOW - t["echeance"]) if t["echeance"] and t["echeance"] < NOW else "", t["regle"], fd(t["avant"]),
          (NOW - t["creee"]).days if t["creee"] else "",
          t["ville"], t["adresse"], t["site"], t["immeuble"], ", ".join(t["boites"]), t["equip"],
          fd(t["planif"]) if t["deplanifie"] else "", t["justif"], " · ".join(dict.fromkeys(t.get("alertes", []))),
          f"G{t['groupe']} : {t['groupe_raisons']}" if t.get("groupe") else "", t.get("voisins", ""), t.get("caser", ""),
          f"T{t['tournee']}" if t.get("tournee") else "", court(t["desc"])] for t in A_PLANIFIER],
        [6, 22, 9, 12, 8, 16, 9, 26, 16, 8, 22, 30, 18, 18, 22, 16, 16, 40, 45, 30, 30, 40, 9, 70], couleur=lambda l: l[0])

# Tickets liés
ws = wb.create_sheet("Tickets liés")
L = []
for i, g in enumerate(GROUPES, 1):
    for t in sorted(g, key=lambda x: x["creee"] or NOW):
        L.append([f"G{i}", t["groupe_raisons"], t["num"], t["statut"], fd(t["planif"]) if t["planif"] else "", t["tech"],
                  fd(t["creee"]), t["ville"], t["adresse"], t["immeuble"], ", ".join(t["boites"]), ", ".join(t["cls"]), court(t["desc"], 200)])
feuille(ws, ["Groupe", "Pourquoi liés", "N°", "Statut", "Planifiée", "Technicien", "Créée", "Ville", "Adresse", "Immeuble", "Boîte(s)",
             "Réf. client", "Description"], L, [8, 30, 22, 13, 16, 20, 16, 22, 30, 18, 24, 18, 70])

# Tournées
ws = wb.create_sheet("Tournées")
L = []
for i, z in enumerate(TOURNEES, 1):
    for t in z:
        L.append([f"T{i}", len(z), t["prio"], t["num"], t["statut"], fd(t["echeance"]), t["ville"], t["adresse"], t["site"],
                  "oui" if t["rdv"] else "", "oui" if t["nacelle"] else "", t.get("caser", ""), court(t["desc"], 160)])
feuille(ws, ["Tournée", "Nb tickets", "Prio", "N°", "Statut", "Échéance GTR", "Ville", "Adresse", "PM (Site)", "RDV client",
             "Nacelle", "À caser avec", "Description"], L, [9, 9, 6, 22, 12, 16, 22, 30, 18, 9, 9, 40, 60], couleur=lambda l: l[2])

# PM chargés
ws = wb.create_sheet("PM chargés")
L = []
for pm, v in sorted(PM_ALERTES.items(), key=lambda kv: -len(kv[1])):
    c = [x["creee"] for x in v if x["creee"]]
    L.append([pm, len(v), sum(1 for x in v if x["a_planifier"]), ", ".join(sorted({x["ville"] for x in v})),
              f"{fd(min(c), False)} → {fd(max(c), False)}" if c else "", ", ".join(sorted({x["equip"] for x in v if x["equip"]})),
              ", ".join(x["num"] for x in sorted(v, key=lambda x: x["creee"] or NOW))])
feuille(ws, ["PM (Site)", "Tickets ouverts", "dont à planifier", "Communes", "Créés entre", "Équipement", "Tickets"], L,
        [22, 10, 10, 30, 26, 30, 90])

# Incohérences
ws = wb.create_sheet("Incohérences")
INC.sort(key=lambda i: (i[0], i[1]["num"]))
feuille(ws, ["Gravité", "N°", "Statut", "Problème", "Détail", "Technicien", "Ville", "Adresse"],
        [[g, t["num"], t["statut"], p, d, t["tech"], t["ville"], t["adresse"]] for g, t, p, d in INC],
        [14, 22, 13, 45, 70, 20, 22, 30], couleur=lambda l: l[0])

# Règles GTR utilisées
ws = wb.create_sheet("Règles GTR")
regles = Counter(t["regle"] for t in T)
feuille(ws, ["Règle", "Délai compté depuis la création", "Tickets concernés"], [
    ["[4H HO] dans la description", "4 heures ouvrées", sum(n for r, n in regles.items() if r.startswith("[4H HO]"))],
    ["[8H HO] dans la description", "8 heures ouvrées", sum(n for r, n in regles.items() if r.startswith("[8H HO]"))],
    ["« OTR xJ OUVRES » dans la description", "x jours ouvrés", sum(n for r, n in regles.items() if r.startswith("GTR écrite"))],
    ["Niveau 1", f"{GTR_NIVEAU[1][1]} heures (24 h/24)", sum(n for r, n in regles.items() if r.startswith("Niveau 1 :"))],
    ["Niveau 2 hors RDV client", f"{GTR_NIVEAU[2][1]} jour(s) ouvré(s)", sum(n for r, n in regles.items() if r.startswith("Niveau 2 hors RDV"))],
    ["Niveau 2 avec RDV client", "Pas de délai fixe (selon le RDV)", regles.get("Niveau 2 avec RDV client : selon le RDV", 0)],
    ["Niveau 3", f"{GTR_NIVEAU[3][1]} jours ouvrés", sum(n for r, n in regles.items() if r.startswith("Niveau 3 :"))],
    ["Niveau 4", f"{GTR_NIVEAU[4][1]} jours ouvrés", sum(n for r, n in regles.items() if r.startswith("Niveau 4 :"))],
    ["Sans niveau (malfaçons, audits…)", "Échéance « À faire avant » de Praxedo", regles.get("Sans niveau : échéance Praxedo", 0)],
    ["", "", ""],
    ["Jours ouvrés", "Du lundi au vendredi, hors jours fériés (+ Vendredi saint et 26 décembre en Moselle)", ""],
    ["Heures ouvrées", f"De {HO_DEBUT} h à {HO_FIN} h les jours ouvrés", ""],
    ["Ordre", "[xH HO], puis « OTR xJ » écrit dans le ticket, puis le niveau", ""],
], [40, 60, 16])

# Colonnes de l'export
ws = wb.create_sheet("Colonnes export")
UTILITE = {
    "N°": "Indispensable", "Statut": "Indispensable", "Type inter.": "Utile", "Planifiée": "Indispensable",
    "Nom technicien": "Indispensable", "Prénom technicien": "Facultatif (le nom suffit)", "Site": "Indispensable : c'est le PM",
    "Equipement": "Utile (mélange catégories et codes de boîte)", "Créée": "Indispensable", "À faire avant": "Indispensable",
    "Adresse": "Indispensable", "CP": "Indispensable", "Ville site": "Indispensable", "Contact": "Utile : c'est l'immeuble",
    "Description": "Indispensable (niveau, GTR, boîte, référence client)", "Client": "Inutile : répète la ville",
    "Drapeaux": "Peu utile : presque toujours « Backbone »", "Equipiers": "Inutile : toujours vide", "INTER OCEL": "Inutile : toujours vide",
}
L = []
for c in COLS:
    v = [r.get(c, "").strip() for r in rows if (r.get(c) or "").strip()]
    L.append([c, f"{len(v)} / {len(rows)}", len(set(v)), UTILITE.get(c, ""), court(Counter(v).most_common(1)[0][0], 60) if v else ""])
feuille(ws, ["Colonne", "Remplie", "Valeurs différentes", "Utilité pour l'analyse", "Valeur la plus fréquente"], L, [22, 12, 12, 50, 60])

os.makedirs(OUT, exist_ok=True)
jour = NOW.strftime("%Y-%m-%d")
xlsx = os.path.join(OUT, f"Analyse_tickets_{jour}.xlsx")
wb.save(xlsx)


# ---------- carte ----------

pts = []
for t in T:
    if not t["geo"]:
        continue
    if t["a_planifier"]:
        k = "retard" if t.get("prio") == 1 else "aplanifier"
    else:
        k = "planifie"
    cp_ = t.get("caser_p")
    cs = dict(lat=cp_["geo"]["lat"], lon=cp_["geo"]["lon"], n=cp_["num"], p=fd(cp_["planif"]), tech=cp_["tech"] or "sans technicien",
              d=fdist(t["caser_d"])) if cp_ else None
    pts.append(dict(lat=t["geo"]["lat"], lon=t["geo"]["lon"], k=k, n=t["num"], s=t["statut"], ap=t["a_planifier"],
                    c=int(t["creee"].replace(tzinfo=ZoneInfo("Europe/Paris")).timestamp() * 1000) if t["creee"] else None,
                    cr=fd(t["creee"]), cs=cs, j=t["justif"] if t["statut"] in ("Qualifiée", "Pré-planifiée", "Planifiée", "Synchronisée") else "",
                    p=fd(t["planif"]) if t["planif"] else "", tech=t["tech"], av=fd(t["echeance"]), v=t["ville"], a=t["adresse"],
                    pm=t["site"], al=" · ".join(dict.fromkeys(t.get("alertes", []))), g=t.get("groupe", ""), tr=t.get("tournee", ""),
                    d=court(t["desc"], 220)))
html = """<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Carte des tickets</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
html,body{margin:0;height:100%;font:14px system-ui,sans-serif;background:#fff;color:#1a1a1a}
#map{position:absolute;inset:0}
.leg{position:absolute;z-index:1000;top:10px;right:10px;background:#fff;padding:10px 12px;border-radius:8px;box-shadow:0 1px 6px #0003;
  width:270px;max-height:calc(100% - 40px);overflow:auto;box-sizing:border-box}
.leg h1{font-size:15px;margin:0 0 2px}.leg label{display:block;margin:3px 0;cursor:pointer}
.leg h2{font-size:13px;margin:10px 0 4px;color:#6c2e91}
.dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:-1px;border:1px solid #0006}
.nv{display:block;width:100%;text-align:left;font:inherit;font-size:12.5px;background:#f4eefa;border:1px solid #e0d2ee;border-radius:6px;
  padding:5px 7px;margin:4px 0;cursor:pointer;color:inherit}
.nv:hover,.nv:focus-visible{border-color:#8e44ad;outline:none}
.nv b{color:#6c2e91}.nv span{display:block;color:#555}
.pop b{font-size:14px}.pop small{color:#555}.pop .foulee{color:#6c2e91;font-weight:600}
#plier{float:right;border:0;background:none;font:inherit;color:#555;cursor:pointer;padding:0 0 0 8px}
@media (max-width:560px){.leg{top:auto;bottom:10px;left:10px;right:10px;width:auto;max-height:42%}}
</style></head><body><div id="map"></div>
<div class="leg" id="leg"><button id="plier" type="button" aria-expanded="true">Masquer</button><h1>Tickets Praxedo</h1><small>__DATE__</small>
<div id="corps">
<label><input type="checkbox" data-k="nouveau" checked><span class="dot" style="background:#8e44ad"></span>Nouveau (moins de __NOUVEAU_H__ h) <b id="nbnv"></b></label>
<label><input type="checkbox" data-k="justifie" checked><span class="dot" style="background:#8b939c"></span>Justifié / en attente <b id="nbj"></b></label>
<label><input type="checkbox" data-k="retard" checked><span class="dot" style="background:#d62728"></span>À planifier, prioritaire</label>
<label><input type="checkbox" data-k="aplanifier" checked><span class="dot" style="background:#ff9f1c"></span>À planifier</label>
<label><input type="checkbox" data-k="planifie" checked><span class="dot" style="background:#1f77b4"></span>Déjà planifié</label>
<small>Un nouveau prioritaire a un contour rouge. Le trait violet le relie à l'intervention déjà planifiée la plus proche (moins de __RAYON__).</small>
<div id="nouveaux"></div>
</div></div>
<script>
const P=__DATA__;
const NOUVEAU_MS=__NOUVEAU_H__*3600e3, maintenant=Date.now();
const C={nouveau:'#8e44ad',justifie:'#8b939c',retard:'#d62728',aplanifier:'#ff9f1c',planifie:'#1f77b4'};
const map=L.map('map');
// Fond Plan IGN (Géoplateforme) : libre, sans clé. Les serveurs d'OpenStreetMap refusent les pages ouvertes depuis un fichier.
L.tileLayer('https://data.geopf.fr/wmts?SERVICE=WMTS&REQUEST=GetTile&VERSION=1.0.0&LAYER=GEOGRAPHICALGRIDSYSTEMS.PLANIGNV2&STYLE=normal&TILEMATRIXSET=PM&TILEMATRIX={z}&TILEROW={y}&TILECOL={x}&FORMAT=image/png',
  {maxZoom:19,attribution:'&copy; <a href="https://www.ign.fr/">IGN</a> Plan IGN'}).addTo(map);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const layers={planifie:L.layerGroup().addTo(map),justifie:L.layerGroup().addTo(map),aplanifier:L.layerGroup().addTo(map),retard:L.layerGroup().addTo(map),nouveau:L.layerGroup().addTo(map)};
const b=[], nouveaux=[]; let nbj=0;
for(const p of P){
  b.push([p.lat,p.lon]);
  // « nouveau » est recalculé à l'ouverture : la couleur s'en va toute seule 24 h après la création du ticket
  // un ticket justifié (note « GTR justifiée », « poteau à remplacer »…) passe en gris, même s'il est nouveau ou en retard
  const justifie=p.ap&&!!p.j;
  const nouveau=!justifie&&p.ap&&p.c&&maintenant-p.c<NOUVEAU_MS;
  const k=justifie?'justifie':nouveau?'nouveau':p.k;
  if(justifie)nbj++;
  const m=L.circleMarker([p.lat,p.lon],{radius:k==='planifie'?6:(nouveau?10:8),color:nouveau&&p.k==='retard'?C.retard:'#222',
    weight:nouveau&&p.k==='retard'?3:1,fillColor:C[k],fillOpacity:.92});
  const foulee=p.cs?`<br><span class="foulee">À programmer avec ${esc(p.cs.n)} le ${esc(p.cs.p)} · ${esc(p.cs.tech)} · à ${esc(p.cs.d)}</span>`:'';
  m.bindPopup(`<div class="pop"><b>${esc(p.n)}</b>${nouveau?' <span class="foulee">· nouveau</span>':''}<br>${esc(p.s)}${p.p?' · '+esc(p.p):''}${p.tech?' · '+esc(p.tech):''}<br>
  ${esc(p.a)}, ${esc(p.v)}<br><small>Créé le ${esc(p.cr)} · PM ${esc(p.pm)} · échéance GTR ${esc(p.av||'—')}${p.g?' · groupe G'+p.g:''}${p.tr?' · tournée T'+p.tr:''}</small>
  ${p.j?'<br><b style="color:#4a535c">Justifié · '+esc(p.j)+'</b>':''}${foulee}${p.al?'<br><b style="color:#b00">'+esc(p.al)+'</b>':''}<br><small>${esc(p.d)}</small></div>`,{maxWidth:320});
  m.addTo(layers[k]);
  if(nouveau&&p.cs)L.polyline([[p.lat,p.lon],[p.cs.lat,p.cs.lon]],{color:C.nouveau,weight:3,dashArray:'6 5'}).addTo(layers.nouveau);
  if(k==='retard')L.circle([p.lat,p.lon],{radius:__RAYONM__,color:C.retard,weight:1,dashArray:'4 4',fill:false}).addTo(layers.retard);
  if(nouveau)nouveaux.push({p,m});
}
document.getElementById('nbnv').textContent='('+nouveaux.length+')';
document.getElementById('nbj').textContent='('+nbj+')';
const liste=document.getElementById('nouveaux');
if(nouveaux.length){
  liste.innerHTML='<h2>Nouveaux à programmer</h2>';
  nouveaux.sort((x,y)=>(!!y.p.cs-!!x.p.cs)||(y.p.c-x.p.c));
  for(const {p,m} of nouveaux){
    const bt=document.createElement('button');bt.type='button';bt.className='nv';
    bt.innerHTML=`<b>${esc(p.n)}</b> · ${esc(p.v)}<span>${p.cs?'À côté de '+esc(p.cs.tech)+' le '+esc(p.cs.p)+' ('+esc(p.cs.d)+')':'Aucune intervention planifiée à côté'}</span>`;
    bt.onclick=()=>{map.setView([p.lat,p.lon],16);m.openPopup();};
    liste.appendChild(bt);
  }
}
if(b.length)map.fitBounds(b,{padding:[30,30]});else map.setView([48.69,6.18],11);
document.querySelectorAll('.leg input').forEach(i=>i.onchange=()=>i.checked?layers[i.dataset.k].addTo(map):layers[i.dataset.k].remove());
const plier=document.getElementById('plier'),corps=document.getElementById('corps');
plier.onclick=()=>{corps.hidden=!corps.hidden;plier.textContent=corps.hidden?'Afficher':'Masquer';plier.setAttribute('aria-expanded',String(!corps.hidden));};
</script></body></html>"""
carte = os.path.join(OUT, f"Carte_tickets_{jour}.html")
open(carte, "w", encoding="utf-8").write(html.replace("__DATA__", json.dumps(pts, ensure_ascii=False).replace("</", "<\\/")).replace("__DATE__", f"Analyse du {fd(NOW)}").replace("__RAYONM__", str(CASER_M)).replace("__RAYON__", fdist(CASER_M)).replace("__NOUVEAU_H__", str(NOUVEAU_H)))

print(f"{len(T)} interventions · {len(A_PLANIFIER)} à planifier · {len(NOUVEAUX)} nouveaux · {len(GROUPES)} groupes liés · {len(TOURNEES)} tournées · {len(INC)} incohérences")
print(xlsx)
print(carte)
