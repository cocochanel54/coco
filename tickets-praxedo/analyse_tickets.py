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


def jours_ouvres(depart, n):
    d = depart
    while n > 0:
        d += timedelta(1)
        if d.weekday() < 5 and d.date() not in feries(d.year, ""):
            n -= 1
    return d


T = []
for r in rows:
    num = (r.get("N°") or "").strip()
    if not num:
        continue
    desc = txt(r.get("Description"))
    tete = desc[:250]
    tags = [t.strip().upper() for t in re.findall(r"\[([^\]]+)\]", tete)]
    niv = re.search(r"niveau\s*(\d)", tete, re.I)
    gtr = None
    if re.search(r"\[\s*8H HO\s*\]", tete, re.I):
        gtr = ("8 h ouvrées", None)
    m = re.search(r"OTR\s*(\d+)\s*J\s*OUVR", desc, re.I)
    if m:
        gtr = (f"{m.group(1)} jours ouvrés", int(m.group(1)))
    contact = (r.get("Contact") or "").strip()
    equip = (r.get("Equipement") or "").strip()
    nom = " ".join(x for x in ((r.get("Nom technicien") or "").strip(), (r.get("Prénom technicien") or "").strip()) if x)
    base = num.split("#")[0].upper()
    info = re.search(r"Information du client\s*:\s*([A-Z0-9-]+)\s*/\s*([^\n]*)", desc, re.I)
    T.append(dict(
        num=num, base=base, iteration=int(num.split("#")[1]) if "#" in num and num.split("#")[1].isdigit() else 0,
        statut=(r.get("Statut") or "").strip(), type=(r.get("Type inter.") or "").strip(),
        drapeaux=(r.get("Drapeaux") or "").strip(), planif=dt(r.get("Planifiée")), tech=nom,
        client=(r.get("Client") or "").strip(), site=(r.get("Site") or "").strip(), equip=equip,
        creee=dt(r.get("Créée")), avant=dt(r.get("À faire avant")),
        adresse=(r.get("Adresse") or "").strip(), cp=(r.get("CP") or "").strip(), ville=(r.get("Ville site") or "").strip(),
        contact=contact, immeuble=contact.upper() if RE_IMMEUBLE.match(contact) else "",
        desc=desc, tags=tags, niveau=int(niv.group(1)) if niv else None, gtr=gtr,
        rdv="RDVCLI" in tags, nacelle=bool(re.search(r"nacelle", desc, re.I)),
        cls=sorted(set(RE_CLS.findall(desc + " " + (r.get("Client") or "")))),
        pto=sorted(set(RE_PTO.findall(desc))),
        boites=sorted({b.upper() for b in RE_BOITE.findall(desc + " " + equip)}),
        refs=sorted({"BB-" + x.upper() for x in RE_REF.findall(desc)} - {base}),
        info_site=info.group(1).upper() if info else "", info_adr=info.group(2).strip() if info else "",
    ))

for t in T:
    t["ouvert_non_planifie"] = t["statut"] == "Qualifiée"
    t["a_planifier"] = t["statut"] in ("Qualifiée", "Pré-planifiée")
    t["deplanifie"] = t["statut"] == "Qualifiée" and t["planif"] is not None
    t["dept"] = t["cp"][:2] if t["cp"] else (t["site"].split("-")[1] if re.match(r"^(SRO|PMU|NRO)-\d{2}", t["site"]) else "54")


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


ref_niveau = defaultdict(Counter)  # délai habituel création -> échéance par niveau
for t in T:
    if t["creee"] and t["avant"] and t["niveau"]:
        ref_niveau[t["niveau"]][(t["avant"] - t["creee"]).days] += 1

for t in T:
    t.setdefault("alertes", [])
    ouvert = t["statut"] in ("Qualifiée", "Pré-planifiée", "Planifiée", "Synchronisée")
    if ouvert and t["avant"] and t["avant"] < NOW:
        j = (NOW - t["avant"]).days
        g = "1 - Urgent" if t["statut"] == "Qualifiée" else "2 - Important"
        inc(g, t, f"Échéance dépassée de {j} j" if j else "Échéance dépassée aujourd'hui",
            f"À faire avant le {fd(t['avant'])} · statut {t['statut']}" + (f" · planifiée le {fd(t['planif'])}" if t["planif"] and not t["deplanifie"] else ""))
    elif t["a_planifier"] and t["avant"] and t["avant"] < NOW + timedelta(days=2):
        inc("1 - Urgent", t, "Échéance dans moins de 48 h et pas planifié", f"À faire avant le {fd(t['avant'])}")
    if t["gtr"] and t["gtr"][1] and t["creee"] and t["avant"]:
        attendu = jours_ouvres(t["creee"], t["gtr"][1])
        if abs((t["avant"] - attendu).days) > 3:
            inc("3 - À vérifier", t, f"Échéance qui ne colle pas à la GTR ({t['gtr'][0]})",
                f"GTR écrite dans la description : {t['gtr'][0]}, soit vers le {fd(attendu, False)} · échéance Praxedo : {fd(t['avant'], False)}")
    if t["gtr"] and t["gtr"][0] == "8 h ouvrées" and t["statut"] in ("Qualifiée", "Pré-planifiée") and t["creee"] and NOW - t["creee"] > timedelta(days=1):
        inc("1 - Urgent", t, "GTR 8 h ouvrées dépassée et pas planifié",
            f"Créé le {fd(t['creee'])}, toujours « {t['statut']} »" + (f" · échéance Praxedo au {fd(t['avant'], False)}, sans rapport avec 8 h" if t["avant"] and (t["avant"] - t["creee"]).days > 5 else ""))
    elif t["gtr"] and t["gtr"][0] == "8 h ouvrées" and t["avant"] and t["creee"] and (t["avant"] - t["creee"]).days > 5:
        inc("3 - À vérifier", t, "Échéance qui ne colle pas à la GTR (8 h ouvrées)", f"Échéance Praxedo au {fd(t['avant'], False)}")
    if t["niveau"] in (3, 4) and t["creee"] and t["avant"] and not t["gtr"]:
        delai = (t["avant"] - t["creee"]).days
        habituel = ref_niveau[t["niveau"]].most_common(1)[0][0] if ref_niveau[t["niveau"]] else None
        if habituel is not None and habituel < 30 and delai > 60:
            inc("3 - À vérifier", t, f"Échéance inhabituelle pour un niveau {t['niveau']}",
                f"{delai} jours entre création et échéance, contre {habituel} habituellement pour ce niveau")
    if ouvert and not t["avant"]:
        inc("3 - À vérifier", t, "Pas d'échéance (« À faire avant » vide)")
    if t["deplanifie"]:
        inc("2 - Important", t, "Déplanifié : revenu en « Qualifiée »",
            f"Était planifié le {fd(t['planif'])}, n'a plus de technicien" + (" · déjà en retard à cette date" if t["avant"] and t["planif"] > t["avant"] else ""))
    if t["statut"] in ("Planifiée", "Synchronisée") and t["planif"]:
        if t["planif"] < TODAY:
            inc("2 - Important", t, "Date planifiée passée, intervention pas clôturée",
                f"Prévue le {fd(t['planif'])} ({t['tech']}), toujours « {t['statut']} »")
        if t["avant"] and t["planif"] > t["avant"]:
            inc("2 - Important", t, "Planifiée après l'échéance",
                f"Prévue le {fd(t['planif'])}, à faire avant le {fd(t['avant'])}")
    if t["statut"] == "Pré-planifiée" and t["planif"] and t["avant"] and t["planif"] > t["avant"]:
        inc("2 - Important", t, "Pré-planifiée après l'échéance", f"Prévue le {fd(t['planif'])}, à faire avant le {fd(t['avant'])}")
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
    voisins = sorted([(dist(t, u), u) for u in T if u is not t and u["a_planifier"] and dist(t, u) is not None and dist(t, u) <= PROCHE_M], key=lambda c: c[0])
    t["voisins"] = ", ".join(f"{u['num']} ({fdist(d)})" for d, u in voisins)


def priorite(t):
    # 1 : en retard ou sur le point de l'être · 2 : client coupé, déplanifié, niveau 3 · 3 : le reste
    a = t.get("alertes", [])
    if any(x.startswith(("Échéance dépassée", "Échéance dans moins", "GTR 8 h")) for x in a):
        return 1
    if t["niveau"] in (2, 3) and t["equip"] in ("Service Interrompu", "Service Dégradé") or t["deplanifie"] or t["niveau"] == 3:
        return 2
    return 3


A_PLANIFIER = [t for t in T if t["a_planifier"]]
for t in A_PLANIFIER:
    t["prio"] = priorite(t)
# en retard : le plus en retard d'abord ; sinon le plus ancien d'abord (les échéances à 3 mois ne sont pas fiables)
A_PLANIFIER.sort(key=lambda t: (t["prio"], (t["avant"] or datetime.max) if t["prio"] == 1 else (t["creee"] or NOW)))

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
COUL = {1: "F8CBAD", 2: "FFE699", 3: "E2EFDA", "1 - Urgent": "F8CBAD", "2 - Important": "FFE699", "3 - À vérifier": "DDEBF7", "4 - Info": "EDEDED"}


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
    ("TICKETS À PLANIFIER", ""),
    ("Qualifiée (jamais planifiés)", sum(1 for t in qual if not t["deplanifie"])),
    ("Qualifiée mais déjà planifiés une fois (déplanifiés)", sum(1 for t in qual if t["deplanifie"])),
    ("Pré-planifiée (à confirmer)", nb.get("Pré-planifiée", 0)),
    ("dont priorité 1 (échéance dépassée ou dans moins de 48 h, GTR 8 h dépassée)", sum(1 for t in A_PLANIFIER if t["prio"] == 1)),
    ("dont priorité 2 (client coupé, déjà déplanifié, niveau 3)", sum(1 for t in A_PLANIFIER if t["prio"] == 2)),
    ("dont échéance dépassée", sum(1 for t in A_PLANIFIER if t["avant"] and t["avant"] < NOW)),
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

# À planifier
ws = wb.create_sheet("À planifier")
feuille(ws, ["Prio", "N°", "Statut", "Niveau", "À faire avant", "Retard (j)", "Âge (j)", "Ville", "Adresse", "PM (Site)", "Immeuble",
             "Boîte(s)", "Équipement", "Ancienne date planifiée", "Alertes", "Tickets liés (groupe)", "Tout proches (< 300 m)",
             "À caser avec (intervention déjà planifiée)", "Tournée", "Description"],
        [[t["prio"], t["num"], t["statut"], t["niveau"] or "", fd(t["avant"]),
          (NOW - t["avant"]).days if t["avant"] and t["avant"] < NOW else "", (NOW - t["creee"]).days if t["creee"] else "",
          t["ville"], t["adresse"], t["site"], t["immeuble"], ", ".join(t["boites"]), t["equip"],
          fd(t["planif"]) if t["deplanifie"] else "", " · ".join(dict.fromkeys(t.get("alertes", []))),
          f"G{t['groupe']} : {t['groupe_raisons']}" if t.get("groupe") else "", t.get("voisins", ""), t.get("caser", ""),
          f"T{t['tournee']}" if t.get("tournee") else "", court(t["desc"])] for t in A_PLANIFIER],
        [6, 22, 12, 8, 16, 9, 8, 22, 30, 18, 18, 22, 16, 16, 45, 30, 30, 40, 9, 70], couleur=lambda l: l[0])

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
        L.append([f"T{i}", len(z), t["prio"], t["num"], t["statut"], fd(t["avant"]), t["ville"], t["adresse"], t["site"],
                  "oui" if t["rdv"] else "", "oui" if t["nacelle"] else "", t.get("caser", ""), court(t["desc"], 160)])
feuille(ws, ["Tournée", "Nb tickets", "Prio", "N°", "Statut", "À faire avant", "Ville", "Adresse", "PM (Site)", "RDV client",
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
    pts.append(dict(lat=t["geo"]["lat"], lon=t["geo"]["lon"], k=k, n=t["num"], s=t["statut"],
                    p=fd(t["planif"]) if t["planif"] else "", tech=t["tech"], av=fd(t["avant"]), v=t["ville"], a=t["adresse"],
                    pm=t["site"], al=" · ".join(dict.fromkeys(t.get("alertes", []))), g=t.get("groupe", ""), tr=t.get("tournee", ""),
                    d=court(t["desc"], 220)))
html = """<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Carte des tickets</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
html,body{margin:0;height:100%;font:14px system-ui,sans-serif;background:#fff;color:#1a1a1a}
#map{position:absolute;inset:0}
.leg{position:absolute;z-index:1000;top:10px;right:10px;background:#fff;padding:10px 12px;border-radius:8px;box-shadow:0 1px 6px #0003;max-width:240px}
.leg h1{font-size:15px;margin:0 0 6px}.leg label{display:block;margin:3px 0;cursor:pointer}
.dot{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:-1px;border:1px solid #0006}
.pop b{font-size:14px}.pop small{color:#555}
</style></head><body><div id="map"></div>
<div class="leg"><h1>Tickets Praxedo</h1><small>__DATE__</small>
<label><input type="checkbox" data-k="retard" checked><span class="dot" style="background:#d62728"></span>À planifier, prioritaire</label>
<label><input type="checkbox" data-k="aplanifier" checked><span class="dot" style="background:#ff9f1c"></span>À planifier</label>
<label><input type="checkbox" data-k="planifie" checked><span class="dot" style="background:#1f77b4"></span>Déjà planifié</label>
<small>Cercle pointillé : rayon de __RAYON__ autour d'un ticket prioritaire.</small></div>
<script>
const P=__DATA__;
const C={retard:'#d62728',aplanifier:'#ff9f1c',planifie:'#1f77b4'};
const map=L.map('map');
L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'&copy; OpenStreetMap'}).addTo(map);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const layers={retard:L.layerGroup().addTo(map),aplanifier:L.layerGroup().addTo(map),planifie:L.layerGroup().addTo(map)};
const b=[];
for(const p of P){
  b.push([p.lat,p.lon]);
  const m=L.circleMarker([p.lat,p.lon],{radius:p.k==='planifie'?6:8,color:'#222',weight:1,fillColor:C[p.k],fillOpacity:.9});
  m.bindPopup(`<div class="pop"><b>${esc(p.n)}</b><br>${esc(p.s)}${p.p?' · '+esc(p.p):''}${p.tech?' · '+esc(p.tech):''}<br>
  ${esc(p.a)}, ${esc(p.v)}<br><small>PM ${esc(p.pm)} · à faire avant ${esc(p.av)}${p.g?' · groupe G'+p.g:''}${p.tr?' · tournée T'+p.tr:''}</small>
  ${p.al?'<br><b style="color:#b00">'+esc(p.al)+'</b>':''}<br><small>${esc(p.d)}</small></div>`);
  m.addTo(layers[p.k]);
  if(p.k==='retard')L.circle([p.lat,p.lon],{radius:__RAYONM__,color:C.retard,weight:1,dashArray:'4 4',fill:false}).addTo(layers.retard);
}
if(b.length)map.fitBounds(b,{padding:[30,30]});else map.setView([48.69,6.18],11);
document.querySelectorAll('.leg input').forEach(i=>i.onchange=()=>i.checked?layers[i.dataset.k].addTo(map):layers[i.dataset.k].remove());
</script></body></html>"""
carte = os.path.join(OUT, f"Carte_tickets_{jour}.html")
open(carte, "w", encoding="utf-8").write(html.replace("__DATA__", json.dumps(pts, ensure_ascii=False).replace("</", "<\\/")).replace("__DATE__", f"Analyse du {fd(NOW)}").replace("__RAYONM__", str(CASER_M)).replace("__RAYON__", fdist(CASER_M)))

print(f"{len(T)} interventions · {len(A_PLANIFIER)} à planifier · {len(GROUPES)} groupes liés · {len(TOURNEES)} tournées · {len(INC)} incohérences")
print(xlsx)
print(carte)
