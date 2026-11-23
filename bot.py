import os
import re
import json
import difflib
import threading
from datetime import datetime

import telebot
import ollama
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# ----------------------------------------------------------------------
# CONFIGURARE
# ----------------------------------------------------------------------
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
PATRON_CHAT_ID = os.environ.get("PATRON_CHAT_ID", "")
MODEL = os.environ.get("MODEL", "llama3.2")

if not TELEGRAM_TOKEN:
    raise SystemExit("Lipsește variabila de mediu TELEGRAM_TOKEN (vezi instrucțiunile din capul fișierului).")

bot = telebot.TeleBot(TELEGRAM_TOKEN)

FISIER_CATALOG = "catalog.txt"
FISIER_EXCEL = "comenzi_cofetarie.xlsx"
FISIER_LOG = "conversatii.log"

TELEFON_COFETARIE = "0700 000 000"

INFO_MAGAZIN = """Magazin: Cofetăria & Brutăria Artizanală „Cofetăria Demo”, București (demo).
Adresă: Strada Exemplu nr. 1.
Program: Luni-Vineri 07:30 - 20:00, Sâmbătă-Duminică 08:30 - 19:00.
Telefon: 0700 000 000.
Plăți acceptate: card, numerar, Apple Pay / Google Pay.
Observații: Torturile întregi se comandă cu cel puțin 24 de ore înainte. Prăjiturile la bucată și cozonacii sunt disponibili zilnic."""

CUVINTE_COMANDA = ("vreau", "comand", "aș dori", "as dori", "doresc", "iau ", "îmi dați", "imi dati", "îmi pregătiți")
CUVINTE_ALERTA = ("operator", "patron", "șef", "sef", "om real", "reclamați", "reclamati",
                  "nemulțumit", "nemultumit", "vorbesc cu cineva")
CUVINTE_DA = ("da", "ok", "confirm", "sigur", "corect", "exact", "desigur", "yes")
CUVINTE_NU = ("nu", "anulez", "anulează", "anuleaza", "renunț", "renunt", "greșit", "gresit")

sesiuni = {}                       # chat_id -> {"stare", "comanda", "istoric"}
lock_excel = threading.Lock()


# ----------------------------------------------------------------------
# CATALOG
# ----------------------------------------------------------------------
def incarca_catalog():
    """Citește catalog.txt (Produs | Preț | Descriere) și returnează o listă de dicționare."""
    produse = []
    if not os.path.exists(FISIER_CATALOG):
        return produse
    with open(FISIER_CATALOG, "r", encoding="utf-8") as f:
        for linie in f:
            linie = linie.strip()
            if not linie or linie.startswith("#"):
                continue
            parti = [p.strip() for p in linie.split("|")]
            if len(parti) < 2:
                continue
            try:
                pret = float(parti[1].replace(",", "."))
            except ValueError:
                continue
            produse.append({
                "nume": parti[0],
                "pret": pret,
                "descriere": parti[2] if len(parti) > 2 else "",
            })
    return produse


def catalog_ca_text(catalog):
    if not catalog:
        return "Catalog indisponibil momentan."
    randuri = []
    for p in catalog:
        rand = f"- {p['nume']}: {p['pret']:g} lei"
        if p["descriere"]:
            rand += f" ({p['descriere']})"
        randuri.append(rand)
    return "\n".join(randuri)


def gaseste_produs(nume, catalog):
    """Potrivește un nume dat de model cu un produs REAL din catalog (altfel None)."""
    nume_l = nume.lower().strip()
    if not nume_l or not catalog:
        return None
    nume_catalog = [p["nume"].lower() for p in catalog]
    if nume_l in nume_catalog:
        return catalog[nume_catalog.index(nume_l)]
    potrivire = difflib.get_close_matches(nume_l, nume_catalog, n=1, cutoff=0.6)
    if potrivire:
        return catalog[nume_catalog.index(potrivire[0])]
    return None


# ----------------------------------------------------------------------
# MODEL AI
# ----------------------------------------------------------------------
def construieste_prompt():
    catalog = incarca_catalog()
    return (
        "Ești asistentul virtual pe Telegram al cofetăriei „Cuptorul cu Poveste” din București.\n"
        "Reguli stricte:\n"
        "1. Scrie în limba română corectă și naturală. Răspunde scurt (2-3 propoziții).\n"
        "2. Folosește STRICT informațiile și prețurile de mai jos. Nu inventa produse, prețuri sau ingrediente.\n"
        "3. Dacă clientul cere ceva ce nu apare în catalog sau cere un om, spune-i politicos să sune la "
        f"{TELEFON_COFETARIE}.\n"
        "4. NU preiei comenzi și NU ceri date personale (nume, telefon, adresă). Dacă clientul vrea să comande, "
        "spune-i să scrie ce produse dorește și în ce cantitate, iar sistemul îi va cere datele.\n"
        "5. NU calcula totaluri de plată. Spune doar prețul pe bucată.\n\n"
        "INFORMAȚII COFETĂRIE:\n" + INFO_MAGAZIN + "\n\n"
        "CATALOG PRODUSE ȘI PREȚURI:\n" + catalog_ca_text(catalog)
    )


def llm(messages, **kw):
    r = ollama.chat(model=MODEL, messages=messages, options={"temperature": 0.2}, **kw)
    return r["message"]["content"]


def extrage_articole(istoric, catalog):
    """Cere modelului produsele dorite, apoi VERIFICĂ fiecare în catalog. Returnează lista de articole."""
    if not catalog:
        return []
    discutie = "\n".join(f"{m['role']}: {m['content']}" for m in istoric[-6:])
    prompt = (
        "Din conversația de mai jos extrage produsele pe care CLIENTUL vrea să le comande, cu cantități. "
        "Folosește DOAR nume de produse exact din CATALOG. Nu inventa produse. "
        'Răspunde JSON: {"articole": [{"produs": "nume exact din catalog", "cantitate": 1}]}. '
        'Dacă nu cere nimic concret: {"articole": []}.\n\n'
        "CATALOG:\n" + catalog_ca_text(catalog) + "\n\nCONVERSAȚIE:\n" + discutie
    )
    try:
        date = json.loads(llm([{"role": "user", "content": prompt}], format="json"))
        articole = date.get("articole", [])
        if not isinstance(articole, list):
            return []
    except Exception as e:
        print("[EROARE EXTRAGERE]", e)
        return []

    combinat = {}
    for a in articole:
        if not isinstance(a, dict):
            continue
        produs = gaseste_produs(str(a.get("produs", "")), catalog)
        if not produs:
            continue
        try:
            cantitate = int(a.get("cantitate", 1))
        except (TypeError, ValueError):
            cantitate = 1
        cantitate = max(1, min(cantitate, 50))
        if produs["nume"] in combinat:
            combinat[produs["nume"]]["cantitate"] += cantitate
        else:
            combinat[produs["nume"]] = {"nume": produs["nume"], "pret": produs["pret"], "cantitate": cantitate}
    return list(combinat.values())


# ----------------------------------------------------------------------
# VALIDARE ȘI FORMATARE
# ----------------------------------------------------------------------
def telefon_valid(text):
    t = re.sub(r"[\s.\-()]", "", text)
    if t.startswith("+40"):
        t = "0" + t[3:]
    elif t.startswith("0040"):
        t = "0" + t[4:]
    return t if re.fullmatch(r"0\d{9}", t) else None


def total_comanda(articole):
    return sum(a["pret"] * a["cantitate"] for a in articole)


def articole_ca_text(articole):
    return "\n".join(f"• {a['cantitate']} x {a['nume']} — {a['pret'] * a['cantitate']:g} lei" for a in articole)


def curata_celula(valoare):
    """Previne formulele malițioase în Excel (text care începe cu = + - @)."""
    s = str(valoare)
    return "'" + s if s[:1] in ("=", "+", "-", "@") else s


def este_raspuns(text, cuvinte):
    t = text.lower().strip(" .!?")
    return any(t == c or t.startswith(c + " ") for c in cuvinte)


# ----------------------------------------------------------------------
# EXCEL / NOTIFICĂRI / LOG
# ----------------------------------------------------------------------
def salveaza_in_excel(c):
    """Returnează True dacă s-a salvat, False dacă nu."""
    produse_text = ", ".join(f"{a['cantitate']} x {a['nume']}" for a in c["articole"])
    with lock_excel:
        try:
            if not os.path.exists(FISIER_EXCEL):
                wb = Workbook()
                ws = wb.active
                ws.title = "Comenzi"
                ws.append(["Data", "Nume Client", "Telefon", "Produse", "Total (lei)", "Livrare / Ridicare"])
                for celula in ws[1]:
                    celula.font = Font(bold=True)
            else:
                wb = load_workbook(FISIER_EXCEL)
                ws = wb.active
            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                curata_celula(c["nume_client"]),
                str(c["telefon"]),
                curata_celula(produse_text),
                total_comanda(c["articole"]),
                curata_celula(c["adresa_livrare"]),
            ])
            wb.save(FISIER_EXCEL)
            return True
        except PermissionError:
            print("[EROARE EXCEL] Închide fișierul comenzi_cofetarie.xlsx ca să pot salva comenzile!")
            return False
        except Exception as e:
            print("[EROARE EXCEL]", e)
            return False


def trimite_patronului(text):
    """Fără parse_mode, deci nu poate crăpa din cauza caracterelor speciale."""
    if PATRON_CHAT_ID.isdigit():
        try:
            bot.send_message(int(PATRON_CHAT_ID), text)
        except Exception as e:
            print("[EROARE TELEGRAM]", e)


def log_conversatie(chat_id, rol, text):
    """Salvează conversațiile ca să vezi unde greșește botul (cere acordul clienților în varianta reală)."""
    try:
        with open(FISIER_LOG, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} | {chat_id} | {rol} | {text}\n")
    except Exception:
        pass


# ----------------------------------------------------------------------
# SESIUNI
# ----------------------------------------------------------------------
def sesiune(chat_id):
    if chat_id not in sesiuni:
        sesiuni[chat_id] = {"stare": "chat", "comanda": {}, "istoric": []}
    return sesiuni[chat_id]


def reseteaza_comanda(s):
    s["stare"] = "chat"
    s["comanda"] = {}


# ----------------------------------------------------------------------
# HANDLERE TELEGRAM
# ----------------------------------------------------------------------
@bot.message_handler(commands=["start"])
def salut(message):
    sesiuni.pop(message.chat.id, None)
    bot.reply_to(
        message,
        "Bună ziua! Bine ați venit la Cuptorul cu Poveste. 🍰\n"
        "Vă pot ajuta cu produse, prețuri, program sau o comandă. Cu ce vă îndulcim astăzi?\n\n"
        "Notă: dacă plasați o comandă, vă vom cere numele și telefonul, folosite doar pentru acea comandă. "
        "Puteți scrie oricând /anuleaza.",
    )


@bot.message_handler(commands=["anuleaza", "anulează"])
def anuleaza(message):
    s = sesiune(message.chat.id)
    reseteaza_comanda(s)
    bot.reply_to(message, "Am anulat comanda în curs. Cu ce vă mai pot ajuta?")


@bot.message_handler(content_types=["text"])
def raspunde(message):
    chat_id = message.chat.id
    text = (message.text or "").strip()[:500]
    if not text:
        return
    s = sesiune(chat_id)
    c = s["comanda"]
    mic = text.lower()
    log_conversatie(chat_id, "client", text)

    # ---------- pașii comenzii: îi controlează codul, nu modelul ----------
    if s["stare"] == "nume":
        if len(text) < 3 or any(ch.isdigit() for ch in text):
            return bot.reply_to(message, "Vă rog să-mi scrieți numele complet.")
        c["nume_client"] = text
        s["stare"] = "telefon"
        return bot.reply_to(message, "Mulțumesc! Care este numărul dumneavoastră de telefon?")

    if s["stare"] == "telefon":
        tel = telefon_valid(text)
        if not tel:
            return bot.reply_to(message, "Numărul nu pare valid. Exemplu: 0722123456")
        c["telefon"] = tel
        s["stare"] = "livrare"
        return bot.reply_to(
            message,
            "Ridicați din cofetărie sau doriți livrare? Scrieți „ridicare” sau adresa completă de livrare.",
        )

    if s["stare"] == "livrare":
        if "ridic" in mic:
            c["adresa_livrare"] = "Ridicare din cofetărie"
        elif len(text) < 8:
            return bot.reply_to(message, "Vă rog să scrieți adresa completă (stradă, număr, oraș) sau „ridicare”.")
        else:
            c["adresa_livrare"] = text
        s["stare"] = "confirmare"
        return bot.reply_to(
            message,
            "Vă rog să verificați comanda:\n\n"
            f"{articole_ca_text(c['articole'])}\n\n"
            f"Total produse: {total_comanda(c['articole']):g} lei\n"
            f"Nume: {c['nume_client']}\nTelefon: {c['telefon']}\nLivrare: {c['adresa_livrare']}\n\n"
            "Confirmați comanda? (da / nu)",
        )

    if s["stare"] == "confirmare":
        if este_raspuns(text, CUVINTE_DA):
            if salveaza_in_excel(c):
                trimite_patronului(
                    "🎂 COMANDĂ NOUĂ\n"
                    f"Client: {c['nume_client']}\nTelefon: {c['telefon']}\n"
                    f"Produse:\n{articole_ca_text(c['articole'])}\n"
                    f"Total produse: {total_comanda(c['articole']):g} lei\n"
                    f"Livrare: {c['adresa_livrare']}"
                )
                reseteaza_comanda(s)
                return bot.reply_to(
                    message,
                    "Comanda a fost înregistrată. Vă contactăm în scurt timp pentru confirmare. Mulțumim! 🍰",
                )
            trimite_patronului(f"❗ EROARE la salvarea comenzii lui {c['nume_client']} ({c['telefon']}). Verifică Excelul!")
            reseteaza_comanda(s)
            return bot.reply_to(
                message,
                f"Din păcate nu am putut înregistra comanda. Vă rugăm să sunați la {TELEFON_COFETARIE}.",
            )
        if este_raspuns(text, CUVINTE_NU):
            reseteaza_comanda(s)
            return bot.reply_to(message, "Am anulat comanda. Dacă doriți altceva, scrieți-mi ce produse vă interesează.")
        return bot.reply_to(message, "Vă rog să răspundeți cu „da” pentru confirmare sau „nu” pentru anulare.")

    # ---------- conversație obișnuită ----------
    s["istoric"].append({"role": "user", "content": text})
    s["istoric"] = s["istoric"][-12:]  # memorie limitată

    if any(w in mic for w in CUVINTE_ALERTA):
        trimite_patronului(f"⚠️ ALERTĂ CLIENT: {message.from_user.first_name}: {text}")

    if any(w in mic for w in CUVINTE_COMANDA):
        catalog = incarca_catalog()
        articole = extrage_articole(s["istoric"], catalog)
        if articole:
            c["articole"] = articole
            s["stare"] = "nume"
            return bot.reply_to(
                message,
                f"Am notat:\n{articole_ca_text(articole)}\n\n"
                f"Total produse: {total_comanda(articole):g} lei.\n"
                "Pentru comandă am nevoie de câteva date. Cum vă numiți?",
            )

    bot.send_chat_action(chat_id, "typing")
    try:
        raspuns = llm([{"role": "system", "content": construieste_prompt()}] + s["istoric"])
    except Exception as e:
        print("[EROARE MODEL]", e)
        return bot.reply_to(message, f"Îmi cer scuze, am o problemă tehnică. Vă rog să sunați la {TELEFON_COFETARIE}.")
    s["istoric"].append({"role": "assistant", "content": raspuns})
    log_conversatie(chat_id, "bot", raspuns)
    bot.reply_to(message, raspuns)


print(f"Botul Cofetăriei „Cuptorul cu Poveste” a pornit (model: {MODEL}). Apasă Ctrl+C pentru oprire.")
bot.infinity_polling()
