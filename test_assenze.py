"""
Test della logica di parsing assenze in turnazione_completa.py
Verifica i bug identificati:
  1. Date come stringhe → silenziosamente scartate
  2. date objects (vs datetime) → silenziosamente scartati
  3. Nomi operatori non corrispondenti → assenza ignorata senza warning
  4. Flusso corretto con absences valide in get_week_absences
"""
import sys
from datetime import datetime, date, timedelta
from pathlib import Path
from io import StringIO
import traceback

sys.path.insert(0, str(Path(__file__).parent))

# ── helper: replica la logica di load_input per il solo blocco Assenze ────────

def parse_absences_current(rows):
    """Logica attuale (buggy) — copia esatta da turnazione_completa.py riga 165-177."""
    from datetime import datetime
    absences = {}
    for row in rows:
        if not row[0] or not row[1]:
            continue
        name = str(row[0]).strip()
        absence_date = row[1]
        if not isinstance(absence_date, datetime):
            continue  # BUG: scarta date e stringhe senza avviso
        absences.setdefault(name, set()).add(
            absence_date.replace(hour=0, minute=0, second=0, microsecond=0)
        )
    return absences


def parse_absences_fixed(rows, known_operators=None):
    """Logica corretta — gestisce str, date, datetime + colonna tipo + warning nomi."""
    from datetime import datetime, date
    absences = {}  # {name: {date: tipo}}
    warnings = []

    for i, row in enumerate(rows, start=2):
        if not row[0] or not row[1]:
            continue
        name = str(row[0]).strip()
        raw_date = row[1]
        tipo = str(row[2]).strip() if len(row) > 2 and row[2] else "Assenza"

        if isinstance(raw_date, datetime):
            absence_date = raw_date.replace(hour=0, minute=0, second=0, microsecond=0)
        elif isinstance(raw_date, date):
            absence_date = datetime(raw_date.year, raw_date.month, raw_date.day)
        elif isinstance(raw_date, str):
            for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
                try:
                    absence_date = datetime.strptime(raw_date.strip(), fmt)
                    break
                except ValueError:
                    continue
            else:
                warnings.append(
                    f"Riga {i}: data non riconosciuta per '{name}': '{raw_date}' — assenza ignorata."
                )
                continue
        else:
            warnings.append(
                f"Riga {i}: tipo data non supportato per '{name}': {type(raw_date)} — assenza ignorata."
            )
            continue

        if known_operators is not None and name not in known_operators:
            warnings.append(
                f"Riga {i}: operatore '{name}' non trovato nel foglio Personale — assenza ignorata."
            )
            continue

        absences.setdefault(name, {})[absence_date] = tipo

    return absences, warnings


# ── get_week_absences — vecchia logica (set) ─────────────────────────────────

def get_week_absences_old(absences, week_days):
    """Vecchia logica: absences = {name: set of dates}. Returns {name: set of day_idx}."""
    week_map = {
        d.replace(hour=0, minute=0, second=0, microsecond=0): i
        for i, d in enumerate(week_days)
    }
    result = {}
    for name, dates in absences.items():
        for d in dates:
            idx = week_map.get(d)
            if idx is not None:
                result.setdefault(name, set()).add(idx)
    return result


# ── get_week_absences — nuova logica (dict) ───────────────────────────────────

def get_week_absences(absences, week_days):
    """Returns {name: {day_idx: tipo}}."""
    week_map = {
        d.replace(hour=0, minute=0, second=0, microsecond=0): i
        for i, d in enumerate(week_days)
    }
    result = {}
    for name, date_map in absences.items():
        for d, tipo in date_map.items():
            idx = week_map.get(d)
            if idx is not None:
                result.setdefault(name, {})[idx] = tipo
    return result


# ── fixture ──────────────────────────────────────────────────────────────────

MON = datetime(2025, 4, 7)  # lunedì
WEEK = [MON + timedelta(days=i) for i in range(7)]
KNOWN_OPS = {"Mario Rossi", "Anna Bianchi", "Luigi Verdi"}

ROWS_MIXED = [
    # (nome, data, tipo)
    ("Mario Rossi",   datetime(2025, 4, 9),  "Ferie"),      # mercoledì — OK
    ("Anna Bianchi",  date(2025, 4, 10),     "Malattia"),   # date object
    ("Luigi Verdi",   "11/04/2025",          "Permesso"),   # stringa
    ("mario rossi",   datetime(2025, 4, 8),  "Ferie"),      # nome sbagliato (case)
    ("Sconosciuto",   datetime(2025, 4, 8),  "Ferie"),      # operatore inesistente
    ("Anna Bianchi",  "data-errata",         "Ferie"),      # stringa non parsabile
    ("",              datetime(2025, 4, 7),  "Ferie"),      # nome vuoto → skip
    ("Mario Rossi",   None,                  "Ferie"),      # data nulla → skip
]


# ── test 1: logica ATTUALE ────────────────────────────────────────────────────

def test_current_logic():
    print("=" * 60)
    print("TEST 1 — Logica ATTUALE (buggy)")
    print("=" * 60)
    absences = parse_absences_current(ROWS_MIXED)
    week_abs = get_week_absences_old(absences, WEEK)

    print(f"Assenze parsed: {absences}")
    print(f"Assenze settimana (day-idx): {week_abs}")

    # Verifica bug
    mario_days = week_abs.get("Mario Rossi", set())
    anna_days  = week_abs.get("Anna Bianchi", set())
    luigi_days = week_abs.get("Luigi Verdi", set())

    print()
    print(f"  Mario Rossi  assente giorni idx: {mario_days}  (atteso: {{2}} = mercoledì)")
    print(f"  Anna Bianchi assente giorni idx: {anna_days}   (atteso: {{3}} = giovedì  ← BUG se vuoto)")
    print(f"  Luigi Verdi  assente giorni idx: {luigi_days}  (atteso: {{4}} = venerdì  ← BUG se vuoto)")

    # Con la vecchia logica absences è Set, non Dict — controlla solo presenza
    bugs = []
    if isinstance(anna_days, set) and 3 not in anna_days:
        bugs.append("BUG: Anna Bianchi (date object) ignorata")
    if isinstance(luigi_days, set) and 4 not in luigi_days:
        bugs.append("BUG: Luigi Verdi (stringa) ignorato")

    if bugs:
        print()
        for b in bugs:
            print(f"  ❌ {b}")
    else:
        print("  ✅ Nessun bug (inatteso)")
    return bool(bugs)


# ── test 2: logica CORRETTA ───────────────────────────────────────────────────

def test_fixed_logic():
    print()
    print("=" * 60)
    print("TEST 2 — Logica CORRETTA (proposta)")
    print("=" * 60)
    absences, warnings = parse_absences_fixed(ROWS_MIXED, known_operators=KNOWN_OPS)
    week_abs = get_week_absences(absences, WEEK)

    print(f"Assenze parsed: {absences}")
    print(f"Assenze settimana (day-idx): {week_abs}")
    print(f"Warning emessi ({len(warnings)}):")
    for w in warnings:
        print(f"  ⚠  {w}")

    mario_map = week_abs.get("Mario Rossi", {})
    anna_map  = week_abs.get("Anna Bianchi", {})
    luigi_map = week_abs.get("Luigi Verdi", {})

    print()
    print(f"  Mario Rossi  : {mario_map}  (atteso: {{2: 'Ferie'}})")
    print(f"  Anna Bianchi : {anna_map}   (atteso: {{3: 'Malattia'}})")
    print(f"  Luigi Verdi  : {luigi_map}  (atteso: {{4: 'Permesso'}})")

    ok = True
    if mario_map.get(2) != "Ferie":
        print("  ❌ Mario Rossi tipo errato o giorno mancante")
        ok = False
    if anna_map.get(3) != "Malattia":
        print("  ❌ Anna Bianchi tipo errato o giorno mancante")
        ok = False
    if luigi_map.get(4) != "Permesso":
        print("  ❌ Luigi Verdi tipo errato o giorno mancante")
        ok = False
    if "mario rossi" in week_abs:
        print("  ❌ 'mario rossi' (case errato) non avrebbe dovuto entrare")
        ok = False
    if "Sconosciuto" in week_abs:
        print("  ❌ 'Sconosciuto' (op inesistente) non avrebbe dovuto entrare")
        ok = False

    # Controlla che warning siano stati emessi per casi problematici
    warning_text = " ".join(warnings)
    if "mario rossi" not in warning_text.lower() and "mario" not in warning_text.lower():
        print("  ❌ Manca warning per 'mario rossi' (nome errato)")
        ok = False
    if "Sconosciuto" not in warning_text:
        print("  ❌ Manca warning per 'Sconosciuto'")
        ok = False

    if ok:
        print("  ✅ Tutti i test passati")
    return ok


# ── test 3: get_week_absences con week fuori range ───────────────────────────

def test_week_out_of_range():
    print()
    print("=" * 60)
    print("TEST 3 — Assenze fuori dal range della settimana (devono essere ignorate)")
    print("=" * 60)
    rows = [
        ("Mario Rossi", datetime(2025, 4, 14), "Ferie"),  # settimana successiva
        ("Anna Bianchi", datetime(2025, 3, 31), "Ferie"), # settimana precedente
    ]
    absences, _ = parse_absences_fixed(rows)
    week_abs = get_week_absences(absences, WEEK)

    print(f"week_abs: {week_abs}  (atteso: {{}} vuoto)")
    if not week_abs:
        print("  ✅ Corretto: assenze fuori range ignorate")
        return True
    else:
        print(f"  ❌ Assenze erroneamente incluse: {week_abs}")
        return False


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    results = []
    results.append(("Test 1 (bug attuale)",   test_current_logic()))
    results.append(("Test 2 (fix proposto)",  test_fixed_logic()))
    results.append(("Test 3 (fuori range)",   test_week_out_of_range()))

    print()
    print("=" * 60)
    print("RIEPILOGO")
    print("=" * 60)
    for name, passed in results:
        icon = "✅" if passed else "❌ BUG CONFERMATO"
        print(f"  {icon}  {name}")
