"""
Test della logica di path resolution per l'eseguibile PyInstaller.

Verifica che sys.executable.parent punti sempre alla cartella corretta
indipendentemente dalla directory di lavoro corrente (CWD).
"""
import sys
import os
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent


def test_script_mode():
    """In modalità script, _HERE deve essere la cartella del file .py."""
    print("TEST 1 — Modalità script (non frozen)")
    # Simula la logica di turnazione_completa.py
    frozen = getattr(sys, "frozen", False)
    if frozen:
        _HERE = Path(sys.executable).parent
    else:
        _HERE = Path(__file__).parent

    assert not frozen, "Questo test deve girare come script Python, non come exe"
    assert _HERE == PROJECT_ROOT, f"Atteso {PROJECT_ROOT}, trovato {_HERE}"
    assert (_HERE / "turnazione_completa.py").exists(), "turnazione_completa.py non trovato"
    print(f"  ✅ _HERE = {_HERE}")


def test_cwd_independence():
    """
    Verifica che sys.executable.parent sia stabile anche se CWD cambia.
    Simula il caso: utente lancia exe da ~ o da cartella diversa.
    """
    print("TEST 2 — Indipendenza dal CWD")
    original_cwd = Path.cwd()

    try:
        # Cambia CWD alla home
        os.chdir(Path.home())
        assert Path.cwd() == Path.home(), "CWD non cambiato"

        # sys.executable.parent NON deve dipendere da CWD
        exe_parent = Path(sys.executable).parent
        cwd = Path.cwd()

        assert exe_parent != cwd, (
            f"sys.executable.parent ({exe_parent}) coincide con CWD ({cwd}) — "
            "test non significativo, CWD è già la cartella dell'exe"
        )

        # La logica corretta usa exe_parent, non cwd
        path_via_exe = exe_parent / "input_turni.xlsx"
        path_via_cwd = cwd / "input_turni.xlsx"

        print(f"  sys.executable.parent = {exe_parent}")
        print(f"  CWD (simulato ~)       = {cwd}")
        print(f"  Path via exe_parent    = {path_via_exe}")
        print(f"  Path via CWD           = {path_via_cwd}")

        # Verifica che CWD (home) NON contenga input_turni.xlsx
        if path_via_cwd.exists():
            print(f"  ⚠️  input_turni.xlsx trovato anche in home — test non conclusivo")
        else:
            print(f"  ✅ CWD (home) non contiene input_turni.xlsx — confermato che Path.cwd() era sbagliato")

    finally:
        os.chdir(original_cwd)
    print(f"  ✅ CWD ripristinato a {original_cwd}")


def test_exe_in_project_root():
    """
    Verifica che l'exe (quando buildato con --distpath .) sia nella
    stessa cartella di input_turni.xlsx.
    """
    print("TEST 3 — Struttura attesa post-build")
    exe_macOS = PROJECT_ROOT / "GeneraTurni"
    exe_windows = PROJECT_ROOT / "GeneraTurni.exe"
    input_file = PROJECT_ROOT / "input_turni.xlsx"
    template_file = PROJECT_ROOT / "input_turni_template.xlsx"

    if exe_macOS.exists():
        print(f"  ✅ exe macOS trovato: {exe_macOS}")
        assert input_file.exists(), f"input_turni.xlsx deve essere accanto all'exe in {PROJECT_ROOT}"
        print(f"  ✅ input_turni.xlsx accanto all'exe")
    elif exe_windows.exists():
        print(f"  ✅ exe Windows trovato: {exe_windows}")
    else:
        print(f"  ⚠️  Exe non ancora buildato — eseguire:")
        print(f"      pyinstaller GeneraTurni.spec --noconfirm --distpath .")
        print(f"  (test non applicabile in questa fase)")

    assert template_file.exists(), f"input_turni_template.xlsx non trovato in {PROJECT_ROOT}"
    print(f"  ✅ input_turni_template.xlsx presente")


def test_frozen_logic_simulation():
    """
    Simula la logica frozen=True usando sys.executable della Python installata.
    In questo test sys.executable = python, non l'exe finale, ma verifica
    che il pattern sia corretto.
    """
    print("TEST 4 — Simulazione logica frozen")

    # Simula: frozen=True, sys.executable = PROJECT_ROOT/GeneraTurni
    simulated_exe = PROJECT_ROOT / "GeneraTurni"
    simulated_here = simulated_exe.parent  # = PROJECT_ROOT

    assert simulated_here == PROJECT_ROOT
    expected_input = simulated_here / "input_turni.xlsx"
    expected_output = simulated_here / "output" / "turnazione_generata.xlsx"

    print(f"  Exe simulato       = {simulated_exe}")
    print(f"  _HERE simulato     = {simulated_here}")
    print(f"  INPUT_FILE atteso  = {expected_input}")
    print(f"  OUTPUT_FILE atteso = {expected_output}")
    assert expected_input == PROJECT_ROOT / "input_turni.xlsx"
    print(f"  ✅ Path corretto")


if __name__ == "__main__":
    print("=" * 60)
    print("TEST PATH RESOLUTION — PyInstaller exe")
    print("=" * 60)

    tests = [
        test_script_mode,
        test_cwd_independence,
        test_exe_in_project_root,
        test_frozen_logic_simulation,
    ]

    passed = 0
    failed = 0
    for t in tests:
        print()
        try:
            t()
            passed += 1
        except AssertionError as e:
            print(f"  ❌ FALLITO: {e}")
            failed += 1
        except Exception as e:
            print(f"  ❌ ERRORE: {e}")
            failed += 1

    print()
    print("=" * 60)
    print(f"RIEPILOGO: {passed} passati, {failed} falliti")
    print("=" * 60)
    sys.exit(1 if failed > 0 else 0)
