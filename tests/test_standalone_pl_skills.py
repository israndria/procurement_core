import importlib.util
import os
from argparse import Namespace
from pathlib import Path


SKILLS = Path(
    os.environ.get("POKJA_DRIVE_ROOT", r"D:\Dokumen\@ POKJA 2026")
) / ".agents" / "skills"


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_tab7_runner_exposes_four_independent_checklists():
    path = SKILLS / "spse-download-kualifikasi-pl" / "scripts" / "run.py"
    source = path.read_text(encoding="utf-8")
    assert "--download" in source
    assert "--parse-hasil" in source
    assert "--parse-evaluasi" in source
    assert "--update-hps" in source

    module = _load(path, "standalone_tab7_runner")
    actions = module._select_actions(
        Namespace(download=False, parse_hasil=True, parse_evaluasi=False, update_hps=False)
    )
    assert actions == {
        "download": False,
        "parse_hasil": True,
        "parse_evaluasi": False,
        "update_hps": False,
    }


def test_load_runner_prefers_canonical_snapshot_and_has_migration_guard(tmp_path):
    path = SKILLS / "pl-snapshot-load-xlsm" / "scripts" / "run.py"
    module = _load(path, "standalone_load_runner")
    folder = tmp_path / "paket"
    canonical = folder / "11. XML Data"
    canonical.mkdir(parents=True)
    legacy = folder / "input_data_snapshot.xml"
    current = canonical / "input_data_snapshot.xml"
    legacy.write_text("legacy", encoding="utf-8")
    current.write_text("canonical", encoding="utf-8")

    assert module._snapshot(folder) == current
    source = path.read_text(encoding="utf-8")
    assert "--template-migration" in source
    assert "LoadDataPL" in source
    assert "_verify_snapshot_cells" in source


def test_standalone_runners_normalize_asisten_pokja_env_root(monkeypatch):
    load_path = SKILLS / "pl-snapshot-load-xlsm" / "scripts" / "run.py"
    tab7_path = SKILLS / "spse-download-kualifikasi-pl" / "scripts" / "run.py"
    load_source = load_path.read_text(encoding="utf-8")
    tab7_source = tab7_path.read_text(encoding="utf-8")
    assert "configured.name.casefold() == \"asisten_pokja\"" in load_source
    assert "configured.name.casefold() == \"asisten_pokja\"" in tab7_source
    assert "def _resolve_live(folders):" in tab7_source
    assert "helper._resolve_package" in tab7_source

    monkeypatch.setenv("POKJA_CODE_ROOT", r"D:\POKJA2026-Code\Asisten_Pokja")
    load_module = _load(load_path, "standalone_load_runner_env")
    tab7_module = _load(tab7_path, "standalone_tab7_runner_env")
    assert load_module.CODE_ROOT.name == "POKJA2026-Code"
    assert tab7_module.CODE_ROOT.name == "POKJA2026-Code"
    assert tab7_module.ASISTEN_ROOT.name == "Asisten_Pokja"


def test_headless_merge_runner_preflights_mergefields_before_preview():
    path = SKILLS / "pl-headless-merge" / "scripts" / "run.py"
    source = path.read_text(encoding="utf-8")
    assert "def _subprocess_preflight" in source
    assert "mergefield valid; belum membuat PDF" in source
    module = _load(path, "standalone_headless_merge_runner")
    assert module.CORE.name == "procurement_core"
    assert module._source_sheet(Path("0. BAPLPK- Paket.xlsm"), "ba-reviu") == "satu_data"
    assert module._source_sheet(Path("0. BAPLJKK - Paket.xlsm"), "ba-reviu") == "list_reviu"
