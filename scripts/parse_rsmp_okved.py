#!/usr/bin/env python3
"""Stream-parse FNS RSMP dump and extract OKVED for client INNs."""
from __future__ import annotations

import json
import re
import time
import zipfile
from pathlib import Path

ROOT = Path("/workspace")
CLIENT_INNS = ROOT / "data" / "client_inns.json"
OKVED_PATH = ROOT / "data" / "okved_egrul.json"
RSMP_ZIP = Path("/tmp/rsmp/rsmp.zip")
STILL_PATH = Path("/tmp/okved_still_missing.json")
FOUND_PATH = Path("/tmp/rsmp_found.json")
PROGRESS_PATH = Path("/tmp/rsmp_found_progress.json")

INN_RE = re.compile(r'ИНН(?:ЮЛ|ФЛ)="(\d{10}|\d{12})"')
OKVED_MAIN_RE = re.compile(r"<СвОКВЭДОсн\b([^>/]*)/?>")
OKVED_DOP_RE = re.compile(r"<СвОКВЭДДоп\b([^>/]*)/?>")
CODE_RE = re.compile(r'КодОКВЭД="([^"]+)"')
NAME_RE = re.compile(r'НаимОКВЭД="([^"]*)"')


def load_need_inns() -> set[str]:
    raw = json.loads(CLIENT_INNS.read_text(encoding="utf-8"))
    need: set[str] = set()
    for inn in raw:
        s = re.sub(r"\D", "", str(inn).split(".")[0])
        if len(s) in (10, 12):
            need.add(s)
    return need


def extract_okved(window: str) -> tuple[str, str, list[str]] | None:
    om = OKVED_MAIN_RE.search(window)
    if not om:
        return None
    attrs = om.group(1)
    cm = CODE_RE.search(attrs)
    if not cm:
        return None
    main = cm.group(1)
    main_name = ""
    nm = NAME_RE.search(attrs)
    if nm:
        main_name = nm.group(1)
    extra: list[str] = []
    for dm in OKVED_DOP_RE.finditer(window):
        dc = CODE_RE.search(dm.group(1))
        if dc:
            code = dc.group(1)
            if code != main and code not in extra:
                extra.append(code)
    return main, main_name, extra


def to_record(inn: str, main: str, main_name: str, extra: list[str]) -> dict:
    return {
        "inn": inn,
        "main": main,
        "all": extra,
        "name": main_name or "",
        "status": "ok",
        "source": "rsmp_fns",
    }


def parse_rsmp(need: set[str]) -> dict[str, dict]:
    found: dict[str, dict] = {}
    if PROGRESS_PATH.exists():
        try:
            found = json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
            print(f"resumed progress {len(found)}", flush=True)
        except Exception:
            found = {}
    pending = set(need) - set(found.keys())
    t0 = time.time()
    bad = 0
    with zipfile.ZipFile(RSMP_ZIP) as z:
        names = [n for n in z.namelist() if n.endswith(".xml")]
        print(f"xml files {len(names)} need={len(need)} pending={len(pending)}", flush=True)
        for i, name in enumerate(names, 1):
            if not pending:
                print(f"  all matched at file {i}", flush=True)
                break
            try:
                with z.open(name) as f:
                    raw = f.read()
            except Exception as e:
                bad += 1
                if bad <= 5 or bad % 50 == 0:
                    print(f"  skip bad ({bad}): {name}: {e}", flush=True)
                continue
            data = raw.decode("utf-8", errors="ignore")
            for m in INN_RE.finditer(data):
                inn = m.group(1)
                if inn not in pending:
                    continue
                window = data[m.start() : m.start() + 4000]
                parsed = extract_okved(window)
                if not parsed:
                    continue
                main, main_name, extra = parsed
                found[inn] = to_record(inn, main, main_name, extra)
                pending.discard(inn)
            if i % 200 == 0 or i == len(names):
                PROGRESS_PATH.write_text(
                    json.dumps(found, ensure_ascii=False), encoding="utf-8"
                )
                print(
                    f"  {i}/{len(names)} found={len(found)} pending={len(pending)} "
                    f"bad={bad} elapsed={time.time() - t0:.0f}s",
                    flush=True,
                )
    PROGRESS_PATH.write_text(json.dumps(found, ensure_ascii=False), encoding="utf-8")
    print(f"bad files skipped: {bad}", flush=True)
    return found


def has_okved(rec: dict) -> bool:
    if rec.get("status") != "ok":
        return False
    return bool(rec.get("main") or rec.get("okved_code"))


def normalize_existing(rec: dict, inn: str) -> dict:
    """Keep EGRUL schema; map legacy okved_code if present."""
    if rec.get("main"):
        out = dict(rec)
        out.setdefault("inn", inn)
        out.setdefault("all", out.get("all") or [])
        out.setdefault("name", out.get("name") or "")
        out.setdefault("status", "ok")
        return out
    if rec.get("okved_code"):
        return {
            "inn": inn,
            "main": rec["okved_code"],
            "all": list(rec.get("all") or []),
            "name": rec.get("okved_name") or rec.get("name") or "",
            "status": "ok",
            "source": rec.get("source") or "rsmp_fns",
        }
    return rec


def main() -> None:
    need = load_need_inns()
    print(f"need {len(need)}", flush=True)
    existing: dict = {}
    if OKVED_PATH.exists():
        existing = json.loads(OKVED_PATH.read_text(encoding="utf-8"))
    print(f"existing {len(existing)}", flush=True)

    found = parse_rsmp(need)
    print(f"RSMP matched {len(found)} of {len(need)}", flush=True)
    FOUND_PATH.write_text(json.dumps(found, ensure_ascii=False), encoding="utf-8")

    merged: dict[str, dict] = {}
    for inn, rec in existing.items():
        if has_okved(rec) or rec.get("main") or rec.get("okved_code"):
            merged[inn] = normalize_existing(rec, inn)
        else:
            merged[inn] = rec

    # Prefer existing good EGRUL (richer all[]), fill gaps from RSMP
    for inn, v in found.items():
        prev = merged.get(inn)
        if prev and has_okved(prev) and prev.get("source") != "rsmp_fns":
            # keep EGRUL, but ensure main present
            continue
        if prev and has_okved(prev) and len(prev.get("all") or []) > len(v.get("all") or []):
            continue
        merged[inn] = v

    OKVED_PATH.parent.mkdir(parents=True, exist_ok=True)
    OKVED_PATH.write_text(
        json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"merged total {len(merged)}", flush=True)

    still = [i for i in sorted(need) if i not in merged or not has_okved(merged[i])]
    STILL_PATH.write_text(json.dumps(still), encoding="utf-8")
    print(f"still missing {len(still)}", flush=True)
    print(f"sample missing {still[:10]}", flush=True)
    ok_in_need = sum(1 for i in need if i in merged and has_okved(merged[i]))
    print(f"okved coverage in need: {ok_in_need}/{len(need)}", flush=True)


if __name__ == "__main__":
    main()
