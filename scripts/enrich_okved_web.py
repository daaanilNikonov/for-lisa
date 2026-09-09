#!/usr/bin/env python3
"""Enrich missing OKVED via SBIS pages, fallback to EGRUL PDF extract."""
from __future__ import annotations

import io
import json
import re
import sys
import time
from pathlib import Path

import requests
from pypdf import PdfReader

CACHE = Path("/tmp/okved_cache")
CACHE.mkdir(exist_ok=True)
OUT = Path("/workspace/data/okved_egrul.json")
STILL = Path("/tmp/okved_still_missing.json")

UA = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}

MAIN_RE = re.compile(
    r"Основной вид деятельности.{0,120}?(\d{2}\.\d{2}(?:\.\d{1,2})?)",
    re.I | re.S,
)
CODE_RE = re.compile(r"\b(\d{2}\.\d{2}(?:\.\d{1,2})?)\b")


def load_merged() -> dict:
    if OUT.exists():
        return json.loads(OUT.read_text(encoding="utf-8"))
    return {}


def save_cache(inn: str, rec: dict) -> None:
    (CACHE / f"{inn}.json").write_text(
        json.dumps(rec, ensure_ascii=False), encoding="utf-8"
    )


def fetch_sbis(session: requests.Session, inn: str) -> dict | None:
    r = session.get(f"https://sbis.ru/contragents/{inn}", timeout=35)
    if r.status_code != 200 or len(r.text) < 1000:
        return None
    m = MAIN_RE.search(r.text)
    if not m:
        # sometimes escaped differently
        m = re.search(
            r"Основной вид деятельности[^0-9]{0,200}(\d{2}\.\d{2}(?:\.\d{1,2})?)",
            r.text,
            re.I | re.S,
        )
    if not m:
        return None
    main = m.group(1)
    # collect nearby additional codes after "Дополнительные"
    extra: list[str] = []
    di = r.text.find("Дополнительн")
    if di > 0:
        chunk = r.text[di : di + 4000]
        for c in CODE_RE.findall(chunk):
            if c != main and c not in extra and not c.startswith("20"):  # skip years
                # filter plausible okved: first two digits 01-99
                if 1 <= int(c[:2]) <= 99:
                    extra.append(c)
            if len(extra) >= 30:
                break
    name = None
    nm = re.search(r'"name"\s*:\s*"([^"]{3,200})"', r.text)
    if nm:
        name = nm.group(1)
    return {
        "inn": inn,
        "main": main,
        "all": extra,
        "name": name or "",
        "status": "ok",
        "source": "sbis",
    }


def fetch_egrul(session: requests.Session, inn: str) -> dict:
    last = "unknown"
    for attempt in range(5):
        try:
            j = session.post(
                "https://egrul.nalog.ru/", data={"query": inn}, timeout=40
            ).json()
            if j.get("captchaRequired"):
                time.sleep(20 + attempt * 20)
                continue
            tok = j["t"]
            time.sleep(0.3)
            rows = (
                session.get(
                    f"https://egrul.nalog.ru/search-result/{tok}", timeout=40
                ).json().get("rows")
                or []
            )
            if not rows:
                return {
                    "inn": inn,
                    "main": None,
                    "all": [],
                    "name": None,
                    "status": "not_found",
                    "source": "egrul",
                }
            t = rows[0]["t"]
            name = rows[0].get("c") or rows[0].get("n")
            session.get(f"https://egrul.nalog.ru/vyp-request/{t}", timeout=40)
            for _ in range(40):
                time.sleep(0.3)
                st = session.get(
                    f"https://egrul.nalog.ru/vyp-status/{t}", timeout=40
                ).json()
                if st.get("status") == "ready":
                    break
            pdf = session.get(
                f"https://egrul.nalog.ru/vyp-download/{t}", timeout=90
            ).content
            if not pdf.startswith(b"%PDF"):
                last = "not_pdf"
                time.sleep(3 + attempt * 3)
                continue
            text = "\n".join(
                (p.extract_text() or "") for p in PdfReader(io.BytesIO(pdf)).pages
            )
            main = None
            msec = re.search(
                r"Сведения об основном виде деятельности(.*?)(?:Сведения о дополнительных видах деятельности|\Z)",
                text,
                re.I | re.S,
            )
            if msec:
                mm = re.search(
                    r"Код и наименование вида деятельности\s+(\d{2}\.\d{2}(?:\.\d{1,2})?)",
                    msec.group(1),
                )
                if mm:
                    main = mm.group(1)
            codes = re.findall(
                r"Код и наименование вида деятельности\s+(\d{2}\.\d{2}(?:\.\d{1,2})?)",
                text,
            )
            if not main and codes:
                main = codes[0]
            add = [c for c in codes if c != main]
            return {
                "inn": inn,
                "main": main,
                "all": add,
                "name": name,
                "status": "ok" if main else "no_okved",
                "source": "egrul",
            }
        except Exception as e:
            last = str(e)
            time.sleep(2 + attempt * 2)
    return {
        "inn": inn,
        "main": None,
        "all": [],
        "name": None,
        "status": f"error:{last}",
        "source": "egrul",
    }


def process_one(session: requests.Session, inn: str) -> dict:
    cache = CACHE / f"{inn}.json"
    if cache.exists():
        rec = json.loads(cache.read_text(encoding="utf-8"))
        if rec.get("main") and rec.get("status") == "ok":
            return rec
    try:
        rec = fetch_sbis(session, inn)
    except Exception:
        rec = None
    if not rec or not rec.get("main"):
        rec = fetch_egrul(session, inn)
    save_cache(inn, rec)
    return rec


def main() -> None:
    inns = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    wid = sys.argv[2] if len(sys.argv) > 2 else "0"
    print(f"worker{wid} start n={len(inns)}", flush=True)
    session = requests.Session()
    session.headers.update(UA)
    ok = 0
    for n, inn in enumerate(inns, 1):
        rec = process_one(session, inn)
        if rec.get("main"):
            ok += 1
        if n % 25 == 0 or n == len(inns):
            print(
                f"worker{wid} {n}/{len(inns)} ok={ok} "
                f"last={inn}:{rec.get('main')}:{rec.get('status')}:{rec.get('source')}",
                flush=True,
            )
        time.sleep(0.35)
    print(f"worker{wid} DONE ok={ok}/{len(inns)}", flush=True)


if __name__ == "__main__":
    main()
