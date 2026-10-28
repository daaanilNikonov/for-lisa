#!/usr/bin/env python3
"""Slow OKVED enrich via EGRUL PDF only (captcha-aware)."""
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

UA = {"User-Agent": "Mozilla/5.0 (compatible; research-bot/1.0)"}


def fetch(session: requests.Session, inn: str) -> dict:
    last = "unknown"
    for attempt in range(6):
        try:
            j = session.post(
                "https://egrul.nalog.ru/", data={"query": inn}, timeout=45
            ).json()
            if j.get("captchaRequired") or j.get("ERRORS"):
                wait = 90 + attempt * 45
                print(f"captcha {inn} sleep {wait}", flush=True)
                time.sleep(wait)
                continue
            tok = j["t"]
            time.sleep(0.35)
            rows = (
                session.get(
                    f"https://egrul.nalog.ru/search-result/{tok}", timeout=45
                )
                .json()
                .get("rows")
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
            session.get(f"https://egrul.nalog.ru/vyp-request/{t}", timeout=45)
            for _ in range(40):
                time.sleep(0.25)
                st = session.get(
                    f"https://egrul.nalog.ru/vyp-status/{t}", timeout=45
                ).json()
                if st.get("status") == "ready":
                    break
            pdf = session.get(
                f"https://egrul.nalog.ru/vyp-download/{t}", timeout=90
            ).content
            if not pdf.startswith(b"%PDF"):
                last = f"not_pdf:{pdf[:40]!r}"
                time.sleep(5 + attempt * 5)
                continue
            text = "\n".join(
                (p.extract_text() or "") for p in PdfReader(io.BytesIO(pdf)).pages[:3]
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
            add = []
            for c in codes:
                if c != main and c not in add:
                    add.append(c)
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
            time.sleep(4 + attempt * 4)
    return {
        "inn": inn,
        "main": None,
        "all": [],
        "name": None,
        "status": f"error:{last}",
        "source": "egrul",
    }


def main() -> None:
    inns = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    wid = sys.argv[2] if len(sys.argv) > 2 else "0"
    print(f"egrul-w{wid} start n={len(inns)}", flush=True)
    time.sleep(5 + int(wid) * 8)
    session = requests.Session()
    session.headers.update(UA)
    ok = 0
    for n, inn in enumerate(inns, 1):
        cache = CACHE / f"{inn}.json"
        if cache.exists():
            rec = json.loads(cache.read_text(encoding="utf-8"))
            if rec.get("main") and rec.get("status") == "ok":
                ok += 1
                if n % 20 == 0:
                    print(f"egrul-w{wid} {n}/{len(inns)} ok={ok} cached {inn}", flush=True)
                continue
        rec = fetch(session, inn)
        cache.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
        if rec.get("main"):
            ok += 1
        if n % 10 == 0 or n == len(inns):
            print(
                f"egrul-w{wid} {n}/{len(inns)} ok={ok} last={inn}:{rec.get('main')}:{rec.get('status')}",
                flush=True,
            )
        time.sleep(1.2)
    print(f"egrul-w{wid} DONE ok={ok}/{len(inns)}", flush=True)


if __name__ == "__main__":
    main()
