#!/usr/bin/env python3
"""Enrich OKVED via list-org.com (search by INN → company page)."""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

import requests

CACHE = Path("/tmp/okved_cache")
CACHE.mkdir(exist_ok=True)

UA = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}

COMPANY_RE = re.compile(r"company/(\d+)")
MAIN_PREFIX_RE = re.compile(
    r"Основной \(по коду ОКВЭД ред\.2\):</span></i>\s*<a[^>]*>(\d{2})</a>",
    re.I,
)
ROW_RE = re.compile(r"<td>(\d{2}\.\d{2}(?:\.\d{1,2})?)</td><td>([^<]*)</td>")


def fetch(session: requests.Session, inn: str) -> dict:
    last = "unknown"
    for attempt in range(4):
        try:
            r = session.get(
                f"https://www.list-org.com/search?type=inn&val={inn}",
                timeout=35,
            )
            if r.status_code == 429:
                time.sleep(30 + attempt * 20)
                continue
            ids = COMPANY_RE.findall(r.text)
            if not ids:
                return {
                    "inn": inn,
                    "main": None,
                    "all": [],
                    "name": None,
                    "status": "not_found",
                    "source": "list_org",
                }
            cid = ids[0]
            time.sleep(0.4)
            r2 = session.get(f"https://www.list-org.com/company/{cid}", timeout=45)
            if r2.status_code == 429:
                time.sleep(30 + attempt * 20)
                continue
            text = r2.text
            main = None
            # table under Виды деятельности — first row is main
            vi = text.find("Виды деятельности")
            if vi >= 0:
                rows = ROW_RE.findall(text[vi : vi + 8000])
                if rows:
                    main = rows[0][0]
                    extra = [c for c, _ in rows[1:] if c != main]
                else:
                    extra = []
            else:
                extra = []
            if not main:
                m = MAIN_PREFIX_RE.search(text)
                if m:
                    main = m.group(1)
            name = None
            nm = re.search(r"<h1[^>]*>([^<]{3,200})</h1>", text)
            if nm:
                name = re.sub(r"\s+", " ", nm.group(1)).strip()
            return {
                "inn": inn,
                "main": main,
                "all": extra[:40],
                "name": name or "",
                "status": "ok" if main else "no_okved",
                "source": "list_org",
            }
        except Exception as e:
            last = str(e)
            time.sleep(3 + attempt * 3)
    return {
        "inn": inn,
        "main": None,
        "all": [],
        "name": None,
        "status": f"error:{last}",
        "source": "list_org",
    }


def main() -> None:
    inns = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    wid = sys.argv[2] if len(sys.argv) > 2 else "0"
    print(f"listorg-w{wid} start n={len(inns)}", flush=True)
    time.sleep(2 + int(wid) * 5)
    session = requests.Session()
    session.headers.update(UA)
    ok = 0
    for n, inn in enumerate(inns, 1):
        cache = CACHE / f"{inn}.json"
        if cache.exists():
            rec = json.loads(cache.read_text(encoding="utf-8"))
            if rec.get("main") and rec.get("status") == "ok":
                ok += 1
                continue
        rec = fetch(session, inn)
        cache.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
        if rec.get("main"):
            ok += 1
        if n % 20 == 0 or n == len(inns):
            print(
                f"listorg-w{wid} {n}/{len(inns)} ok={ok} "
                f"last={inn}:{rec.get('main')}:{rec.get('status')}",
                flush=True,
            )
        time.sleep(0.9)
    print(f"listorg-w{wid} DONE ok={ok}/{len(inns)}", flush=True)


if __name__ == "__main__":
    main()
