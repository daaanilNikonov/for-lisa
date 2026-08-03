# AGENTS.md

## Cursor Cloud specific instructions

This repository is a **presentation-generation project**, not a long-running service. Python scripts
under `scripts/` use `python-pptx` + `lxml` to build branded PowerPoint decks about the SaaS product
«1С:Кабинет сотрудника» (ГК Форус dark template). There is no server, database, port, or env var.

### Build / run (the core workflow)

The dependencies (`python-pptx`, `lxml`; see `requirements.txt`) are installed by the startup update
script. To (re)generate the decks:

- Pre-demo deck (14 slides): `python3 scripts/build_kedo_presentation.py`
  → writes `presentation/1С_Кабинет_сотрудника_Преддемонстрация.pptx` and a copy under
  `для презентаций по кабинету сотрудника/`.
- Pilot deck (7 slides, archived): `python3 scripts/build_presentation.py`
  → writes `presentation/Продвижение_1С_Кабинет_сотрудника.pptx`.

Both scripts resolve paths relative to the repo root via `Path(__file__).resolve().parents[1]`, so they
can be run from any working directory. They print `Saved:` / `Slides:` on success (exit 0).

### Non-obvious caveats

- The build reads a required brand template and image assets. If they are missing the scripts exit with
  a clear error: template `для презентаций по кабинету сотрудника/Презентация ГК Форус темный шаблон 16х9 (1).pptx`,
  icons in `presentation/assets/icons/`, screenshots in `presentation/assets/screens/`. Keep these committed.
- There are **no tests, no linter, and no CI** configured. "End-to-end verification" = run the two build
  scripts and confirm the `.pptx` files open (e.g. reopen with `python-pptx` and check slide counts: 14 and 7).
- To render a deck to images for visual review (not required to build), LibreOffice + poppler are handy but
  are **not** part of the dev dependencies:
  `soffice --headless --convert-to pdf --outdir /tmp <file>.pptx` then `pdftoppm -png -r 90 <file>.pdf out`.
