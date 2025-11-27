# Crypto Exam Generator

PySide6 aplikace pro správu a export zkušebních otázek (Kryptologie). Jednosouborové GUI (`main.py`).

## Novinky – verze 1.8e (2025-11-27)
- Export DOCX **zachovává číslování** – upravují se jen *runs*, `w:pPr` (včetně `numPr`) zůstává.
- Placeholdery se nahrazují i pokud jsou **rozsekané do více `w:t`**.
- INLINE i BLOCK náhrady `<OtázkaX>/<BONUSX>`; `<BONUS3>` se detekuje korektně.
- Výchozí cesty: šablona `data/Šablony/template_AK3KR.docx`, výstup `data/Vygenerované testy/Test_YYYYMMDD_HHMM.docx`.

## Instalace (macOS)
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install PySide6
python3 main.py
```

## Git
```bash
git add main.py README.md
git commit -m "fix(export): zachování číslování + robustní placeholdery (v1.8e)"
git tag v1.8e
git push && git push --tags
```
