# Crypto Exam Generator

PySide6 aplikace pro správu a export zkušebních otázek (Kryptologie). Jednosouborové GUI (`main.py`).

## Verze 1.8e-a (2025-11-27)
- **Revert na 1.8e** + jediný hotfix: doplněna metoda `_choose_data_file` (pad v menu). 
- Zachován export DOCX s 1:1 substitucí (číslování se neničí) a import z DOCX.
- Zachováno vizuální rozlišení BONUS otázek v seznamu („Typ / body“ = `BONUS | +X.XX/ Y.YY`).

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
git commit -m "fix: revert na 1.8e + doplněna _choose_data_file (v1.8e-a)"
git tag v1.8e-a
git push && git push --tags
```
