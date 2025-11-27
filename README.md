# Crypto Exam Generator

PySide6 aplikace pro správu a export zkušebních otázek (Kryptologie). Jednosouborové GUI (`main.py`).

## Novinky – verze 1.8f (2025-11-27)
- **HOTFIX**: doplněny chybějící metody `_choose_data_file` a `_bulk_delete_selected`, připojena tlačítka hromadných akcí.
- Ostatní: v1.8e robustní export DOCX (zachování číslování, placeholdery přes více runů).

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
git commit -m "fix(ui): doplněny chybějící metody a napojení tlačítek (v1.8f)"
git tag v1.8f
git push && git push --tags
```
