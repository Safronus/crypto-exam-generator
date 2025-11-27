# Crypto Exam Generator

PySide6 aplikace pro správu a export zkušebních otázek (Kryptologie). Jednosouborové GUI (`main.py`).

## Verze 1.8e-b (2025-11-27)
- Oprava SyntaxError (`nebo` -> `or`) v `_add_subgroup`.
- Odstraněna rekurze voláním `save_data()` z `_save_current_q` (nyní se ukládá jen z tlačítka a dalších akcí).
- "Uložit změny otázky" ukládá okamžitě na disk.
- Obnoven drag & drop (`DnDTree`) beze změny chování.

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
git commit -m "fix: SyntaxError (nebo->or), zamezení rekurzi a návrat DnD (v1.8e-b)"
git tag v1.8e-b
git push && git push --tags
```
