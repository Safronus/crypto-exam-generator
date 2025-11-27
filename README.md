# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny s **libovolnou hloubkou**, **klasické** a **BONUS** otázky (BONUS s dvojdesetinnými body), **rich‑text editor**, **drag&drop**, **multiselect**, **filtr**, import **DOCX** a **export DOCX ze šablony**.

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.8c
- **Ikona aplikace**: pokud v projektu existuje `icon/icon.png`, nastaví se jako ikona aplikace i hlavního okna (macOS: zobrazí se v Docku pro běžící okno).
- **Fix**: doplněna chybějící metoda `MainWindow._import_from_docx` (menu *Soubor → Import z DOCX…*).
- **Import DOCX**: drobné vylepšení – přenos typu číslování (`decimal` / `lowerLetter` / `upperLetter` / `bullet`) do výsledného HTML (vizuelně věrnější seznamy).

> Pozn.: Pro precizní Word číslování by bylo nutné generovat `numbering.xml`; zde je zvoleno minimalisticky zachovat vzhled pomocí `<ol>/<ul>`.

---

## Instalace (macOS)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install PySide6
python3 main.py
```

### Ikona aplikace
Vytvoř složku `icon/` a přidej soubor `icon.png`:
```
Crypto Exam Generator/
├─ main.py
├─ README.md
└─ icon/
   └─ icon.png
```

---

## Verze

- **Aktuální verze:** `1.8c` (release: 2025-11-27)
- Changelog (výběr):
  - `1.8c` – app ikona, oprava importu DOCX, věrnější číslování seznamů.
  - `1.8b` – export DOCX zachovává formátování (tučné, kurzíva, barvy, zarovnání, odrážky).
  - `1.8a` – průvodce exportem DOCX (šablona, výběr otázek, výpočet Min/Max).
  - `1.7d` – skrytý náhled + fix bold na 1. klik.
  - `1.7c` – autosave všech změn otázky (debounce 1.2 s).

---

## Git

```bash
git add main.py README.md icon/icon.png
git commit -m "feat(ui): app ikona + fix importu DOCX (v1.8c)"
git tag v1.8c
git push && git push --tags
```
