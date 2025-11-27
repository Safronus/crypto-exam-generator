# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny s **libovolnou hloubkou**, **klasické** a **BONUS** otázky (BONUS s dvojdesetinnými body), **rich‑text editor**, **drag&drop**, **multiselect**, **filtr**, import **DOCX** a **export DOCX ze šablony**.

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.8b

- **Export DOCX se zachováním formátování**: tučné, kurzíva, podtržení, barvy, zarovnání odstavců a odrážky/číslování (jednoduše jako `• ` / `1. ` prefix) se nyní promítnou z editoru do výsledného DOCX.
- Stále platí: v šabloně se nahrazují **jen** placeholdery v ostrých závorkách `<>` – ostatní obsah dokumentu zůstává nedotčen.
- Průvodce (3 kroky): šablona + parametry → výběr otázek pro `<Otázka1..N>` a `<BONUS1..M>` → souhrn + generování (`<MinBody>`, `<MaxBody>`).

### Poznámka ke šablonám
- Je vhodné, aby každý placeholder (`<Otázka1>` apod.) stál **samostatně v odstavci**. Při exportu se celý odstavec nahradí odpovídajícím formátovaným obsahem otázky.

---

## Instalace (macOS)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install PySide6
python3 main.py
```

---

## Verze

- **Aktuální verze:** `1.8b` (release: 2025-11-27)
- Changelog (výběr):
  - `1.8b` – export DOCX zachovává formátování (tučné, kurzíva, barvy, zarovnání, odrážky).
  - `1.8a` – průvodce exportem DOCX (šablona, výběr otázek, výpočet Min/Max).
  - `1.7d` – skrytý náhled + fix bold na 1. klik.
  - `1.7c` – autosave všech změn otázky (debounce 1.2 s).

---

## Git

```bash
git add main.py README.md
git commit -m "feat(export): DOCX šablona s uchováním formátování otázek (v1.8b)"
git tag v1.8b
git push && git push --tags
```
