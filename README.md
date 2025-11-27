# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny s **libovolnou hloubkou**, **klasické** a **BONUS** otázky (BONUS s dvojdesetinnými body), **rich‑text editor**, **drag&drop**, **multiselect**, **filtr**, import **DOCX** a **export DOCX ze šablony**.

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.8d
- **Export DOCX – robustní náhrada placeholderů** (inline i blokově; funguje i v záhlaví/zápatí).  
- `<PoznamkaVerze>`, `<DatumČas>`, `<MinBody>`, `<MaxBody>` se nyní nahrazují i když jsou **rozsekané do více runů**.  
- `<OtázkaX>/<BONUSX>`: když je token **samostatně v odstavci**, vloží se plné formátované HTML otázky (více odstavců/listů). Když je **inline v řádku**, vloží se **plain‑text** varianta otázky, aby se nerozbily okolní styly.  
- **Wizard**: sken placeholderů je tolerantní na mezery uvnitř závorek (např. `< BONUS3 >`).

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

- **Aktuální verze:** `1.8d` (release: 2025-11-27)
- Changelog (výběr):
  - `1.8d` – robustní náhrady placeholderů (inline/blokově), fixy pro hlavičky/patičky a rozsekané runy.
  - `1.8c` – app ikona, oprava importu DOCX, věrnější číslování seznamů.
  - `1.8b` – export DOCX zachovává formátování (tučné, kurzíva, barvy, zarovnání, odrážky).
  - `1.8a` – průvodce exportem DOCX (šablona, výběr otázek, výpočet Min/Max).
  - `1.7d` – skrytý náhled + fix bold na 1. klik.
  - `1.7c` – autosave všech změn otázky (debounce 1.2 s).

---

## Git

```bash
git add main.py README.md icon/icon.png
git commit -m "fix(export): robustní náhrada placeholderů + inline/hlavičky (v1.8d)"
git tag v1.8d
git push && git push --tags
```
