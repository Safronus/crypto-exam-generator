# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.7c

- **Autosave všech změn:** jakýkoliv zásah do otázky (text, formát, zarovnání, název, typ, body) se do **JSON ukládá automaticky** (s jemným zpožděním ~1.2 s). Není nutné používat **Uložit vše**.
- **Uložit změny otázky**: okamžité trvalé uložení dané otázky.

### Poznámky k výkonu
- Autosave používá **debounce** (časovač 1.2 s), aby se nepsalo na disk při každém úhozu.
- Při akcích DnD / mazání / přejmenování skupin se ukládá okamžitě jako dříve.

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

- **Aktuální verze:** `1.7c` (release: 2025-11-27)
- Changelog (výběr):
  - `1.7c` – autosave všech změn otázky (debounce 1.2 s).
  - `1.7b` – BONUS body s přesností na dvě desetinná místa.
  - `1.7a` – vizuální odlišení BONUS vs. klasická + tučné skupiny.

---

## Git

```bash
git add main.py README.md
git commit -m "feat(save): autosave všech změn otázky + okamžité uložení tlačítkem (v1.7c)"
git tag v1.7c
git push && git push --tags
```
