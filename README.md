# Crypto Exam Generator

Jednoduchá macOS-friendly aplikace v **PySide6** pro správu zkušebních otázek z předmětu *Kryptologie*.
Podporuje **skupiny a podskupiny** (nově libovolná **hierarchie**), dva typy otázek (**klasická** a **BONUS**) a **bohaté formátování** textu.
Data se ukládají do **JSON** a jsou ve výchozím nastavení mimo git (`/data` je v `.gitignore`).

> Cíl: minimum souborů – celý GUI kód je v jediném souboru `main.py`.

---

## Novinky ve verzi 1.2

- **Hierarchické podskupiny**: podskupina může obsahovat další podskupiny (strom jako složky).
- **Drag & drop**: v levém stromu lze **přesouvat a řadit** podskupiny i otázky. Změny se ihned ukládají.
- Zachováno: import z **DOCX**, editor formátování, JSON úložiště.

---

## Import z DOCX

### Kde najdu import?
- **macOS**: menu **Soubor → Import z DOCX…** je v horní systémové liště (vedle názvu aplikace).
- Navíc je v okně aplikace **tlačítko v toolbaru „Import“**.
- Klávesová zkratka: **Ctrl+I** (na macOS `⌘I` pokud máš mapování Cmd).



- V menu **Soubor → Import z DOCX…** vyber .docx testy (např. export z Wordu).
- Aplikace automaticky najde otázky:
  - **Klasické** – číslované (např. „1. …“); bodování v textu je **ignorováno** a nastaví se **1 bod**.
  - **BONUS** – bloky obsahující „**Otázka <číslo>**“ nebo slovo **BONUS**.
- Importované otázky se uloží do skupiny **„Neroztříděné“** (vytvoří se automaticky).

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

## Struktura dat (JSON)

```json
{
  "groups": [
    {
      "id": "uuid",
      "name": "Název skupiny",
      "subgroups": [
        {
          "id": "uuid",
          "name": "Název podskupiny",
          "subgroups": [ ... ],   // libovolná hloubka
          "questions": [
            {
              "id": "uuid",
              "type": "classic",
              "text_html": "<p>...</p>",
              "points": 1,
              "bonus_correct": 0,
              "bonus_wrong": 0,
              "created_at": "YYYY-MM-DDTHH:MM:SS"
            }
          ]
        }
      ]
    }
  ]
}
```

---

## Verze

- **Aktuální verze:** `1.2a` (release: 2025-11-27)
- Changelog:
  - `1.2` – Hierarchické podskupiny + drag & drop (autosave po přesunu).
  - `1.1a` – Autosave prázdných skupin/podskupin po přidání/rename/mazání.
  - `1.1` – Import z DOCX, přesun otázek mezi skupinami.
  - `1.0` – První verze: GUI, skupiny/podskupiny, typy otázek, editor formátování, JSON úložiště.

---

## Git

```bash
git add main.py README.md
git commit -m "feat: hierarchické podskupiny a drag&drop (v1.2)"
git tag v1.2
git push && git push --tags
```

## Licence

Zvol dle potřeby (např. MIT).
