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

- **Aktuální verze:** `1.5` (release: 2025-11-27)
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


---

## Novinky ve verzi 1.3

- **Import DOCX (oprava):** korektní rozdělení očíslovaných otázek **1..10** na samostatné položky; 
  ignoruje se klasifikační škála „**A -> <...> bodů**“ a podobné instrukce.
- **Multiselect + hromadné akce:** nad stromem je **filtr** (název/obsah) a tlačítka **Přesunout vybrané…** a **Smazat vybrané**.
- **Filtr:** hledá v názvech skupin/podskupin i v textu otázek (HTML se převádí na čistý text).


---

## Novinky ve verzi 1.4

- **Import DOCX (fix):** opravena chyba `name 'text' is not defined` a spolehlivější detekce **1..10** podle wordového číslování (`w:numPr`).  
- **Ignorace škály A–F:** řádky typu `A -> <...> bodů` až `F -> ...` se vynechají.  
- **Zachování odrážek/číslování:** následné odstavce s odrážkami nebo číslovanými položkami se převádějí do HTML seznamů (`<ul>`, `<ol type="a">`, `<ol>`).  
- **Výsledek:** Každá očíslovaná otázka je **samostatná položka**, BONUS otázky zachovány.


---

## Novinky ve verzi 1.4a

- **Import DOCX (kompatibilita):** odstraněny staré duplikáty extractorů; parser nyní akceptuje i starý formát `list[tuple]`, takže chyba `tuple has no attribute get` je vyřešena.


---

## Novinky ve verzi 1.5

- **Drag & Drop fix:** po přesunu otázky se strom **znovu vykreslí** a otázka se **znovu vybere**, takže se hned ukáže její obsah v editoru.
- **Výběr cíle (strom):** při **Přesunout vybrané…** i **Přesunout otázku…** se otevře **strom skupin/podskupin** pro výběr cíle.
- **Qt6 kompatibilita:** používá se `QFont.Weight.Bold/Normal` místo `Qt.Bold/Normal` (odstraňuje pády v _sync_toolbar_to_cursor).
