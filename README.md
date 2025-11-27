# Crypto Exam Generator

Jednoduchá macOS-friendly aplikace v **PySide6** pro správu zkušebních otázek z předmětu *Kryptologie*.
Podporuje **skupiny a podskupiny**, dva typy otázek (**klasická** a **BONUS**) a **bohaté formátování** textu (tučné, kurzíva,
podtržení, barvy, odrážky). Data se ukládají do **JSON** a jsou ve výchozím nastavení mimo git (`/data` je v `.gitignore`).

> Cíl: minimum souborů – celý GUI kód je v jediném souboru `main.py`.

---

## Funkce

- Stromová struktura: **Skupiny → Podskupiny → Otázky**
- **Typy otázek**
  - *Klasická*: text + **body** (default 1)
  - *BONUS*: text + **body za správně** a **body za špatně** (mohou být záporné)
- **Editor textu otázky** s formátováním:
  - **Tučné**, *kurzíva*, <u>podtržení</u>, **barvy**, • **odrážky**
- Přidání, editace, **změna druhu** otázky, změna bodového ohodnocení
- Uložení/načtení z **`data/questions.json`** (lze změnit v UI)
- **Dark theme** (Qt Fusion), HiDPI/Retina

---

## Import z DOCX (nové ve verzi 1.1)

- V menu **Soubor → Import z DOCX…** vyber .docx testy (např. export z Wordu).
- Aplikace z dokumentu automaticky najde otázky:
  - **Klasické** – číslované (např. „1. …“) nebo obsahují bodové ohodnocení v textu (ignoruje se, nastaví se **1 bod**).
  - **BONUS** – bloky obsahující „**Otázka 111**“, „**Otázka 666**“ nebo slovo **BONUS**.
- Importované otázky se uloží do skupiny **„Neroztříděné“** (vytvoří se automaticky).

> Pozn.: U BONUS otázek se pro import nastavuje výchozí **+1 / 0** (správně / špatně).

## Přesun otázek mezi skupinami

- V menu **Úpravy → Přesunout otázku…** lze aktuálně vybranou otázku přesunout do jiné skupiny / podskupiny.

---

## Instalace (macOS)

```bash
# 1) Vytvoř a aktivuj venv (doporučeno)
python3 -m venv .venv
source .venv/bin/activate

# 2) Nainstaluj PySide6
pip install --upgrade pip
pip install PySide6

# 3) Spusť aplikaci
python3 main.py
```

> Pozn.: Aplikace vytvoří složku `data/` (pokud neexistuje) a soubor `questions.json`.
> V menu je možné **zvolit jiný JSON soubor** umístěný např. mimo repository.

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

> **text_html** obsahuje HTML z `QTextEdit` (rich text).

---

## Git & GitHub – vytvoření repozitáře

> Níže jsou příkazy pro **nový projekt** s názvem **Crypto Exam Generator**. Vyžaduje GitHub CLI (`gh`) a přihlášení.

```bash
# 0) Přihlášení do GitHubu (jednorázově, interaktivně v prohlížeči)
gh auth login -w -s 'repo,workflow'

# 1) Inicializace gitu
git init

# 2) Vytvoření .gitignore (viz níže) a přidání souborů
git add main.py README.md .gitignore

# 3) První commit
git commit -m "feat: initial commit for Crypto Exam Generator v1.1"

# 4) Vytvoření vzdáleného repozitáře a push (privátní)
gh repo create crypto-exam-generator --private --source=. --remote=origin --push
#   - nebo pokud už repo existuje:
# gh repo create <org>/crypto-exam-generator --private
# git remote add origin git@github.com:<org>/crypto-exam-generator.git
# git push -u origin main
```

---

## .gitignore (doporučené)

- Ignoruje složku `data/` (kde jsou otázky), virtuální prostředí, cache a macOS artefakty.

Viz přiložený `.gitignore`.

---

## Verze

- **Aktuální verze:** `1.1` (release: 2025-11-27)
- Changelog:
  - `1.1` – Import z DOCX, přesun otázek mezi skupinami.
  - `1.0` – První verze: GUI, skupiny/podskupiny, typy otázek, editor formátování, JSON úložiště.

---

## Licence

Zvol dle potřeby (např. MIT).
