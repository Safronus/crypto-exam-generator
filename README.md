# Crypto Exam Generator

## v8.4.3b — 2026-01-06
- Filtr: po vymazání filtru se obnoví stav sbalení/rozbalení stromu tak,
  jak byl před prvním použitím filtru.

## v8.4.3 — 2026-01-06
- Filtr otázek: pokud filtr najde otázky v **sbalených** podskupinách, tyto se nyní **automaticky rozbalí**,
  aby byly všechny shody okamžitě viditelné. Původní logika skrývání se nemění.
  
## v8.4.2a — 2026-01-06
- Start aplikace – volba DB: vylepšené popisky file dialogů.
- Volba „Nahrát jinou DB a předtím uložit aktuální DB“ nyní otevře druhý dialog **Uložit jako…**,
  kde lze zvolit **umístění i název** zálohy. Výchozí návrh je `data/backups/<název>-backup-<timestamp>.json`.

## v8.4.1 — 2026-01-06
- Import z DOCX: nová volba **rozsahu kontroly duplicit** – buď proti **celé databázi** (původní chování),
  nebo **jen proti cílové podskupině**. Výběr v jednoduchém dialogu.

## v8.4.0 — 2026-01-06
- Strom „Otázky“: sloupec **Typ / body** se správně vyplní hned po startu aplikace.
- Fix: Po editaci otázky se ve sloupci **Typ / body** u klasických otázek ztrácelo `b.`.
  Nyní se automaticky doplní, pokud chybí (např. „Klasická | 1“ → „Klasická | 1 b.“).

## v8.3.7d — 2026-01-06
- Drag&drop („vyhození do koše“): po smazání se **zachová** stav sbalení/rozbalení stromu,
  během obnovy UI je potlačeno auto-rozbalení.
  
## v8.3.7c — 2026-01-06
- Drag&drop: po přesunu se **zachová** stav sbalení/rozbalení stromu (žádné auto-rozbalení).

## v8.3.7b — 2026-01-06
- Přesun vybraných (tlačítko i kontextová akce): po přesunu se **zachová** stav sbalení/rozbalení stromu.
- Žádné automatické rozbalení dalších větví (strom vypadá stejně jako před akcí).

## v8.3.6 — 2026-01-06
- Import z DOCX: po importu se zachová stav sbalení/rozbalení stromu.
- Rozbalí se **pouze** podskupina, do které byl import proveden (tlačítkem *Import* i přes kontextovou akci).

## v8.3.5 — 2026-01-06
- Kontextové menu podskupiny: přidána akce **„Import z DOCX do této podskupiny…“**.
- Import z DOCX nyní umí jednorázové **předvolení cíle**, takže lze přeskočit dialog s výběrem skupiny/podskupiny.

## v8.3.4 — 2026-01-06
- Strom otázek: přidána kontextová akce **„Přejmenovat…“** pro skupiny i podskupiny.
  Přejmenování probíhá s uchováním stavu sbalení/rozbalení stromu.
  
## v8.3.3 — 2026-01-06
- Strom otázek: při přejmenování skupiny/podskupiny se zachová stav sbalení/rozbalení (žádné hromadné auto-rozbalení).

## v8.3.2 — 2026-01-06
- Při přidání podskupiny se po obnově stavu rozbalí **jen rodičovská větev** (skupina/podskupina),
  pokud byla před akcí **sbalená**. Ostatní větve zůstávají beze změny.

**Verze:** v8.3.0 · **Platforma:** macOS · **GUI:** PySide6

Crypto Exam Generator je desktopová aplikace pro správu otázek (skupiny → podskupiny → otázky) a generování **DOCX** dokumentů ze šablon.
Zaměřuje se na **konzistentní vzhled** (dark theme), **spolehlivý export** (dědění formátu šablony), a **rychlou práci** se stromem otázek
(duplikace, přesuny, hromadné mazání, historie exportů, „vtipné odpovědi“ aj.).

---

## ✨ Klíčové funkce

- **Strom otázek** se strukturou *Skupina → Podskupina → Otázka*.
- Kontextové akce nad **otázkami**:
  - **Duplikovat otázku** (vkládá kopii do stejné podskupiny).
  - **Duplikovat do podskupiny** (vybereš cílovou podskupinu v dialogu).
  - **Přidat otázku** (do vybrané skupiny/podskupiny; automaticky vytvoří „Default“ podskupinu, pokud chybí).
  - **Smazat vybrané** (hromadně – otázky/podskupiny/skupiny; přesune do „koše“).
  - (Volitelné) **Přesunout do…** – přes dialog výběru cíle (pokud je v projektu aktivní).
- **Zachování rozbalení stromu** při duplikaci/přidání/smazání: strom zůstane ve **stejném stavu**, případně se rozbalí jen **cílová podskupina**.
- **Perzistence rozbalení** stromu mezi **restarty aplikace** (per projekt; QSettings).
- **Export do DOCX** se zachováním **fontu a velikosti** písma ze šablony (placeholderu):
  - Inline i blokové nahrazování (včetně odrážek/číslování).
  - Přenášejí se styly **b/i/u/barva** z obsahu, ale font a velikost **přebírá šablona**.
  - Zachování **page breaků**.
  - Obrázky: vložení s volitelnou velikostí (cm); **HEIC/HEIF** se na macOS převádí přes `sips`.
- **Historie exportů**: přehled v záložce **Historie** se sloupci *Typ, Cílový soubor, Hash, Časová stopa*,
  tříděno **podle „Časová stopa“ sestupně** (nejnovější nahoře).
- **Koš**: smazané otázky se evidují v interní struktuře (pro pozdější kontrolu/diagnostiku).

---

## 🧩 Instalace (macOS)

> Doporučeno: **Python 3.10+** (ověřeno na 3.11).

1) Vytvoř a aktivuj virtuální prostředí:
```bash
python3 -m venv .venv
source .venv/bin/activate
```

2) Nainstaluj závislosti (minimum):
```bash
pip install -U pip
pip install PySide6 python-docx
```
> `python-docx` vyžaduje `lxml`, které se nainstaluje automaticky.

3) (Volitelně) Ulož požadavky:
```bash
pip freeze > requirements.txt
```

---

## ▶️ Spuštění

```bash
source .venv/bin/activate   # pokud ještě neběží
python3 main.py
```

Aplikace používá **dark theme** a je optimalizovaná pro **HiDPI/Retina** na macOS.

---

## 📂 Struktura projektu (orientačně)

```
.
├── main.py                 # Celá aplikace (GUI, logika exportu, práce se stromem, aj.)
├── data/
│   ├── history.json        # Historie exportů
│   └── ...                 # Další data projektu
├── templates/
│   └── template.docx       # Šablona(y) pro export
└── README.md               # Tento soubor
```

> **Pozn.:** Per-projektová perzistence rozbalení stromu využívá `QSettings` (klíč podle hash cesty projektu).

---

## 🌳 Práce se stromem (Skupiny/Podskupiny/Otázky)

- Strom zobrazuje **Skupiny** (top-level), jejich **Podskupiny** (libovolně do hloubky) a v nich **Otázky**.
- **Kontextové menu** nad *otázkou* obsahuje:
  - **Duplikovat otázku** – vytvoří kopii v **téže** podskupině.
  - **Duplikovat do podskupiny** – vybereš cílovou skupinu/podskupinu v dialogu (kopie se vloží tam).
  - **Přidat otázku** – vloží novou „classic“ otázku; pokud má skupina 0 podskupin, vytvoří se „Default“.
  - **Smazat vybrané** – hromadně (otázky/podskupiny/skupiny). Záznamy jdou do interního „koše“.
- Při **duplikaci/přidání/smazání**:
  - Aplikace dočasně potlačí „auto-expand“ během obnovy stromu, **obnoví původní rozbalení** a **případně rozbalí jen cílovou podskupinu**,
    pokud byla před akcí sbalená.
- Při **zavření a znovu otevření** aplikace se stav rozbalení stromu **obnoví** tak, jak byl před zavřením.

---

## 📝 Export do DOCX (šablony, placeholdery)

### Princip
- Aplikace načte DOCX **šablonu** a nahradí **placeholdery** (např. `<Otazka1>`) konkrétním obsahem.
- Nahrazení zvládá **inline** i **blokové** formy (tzn. i odrážky/číslované seznamy).

### Důležité o formátu písma
- Text vkládaný na místo placeholderu **přebírá font a velikost písma šablony** (tj. **font-family** a **size** určené v placeholderu/stylu).
- Vložený obsah může mít **b**, **i**, **u** a **barvu** – tyto styly se uplatní, ale **font a velikost řídí šablona**.
- U **odrážek/číslování** se zachovává číslování a formát (kopíruje se úroveň/numPr z placeholderu).

### Obrázky
- Vloží se jako `add_picture(...)` s volitelnou **šířkou/výškou v cm**.
- Soubory **HEIC/HEIF** se na **macOS** převádějí nástrojem `sips` na JPEG automaticky.

### Page breaky
- Stránkové zlomky z místa placeholderu se **zachovají** (extrakce a obnova před/po změnách odstavce).

---

## 📜 Historie exportů

- Záložka **Historie** ukazuje přehled exportů se sloupci: **TYP**, **CÍLOVÝ SOUBOR**, **DIGITÁLNÍ OTISK (HASH)**, **ČASOVÁ STOPA**.
- Záznamy jsou **seřazeny** podle **„Časová stopa“ sestupně** (nejnovější nahoře).
- Záznamy „balíků“ exportů (více souborů) zobrazují počet kusů `NNx` a `hash` jako „(více)“.

---

## 🗑️ Koš (Trash)

- Při hromadném mazání (`Smazat vybrané`) se dotčené otázky **zapisují do koše** s metadaty (čas smazání, zdrojová skupina/podskupina aj.).
- Slouží k auditu/diagnostice. (Obnova může být projektově specifická.)

---

## ⚙️ Nastavení & perzistence

- **QSettings** (per projekt – klíč z hash cesty) ukládá **sadu viditelně rozbalených** uzlů stromu.
- Při startu se stav **obnoví** po prvním vykreslení okna (bez rebuild UI).

---

## 🧰 Požadavky a závislosti

- **macOS** (doporučeno 12+), **Python 3.10+**.
- Základní balíčky:
  - `PySide6` – GUI.
  - `python-docx` – generování DOCX.
  - (automaticky) `lxml` – XML vrstvy pro `python-docx`.
- **sips** (součást macOS) pro konverzi **HEIC/HEIF → JPEG**.

---

## 🐞 Troubleshooting

- **Vložený text „skáče“ na 12 pt:** Zkontroluj, že placeholder/odstavec má v šabloně **explicitní** velikost (`w:sz`) nebo styl s velikostí. Aplikace kopíruje `rPr` a výslovně nastavuje `font.size` z placeholderu.
- **HEIC se nevloží:** Ověř, že na macOS je dostupný `sips` (standardně je), případně vlož JPEG/PNG.
- **Strom se po akci „rozsype“:** Zachování stavu je aktivní při **duplikaci/přidání/smazání** i **mezi relacemi**. Pokud by se choval jinak, zkontroluj, zda nebyla ručně volána metoda, která strom hromadně expanduje.

---

## 🧪 Smoke test (ruční)

1. Spusť aplikaci, vytvoř **Skupinu** a **Podskupinu**, přidej **Otázku**.
2. V šabloně nastav **placeholder** s konkrétní velikostí písma (např. Calibri 9) a proveď **export** – ve výstupu musí být **stejná velikost**.
3. **Duplikuj otázku** a **Duplikuj do podskupiny** – strom se **nezmění** (rozbalí se jen cílová větev, pokud byla sbalená).
4. **Smaž vybrané** otázky – strom si **zachová** rozbalení; položky se objeví v **Koši**.
5. Zavři a znovu otevři aplikaci – strom je **ve stejném** stavu rozbalení.

---

## 📦 Changelog (výběr)

- **v8.3.0** — Perzistence rozbalení stromu mezi relacemi (QSettings; per projekt).  
- **v8.2.5** — Přidání/Smazání otázek zachovává rozbalení; případně rozbalí jen cílovou podskupinu.  
- **v8.2.4** — Oprava ukládání stavu (ignorují se potomci sbalených uzlů).  
- **v8.2.3** — Obnovení přesně původního rozbalení po duplikaci.  
- **v8.2.2** — Potlačení auto-expand při obnově stromu; rozbalení jen cílové podskupiny.  
- **v8.2.0** — Nová akce **„Duplikovat do podskupiny“** v kontextovém menu u otázky.  
- **v8.1.3** — **Historie**: řazení podle **„Časová stopa“** sestupně (nejnovější první).  
- **v8.1.x** — Export DOCX: přebírání **fontu a velikosti** ze šablony (řeší výchozí 12 pt).

---

## 📝 Licence

Interní / dle projektu. (Neuvedeno v repozitáři.)

---

## 🙋 Podpora

Návrhy a bugreporty prosím posílej s co nejkratším popisem kroků + šablonou/ukázkou,
aby bylo možné problém rychle reprodukovat.
