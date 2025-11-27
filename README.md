# Crypto Exam Generator – Patch exportu DOCX (pro 1.8d)

Co je opraveno (jen export DOCX, nic dalšího):
- placeholdery se detekují ve **všech** `word/*.xml` (včetně header/footer) a nahradí se **1:1**
- `<OtázkaX>` a `<BONUSX>` se vkládají **do stejného odstavce** – číslování zůstane (více odstavců otázky → `w:br`)
- inline tokeny jako `<PoznamkaVerze>`, `<DatumČas>`, `<MinBody>`, `<MaxBody>` se nahradí i když jsou Wordem rozsekané do více `w:t`
- výchozí cesty šablony/výstupu zachovány

Soubor ke stažení a nahrazení:
- [Download main_patched.py](sandbox:/mnt/data/main_patched.py) → u sebe přepište `main.py`

