# Vaulti kaustade struktuur

## Probleem

**Inventor ei luba hoida parameetrilisest mudelist mitut varianti erinevate parameetritega.**

Kui muudame parameetrilise istme laiuseks 900 mm, siis muutuvad *kõik* sellele viitavad koostud — ka need, mis peaksid jääma 800 mm laiuseks.

See tähendab, et ühe parameetrilise mudeli põhjal ei saa Vaultis hallata mitut tootevarianti korraga.

---

## Lahendus: Aluselemendid vs Elemendid

```
$/Tooted/<Tootepere>/
│
├── Aluselemendid/                 ← Parameetrilised tööfailid
│   ├── Käetugi/
│   │   ├── elemendid.xlsx         (väljastamise definitsioonid)
│   │   └── *.ipt, *.iam, *.idw    (parameetrilised mudelid)
│   ├── Selg/
│   └── Iste/
│
├── Elemendid/                     ← Väljastatud elemendid tootmiseks
│   ├── Ühine/                     (jagatud detailid elementide vahel)
│   ├── Käetugi_V/                 (vasak käetugi - fikseeritud)
│   │   ├── *.ipt, *.iam           (fikseeritud parameetritega)
│   │   ├── *.idw                  (joonised)
│   │   └── *.xlsx                 (BOMid)
│   ├── Käetugi_P/                 (parem käetugi - fikseeritud)
│   └── Iste_110/                  (iste variant 110 - fikseeritud)
│
├── Alusmoodulid/                  ← Parameetrilised moodulite tööfailid
│   └── ...
│
└── Moodulid/                      ← Väljastatud moodulid tootmisele
    └── ...
```

---

## Põhimõte

| Kaust | Otstarve | Parameetrid |
|-------|----------|-------------|
| **Aluselemendid** | Parameetrilised mudelid, millest luuakse konkreetseid elemente | Muudetavad |
| **Elemendid** | Tootmisele väljastatud elemendid | Fikseeritud |
| **Alusmoodulid** | Parameetrilised moodulite kooslused | Muudetavad |
| **Moodulid** | Tootmisele väljastatud moodulite kooslused | Fikseeritud |

**Väljastamine** = parameetrilisest mudelist luuakse fikseeritud koopia kindlate parameetriväärtustega.

- Iga väljastatud variant saab oma revisioninumbri
- Elemendid on üksteisest sõltumatud
- Parameetrilise mudeli muutmine ei mõjuta juba väljastatud elemente

---

## Miks just selline struktuur?

### 1. Selge eraldus töö- ja tootmisfailide vahel

- **Aluselemendid/Alusmoodulid** — siin toimub arendustöö
- **Elemendid/Moodulid** — siit läheb tootmisele

### 2. Revisjonihaldus

Väljastatud variandid saavad Vaultis korralikud revisioninumbrid. Parameetrilisi alusmudeleid ei revisjonita samamoodi — need on "töödokumendid".

### 3. Sõltumatus

Kui väljastame Iste_110, siis see jääb igavesti Iste_110-ks. Hiljem võime luua Iste_120 ilma, et vana variant muutuks.

---

## Sisemine struktuur (Aluselemendi sees)

```
Aluselemendid/<ElementName>/
├── Eskiis/          ← Skelett-osad ja visandid
├── Karkass/
│   ├── Detailid/
│   └── Joonised/
├── Poroloon/
│   ├── Detailid/
│   └── Joonised/
└── elemendid.xlsx   ← Väljastamise definitsioonid
```

Karkassi ja porolooni eraldamine on praktilisuse küsimus — neid töötatakse sageli eraldi.

**Eskiis** on Inventori spetsiifiline: skelett-osad, millele teised komponendid viitavad. Väljastatud elementides püüame selle kas peita või eemaldada.

---

## Kokkuvõte

> Aluselemendid sisaldavad parameetrilisi mudeleid, Elemendid sisaldavad fikseeritud tootmisvariante — see lahendab Inventori piirangu, et parameetriline mudel saab korraga olla ainult ühes olekus.
