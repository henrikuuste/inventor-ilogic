# Elementide väljastamine (Element Release)

See kaust sisaldab iLogic skripte elementide väljastamiseks Vaulti.

**Terminoloogia uuendatud 2026-05-12** (vt `docs/UBIQUITOUS_LANGUAGE.md`):
- "Alusmoodul" (vana) → "Aluselement" (base element)
- "Moodul" (vana) → "Väljastatud element" (released element)

## Kasutamine

1. Ava aluselemendi põhikoost (nt `Aluselemendid/Kapp-A/Kapp-A.iam`)
2. Veendu, et Excel fail `elemendid.xlsx` on kaustas olemas
3. Käivita reegel **Loo elemendid**
4. Kontrolli plaani ja kinnita väljastamine
5. Failid luuakse kausta `Elemendid/`

## Excel faili formaat

Excel fail peab olema nimega `elemendid.xlsx` ja asuma aluselemendi kaustas (`Aluselemendid/{AlusElement}/elemendid.xlsx`).

Tagasiühilduvuse jaoks toetatakse ka vana nimekuju `moodulid.xlsx`.

### Kohustuslik veerg

| Veeru nimi | Kirjeldus |
|------------|-----------|
| `Element` | Elemendi nimi (kasutatakse kausta nimeks) |

Tagasiühilduvuse jaoks toetatakse ka vana veerunime `MooduliNimi`.

### Parameetri veerud

Ülejäänud veerud on parameetri nimed, mis vastavad Inventori parameetritele.

| Veeru nimi | Kirjeldus |
|------------|-----------|
| `Laius` | Parameetri `Laius` väärtus |
| `Kõrgus` | Parameetri `Kõrgus` väärtus |
| `Sügavus` | Parameetri `Sügavus` väärtus |
| `_Märkus` | Erikoodid (algavad `_`) jäetakse vahele |

**NB:** Part Number genereeritakse automaatselt Vault numbriseemist, seda ei määrata Excelis.

### Näide

| Element | Laius | Kõrgus | Sügavus | _Märkus |
|--------------|-------|--------|---------|---------|
| Kapp-800 | 800 mm | 2100 mm | 600 mm | Standard |
| Kapp-1000 | 1000 mm | 2100 mm | 600 mm | Lai |
| Kapp-1200 | 1200 mm | 2400 mm | 600 mm | Kõrge |

### Ühe elemendiga aluselement

Kui aluselemendil on ainult üks väljastatud element (parameetreid ei muudeta), loo Excel fail ühe reaga:

| Element |
|--------------|
| Kapp-A |

## Väljundfailide struktuur

```
Elemendid/
├── Ühine/                    # Jagatud detailid (sama geomeetria kõigis elementides)
│   ├── 00001.ipt
│   └── 00002.ipt
├── Kapp-800/                 # Elemendispetsiifilised failid
│   ├── 00003.ipt
│   ├── 00004.iam
│   └── 00005.idw
├── Kapp-1000/
│   ├── 00006.ipt
│   ├── 00007.iam
│   └── 00008.idw
└── _manifest.json            # Väljastamise manifest
```

## Mall

Kasuta malli `_elemendid_template.xlsx` uue Excel faili loomiseks.
