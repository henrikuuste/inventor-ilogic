# Moodulite väljastamine (Module Release)

See kaust sisaldab iLogic skripte moodulite väljastamiseks Vaulti.

## Kasutamine

1. Ava alusmooduli põhikoost (nt `Alusmoodulid/Kapp-A/Kapp-A.iam`)
2. Veendu, et Excel fail `moodulid.xlsx` on kaustas olemas
3. Käivita reegel **Loo moodulid**
4. Kontrolli plaani ja kinnita väljastamine
5. Failid luuakse kausta `Moodulid/`

## Excel faili formaat

Excel fail peab olema nimega `moodulid.xlsx` ja asuma alusmooduli kaustas (`Alusmoodulid/{AlusmooduliNimi}/moodulid.xlsx`).

### Kohustuslik veerg

| Veeru nimi | Kirjeldus |
|------------|-----------|
| `MooduliNimi` | Mooduli nimi (kasutatakse kausta nimeks) |

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

| MooduliNimi | Laius | Kõrgus | Sügavus | _Märkus |
|-------------|-------|--------|---------|---------|
| Kapp-800 | 800 mm | 2100 mm | 600 mm | Standard |
| Kapp-1000 | 1000 mm | 2100 mm | 600 mm | Lai |
| Kapp-1200 | 1200 mm | 2400 mm | 600 mm | Kõrge |

### Ühe mooduliga alusmoodul

Kui alusmoodulil on ainult üks moodul (parameetreid ei muudeta), loo Excel fail ühe reaga:

| MooduliNimi |
|-------------|
| Kapp-A |

## Väljundfailide struktuur

```
Moodulid/
├── Ühine/                    # Jagatud detailid (sama geomeetria kõigis moodulites)
│   ├── 00001.ipt
│   └── 00002.ipt
├── Kapp-800/                 # Moodulispetsiifilised failid
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

Kasuta malli `_moodulid_template.xlsx` uue Excel faili loomiseks.
