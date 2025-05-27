# Filmu dienasgrāmata — automatizēta statistika

## Projekta uzdevums

Šī programma automatizē filmu dienasgrāmatas pārvaldību un analīzi, izmantojot Microsoft Excel failu. Katru reizi, kad tiek pievienoti jauni dati par noskatītajām filmām, programma saglabā esošo ierakstu kopiju ar šodienas datumu un analizē:

- Cik filmas ir ierakstītas;
- Cik bieži lietotāja vērtējums sakrīt, atšķiras vai pārsniedz IMDb reitingu;
- Populārāko žanru skatīto filmu vidū.

Tas palīdz sekot saviem kinoieradumiem un analizēt, kā tie mainās laika gaitā.

## Izmantotās Python bibliotēkas

### `openpyxl`
Izmanto Excel failu atvēršanai, rakstīšanai un rediģēšanai:
- `load_workbook()` – lai ielādētu esošo Excel dokumentu.
- `copy_worksheet()` – lai veidotu datuma kopiju.
- Datu iterācija, vērtību pievienošana un notīrīšana.

### `datetime`
- Nodrošina šodienas datumu, lai izmantotu to kā lapas nosaukumu.

### `collections.Counter`
- Tiek izmantots, lai noskaidrotu, kurš žanrs tiek skatīts visbiežāk.

## Pašdefinētas datu struktūras

Programmas loģikā izmantots:
- Žanru skaitītājs (`Counter`)
- Loģika, kas klasificē IMDb vs lietotāja reitingus

## Lietošanas instrukcija

1. Atver Excel failu ar nosaukumu `MovieDiary.xlsx`, un ievadi filmas datus (Žanrs, Nosaukums, IMDb reitings, Tavs reitings, Noskatīšanās datums).
2. Palaid Python skriptu `movie_diary.py`.
3. Programma:
   - Izveidos jaunu lapu ar datuma nosaukumu;
   - Nokopēs datus;
   - Analizēs rezultātus;
   - Notīrīs vecos ierakstus, lai vari ievadīt nākamos.

## Rezultāts

- Tiek saglabātas visas dienas vēstures lapas;
- Statistika redzama kolonnās H un I jaunajā lapā;
- Ērts veids, kā sekot līdzi skatīšanās paradumiem.

