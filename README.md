# Filmu dienasgrāmata – skatīšanās statistika ar Excel

## Projekta uzdevums

Šī programma automatizē filmu dienasgrāmatas pārvaldību un statistikas analīzi, izmantojot Excel failu. Tā palīdz sekot līdzi skatīšanās paradumiem un analizēt, kā tie mainās laika gaitā.

Programmas iespējas:
- Saglabā dienas ierakstu kopijas ar datumu;
- Aprēķina kopējo filmu skaitu;
- Salīdzina IMDb un lietotāja reitingus;
- Nosaka populārāko žanru.

## Izmantotās Python bibliotēkas

- `openpyxl` – Excel failu apstrādei (`load_workbook`, `copy_worksheet`, šūnu iterācija).
- `datetime` – datuma iegūšanai jaunās lapas nosaukumam.
- `collections.Counter` – populārākā žanra aprēķinam.

## Pašdefinētas datu struktūras

- Skaitītāji (`film_count`, `higher_imdb`, `higher_user`, `equal_scores`);
- Žanru biežuma analizators (`Counter`);
- Loģika IMDb un lietotāja reitingu salīdzināšanai.

## Lietošanas instrukcija

1. Atver Excel failu `MovieDiary.xlsx`.
2. Aizpildi rindas ar šādiem datiem:
   - Žanrs
   - Filmas nosaukums
   - Gads
   - IMDb reitings
   - Tavs reitings
3. Palaid Python skriptu `movie_diary.py`.
4. Programma:
   - Izveidos jaunu lapu ar šodienas datumu;
   - Nokopēs datus;
   - Aprēķinās statistiku;
   - Ievietos to kolonnās F un G;
   - Notīrīs vecos datus (saglabājot galvenes).

## Rezultāts

- Katra diena tiek saglabāta atsevišķā lapā;
- Statistika redzama kolonnās F un G:
  - Filmu skaits
  - IMDb > Tavs reitings
  - Tavs > IMDb
  - Vienādi vērtējumi
  - Populārākais žanrs

---

