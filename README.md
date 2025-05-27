Filmu dienasgrāmatas pārvaldības sistēma
Projekta apraksts
Šis projekts ir izstrādāts kursa "Datu struktūras un algoritmi" ietvaros, lai automatizētu filmu skatīšanās dienasgrāmatas pārvaldību, izmantojot Excel failu "MovieDiary.xlsx". Programmas galvenais mērķis ir:

Ielasīt datus par skatītajām filmām (nosaukums, žanrs, gads, IMDB vērtējums, lietotāja vērtējums) no Excel faila.
Aprēķināt statistiku:
Kopējais skatīto filmu skaits.
Vidējais lietotāja vērtējums.
Populārākais žanrs (balstoties uz visbiežāk sastopamo žanru).


Saglabāt statistiku Excel failā attiecīgajās šūnās (G2–G4).
Izveidot jaunu Excel lapu ar šodienas datumu, kurā saglabāta tabulas struktūra bez datiem, lai lietotājs varētu ievadīt jaunas filmas.
Saglabāt izmaiņas Excel failā, nodrošinot datu integritāti un lietošanas ērtumu.

Programma ir izstrādāta, lai būtu vienkārši lietojama un palīdzētu lietotājam sekot līdzi skatītajām filmām un to vērtējumiem. Tā izmanto pielāgotu datu struktūru, lai apstrādātu un validētu datus, kā arī nodrošina izturību pret kļūdainiem vai nepilnīgiem ierakstiem.
Izmantotās Python bibliotēkas
Projekta izstrādē tika izmantotas šādas Python bibliotēkas:

openpyxl: Šī bibliotēka tiek izmantota, lai lasītu un rakstītu Excel failus. Tā nodrošina funkcionalitāti, lai ielādētu esošu Excel failu (load_workbook), izveidotu jaunas lapas (copy_worksheet), manipulētu ar šūnām un saglabātu izmaiņas. openpyxl tika izvēlēts tā plašā atbalsta un ērtās lietošanas dēļ.
datetime: Izmantots, lai ģenerētu šodienas datumu (datetime.now()), kas tiek izmantots kā jaunās Excel lapas nosaukums. Tas nodrošina, ka katra jauna dienasgrāmata tiek saglabāta ar unikālu datumu.
openpyxl.styles: Izmantots, lai definētu un pielietotu formatēšanas stilus Excel šūnām (piemēram, treknrakstu galvenēm). Tas uzlabo tabulas lasāmību un vizuālo noformējumu.
collections.Counter: Izmantots, lai noteiktu populārāko žanru, analizējot filmu žanru biežumu. Šī bibliotēka nodrošina efektīvu veidu, kā saskaitīt elementus sarakstā.

Pielāgotās datu struktūras
Projektā tika izveidota pielāgota datu struktūra MovieDiary klase, kas apvieno sarakstu un vārdnīcu, lai uzglabātu un apstrādātu filmu datus. Šī struktūra:

Uzglabā katras filmas nosaukumu, žanru, izdošanas gadu, IMDB vērtējumu un lietotāja vērtējumu.
Validē datus, izslēdzot nepilnīgus ierakstus (piemēram, filmas bez nosaukuma vai lietotāja vērtējuma).
Nodrošina metodes statistikas aprēķināšanai, tostarp skatīto filmu skaita, vidējā lietotāja vērtējuma un populārākā žanra noteikšanu.

Datu struktūras izmantošana ļauj efektīvi apstrādāt datus pirms to ierakstīšanas Excel failā, samazinot kļūdu iespējamību un uzlabojot koda modularitāti.
Programmatūras izmantošanas metodes
Lai izmantotu programmu, veiciet šādas darbības:

Sagatavošana:

Pārliecinieties, ka Python vidē ir instalētas nepieciešamās bibliotēkas. Instalējiet tās, izmantojot komandu:pip install openpyxl


Nodrošiniet, ka fails "MovieDiary.xlsx" atrodas tajā pašā direktorijā, kurā tiek izpildīts Python skripts.


Datu ievade:

Atveriet "MovieDiary.xlsx" failu un aktīvajā lapā ievadiet datus par skatītajām filmām:
A kolonna: Filmas nosaukums (piemēram, "Dune: Part Two").
B kolonna: Žanrs (piemēram, "Zinātniskā fantastika").
C kolonna: Izdošanas gads (piemēram, 2024).
D kolonna: IMDB vērtējums (piemēram, 8.6).
E kolonna: Lietotāja vērtējums (piemēram, 8).


Saglabājiet failu.


Programmas palaišana:

Izpildiet Python skriptu movie_diary.py, izmantojot komandu:python movie_diary.py


Programma:
Ielasīs datus no aktīvās lapas.
Aprēķinās statistiku: skatīto filmu skaitu, vidējo lietotāja vērtējumu un populārāko žanru.
Ierakstīs statistiku šūnās G2 (skatītās filmas), G3 (vidējais vērtējums) un G4 (populārākais žanrs).
Izveidos jaunu Excel lapu ar šodienas datumu un nokopēs tabulas struktūru (bez datiem).
Saglabās izmaiņas failā "MovieDiary.xlsx".


Rezultātu apskate:

Atveriet "MovieDiary.xlsx" un pārbaudiet:
Aktīvajā lapā (ar šodienas datumu) būs tukša tabula jaunu filmu datu ievadei.
Iepriekšējā lapā būs saglabāta statistika šūnās G2–G4.



