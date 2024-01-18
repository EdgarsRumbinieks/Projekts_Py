# Projekts priekšmetā "Lietojumprogrammatūras automatizēšanas rīki"
Programatūrai ir nepieciešams saņemt csv failu ar kādiem statistiskiem datiem (vēlams no [stat.gov.lv](https://stat.gov.lv/lv)) un jāizveido xlsx failu ar 3-4 tabulām, kurās būs attēlotas aprēķinātas vērtības saistītas ar sekojošam nākošas vērtības prognozēšanas metodēm:
*Slīdošās vidējas metode
*Eksponenciālas noguldīšanas metode
*Tendences metode
*Sezonālais modelis
Kā arī katrai metodei ir jāizveido papildus tabulu ar vidējo kvadrātisko kļūdu.
Papildus programatūra veido rezultātu tabulu, kurā, balstoties uz aprēķiniem, **cenšas** prognozēt ievietoto datu nākošas vērtības intervālu.
Svarīgi piebalstīt, ka programatūras uzdevums **nav** prognozēt nākošo vērtību, bet automatizēt aprēķinus un apkopot tos tabulās, lai, balstoties un tiem, varētu veikt turpmāku analizēšanu.

Zinot programatūras nolūku, var aprakstīt nepieciešamos projekta uzdevumus.
## Projekta uzdevumi
1. Uzzināt csv faila datu izvietošanas veidu (rindā vai kolonnā) ([stat.gov.lv](https://stat.gov.lv/lv) datos var tikties abi varianti).
2. Balstoties uz iepriekšējā punktā iegūtu informāciju, nolasīt no csv faila datus masīvā "data" un katras datas laiku masīvā "date".
3. Izveidot xslx failu.
4. Veikt slīdošās vidējas metodes aprēķinus un ievietot tos tabulā.
_**Slīdošas vidējas metode:** S ir kāpe, kura nevar būt lielāka par pusi no datu daudzuma, katrā nākoša vērtība aprēķinās, saskaitot vidējo aritmētisko no S iepriekšējam vērtībām. Piemēram data = [1,2,3,4], kāpe S ir 2, tad 5. vērtība ar tādu kāpi būs (data[4-2] + data[4-1])/2 = (3+4)/2  = 3.5_
5. Saglābāt nākoša perioda aprēķinātus datus masīvā "slid_data"
6. Aprēķināt slīdošās vidējas metodes vidējo kvadrātisko kļūdu un ievietot tabulā.
_**Vidēja kvadrātiskā kļūda:** salīdzina aprēķinātos metodē datus ar reālajiem, vērtējot metodes precizitāti. Aprēķināmā pēc formulas : (1/aprēķināto vērtību daudzums)×∑((reāla vērtība - aprēķinātā vērtība)^2)_
7. Saglabāt vidējo kvadrātisko kļūdu katrai kāpei(S) masīvā "slid_data_klud"
8. Veikt eksponenciālas noguldīšanas metodes aprēķinus un ievietot tos tabulā.
_**Eksponenciālas noguldīšanas metode:** Vērtība = α*iepriekšēja reālā vērtība + (1-α)×aiziepriekšēja reāla vērtība, kur α ir koeficients, kurš pieder intervālam (0;1). Piemēram data = [1,2,3,4], α = 0.1, tad 5. vērtība ar tādu α būs 0.1×data[4-1] + 0.9×data[4-2] = 0.1×4 + 0.9×3 = 0.4 + 2.7 = 3.1_
9. Saglābāt nākoša perioda aprēķinātus datus masīvā "ekspon_data".
10. Aprēķināt eksponenciālas noguldīšanas metodes vidējo kvadrātisko kļūdu un ievietot tabulā.
11. Saglabāt vidējo kvadrātisko kļūdu katram α masīvā "ekspon_data_klud"
12. Veikt tendences metodes aprēķinus un ievietot tos tabulā. 
_**Tendences metode:** Tendences līknes (trendline) izveidošana (polinomiska, logoritmiska, exponinciāla... funkcijas). Izveidot sarakstu ar katras reālas vērtības kārtas numuru un ievietot to iegūtajā funkcija, lai atrastu vērtību. Piemēram data = [0.3,21,52.4,99.6], ieveidojot tndlinājās polinometrisku(2) funkciju, saņemsim y = 6.2x^2 + 2.1x - 8, no tās var aprēķināt 5. vērtību: 6.2×5^2 + 2.1×5 - 8 = 157.5_
13. Saglābāt nākoša perioda aprēķinātus datus masīvā "trend_data_klud".
14. Atrēķināt tendences metodes vidējo kvadrātisko kļūdu un ievietot tabulā.
15. Saglabāt vidējo kvadrātisko kļūdu katrai funkcijai masīvā "trend_data_klud"
16. Pajautājot lietotajam informāciju par laika posmu starp katru periodu (mēnesis, kavartile vai gads).
17. Pajautājot lietotājam informāciju par pirmā perioda mēnesi vai kvartili (7. mēnesis / 2. kvartile...)
18. Balstoties uz iegūtiem iepriekšējos 2 jautājumos datiem, veikt sezonālas modeļa aprēķinus un ievietot tos tabulā (ja laika posms starp periodiem nav gads)
_**Sezonālais modelis:** izveidot lineāro tendences līkni un aprēķināt individuālo indeksu katrai datai, lineāro vērtību dalot ar reālo. Pēc tām saskaitīt visus individuālus indeksus, kuriem ir viena un tā pati sezona (mēnesis/kvartile) un padalīt summu ar summējošo vērtību skaitu, aprēķinot vidējo aritmētisko (sezonas indeksu). Perioda vērtību iegūst sareizinot lineāro vērtību ar tai atbilstošas sezonas indeksu._
19. Saglābāt nākoša perioda vērtību mainīgajā "sezon_data"
20. Aprēķināt sezonāla moduļa vidējo kvadrātisko kļūdu un saglabāt to mainīgajā "sezon_data_klud"
21. Saglabāt atsevišķajā mainīgajā minimālas vidējas kvadrātiskas kļūdas indeksu no katra masīva (kurā bija saglabātas kļūdas)
22. Atrast katras metodes labāku vienību (vienību ar vismazāko vidējo kvadrātisko kļūdu).
23. Ievadīt visas labākas vērtības, kā arī to kļūdas rezultātu tabulā.
24. Aprēķināt nākošas vērtības intervālu (vērtība - sqrt(vidēja kvadrātiska kļūda) ; vērtība + sqrt(vidēja kvadrātiskā kļūda))

## Projektā izmantotas bibliotēkas
Lai izstrādātu programatūru bija izmantotas 3 bibliotēkas:
* **openpyxl** - bibliotēka ļauj veidot un rediģēt xslx failus, tāpēc tā ir nepieciešama projektam, kūtra gala rezultāts ir izveidot xlsx failu.
* **math** - bibliotēka ļauj izmantot vairākas matemātiskas funkcijas, tomēr šajā programatūrā izmantojas tikai funkcija _math.floor()_
* **numpy** - bibliotēka ļauj izmantot vairākas saistītas ar matemātiku funkcijas. Bibliotēkas funkcijas _numpy.array()_ un _numpy.polyfit()_ atrodas Tendences metodes pamatā, jo ļauj veidot tendences līniju vienādojumus. Funkcija _numpy.argmin()_ atviegloja masīva minimālas vērtības indeksa atrašanu.

## Programatūrai uzrakstītas funkcijas
Failā "funkcijas.py" ir saglabātas vairākas funkcijas:
* **Xlsx_Letter(number):** Ievadot skaitli, atgriež burtu kombināciju. Ekselē x-ass virzienā atrodas burtu kombinācijas ar kuram nevar veikt algebriskas darbības, tāpēc bija izstrādāta funkcija, ar kuras palīdzību vatēs noradīt Excel rūtiņas koordinātu, kā divu skaitļu kombināciju. Piemēram workbook['AB12'] vietā var rakstīt wokbook[Xlsx_Letter(28) + '12']. Tāda forma ļauj viegli veikt algebriskas darbības ar abam rūtiņas koordinātām. Funkcija spēj konvertēt skaitli burtu kombinācijā kopš līdz 18278 vai 'ZZZ'
* **DataDisplay(workbook,date,data,x,y):** Funkcija ieraksta iedotajā workbookā kolonu no 'date' datiem un kolonu no 'data' datiem. Koordināta (x;y) ir tabulas kreisā augšēja vieta, kur atrodas pirmā vērtība no 'date' masīva. Funkcija arī ieraksta kolonas 'data' nosaukumu 'y' virs minēta stabiņa. DataDisplay ir nepieciešama kā papildelements katrai tabulai, lai Excel failā varētu viegli saprast kādam laika posmam un kādai vērtībai attiecas prognoze / kļūda.
* **SlidPapild(workbook,x,y,number):** papildus funkcija slīdošas vidējas metodei ar kuru palīdzību tik skaitītas visas nepieciešamas metodei vērtības. x ir x koordināta, kur atrodas visi 'data' masīva elementi ierakstīti ar DataDisplay palīdzību, y ir y koordināta, kura ir uz 1 mazākā nekā vieta, kur ir jāieraksta aprēķināto vērtību, number ir kāpe (S), par kuru bija uzrakstīts nodaļā _"Projekta uzdevumi"_.
* **VidSlid(workbook,date,data,x,y):** Izveido slīdošas vidējas metodes tabulu un ievieto to tā, lai koordināte (x,y) būtu kreisa augšēja tabulas daļa, kur ir ierakstīta pirmā vērtība. Kā arī ievieto katras kolonas nosaukumu (kāpi) virs tai atbilstošai kolonai un saglabā pēdējā (nākošā) perioda prognozi masīvā.
* **TendencesMetode(workbook,date,data,x,y):** Izveido tendences metodes tabulu, kā arī kolonu, kurā ir ierakstīs katras masīva 'data' katra elementa kārtas numurs un ievieto tās tā, lai koordināte (x,y) būtu kreisa augšēja tabulas daļa, kur ir ierakstīta pirmā vērtība (pirmais kārta numura: '1'). Vērtības funkcija iegūst, ievietojot atbilstošo kārtas numuru, ar numpy bibliotēkas palīdzību izveidotājā tendences līnijas vienādojuma. Kā arī ievieto katras kolonas nosaukumu (polinoma pakāpi) virs tai atbilstošai kolonai. Virs kolonas ar kārtas numuriem ieraksta nosaukumu 't'. Saglabā pēdējā (nākošā) perioda prognozi masīvā.
* **EksponNoglud(workbook,date,data,x,y):** Izveido eksponenciālas noguldīšanas metodes tabulu un ievieto to tā, lai koordināte (x,y) būtu kreisa augšēja tabulas daļa, kur ir ierakstīta pirmā vērtība. Kā arī ievieto katras kolonas nosaukumu (α) virs tai atbilstošai kolonai un saglabā pēdējā (nākošā) perioda prognozi masīvā.
* **SezonMet(workbook,date,data,number,first_sezon,x,y):** Izveido sezonāla modeļa tabulu un ievieto to tā, lai koordināte (x,y) būtu kreisa augšēja tabulas daļa, kur ir ierakstīta pirmā vērtība. Funkcija izveido tādas kolonnas:
1. lineāras tendences līknes vērtības ar nosaukumu 'linear'
2. tekoša sezona katrai vērtībai (aprēķināšanai ir nepieciešams zināt laika periodu starp katru vērtību (ja mēnesis number = 12, ja kvartile number = 4), kā arī pirmo sezonu (first_sezon)). Kolonas nosaukums 'Sezona'
3. visus individuālos indeksus ar nosaukumu 'i'
4. katras sezonas indeksus ar nosaukumu 'i sez'
5. prognozes vērtības ar nosaukumu 'y~'
Funkcija saglabā pēdējā (nākošā) perioda vērtību mainīgajā.
* **KvadKlud(workbook,date,data,x_sak,y_sak,x_beig,y_beig):** Visam vērtībām, kas atrodas taisnstūrī, kuru var aprakstīt ar divām koordinātēm (x_sak ; y_sak) un (x_beig ; y_beig), tiek aprēķinātā kvadrātisko kļūdu un ievietoto to tabulā, kura atrodas 4 x vērtības pa labi no x_beig. Funkcija arī ievieto katras kolonas nosaukumu, kurš izskatās tā: 'Q' + nosaukums virs kolonnas, kas atrodas iezīmētājā taisnstūrī. Zem izveidotas tabulas paradās rinda, ar nosaukumu 'AVERANGE', kurā ir ierakstīta vidēja vērtībā no visas kolonas ar kvadrātiskam kļūdām. Šīs rindas vērtības tiek saglabātas masīvā.
* **sezon_input(), first_sezon_input() un try_to_name():** funkcijas, kas cenšas novērst iespējamas kļūdas saistītas ar lietotāja mijiedarbību ar teksta ievadi. Funkcijas neļauj lietotājam ievadīt neprasītas vērtības.
## Izmantošanas instrukcija
1. Lejupladēt kādu statistiku (vēlams no [stat.gov.lv](https://stat.gov.lv/lv)) **csv** formātā failā
2. Ievietot lejupladēto csv failu programatūras mapē un **pārdēvēt** to par "data.csv"
3. Pārbaudīt datu struktūru:
**Ja dati ir novietoti rindā:**
* Pārliecināties, ka datumi ir 3-ajā rinda
* Pārliecināties, ka dati ir 4-ajā rindā
**Ja dati ir novietoti kollonā:**
* Pārliecināties, ka datumi ar datiem sākas no 4. rindas (neskaitot nosaukumu)
* Pārliecināties, ka katram datumam ir piešķirts skaitlis (ja kādam datumam nav vērtības, tad ir jānodzēš visas rindas pirms tās)
4. Palaist failu 'Galvenais.py'
5. Kad programma paprasīs: ievadīt '1' ja dati ir rindā, ievadīt '2' ja dati ir kolonā
6. Kad programma paprasīs: ievadīt vēlamo polinoma pakāpi līdz kuram veikt tendences metodes rēķināšanu. Rekomendējas neievadīt skaitli lielāko par 6, tomēr, ja ir daudz datu, iespēja ievadīt lielākās pakāpes var būt noderīgā.
7. Kad programma paprasīs: ievadīt intervālu starp datiem - "1", ja mēnesis, "2", ja kvartile, "3", ja gads
8. Gadījumā, ja 7. punktā bija ievadīts "1" vai "2", kad programma paprasīs: ievadīt pirmā perioda numuru (ja februāris - 2, ja 3. kvartile - 3 ...)
9. Kad programma paprasīs: atbildēt uz jautājumu: "Vai saglabāt xslx failu", ierakstot "yes" vai "no"
10. Ja 9. punktā bija atbildēts "yes", kad programma paprasīs: ievadīt rezultātu xlsx faila nosaukumu
11. Analizēt rezultāta ticamību (itīpaši, ja metode ar mazāko vidējo kvadrātisko kļūdu ir tendences ar lielas pakāpes polinomu):
Ja uzskatāt, ka rezultāts nav ticams: ir jāatver saglabāto xslx failu un dziļāk jāanalizē rezultātus.
## Programatūras darbības video demonstrācija
CSV faila dati ir rindā: [links](https://youtu.be/QojjvgCOhhM)
CSV faila dati ir kolonnā: [links](https://youtu.be/1PphHRiI83s)

Programatūras demonstrācija: [links](https://youtu.be/mQZy5TFIO2Q)