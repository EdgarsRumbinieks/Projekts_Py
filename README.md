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
_**Vidēja kvadrātiskā kļūda:** salīdzina aprēķinātos metodē datus ar reālajiem, vērtējot metodes precizitāti. Aprēķināmā pēc formulas : (1/aprēķināto vērtību daudzums)*∑((reāla vērtība - aprēķinātā vērtība)^2)_
7. Saglabāt vidējo kvadrātisko kļūdu katrai kāpei(S) masīvā "slid_data_klud"
8. Veikt eksponenciālas noguldīšanas metodes aprēķinus un ievietot tos tabulā.
_**Eksponenciālas noguldīšanas metode:** Vērtība = α*iepriekšēja reālā vērtība + (1-α)*aiziepriekšēja reāla vērtība, kur α ir koeficients, kurš pieder intervālam (0;1). Piemēram data = [1,2,3,4], α = 0.1, tad 5. vērtība ar tādu α būs 0.1*data[4-1] + 0.9*data[4-2] = 0.1*4 + 0.9*3 = 0.4 + 2.7 = 3.1_
9. Saglābāt nākoša perioda aprēķinātus datus masīvā "ekspon_data".
10. Aprēķināt eksponenciālas noguldīšanas metodes vidējo kvadrātisko kļūdu un ievietot tabulā.
11. Saglabāt vidējo kvadrātisko kļūdu katram α masīvā "ekspon_data_klud"
12. Veikt tendences metodes aprēķinus un ievietot tos tabulā. 
_**Tendences metode:** Tendences līknes (trendline) izveidošana (polinomiska, logoritmiska, exponinciāla... funkcijas). Izveidot sarakstu ar katras reālas vērtības kārtas numuru un ievietot to iegūtajā funkcija, lai atrastu vērtību. Piemēram data = [0.3,21,52.4,99.6], ieveidojot tndlinājās polinometrisku(2) funkciju, saņemsim y = 6.2*x^2 + 2.1*x - 8, no tās var aprēķināt 5. vērtību: 6.2*5^2 + 2.1*5 - 8 = 157.5_
13. Saglābāt nākoša perioda aprēķinātus datus masīvā "trend_data_klud".
14. Atrēķināt tendences metodes vidējo kvadrātisko kļūdu un ievietot tabulā.
15. Saglabāt vidējo kvadrātisko kļūdu katrai funkcijai masīvā "trend_data_klud"
16. Pajautājot lietotajam, uzzināt laika posmu starp katru periodu (mēnesis, kavartele vai gads).
17. Pajautājot lietotājam, uzzināt pirmā perioda mēnesi vai kvarteli (7. mēnesis / 2. kvartele...)
18. Balstoties uz iegūtiem iepriekšējos 2 jautājumos datiem, veikt sezonālas modeļa aprēķinus un ievietot tos tabulā (ja laika posms starp periodiem nav gads)
_**Sezonālais modelis:** izveidot lineāro tendences līkni un aprēķināt individuālo indeksu katrai datai, lineāro vērtību dalot ar reālo. Pēc tām saskaitīt visus individuālus indeksus, kuriem ir viena un tā pati sezona (mēnesis/kvartele) un padalīt summu ar summējošo vērtību skaitu, aprēķinot vidējo aritmētisko (sezonas indeksu). Perioda vērtību iegūst sareizinot lineāro vērtību ar tai atbilstošas sezonas indeksu._
19. Saglābāt nākoša perioda vērtību mainīgajā "sezon_data"
20. Aprēķināt sezonāla moduļa vidējo kvadrātisko kļūdu un saglabāt to mainīgajā "sezon_data_klud"
21. Saglabāt atsevišķajā mainīgajā minimālas vidējas kvadrātiskas kļūdas indeksu no katra masīva (kurā bija saglabātas kļūdas)
22. Atrast katras metodes labāku vienību (vienību ar vismazāko vidējo kvadrātisko kļūdu).
23. Ievadīt visas labākas vērtības, kā arī to kļūdas rezultātu tabulā.
24. Aprēķināt nākošas vērtības intervālu (vērtība - sqrt(vidēja kvadrātiska kļūda) ; vērtība + sqrt(vidēja kvadratīska kļūda))

## Projektā izmantotas bibliotēkas
Lai izstrādātu programatūtu bija izmantotas 3 bibliotēkas:
* **openpyxl** - bibliotēka ļauj veidot un redigēt xslx failus, tāpēc tā ir nepieciešama projektam, kutra gala rezultāts ir izveidot xlsx failu.
* **math** - bibliotēka ļauj izmantot vairākas matematiskas funkcijas, tomēr šajā kodā izmantojas tikai funkcija _math.floor()_
* **numpy** - bibliotēka ļauj izmantot vairākas saistītas ar matematīsku funkcijas. Bibliotēkas funkcijas _numpy.array()_ un _numpy.polyfit()_ atrodas Tendences metodes pamatā, jo ļauj veidot tendences līniju vienadojumus.