
Synthèse statistique des [prénoms mixtes](https://fr.wikipedia.org/wiki/Prénom_mixte)
---

Il n'existe pas de dictionnaire académique des prénoms (il existe seulement un « Officiel des Prénoms », maintenant disponible sur [Wikipédia](https://fr.wikipedia.org/wiki/Liste_de_prénoms_en_français)), car il s'agit d'une question de multiculturalisme lié à l'immigration. On pourrait établir un dictionnaire des prénoms français, qui définirait l'orthographe recommandée pour chaque prénom. Au niveau de la francophonie, ces dictionnaires pourraient éventuellement converger vers un consensus, ... ou pas ! Et au niveau des cultures en dehors de la francophonie, on ne pourrait tout simplement pas définir une norme unique pour toutes ces cultures. Imaginons que M. Groçon émigre à l'étranger, il voudra absolument que son ç cédille soit préservé, n'est-ce pas ? Mais est-ce qu'au moins on pourra saisir ce ç cédille dans le système informatique de l'état en question ? Il en est de même pour les prénoms d'autres cultures qui arrivent en France. Bref, c'est probablement la raison pour laquelle il y a tant d'orthographes fantaisistes comme on va le voir ici, dans un pays qui aime pourtant tellement les normes et les standards...

<!-- TOC -->

- [Téléchargement](#t%C3%A9l%C3%A9chargement)
- [Synthèse statistique des prénoms mixtes épicènes](#synth%C3%A8se-statistique-des-pr%C3%A9noms-mixtes-%C3%A9pic%C3%A8nes)
- [Synthèse statistique des prénoms mixtes homophones](#synth%C3%A8se-statistique-des-pr%C3%A9noms-mixtes-homophones)
- [Synthèse statistique des prénoms similaires](#synth%C3%A8se-statistique-des-pr%C3%A9noms-similaires)
- [Synthèse statistique des prénoms unigenres](#synth%C3%A8se-statistique-des-pr%C3%A9noms-unigenres)
- [Synthèse statistique des prénoms accentués](#synth%C3%A8se-statistique-des-pr%C3%A9noms-accentu%C3%A9s)
- [Synthèse statistique des prénoms fréquents](#synth%C3%A8se-statistique-des-pr%C3%A9noms-fr%C3%A9quents)
- [Liens externes](#liens-externes)

<!-- /TOC -->

# Téléchargement

Le fichier nat2020.csv peut être téléchargé ici :

https://www.insee.fr/fr/statistiques/2540004

https://www.insee.fr/fr/statistiques/fichier/2540004/nat2020_csv.zip

et doit être placé dézipé dans le dossier PrenomMixte\bin de la solution.

# Synthèse statistique des prénoms mixtes épicènes

On ne peut pas définir le type mixte sans prendre en compte au moins un seuil. D'une part un seuil concernant le nombre minimal d'occurrences du prénom parmi les 86 millions de prénoms attribués en France en un peu plus d'un siècle, et d'autre part la fréquence relative minimale entre masculin et féminin : un prénom vraiment mixte devrait au moins concerner par exemple plus de 1% des cas (1% de masculin et 99% de féminin, ou alors l'inverse), ou peut-être 10 % des cas (10% de masculin et 90% de féminin) en étant plus strict, pour avoir une véritable signification statistique. Si on n'applique aucun seuil, on arrive à près de 2 000 prénoms mixtes épicènes, mais chaque occurrence devient pratiquement une exception en soi. Voici les résultats avec un réglage du premier seuil à 2 000 occurrences minimum (avec une fréquence relative minimale de 1 %), soit 217 prénoms mixtes épicènes (voir le résultat [ici](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_mixtes_épicènes) pour une liste sans seuil minimum d'occurrences avec 1460 prenoms).

Le fichier de l'INSEE commence en 1900 et fini en 2020, certains prénoms n'ont pas les accents, d'autres n'ont pas de date (on trouve 'XXXX' à la place, soit 1 % des cas), ou bien ne sont pas identifiés (on trouve '_PRENOMS_RARES' à la place du prénom, soit 1.9 % des cas). On a 84 107 375 prénoms correctement identifiés sur 86 605 605 naissances répertoriées en France sur un peu plus d'un siècle.

Les colonnes sont Sexe (1:Garçon, 2:Fille);Prénom usuel;Année;Nombre d'occurrence, il contient 667 365 lignes de prénoms, par exemple :
2;BERANGERE;2010;5, soit 5 naissances de Bérangère en 2010. L'année moyenne est logiquement pondérée par le nombre d'occurrences pour chaque année. On corrige toujours les accents dans ces listes, même si parfois la version sans accent est majoritaire (mais on ne les distingue pas pour le moment).

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 30 000

Fréquence relative min. genre = 1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|
|1|2 259 111|Marie|1934|1920|1934|26 873|2 232 238|2.609%|1%|99%
|2|468 384|Claude|1944|1944|1945|411 739|56 645|0.541%|88%|12%
|3|409 960|Dominique|1958|1958|1958|241 941|168 019|0.473%|59%|41%
|4|282 947|Camille|1979|1949|1990|77 734|205 213|0.327%|27%|73%
|5|91 677|Yannick|1973|1973|1962|86 469|5 208|0.106%|94%|6%
|6|83 926|Irène|1934|1928|1934|1 815|82 111|0.097%|2%|98%
|7|61 708|Maxence|2004|2004|1991|60 992|716|0.071%|99%|1%
|8|58 986|José|1963|1963|1946|57 405|1 581|0.068%|97%|3%
|9|56 927|Sacha|2009|2009|2005|53 561|3 366|0.066%|94%|6%
|10|56 580|Andréa|1974|2008|1966|9 967|46 613|0.065%|18%|82%
|11|47 903|Lou|2009|2008|2009|3 224|44 679|0.055%|7%|93%
|12|45 842|Élie|1948|1947|1990|44 831|1 011|0.053%|98%|2%
|13|38 407|Cyrille|1970|1970|1968|36 852|1 555|0.044%|96%|4%
|14|36 411|Noa|2010|2010|2007|30 655|5 756|0.042%|84%|16%
|15|35 166|Gaël|1989|1989|1977|34 360|806|0.041%|98%|2%
|16|34 342|Alix|1992|1983|1995|8 529|25 813|0.040%|25%|75%
|17|32 914|Morgan|1995|1995|1994|30 605|2 309|0.038%|93%|7%
|18|31 286|Alex|1987|1987|1995|30 810|476|0.036%|98%|2%
|19|30 009|France|1942|1931|1943|1 809|28 200|0.035%|6%|94%

# Synthèse statistique des prénoms mixtes homophones

Les associations homophones sont programmées explicitement : il faut connaitre la prononciation pour pouvoir regrouper les prénoms qui se prononcent pareil. Les prénoms mixtes **épicènes** sont affichés en gras (si la fréquence relative d'un des genres est par exemple 0.9%, elle est arrondie à 1%, mais reste inférieure au seuil de 1%, comme par exemple Maxime, il n'est donc pas en gras). Si la fréquence d'une variante est inférieure à 1% à celles de ses homophones, elle est retirée et décomptée (il s'agit souvent d'une orthographe fantaisiste).

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 400 000

Fréquence relative min. variante = 1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|Fréq. rel. var.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|--:|
|1.0|1 122 931|Michel, Michèle, Michelle|1948|1947|1949|820 636|302 295|1.297%|73%|27%
|1.1|820 487|Michel|1947|1947|1937|820 353|134|0.947%|100%|0%|73%
|1.2|180 281|Michèle|1951|1974|1951|283|179 998|0.208%|0%|100%|16%
|1.3|122 163|Michelle|1947|0|1947|0|122 163|0.141%|0%|100%|11%
|2.0|928 261|André, Andrée|1930|1931|1927|712 232|216 029|1.072%|77%|23%
|2.1|712 745|André|1931|1931|1924|712 140|605|0.823%|100%|0%|77%
|2.2|215 516|Andrée|1927|1922|1927|92|215 424|0.249%|0%|100%|23%
|3.0|738 377|René, Renée|1927|1928|1925|516 277|222 100|0.853%|70%|30%
|3.1|516 616|René|1928|1928|1922|516 239|377|0.597%|100%|0%|70%
|3.2|221 761|Renée|1925|1919|1925|38|221 723|0.256%|0%|100%|30%
|4.0|697 894|Marcel, Marcelle|1924|1925|1922|468 573|229 321|0.806%|67%|33%
|4.1|468 604|Marcel|1925|1925|1913|468 567|37|0.541%|100%|0%|67%
|4.2|229 290|Marcelle|1922|1919|1922|6|229 284|0.265%|0%|100%|33%
|5.0|678 464|Daniel, Danièle, Danielle|1951|1951|1949|435 725|242 739|0.783%|64%|36%
|5.1|435 794|Daniel|1951|1951|1932|435 675|119|0.503%|100%|0%|64%
|5.2|158 110|Danielle|1949|0|1949|0|158 110|0.183%|0%|100%|23%
|5.3|84 560|Danièle|1950|1982|1950|50|84 510|0.098%|0%|100%|12%
|6.0|497 064|Jack, Jacques|1941|1941|1930|496 964|100|0.574%|100%|0%
|6.1|482 909|Jacques|1941|1941|1930|482 809|100|0.558%|100%|0%|97%
|6.2|14 155|Jack|1949|1949|0|14 155|0|0.016%|100%|0%|3%
|7.0|467 873|Paul, Paule|1945|1947|1927|421 489|46 384|0.540%|90%|10%
|7.1|421 605|Paul|1947|1947|1923|421 489|116|0.487%|100%|0%|90%
|7.2|46 268|Paule|1927|0|1927|0|46 268|0.053%|0%|100%|10%
|8.0|410 513|Pascal, Pascale|1964|1964|1963|307 814|102 699|0.474%|75%|25%
|8.1|307 918|Pascal|1964|1964|1942|307 811|107|0.356%|100%|0%|75%
|8.2|102 595|Pascale|1963|1958|1963|3|102 592|0.118%|0%|100%|25%

# Synthèse statistique des prénoms similaires

Les associations similaires sont programmées explicitement : il faut connaitre l'étymologie pour pouvoir regrouper les prénoms qui ont la même racine. Par exemple les prénoms Maël et Maëlle sont homophones, et les prénoms Maëliss, Maëlisse, Maëllys, Maélys et Maëlys sont aussi homophones. On voit que ces deux séries sont apparentées. Pour regrouper ces deux séries dans la synthèse des prénoms similaires, il faut choisir un prénom pivot. Le plus logique serait de choisir le prénom pivot Maélys dans l'association Maël-Maélys, car c'est le plus fréquent, mais au moment de regrouper les prénoms similaires, on ajoute automatiquement tous les prénoms homophones, pour éviter d'avoir à refaire ces associations. Or l'association ne peut fonctionner que dans un seul sens, via ce qu'on appelle un dictionnaire : il s'agit d'une structure informatique qui associe une clé à une valeur, ici une variante (la clé) à un prénom pivot (la valeur), pour pouvoir avoir plusieurs variantes pour une même valeur pivot. On voit bien qu'en regroupant les dictionnaires homophones et similaires, l'élément Maélys va se retrouver des deux côtés du dictionnaire, c'est-à-dire à la fois clé et valeur, ce qui sera évidement refusé par le logiciel. La solution à ce problème consiste à choisir un autre pivot pour l'une des deux séries homophones, à savoir un variant dont la fréquence est trop faible pour être retenue, ce qui ne faussera pas le décompte. Ici on choisi donc le pivot Maëllis au lieu de Maélys comme pivot homophone (Maëllis a seulement 133 occurrences, soit 0.3% des variantes, il ne sera donc pas inclus dans le décompte). 

Les prénoms mixtes **épicènes** sont affichés en gras, et les prénoms mixtes *homophones* en italique. Un prénom peut être à la fois ***épicène et homophone***. Si la fréquence d'une variante est inférieure à 1% relativement à celles de ses homogènes (similaires), elle est retirée et décomptée. On n'affiche que les prénoms avec au moins 20 000 occurrences, ce qui donne 81 prénoms (voir le résultat [ici](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_similaires) pour une liste avec un seuil minimum d'occurrences à 1000 avec 102 prenoms).

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 800 000

Fréquence relative min. variante = 1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|Fréq. rel. var.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|--:|
|1.0|2 692 675|Jean, Jeanne, Jeannine|1936|1938|1930|1 914 096|778 579|3.109%|71%|29%
|1.1|1 914 572|Jean|1938|1938|1938|1 914 060|512|2.211%|100%|0%|70%
|1.2|559 385|Jeanne|1928|1933|1928|36|559 349|0.646%|0%|100%|21%
|1.3|218 718|Jeannine|1936|0|1936|0|218 718|0.253%|0%|100%|8%
|2.0|1 234 782|*Michel*, *Michèle*, Micheline, *Michelle*|1947|1947|1946|820 636|414 146|1.426%|66%|34%
|2.1|820 487|*Michel*|1947|1947|1937|820 353|134|0.947%|100%|0%|66%
|2.2|180 281|*Michèle*|1951|1974|1951|283|179 998|0.208%|0%|100%|15%
|2.3|122 163|*Michelle*|1947|0|1947|0|122 163|0.141%|0%|100%|10%
|2.4|111 851|Micheline|1937|0|1937|0|111 851|0.129%|0%|100%|9%
|3.0|974 411|Pierre, Pierrette|1942|1942|1938|891 611|82 800|1.125%|92%|8%
|3.1|892 502|Pierre|1942|1942|1931|891 611|891|1.031%|100%|0%|92%
|3.2|81 909|Pierrette|1938|0|1938|0|81 909|0.095%|0%|100%|8%
|4.0|888 995|*Jack*, ***Jackie***, *Jacques*, Jacqueline|1940|1941|1939|514 964|374 031|1.026%|58%|42%
|4.1|482 909|*Jacques*|1941|1941|1930|482 809|100|0.558%|100%|0%|54%
|4.2|372 507|Jacqueline|1939|0|1939|0|372 507|0.430%|0%|100%|42%
|4.3|19 424|***Jackie***|1948|1947|1966|18 000|1 424|0.022%|93%|7%|2%
|4.4|14 155|*Jack*|1949|1949|0|14 155|0|0.016%|100%|0%|2%
|5.0|838 776|*Paul*, *Paule*, Paulette, Pauline|1948|1947|1949|421 489|417 287|0.969%|50%|50%
|5.1|421 605|*Paul*|1947|1947|1923|421 489|116|0.487%|100%|0%|50%
|5.2|213 324|Paulette|1929|0|1929|0|213 324|0.246%|0%|100%|25%
|5.3|157 579|Pauline|1982|0|1982|0|157 579|0.182%|0%|100%|19%
|5.4|46 268|*Paule*|1927|0|1927|0|46 268|0.053%|0%|100%|5%
|6.0|835 300|Louis, Louïse, Louïsa, Louisette|1944|1942|1948|525 245|310 055|0.964%|63%|37%
|6.1|525 347|Louis|1942|1942|1917|525 231|116|0.607%|100%|0%|63%
|6.2|271 460|Louïse|1947|1903|1947|14|271 446|0.313%|0%|100%|32%
|6.3|20 113|Louisette|1939|0|1939|0|20 113|0.023%|0%|100%|2%
|6.4|18 380|Louïsa|1976|0|1976|0|18 380|0.021%|0%|100%|2%
|7.0|800 602|François, Françoise|1947|1946|1948|398 854|401 748|0.924%|50%|50%
|7.1|401 539|Françoise|1948|1922|1948|9|401 530|0.464%|0%|100%|50%
|7.2|399 063|François|1946|1946|1933|398 845|218|0.461%|100%|0%|50%

# Synthèse statistique des prénoms unigenres

Les prénoms unigenres sont tous ceux qui ne sont dans aucune des listes précédentes : ni mixtes épicènes, ni mixtes homophones, ni similaires. Ce sont les prénoms qui sont que masculins, ou bien que féminins.

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 200 000

Fréquence relative min. variante = 1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|
|1|506 923|Alain|1954|1954|1952|506 910|13|0.585%|100%|0%
|2|423 748|Roger|1929|1929|1936|423 592|156|0.489%|100%|0%
|3|400 176|Monique|1944|1932|1944|419|399 757|0.462%|0%|100%
|4|395 098|Patrick|1960|1960|1959|395 095|3|0.456%|100%|0%
|5|394 850|Cathérine|1958|1937|1958|173|394 677|0.456%|0%|100%
|6|382 977|Nathalie|1970|1934|1970|76|382 901|0.442%|0%|100%
|7|379 660|Gérard|1947|1947|1936|379 421|239|0.438%|100%|0%
|8|305 081|Madeleine|1924|1920|1924|13|305 068|0.352%|0%|100%
|9|304 415|Sébastien|1980|1980|1936|304 340|75|0.351%|100%|0%
|10|289 871|Suzanne|1925|1933|1925|44|289 827|0.335%|0%|100%
|11|289 555|Thierry|1966|1966|1938|289 537|18|0.334%|100%|0%
|12|281 502|Hélène|1946|1936|1946|114|281 388|0.325%|0%|100%
|13|280 120|Olivier|1972|1972|1933|280 100|20|0.323%|100%|0%
|14|279 624|Thomas|1993|1993|1943|279 343|281|0.323%|100%|0%
|15|271 226|Marguerite|1920|1932|1920|103|271 123|0.313%|0%|100%
|16|265 904|Guy|1943|1943|1939|265 720|184|0.307%|100%|0%
|17|246 002|Sophie|1975|1936|1975|77|245 925|0.284%|0%|100%
|18|240 242|Sandrine|1974|0|1974|0|240 242|0.277%|0%|100%
|19|239 054|Céline|1977|1943|1977|72|238 982|0.276%|0%|100%
|20|237 848|Marc|1959|1959|1936|237 616|232|0.275%|100%|0%
|21|237 302|Didier|1961|1961|1933|237 170|132|0.274%|100%|0%
|22|235 934|Véronique|1965|1936|1966|141|235 793|0.272%|0%|100%
|23|232 702|Vincent|1978|1978|1934|232 361|341|0.269%|100%|0%
|24|220 720|Bruno|1965|1965|1939|220 605|115|0.255%|100%|0%
|25|214 563|Jean-pierre|1954|1954|0|214 563|0|0.248%|100%|0%
|26|209 130|Léa|1991|1927|1991|100|209 030|0.241%|0%|100%
|27|202 704|Brigitte|1957|1941|1957|124|202 580|0.234%|0%|100%

# Synthèse statistique des prénoms accentués

Cette synthèse permet d'analyser qu'elle est l'orthographe la plus fréquente d'un prénom selon l'accentuation. Cependant, l'absence d'accent n'est pas forcément significative, dans la mesure où cela peut être simplement une contrainte ou procédure de saisie des prénoms sans possibilité d'accents, sur les anciens systèmes informatiques. Par exemple pour le prénom Jérome on a le nombre d'occurrences suivantes :

Jerome : 182 506

Jérome : 22 679

Jerôme : 2 097

Jérôme : 1 109

On se doute que Jerome sans accent n'est pas significatif, c'est sûrement lié à une question de saisie informatique. Par contre, la proportion entre les 3 autres formes est davantage significative, quoique Jerôme soit tout de même étonnant, et on peut présumer que Jerome se réfère probablement à Jérome, mais sans certitude (car MS-Word corrige Jérome en Jérôme), ce qui fausserait les proportions.

Voici les résultats avec un réglage du seuil à 20 000 occurrences minimum, avec une fréquence relative minimale des variantes accentuées de 0.1 %, avec 126 prénoms (voir le résultat [ici](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_accentués) pour une liste avec un seuil minimum d'occurrences à 2000 avec 542 prenoms) :

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 200 000

Fréquence relative min. variante = 0.1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|Fréq. rel. var.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|--:|
|1.0|320 897|Eric, Éric|1967|1967|0|320 897|0|0.371%|100%|0%
|1.1|318 166|Eric|1967|1967|0|318 166|0|0.367%|100%|0%|99.1%
|1.2|2 731|Éric|1972|1972|0|2 731|0|0.003%|100%|0%|0.9%
|2.0|304 415|Sebastien, Sébastien|1980|1980|1936|304 340|75|0.351%|100%|0%
|2.1|304 025|Sébastien|1980|1980|1936|303 950|75|0.351%|100%|0%|99.9%
|2.2|390|Sebastien|2010|2010|0|390|0|0.000%|100%|0%|0.1%
|3.0|239 054|Céline, Celine|1977|1943|1977|72|238 982|0.276%|0%|100%
|3.1|238 727|Céline|1977|1943|1977|72|238 655|0.276%|0%|100%|99.9%
|3.2|327|Celine|2012|0|2012|0|327|0.000%|0%|100%|0.1%
|4.0|209 130|Léa, Lea|1991|1927|1991|100|209 030|0.241%|0%|100%
|4.1|202 278|Léa|1990|1927|1990|100|202 178|0.234%|0%|100%|96.7%
|4.2|6 852|Lea|2012|0|2012|0|6 852|0.008%|0%|100%|3.3%
|5.0|208 391|Jerome, Jérome, Jerôme, Jérôme|1976|1976|1935|208 225|166|0.241%|100%|0%
|5.1|182 506|Jerome|1976|1976|1935|182 340|166|0.211%|100%|0%|87.6%
|5.2|22 679|Jérôme|1978|1978|0|22 679|0|0.026%|100%|0%|10.9%
|5.3|2 097|Jérome|1978|1978|0|2 097|0|0.002%|100%|0%|1.0%
|5.4|1 109|Jerôme|1979|1979|0|1 109|0|0.001%|100%|0%|0.5%

# Synthèse statistique des prénoms fréquents

En appliquant aucun seuil, on trouve 32 346 prénoms distincts, mais un peu moins si on corrige quelques accents. Et 359 prénoms avec au moins 50 000 occurrences (voir le résultat [ici](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_fréquents) pour un seuil à 4000 avec 1378 prenoms) :

Date début = 1900

Date fin   = 2020

Nb. total de prénoms identifiés et datés = 84 107 375

Nb. total de prénoms = 86 605 605

Nb. prénoms ignorés ('_PRENOMS_RARES') = 1 653 266 : 1.9%

Nb prénoms ignorés (date 'XXXX') = 844 964 : 1.0%

Seuil min. = 400 000

Fréquence relative min. variante = 1%


|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|
|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|
|1|2 259 111|**Marie**|1934|1920|1934|26 873|2 232 238|2.609%|1%|99%
|2|1 914 572|Jean|1938|1938|1938|1 914 060|512|2.211%|100%|0%
|3|892 502|Pierre|1942|1942|1931|891 611|891|1.031%|100%|0%
|4|820 487|*Michel*|1947|1947|1937|820 353|134|0.947%|100%|0%
|5|712 745|*André*|1931|1931|1924|712 140|605|0.823%|100%|0%
|6|559 385|Jeanne|1928|1933|1928|36|559 349|0.646%|0%|100%
|7|538 776|Philippe|1961|1961|1934|538 519|257|0.622%|100%|0%
|8|525 347|Louis|1942|1942|1917|525 231|116|0.607%|100%|0%
|9|516 616|*René*|1928|1928|1922|516 239|377|0.597%|100%|0%
|10|506 923|Alain|1954|1954|1952|506 910|13|0.585%|100%|0%
|11|482 909|*Jacques*|1941|1941|1930|482 809|100|0.558%|100%|0%
|12|469 302|Bernard|1945|1945|1940|469 257|45|0.542%|100%|0%
|13|468 604|*Marcel*|1925|1925|1913|468 567|37|0.541%|100%|0%
|14|468 384|**Claude**|1944|1944|1945|411 739|56 645|0.541%|88%|12%
|15|435 794|*Daniel*|1951|1951|1932|435 675|119|0.503%|100%|0%
|16|423 748|Roger|1929|1929|1936|423 592|156|0.489%|100%|0%
|17|421 605|*Paul*|1947|1947|1923|421 489|116|0.487%|100%|0%
|18|419 228|Robert|1931|1931|1927|419 125|103|0.484%|100%|0%
|19|409 960|**Dominique**|1958|1958|1958|241 941|168 019|0.473%|59%|41%
|20|405 784|Henri|1927|1927|1929|405 647|137|0.469%|100%|0%
|21|405 616|Christian|1954|1954|1940|405 548|68|0.468%|100%|0%
|22|405 475|Georges|1929|1929|1934|404 878|597|0.468%|100%|0%
|23|405 347|Nicolas|1983|1983|1932|405 112|235|0.468%|100%|0%
|24|401 539|Françoise|1948|1922|1948|9|401 530|0.464%|0%|100%
|25|400 176|Monique|1944|1932|1944|419|399 757|0.462%|0%|100%

# Liens externes

- [Wikipédia - Prénom mixte](https://fr.wikipedia.org/wiki/Prénom_mixte)
- [Synthèse statistique des prénoms mixtes épicènes (liste complète)](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_mixtes_épicènes)
- [Synthèse statistique des prénoms similaires avec au moins 1000 occurrences](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_similaires)
- [Synthèse statistique des prénoms accentués avec au moins 2000 occurrences](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_accentués)
- [Synthèse statistique des prénoms fréquents avec au moins 4000 occurrences](http://patrice.dargenton.free.fr/wiki/index.php?title=Synthèse_statistique_des_prénoms_fréquents)
