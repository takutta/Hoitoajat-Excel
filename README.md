# Hoitoajat Excel -työkalu

## Mikä?

Tämä työkalu toimii [**Tiedoevryn**](https://www.tietoevry.com/fi) **Edlevon** dataa käyttämällä. Voit luoda sillä ryhmän (tai yhdistettyjen ryhmien) **viikkolistoja** , **esitäytettyjä ruokatilauslappuja** sekä koko päiväkodin **aamu- ja iltalistat** , joista näkee ketkä lapset tulevat milloinkin päiväkotiin, ja milloin sieltä lähtevät. On myös mahdollista luoda **päivälappuja** , joissa näkyvät hoitoajat.

_Meillä Kukkumäen pk:lla titanikko **tulostelee** tiistai-aamupäivisin listat tällä ohjelmalla, **säätää** niiden pohjalta vielä työvuoroja, ja **kiikuttaa** laput lopuksi ryhmiin._

## Miten?

1. Resurssisuunnittelun Läsnäolot -osuudessa **haetaan** koko päiväkodin hoitoajat ja kopioidaan ne Exceliin.
2. **Muokataan** asetuksia.
3. **Luodaan** listat.
4. **Tulostetaan** valmiit listat.

## Huom!

Ohjelma voi kysyä pariinkin kertaan varmistusta. **Ota muokkaus ja sisältö käyttöön**. Muuten ohjelma ei toimi kunnolla.

## Lataus ja päivitys

- **Hoitoajat.xlsm** on varsinainen ohjelma. Tämän käynnistät.
- **Hoitoajat-data.xlsx** on asetustiedosto, johon tallentuvat mm. Erityisruokavaliot. Se on pidettävä samassa kansiossa ohjelman kanssa!

Päivitettäessä ohjelmaa, **korvaa vain Hoitoajat.xlsm tiedosto uudella.**

## Testaus

Voit testata ohjelmaa **demo**-kansiossa olevalla versiolla. Siinä on mukana hatusta vedetty hoitoaikalista, josta ohjelma tietonsa hakee.

## Ohjeet

### 1. Tietojen hakeminen

1. Kirjaudu ensin Chrome/Edge -selaimella Resurssisuunnitteluun
2. Mene läsnäolojen yleisnäkymään
3. Valitse seuraava viikko
4. Hoitoaikojen näkyessä ruudulla, kopioi koko sivun sisältö leikepöydälle
   _(pidä **ctrl**-nappi pohjassa ja paina ensin **a** ja sitten **c**.
5. Siirry takaisin Hoitoajat-työkaluun. Klikkaa **Päiväkoti** -välilehdellä **Lisää hoitoajat** -nappulaa.

Tämä tietojen haku on tehtävä viikoittain esim. tiistaisin.

### 2. Asetusten muokkaus

Hoitoaikojen lisäämisen jälkeen, voit muokata asetuksia, jotka sijaitsevat kolmella eri välilehdellä: **Päiväkoti** , **Ryhmät** ja **Lapset**. Välilehdet sijaitsevat Excelissä vasemmassa alareunassa. Tässä pikainen esittely muokattavista asioista. Näitä asioita voit muokata, **älä koske muihin kohtiin!**

#### Päiväkoti -välilehti

- **Aamu- ja iltalistat**
   - Koko päiväkodin lapsista tehtävät aamu- ja iltalistat.
- **Ruokatilausten lähettmänien sähköpostilla pdf -muodossa**
   - Tällä kaavakkeella voit lähettää kaikki luomasi ruokatilauslaput pdf:nä haluamallesi henkilölle (keittäjälle tai hänen esihenkilölle). _Sinun pitää ensin luoda ruokatilauslaput._
- **Tulostetaanko ruokakoonti keittiölle**
   - Luo listan, jossa näkyy kaikkien luotujen ruokatilauslappujen koonti.
- **Päivälappujen pohja-työkalu**
   - Tällä kaavakkeella voit kopioida haluamasi päivälappu-pohjan muokattavaksi.
      1. Valitse pohja ja kirjoita viereiselle solulle haluamasi ryhmän nimi
      2. Klikkaa "Luo uusi päivälapun pohja" -nappulaa
      3. Uusi pohja-välilehti ilmaantuu, *esim. pl_juolukat*
      4. Voit muokata uuden pohjan ulkonäköä haluamaksesi. Pohja sisältää koodisanoja, joiden paikkaa vaihtamalla voit vaikuttaa nimen ( **pl-nimi** ) ja hoitoajan ( **pl-hoitoaika** ) sijaintiin, sarakkeiden määrään ( **pl-vikasarake** ) ja alaosan tilaan ( **pl-alateksti** ). Jotta pohja mahtuu yhdelle paperille, sinun pitää asettaa **pl-vikarivi** -koodi viimeiselle riville.
      5. **Kokeile!** Ota yhteyttä jos on hankalaa ja törmäät ongelmiin!

#### Ryhmä -välilehti

- **Järjestys**
   - Tässä järjestyksessä laput tulostuvat.
- **Käytössä**
   - Ryhmän voi välillä laittaa pois käytöstä.
- **Nimi**
   - Nimet haetaan suoraan Läsnästä.
- **Viikkolistan aakkosjärjestys**
   - Missä järjestyksessä ryhmän lapset viikkolistalla järjestetään.
- **Boldaukset**
   - Halutessasi hoitoaika **boldataan** jos se on aikaisempi tai myöhäisempi.
- **Puokkariboldaukset**
   - Näiden kellonaikojen sisällä olevat hoitoajat alleviivataan. Näin on kätevä nähdä puolipäiväiset lapset ja muut poikkeavat hoitoajat.
- **Ruokalapun tulostus**
  Erilaisia vaihtoehtoja päivä/ilta- ja viikonloppuruokien tilaukseen.
- **Yhdistä ruokalaput**
  Voit halutessasi yhdistää kahden tai useamman ryhmän ruokalaput yhteen lappuun. Yhdelle lapulle mahtuu korkeintaan 18 erityisruokavaliota. Ryhmät erotellaan pilkulla esim. _Mansikat, Mustikat, Vadelmat._
- **Ruokailujen ajankohdat**
  Näitä aikoja käytetään ruokatilauslapun automaattiseen täyttämiseen hoitoaikojen pohjalta.
   - _Ensimmäinen aika kertoo **mihin mennessä** lapsi vielä ehtii ruokailuun.
   - _Toinen aika kertoo milloin lapsi vielä ehtii ruokailuun **ellei häntä haeta**.
   - _Esim. Jos aamupalan ajat ovat 8:00 ja 8:15:_
      - _Lapsen tarvitsee tulla klo 8 mennessä päiväkotiin ehtiäkseen aamupalalle_
      - _Jos hänet haetaan aikaisintaan 8:15, hän ehtii vielä syödä aamupalan._
- **Viikkolistan tulostus**
   - Mahdollisuus valita myös viikonloput.
- **Yhdistä viikkolistat** (Ryhmän nimi tai ryhmien nimet pilkulla erotettuna _esim. Mansikat, Mustikat, Karpalot_)
   - Voit halutessasi yhdistää kahden tai useamman ryhmän lapset yhteen viikkolistaan.
- **Yhdistetyn listan nimi**
   - Yhdistetty viikkolista tarvitsee nimen. Se ei saa olla liian pitkä, tai menee päivämäärien yli.
- **Yhdistettyjen listojen tyylit**
   - On mahdollista asettaa esim. vain viikkolista yhdistetyksi listaksi.
   - 2-puoliset päivälaput: tulostaa 2 eri ryhmää eri lapuille. Niiden tulostaminen 2-puoliseksi päivälapuksi ei vielä toimi, mutta voit valita nämä 2 välilehteä (ctrl-pohjassa hiirellä kliksuttaen) ja tulostaa manuaalisesti 2-puolisen paperin.
- **Päivystys**
   - Vain ne lapset, joiden tietoihin on merkattu tämä ryhmä päivystäväksi ryhmäksi, näkyvät listoilla. On mahdollista pitää useampia päivystys-ryhmiä joissa on eri lapset.
- **Päivälappujen tulostus**
   - Joka päivälle oma lappu, jossa hoitoaika + tilaa kirjoittaa lapsen kuulumisia.
   - Mahdollisuus poistaa poissaolijat listalta (lapset, jotka viikon kaikkina päivinä poissa)
- **Päivälappujen aakkosjärjestys**
- **Päivälappujen pohja**
   - Jos haluat tehdä oman pohjan, valitse pohjaksi Kustomoitu ja mene Päiväkoti-välilehdelle luomaan uusi pohja. Ohjeet siihen löytyvät ylempää.
- **Tyhjät nimirivit**
   - Päivälapun lopussa tyhjien rivien määrä.

#### Lapset -välilehti

- **Kutsumanimi**
   - Lapuissa näkyvä nimi. Ohjelma on arvannut etunimen + sukunimen 1. kirjaimen esim. Kalle P. Halutessasi voit vaihtaa kutsumanimen.
- **Erityisruokavalio**
   - Näiden lasten ruokailut lasketaan erilliselle sarakkeelle tässä järjestyksessä. Tämä on hyvä jos keittäjä haluaa vaikkapa laktoosittomat listan viimeiseksi.
- **Työntekijä**
   - Tällä voit lisätä esimerkkiateriaa syövän työntekijän, jolla on erityisruokavalio. Lisätäksesi työntekijän:
      1. Kirjoita uudelle riville kutsumanimi-sarakkeeseen nimi ja ryhmä-sarakkeeseen oikea ryhmä.
      2. Merkitse sopiva järjestysnumero erityisruokavalio-sarakkeeseen. \* Kirjoita jotakin, ihan mitä vain, Työntekijä-sarakkeeseen.
   - Tällöin ohjelma ymmärtää lisätä työntekijän nimen ruokalapulle. Ohjelma ei kylläkään voi millään tietää milloin työntekijä syö esimerkkiä, joten se on täytettävä manuaalisesti.

- **Päivystys**
   - Merkitse päivystävän ryhmän nimi. Silloin vain nämä lapset näkyvät listalla. Muista myös merkata ryhmän asetuksista päivystys!

### 3. Listojen luominen

Klikkaa Päiväkoti-välilehden **Luo listat** -nappulaa. Ohjelma yrittää nyt rakentaa listat. Jos kaikki menee hyvin, läjä uusia välilehtiä ilmaantuu. Niitä voi katsella ja halutessaan muokatakin.

### 4. Tulostaminen

Klikkaa **Tulosta** -nappulaa (_Päiväkoti-välilehdellä_). Kaikki luomasi laput tulostuvat oletustulostimellasi.
