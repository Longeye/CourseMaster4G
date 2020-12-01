Kurzusmester

Nagyon 1.0 verzió

Ezt a táblázatot a szükség szülte, így kinézete és kényelme vetekszik a kőbunkóéval.
A normálisan használható verzió fejlesztése folyamatban...

Az alábbi linkről lehet megnyitni:
https://docs.google.com/spreadsheets/d/1-vB14LTFsV205glhUPPLbmonl0HIODilaAVIFCemOIw/edit?usp=sharing

Csak azután válik használhatóvá, miután másolatot készítettünk róla a saját drive-unkra.
A script-ek is csak ezután válnak elérhetővé a script-szerkesztőn keresztül.
Megnézni és futtatni lehet bármilyen Google account-tal, de csak Google Workspace rendszergazdai szintű jogosultságú felhasználóval fog megfelelően működni.

A táblázat alapértelmezés szerint a Státusz munkalapot tartalmazza csak a vezérlőelemekkel és a hozzá tartozó Apps Script-ekkel.
Először is ki kell töltenünk a legfőbb paramétereket. A Domain-t, a diákok domain-jét (ez a kettő lehet ugyanaz).
A tanároknak és a diákoknak kialakított szervezeti egységek teljes elérési útját.
Ezután az "Alapadatok frissítése" gombra kell kattintani. Ekkor egy script kigyűjti a Google Workspace (G Suit) rendszerből a kurzusokat, a tanárokat és diákokat.
Létrejönnek sorban a "Kurzusok", "Tanárok" és "Tanulók" munkalapok.
Ezeket érdemes ezután is csak az "Alapadatok frissítése" funkcióval regenerálni.

Funkciók:
Tanárok kurzusonként
Kilistázza a kurzusok neveit és alatta a hozzájuk tartozó tanárok neveit.
Létrehozza a "Kurzusok - tanárok" munkalapot.

Tanulók kurzusonként
Lémyegében megyegyezik az előző funkcióval, csak diákokra.
Létrehozza a "Kurzusok - diákok" munkalapot.

Tanulói kurzusműveletek
Ha a B14-es mezőbe beírjuk egy kurzus nevének kezdetét valameddig és lenyomjuk az Enter billentyűt, akkor mögötte (C14) megjelenik a kurzus ID-je és a teljes neve (D14).
A B15-ös cellába egy tanuló e-mail címének elejét adhatjuk meg. Enter lenyomása után megjelenik a diák teljse e-mail címe.
Ha a két paramétert beállítottuk, akkor használhatjuk a "Tanuló hozzáadása kurzushoz" funkciót.
Ennek a műveletnek az eredményességét ellenőrízhetjük a "Kurzus tanulóinak listázás" funkcióval.
Ez létrehoz egy új munkalapt, aminek a neve a kurzus neve lesz.

Tanári kurzusműveletek
Ha a B22-as mezőbe beírjuk egy kurzus nevének kezdetét valameddig és lenyomjuk az Enter billentyűt, akkor mögötte (C22) megjelenik a kurzus ID-je és a teljes neve (D22).
A B23-as cellába egy tanár e-mail címének elejét adhatjuk meg. Enter lenyomása után megjelenik a tanár teljse e-mail címe.
Ha a két paramétert beállítottuk, akkor használhatjuk a "Tanár hozzáadása kurzushoz" funkciót.
Ennek a műveletnek az eredményességét ellenőrízhetjük a "Kurzus tanárainak listázás" funkcióval.
Ez létrehoz egy új munkalapt, aminek a neve a kurzus neve lesz, kiegészítve a " - tanárok" szöveggel.
