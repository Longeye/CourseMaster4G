A Google Apps Script ebben a táblázatban található:

https://docs.google.com/spreadsheets/d/1Yph5MZUADPBpkjWahhJc27Xu_GVcqxwWxl43kjzPwkg/edit?usp=sharing

Csak olvasására van megosztva, így futtatás előtt másolatot kell készíteni róla.

Alapértelmezetten TESZT üzemmódban fut, tehát csak olyan felhasználóknak és csoportoknak módosítja az e-mail címét, amelyeknek az e-mail címe "teszt" stringgel kezdődik.
Csoportok esetében a csoport nevét vizsgálja.

A "Diákok" nevű munkalapra fel lehet tölteni a diákok neveit és az alapértelmezett jelszavukat. Pl. OM azonosító. Azoknak a felhasználóknak, akiknek megtalálható itt az e-mail címe, beállítja a jelszavát az itt szereplő értékre.
A megadott értékeknek teljeíteni kell a jelszavakkal szemben támasztott biztonsági előírásokat, különben hibaüzenetet kapunk.

Akik ebben a táblázatban nem szerepelnek, azoknak nem módosítja a jelszavát.

Ha éles adatokon szeretnéd futtatni, akkor át kell írni a scriptben az első sort:

var testmode = false;

