A Google Apps Script ebben a táblázatban található:

https://docs.google.com/spreadsheets/d/1_NpAR3qFhq4y815ISzmARQflO-oWInTfki_pVN6DmKA/edit?usp=sharing

Csak olvasására van megosztva, így futtatás előtt másolatot kell készíteni róla.
Az első megnyitáskor inicializálja magát. Létrehozza a "Status" és a "Diákok" munkalapokat (ha még nem léteznek), valamint elkészíti a "Domainváltó" menüpontot a főmenüben, amelyen keresztül elérhetőek a script funkciói. Íly módon elégséges a scriptet átmásolni egy másik (akár üres) táblázatba is.
Script a scriptszerkesztőben.

Csak admin jogosultságú felhasználóval működik, mivel a felhasználók e-mail címét kérdezi le és módosítja.

Alapértelmezetten "TESZT" üzemmódban fut, tehát csak olyan felhasználóknak és csoportoknak módosítja az e-mail címét, amelyeknek az e-mail címe "teszt" stringgel kezdődik.
Csoportok esetében a csoport nevét vizsgálja. A "TESZT" üzemmódot jelzi a "Status" munkalapon is. A Domainváltó menüben lehet ki- és bekapcsolni a "TESZT" üzemmódot.

A "Status" munkalapon követhető hol tart a script tevékenysége.

A "Diákok" nevű munkalapra fel lehet tölteni a diákok neveit és az alapértelmezett jelszavukat. Pl. OM azonosító. Azoknak a felhasználóknak, akiknek megtalálható itt az e-mail címe, beállítja a jelszavát a második oszlopban szereplő értékre.
A megadott értékeknek teljesíteni kell a jelszavakkal szemben támasztott biztonsági előírásokat, különben hibaüzenetet kapunk.

Akik ebben a táblázatban nem szerepelnek, azoknak nem módosítja a jelszavát.

