# 30 let Fantázie – Jubilejní pátrací hra

Interaktivní webová hra vytvořená pro oslavu 30. výročí pobočky Fantázie (SVČ Helceletka Brno).

## 📍 O hře
- **12 stanovišť** rozmístěných po brněnské městské části Slatina.
- Hráči (rodiny s dětmi) navštěvují stanoviště, skenují QR kódy na cedulkách v terénu a odpovídají na otázky.
- Za správné odpovědi získávají dílky puzzle a skládají jubilejní logo Fantázie.
- Po dokončení celé trasy získávají dárkový certifikát k vyzvednutí odměny na pobočce.

## 🔐 Administrace hry
- Testovací funkce a generátor QR kódů pro laser jsou chráněny heslem `Fantoska30@`.
- Běžní hráči nemohou odemykat stanoviště na dálku z mapy (musí navštívit fyzické místo ve Slatině).

## ☁️ Synchronizace
- Propojeno s Google Tabulkou přes Google Apps Script (`google_apps_script.js`).
