# Ergonomianalys KIM 1

Svensk KIM 1-app med Dometics befintliga bedömningsskalor och beräkning.

- Motivering eller kommentar på fråga 1–5.
- Flera bilder med bildtext på fråga 1–5 och punkt 8.
- Dokumentbilagor under punkt 8.
- Tre huvudval: **Exportera PowerPoint**, **Exportera Excel** och **Skriv ut**.
- PowerPoint använder det ursprungliga A4-upplägget med frågorna och bilagorna på separata bilder och en sammanfattning sist. Text, tabeller och foton är redigerbara objekt.
- Dokumenten bäddas in som OLE Package-objekt. Öppna dem i PowerPoint för Windows.
- Excel-exporten är en riktig `.xlsx` med numeriska, redigerbara poäng och formler för riskpoäng/riskklass. Bilder och fullständiga längre kommentarer finns i underlaget. Dokumentbilagorna listas i Excel och bäddas in i PowerPoint.
- Skriv ut ger en onepager följd av bild- och kommentarsunderlag. Webbläsarens utskrift kan också sparas som PDF.
- Analysen, fotona och dokumenten sparas automatiskt i den aktuella webbläsarens IndexedDB. Innehållet skickas inte till någon server. Rensning av webbläsarens data tar bort det lokala utkastet.

Dokument: högst 20 MB per fil och 50 MB totalt. Bilder komprimeras till högst 2000 pixlar på längsta sidan. Inbäddade dokument kan omfattas av organisationens säkerhetsinställningar i Office.

Appen är statisk och använder medföljande PptxGenJS 4.0.1 och JSZip 3.10.1. Licenser finns under `vendor/`. Ingen CDN eller installation behövs för exporten.

## Verifiering

Kör `node tests/export.test.cjs` med Node 18 eller senare. Testet kontrollerar paketets interna referenser, sammanfattning, separata foton, långa kommentarer, inbäddade dokument, Excel-formler och riskgränser.
