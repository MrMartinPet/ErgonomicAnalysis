# Ergonomianalys KIM 1

Svensk KIM 1-app med Dometics befintliga bedömningsskalor och beräkning.

- Motivering eller kommentar på fråga 1–5.
- Flera bilder med bildtext på fråga 1–5 och punkt 8.
- Dokumentbilagor under punkt 8.
- Redigerbar PowerPoint i A4 liggande. Text, tabeller och foton är separata objekt. Sammanfattningen ligger sist.
- Dokumenten bäddas in som OLE Package-objekt. Öppna dem i PowerPoint för Windows. ZIP-exporten innehåller också dokumenten separat för andra program och enheter.
- PDF för utskrift och CSV med poäng, motiveringar och bilageförteckning.
- Analysen, fotona och dokumenten sparas automatiskt i den aktuella webbläsarens IndexedDB. Innehållet skickas inte till någon server. Rensning av webbläsarens data tar bort det lokala utkastet.

Dokument: högst 20 MB per fil och 50 MB totalt. Bilder komprimeras till högst 2000 pixlar på längsta sidan. Inbäddade dokument kan omfattas av organisationens säkerhetsinställningar i Office.

Appen är statisk och använder medföljande PptxGenJS 4.0.1 och JSZip 3.10.1. Licenser finns under `vendor/`. Ingen CDN eller installation behövs för exporten.

## Verifiering

Kör `node tests/export.test.cjs` med Node 18 eller senare. Testet kontrollerar rapportens innehåll, redigerbara tabeller, separata foton, långa kommentarer, sista sammanfattningssidan och oförändrade bilagebytes i små och stora OLE-filer.
