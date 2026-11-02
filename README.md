# Ergonomianalys KIM 1

Svensk KIM 1-app med Dometics befintliga bedömningsskalor och beräkning.

- Motivering eller kommentar på fråga 1–5.
- Flera bilder med bildtext på fråga 1–5 och punkt 8.
- Dokumentbilagor under punkt 8.
- Tre huvudval: **Exportera PowerPoint**, **Exportera Excel** och **Skriv ut**.
- PowerPoint i Dometic-stil, A4 liggande, med Gibson/Gibson Light och befintlig logotyp. Helhetsresultatet ligger först på en onepager med riskklass, riskpoäng, faktorernas valda nivåer och motiveringar. Fyra bilder ryms per underlagssida. Upp till tre dokument läggs på sista bildsidan när det finns plats; annars samlas upp till sex per dokumentsida. Långa kommentarer får tätare fortsättningssidor utan att innehåll försvinner.
- PowerPoint-text, tabeller och foton är separata redigerbara objekt. Gibson behöver finnas installerat på datorn där presentationen öppnas.
- Dokumenten bäddas in som OLE Package-objekt. Öppna dem i PowerPoint för Windows.
- Excel-exporten är en riktig `.xlsx` med numeriska, redigerbara poäng och formler för riskpoäng/riskklass. Bilder och fullständiga längre kommentarer finns i underlaget. Dokumentbilagorna listas i Excel och bäddas in i PowerPoint.
- Skriv ut ger en onepager följd av bild- och kommentarsunderlag. Webbläsarens utskrift kan också sparas som PDF.
- Analysen, fotona och dokumenten sparas automatiskt i den aktuella webbläsarens IndexedDB. Innehållet skickas inte till någon server. Rensning av webbläsarens data tar bort det lokala utkastet.

Dokument: högst 20 MB per fil och 50 MB totalt. Bilder komprimeras till högst 2000 pixlar på längsta sidan. Inbäddade dokument kan omfattas av organisationens säkerhetsinställningar i Office.

Appen är statisk och använder medföljande PptxGenJS 4.0.1 och JSZip 3.10.1. Licenser finns under `vendor/`. Ingen CDN eller installation behövs för exporten.

## Verifiering

Kör `node tests/export.test.cjs` med Node 18 eller senare. Testet kontrollerar onepagerns placering, Gibson, kompakta sidantal, separata foton, fullständiga långa kommentarer, inbäddade dokument, riktiga Excel-formler och riskgränser.
