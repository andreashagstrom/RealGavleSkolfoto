Detta Python-baserade skript är en helautomatiserad lösning för skolans foto- och dokumenthantering. Från rådata i elevregistret genereras först strukturerade CSV-filer som fungerar som ett smart nav för all information – namn, klass och kopplade fotografier. Med hjälp av dessa filer skapas sedan snygga PDF:er både för namnskyltar och klassfoton.

Namnskyltarna anpassas dynamiskt efter elevens namn och klass, med optimerad layout för enkel utskrift och tydlig läsbarhet. Klassfoton genereras på ett elegant sätt, med flera bilder per sida och automatiskt inkluderad sidnumrering, metadata och loggning – allt för att underlätta administration och spårbarhet.

Bakom kulisserna utnyttjar skriptet moderna Python-bibliotek som ReportLab för PDF-hantering, och det loggar automatiskt alla steg i processen, vilket ger full insyn och kontroll. Med några knapptryck kan hela elevfotoprocessen gå från rådata till färdiga dokument, vilket sparar tid, minskar fel och ger ett proffsigt slutresultat.

1. Skapa en textfil för varje klass i mappen Elevdata. Kopiera data direkt från klasslistan på Admentum. Namn och epost.
2. Kör Python-skriptet.
3. Skriv ut namnskyltar.
4. Fota. Namnge och beskär bilderna.
5. Kör Python-skriptet igen för att skapa klassfoton-filer.
6. Ladda upp filerna som skapats i klassfoton.

Observera att du behöver installera Python först. Detta gör du via Windows/MS Store. Sedan kan du behöva installera några libraries. Det kommer terminalen att meddela när du kör skriptet. Exempelvis krävs reportlab. Öppna då CMD och skriv pip install reportlab

/ Andreas Hagström
2025-08-27