# Catalogo Opere (Vite + React)

Applicazione web per impaginare un catalogo di opere (quadri, fotografie, dipinti) in formato libro sfogliabile.

Consente di:
- importare immagini e metadati opere
- impaginare opere e testi su pagine
- gestire copertina / retrocopertina / pagine speciali
- generare pagine finali `Elenco opere`
- esportare/importare JSON
- stampare/esportare PDF tramite `window.print()`

## Requisiti

- Node.js 20+ (consigliato)
- npm
- Docker / Docker Compose (opzionale)

## Avvio locale (sviluppo)

```bash
npm install
npm run dev
```

Apri l’URL mostrato da Vite (di default `http://localhost:5173`).

## Build produzione

```bash
npm run build
npm run preview
```

## Avvio con Docker (porta 6063)

### Docker Compose

```bash
docker compose up --build
```

App disponibile su:

- `http://localhost:6063`

### Docker (senza compose)

```bash
docker build -t catalogo-opere .
docker run --rm -p 6063:6063 catalogo-opere
```

## Funzionalità principali

### Catalogo opere

- Drag & drop immagini sul frame principale
- Import multiplo immagini dal filmstrip (selezione multipla)
- Editor opera con metadati:
  - titolo, autore, anno, tipo, tecnica, dimensioni, inventario, collocazione, note
- Persistenza immagini in `IndexedDB` (evita errore quota `localStorage`)

### Filmstrip (sinistra)

- Selezione opera
- Modifica opera
- Elimina opera (`×`)
- Drag dell’opera sulle pagine

### Libro / impaginazione

- Vista a doppia pagina (spread)
- Vista `libro reale` / `impaginazione tecnica`
- Pan + zoom
- Drag, resize, snap/griglia/guide
- Spostamento opere tra pagine con maniglia `⇄`
- Didascalie spostabili liberamente
- Inline edit dei testi (e didascalie)

### Pagine speciali

- Copertina editabile
- Seconda di copertina (credits) editabile
- Terza di copertina editabile
- Retrocopertina editabile (sintesi, bio autore, codice edizione)

### Elenco opere (pagine finali)

- Numero pagine generato automaticamente in base alle opere
- Lista con:
  - `Titolo - Autore - Anno`
  - `pag. N` (pagina dove compare l’opera)
- Elementi testo editabili (`titolo elenco` + `lista`)
- Comando `Rigenera elenco opere` nel menu overflow

## Tema catalogo

Configurabile dal pannello tema (icona in alto):

- font, peso, dimensione testo/titoli
- colori testo/accento/UI
- margini globali catalogo
- sfondo pagina default (applicato a tutte le pagine)
- default mostra didascalia
- bordo elementi default `%` (applicato globalmente)

## Controlli utili

- `Alt/Ctrl/Cmd + rotellina`: zoom libro
- `Shift + drag` sullo sfondo: pan
- `Alt` durante drag/resize: snap temporaneamente OFF
- `Shift` durante drag didascalia: riattiva snap sulla didascalia
- `Doppio click` su testo/didascalia: modifica inline
- `Esc`: annulla inline edit
- `Ctrl/Cmd + Invio`: salva inline edit

## Export / Import

Dal menu overflow (`⋯`) in topbar:

- `Esporta JSON`
- `Importa JSON`
- `Rigenera elenco opere`
- `Stampa / PDF Catalogo`

### PDF / Stampa

La stampa usa il formato pagina scelto (A4 / quadrato / orizzontale) e renderizza le pagine singole del catalogo.

Per ottenere un PDF più fedele:
- abilita “grafica di sfondo / background graphics” nel dialogo di stampa del browser

## Persistenza dati

- Stato catalogo: `localStorage`
- Immagini: `IndexedDB`

Questo evita errori di quota dovuti a immagini Base64 in `localStorage`.

## Struttura progetto (essenziale)

- `src/App.jsx` - logica principale applicazione/editor
- `src/main.jsx` - bootstrap React
- `src/styles.css` - stili UI/editor/print
- `Dockerfile` - build/serve container
- `docker-compose.yml` - avvio rapido su `6063`

## Note

- Le pagine `spacer` tecniche non sono editabili.
- Le pagine `Elenco opere` sono generate automaticamente ma supportano override manuali (testi editabili).
- Alcune impostazioni del tema (es. sfondo pagina default, bordo elementi default) propagano le modifiche a tutte le pagine/elementi esistenti.
