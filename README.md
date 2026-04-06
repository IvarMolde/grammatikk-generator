# Grammatikk-generator – MBO
## Molde voksenopplæringssenter

Genererer grammatikkoppgaver (Word + PowerPoint) for voksne innvandrere på nivå A0–B2.

---

## Struktur

```
grammatikk-generator/
├── public/
│   └── index.html        ← Nettside med grensesnitt
├── api/
│   ├── lag-docx.js       ← Lager Word-dokument
│   └── lag-pptx.js       ← Lager PowerPoint
├── vercel.json
└── package.json
```

---

## Deploy til Vercel

### Steg 1: Last opp til GitHub
1. Pakk ut ZIP-filen
2. Gå til github.com → New repository → `grammatikk-generator`
3. Last opp alle filene

### Steg 2: Koble til Vercel
1. Gå til vercel.com → New Project
2. Velg GitHub-repoet `grammatikk-generator`
3. Klikk Deploy

### Steg 3: Legg inn miljøvariabler (ikke nødvendig – bruker API-nøkkel direkte i nettleseren)

> Merk: Denne appen lar læreren taste inn sin Google Gemini API-nøkkel direkte i nettleseren.
> Nøkkelen lagres kun i sessionStorage og sendes kun til Google sitt API.
> Ingen nøkkel lagres på serveren.

---

## Bruk

1. Åpne nettsiden på Vercel-URL
2. Lim inn din Google Gemini API-nøkkel (AIza...)
3. Velg CEFR-nivå (A0–C1)
4. Velg grammatikktema (50 å velge mellom)
5. Klikk «Generer grammatikkoppgaver»
6. Last ned Word-dokument og/eller PowerPoint

---

## API-nøkkel

Hent nøkkel på: **aistudio.google.com/app/api-keys**
Nøkkelen starter med `AIza...` 

---

## Teknologi

- **Frontend**: Ren HTML/CSS/JavaScript (ingen rammeverk)
- **Backend**: Vercel serverless functions (Node.js)
- **AI**: Google Gemini 2.0 Flash (`generativelanguage.googleapis.com`)
- **Word**: `docx` npm-pakke
- **PowerPoint**: `pptxgenjs` npm-pakke
- ## Oppdatert
