// api/generer.js
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ feil: "Kun POST er tillatt" });

  const { passord, tema, niva } = req.body;

  const riktigPassord = process.env.APP_PASSORD || "mbo2026";
  if (passord !== riktigPassord) {
    return res.status(401).json({ feil: "Feil passord" });
  }

  const apiKey = process.env.GOOGLE_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ feil: "Google API-nøkkel ikke konfigurert på serveren" });
  }

  // Ping-sjekk (ingen tema/niva = bare passord-test)
  if (!tema || !niva) {
    return res.status(200).json({ ok: true });
  }

  const prompt = `Du er en norsklærer som lager grammatikkoppgaver for voksne innvandrere på nivå ${niva}.
Lag materiale om: ${tema}

Svar KUN med dette JSON-objektet. Ingen tekst før eller etter. Ingen backticks. Kun JSON:

{
  "tema": "${tema}",
  "niva": "${niva}",
  "forklaring": "Skriv 5-7 setninger som forklarer ${tema} på enkel norsk bokmål tilpasset ${niva}-nivå. Gi konkrete eksempler.",
  "grammatikkForklaring": "Skriv 4-6 mønstre eller regler for ${tema}. Bruk bindestrek foran hvert punkt. Vis gjerne eksempler med pil (->).",
  "lesetekst": "Skriv en naturlig norsk tekst på ca 120 ord om hverdagslivet som inneholder mange eksempler på ${tema}. Tilpass til ${niva}-nivå.",
  "oppgaver": [
    {
      "type": "Fyll inn",
      "instruksjon": "Fyll inn riktig form i setningene. Velg fra ordene i parentes.",
      "innhold": [
        "1. Skriv en setning med blank relatert til ${tema}",
        "2. Skriv en setning med blank relatert til ${tema}",
        "3. Skriv en setning med blank relatert til ${tema}",
        "4. Skriv en setning med blank relatert til ${tema}",
        "5. Skriv en setning med blank relatert til ${tema}"
      ],
      "skrivelinje": 0
    },
    {
      "type": "Riktig eller galt",
      "instruksjon": "Er setningen riktig (R) eller gal (G)? Rett de gale setningene.",
      "innhold": [
        "1. ___ Skriv en setning om ${tema}",
        "2. ___ Skriv en setning om ${tema}",
        "3. ___ Skriv en setning om ${tema}",
        "4. ___ Skriv en setning om ${tema}",
        "5. ___ Skriv en setning om ${tema}"
      ],
      "skrivelinje": 0
    },
    {
      "type": "Sett i riktig rekkefølge",
      "instruksjon": "Sett ordene i riktig rekkefølge og skriv hele setningen.",
      "innhold": [
        "1. ord / ord / ord / ord",
        "2. ord / ord / ord / ord",
        "3. ord / ord / ord / ord",
        "4. ord / ord / ord / ord",
        "5. ord / ord / ord / ord"
      ],
      "skrivelinje": 5
    },
    {
      "type": "Skriveoppgave",
      "instruksjon": "Skriv 5 egne setninger der du bruker ${tema} riktig. Skriv om deg selv eller hverdagen din.",
      "innhold": [],
      "skrivelinje": 6
    },
    {
      "type": "Velg riktig alternativ",
      "instruksjon": "Sett ring rundt riktig alternativ i parentes.",
      "innhold": [
        "1. Setning med (alternativ1 / alternativ2 / alternativ3)",
        "2. Setning med (alternativ1 / alternativ2 / alternativ3)",
        "3. Setning med (alternativ1 / alternativ2 / alternativ3)",
        "4. Setning med (alternativ1 / alternativ2 / alternativ3)",
        "5. Setning med (alternativ1 / alternativ2 / alternativ3)"
      ],
      "skrivelinje": 0
    }
  ],
  "fasit": "Oppgave A: 1. svar 2. svar 3. svar 4. svar 5. svar. Oppgave B: 1. R 2. G-rettelse 3. R 4. G-rettelse 5. R. Oppgave C: 1. Korrekt setning. 2. Korrekt setning. 3. Korrekt setning. 4. Korrekt setning. 5. Korrekt setning. Oppgave D: Åpen oppgave. Oppgave E: 1. riktig 2. riktig 3. riktig 4. riktig 5. riktig"
}

VIKTIG: Erstatt alle plassholdertekster med ekte innhold om ${tema} på ${niva}-nivå. Fasit må stemme med oppgavene du skriver.`;

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const svar = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          maxOutputTokens: 8000,
          temperature: 0.5,
          responseMimeType: "application/json"
        }
      })
    });

    if (!svar.ok) {
      const feil = await svar.json();
      return res.status(svar.status).json({
        feil: feil.error?.message || "Feil fra Google Gemini API"
      });
    }

    const data = await svar.json();
    const tekst = data.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!tekst) return res.status(500).json({ feil: "Tomt svar fra Gemini" });

    // Parse JSON direkte
    try {
      const innhold = JSON.parse(tekst);
      return res.status(200).json({ innhold });
    } catch (parseErr) {
      // Prøv å rense og parse på nytt
      let renset = tekst.trim();
      renset = renset.replace(/^```[a-z]*\n?/i, '').replace(/\n?```$/, '').trim();
      const start = renset.indexOf('{');
      const slutt = renset.lastIndexOf('}');
      if (start !== -1 && slutt !== -1) {
        const innhold = JSON.parse(renset.slice(start, slutt + 1));
        return res.status(200).json({ innhold });
      }
      return res.status(500).json({ feil: "Kunne ikke tolke svar fra Gemini. Prøv igjen." });
    }

  } catch (feil) {
    console.error("Serverfeil:", feil);
    return res.status(500).json({ feil: "Serverfeil: " + feil.message });
  }
};
