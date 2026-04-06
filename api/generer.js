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

  if (!tema || !niva) {
    return res.status(200).json({ ok: true });
  }

  // CEFR-beskrivelser og pedagogiske krav per nivå
  const cefrProfil = {
    "A0": {
      beskrivelse: "Nybegynner uten norskkunnskaper. Forstår og bruker enkeltord og svært korte fraser. Setninger på maks 4-5 ord. Alltid bildestøtte og ordbank.",
      tekstlengde: "40-60 ord",
      setningslengde: "maks 5 ord per setning",
      ordforrad: "kun de mest grunnleggende hverdagsord",
      oppgavetyper: "matching, avkrysning, bildestøtte, fyll inn med ordbank",
      grammatikk: "kun presens av være og ha, enkle substantiv"
    },
    "A1": {
      beskrivelse: "Kan forstå og bruke kjente ord og svært enkle fraser om konkrete behov. Enkle setninger på 5-8 ord. Alltid ordbank i oppgavene.",
      tekstlengde: "60-80 ord",
      setningslengde: "5-8 ord per setning",
      ordforrad: "grunnleggende hverdagsord om familie, mat, jobb, tid",
      oppgavetyper: "fyll inn med ordbank, sant/usant, koble ord til definisjon",
      grammatikk: "presens, enkle substantiv, personlige pronomen, tall"
    },
    "A2": {
      beskrivelse: "Kan forstå setninger og hyppig brukte uttrykk om nære temaer. Setninger på 8-12 ord. Ordbank kan brukes, men er ikke påkrevd.",
      tekstlengde: "80-110 ord",
      setningslengde: "8-12 ord per setning",
      ordforrad: "hverdagsspråk om arbeid, lokalmiljø, handel, familie",
      oppgavetyper: "fyll inn, sant/usant, sett i rekkefølge, enkel skriveoppgave",
      grammatikk: "presens og preteritum, V2-regelen, adjektivbøying, preposisjoner"
    },
    "B1": {
      beskrivelse: "Kan forstå hovedinnholdet i klart språk om kjente emner. Setninger på 10-15 ord. Mer kompleks grammatikk og lengre tekster.",
      tekstlengde: "110-150 ord",
      setningslengde: "10-15 ord per setning",
      ordforrad: "arbeidsliv, samfunn, aktuelle hendelser, mer abstrakt språk",
      oppgavetyper: "skriveoppgave, argumentasjon, setningsanalyse, omformulering",
      grammatikk: "perfektum, modalverb, leddsetninger, sammensatte ord"
    },
    "B2": {
      beskrivelse: "Kan forstå hovedinnholdet i komplekse tekster. Setninger på 15-20 ord. Avansert grammatikk og nyansert språkbruk.",
      tekstlengde: "150-200 ord",
      setningslengde: "15-20 ord per setning",
      ordforrad: "fagspråk, formelt og uformelt register, nyanser og idiomer",
      oppgavetyper: "analyse, diskusjon, omskriving, tekstproduksjon, feilretting",
      grammatikk: "passiv, kondisjonalis, avansert setningsbygging, register"
    },
    "C1": {
      beskrivelse: "Kan uttrykke seg flytende og spontant. Setninger på 20+ ord. Komplekst og nyansert språk på linje med morsmålsbrukere.",
      tekstlengde: "200-250 ord",
      setningslengde: "20+ ord, komplekse setningskonstruksjoner",
      ordforrad: "akademisk og faglig språk, idiomer, stilistiske nyanser",
      oppgavetyper: "akademisk skriving, analyse av autentiske tekster, diskusjon",
      grammatikk: "alle konstruksjoner inkl. stilistiske varianter og sjeldne former"
    }
  };

  const profil = cefrProfil[niva] || cefrProfil["A2"];

  const prompt = `Du er en erfaren norsklærer og pedagog med ekspertise i CEFR-rammeverket. Du lager grammatikkmateriell for voksne innvandrere.

NIVÅ: ${niva}
TEMA: ${tema}

CEFR-PROFIL FOR ${niva}:
${profil.beskrivelse}
- Tekstlengde: ${profil.tekstlengde}
- Setningslengde: ${profil.setningslengde}
- Ordforråd: ${profil.ordforrad}
- Anbefalte oppgavetyper: ${profil.oppgavetyper}
- Grammatikkfokus: ${profil.grammatikk}

ABSOLUTTE KRAV TIL SPRÅKKVALITET:
1. All norsk tekst skal følge bokmålsnormen (Bokmålsordboka, Språkrådet).
2. Korrekt bruk av stor og liten forbokstav: stor bokstav kun etter punktum, spørsmålstegn, utropstegn og i egennavn.
3. Korrekt tegnsetting: komma foran leddsetninger (fordi, når, som, at, hvis), punktum etter fullstendige setninger.
4. Korrekt bøying: substantivbøying (ubestemt/bestemt entall/flertall), verbkonjugasjon, adjektivkongruens.
5. Korrekt ordstilling: V2-regelen i helsetninger, verb sist i leddsetninger.
6. Ingen sammenblanding av nynorsk og bokmål.
7. Fasit MÅ stemme 100% med oppgavene – sjekk nøye at svarene er riktige.

Svar KUN med dette JSON-objektet. Ingen tekst utenfor. Ingen backticks. Kun JSON:

{
  "tema": "${tema}",
  "niva": "${niva}",
  "forklaring": "Skriv en pedagogisk forklaring på ${niva}-nivå om ${tema}. Bruk enkelt, korrekt bokmål. Lengde tilpasset nivå: ${profil.tekstlengde}. Gi 2-3 konkrete eksempelsetninger. Forklar regelen med egne ord slik en god lærer ville gjort det.",
  "grammatikkForklaring": "List opp 4-6 grammatiske mønstre eller regler for ${tema}. Bruk bindestrek foran hvert punkt. Tilpass kompleksitet til ${niva}. Vis alltid eksempel etter regelen med pil (->). Eksempel format: - Regelbeskrivelse -> Eksempelsetning",
  "lesetekst": "Skriv en sammenhengende, naturlig norsk tekst på ${profil.tekstlengde} om hverdagslivet eller arbeidslivet. Teksten skal inneholde mange eksempler på ${tema}. Tilpass vokabular og setningslengde (${profil.setningslengde}) til ${niva}. Teksten skal leses flytende og ikke virke konstruert.",
  "oppgaver": [
    {
      "type": "VELG PASSENDE OPPGAVETYPE for ${niva}: ${profil.oppgavetyper}",
      "instruksjon": "Klar og tydelig instruksjon tilpasset ${niva}-nivå. Bruk enkle ord på lavere nivåer.",
      "innhold": ["Skriv 5 oppgavepunkter direkte relatert til ${tema}. Vanskelighetsgrad tilpasset ${niva}."],
      "skrivelinje": 0
    },
    {
      "type": "VELG ANNEN PASSENDE OPPGAVETYPE",
      "instruksjon": "Instruksjon",
      "innhold": ["5 oppgavepunkter"],
      "skrivelinje": 0
    },
    {
      "type": "VELG PASSENDE OPPGAVETYPE",
      "instruksjon": "Instruksjon",
      "innhold": ["5 oppgavepunkter"],
      "skrivelinje": 3
    },
    {
      "type": "Skriveoppgave",
      "instruksjon": "Tilpass skriveoppgaven til ${niva}. A0-A1: skriv 2-3 ord eller setninger. A2: skriv 3-5 setninger. B1: skriv et avsnitt. B2-C1: skriv en lengre tekst med argumentasjon.",
      "innhold": [],
      "skrivelinje": 5
    },
    {
      "type": "VELG PASSENDE OPPGAVETYPE",
      "instruksjon": "Instruksjon",
      "innhold": ["5 oppgavepunkter"],
      "skrivelinje": 0
    }
  ],
  "fasit": "Oppgave A:\\n1. [korrekt svar]\\n2. [korrekt svar]\\n3. [korrekt svar]\\n4. [korrekt svar]\\n5. [korrekt svar]\\n\\nOppgave B:\\n1. [korrekt svar]\\n2. [korrekt svar]\\n3. [korrekt svar]\\n4. [korrekt svar]\\n5. [korrekt svar]\\n\\nOppgave C:\\n1. [korrekt svar]\\n2. [korrekt svar]\\n3. [korrekt svar]\\n4. [korrekt svar]\\n5. [korrekt svar]\\n\\nOppgave D: Åpen oppgave – se vurderingskriterier i lærerveiledning.\\n\\nOppgave E:\\n1. [korrekt svar]\\n2. [korrekt svar]\\n3. [korrekt svar]\\n4. [korrekt svar]\\n5. [korrekt svar]",
  "bilde1": "2-3 engelske ord for et relevant fotografi til tittelsliden for temaet ${tema}",
  "bilde2": "2-3 engelske ord for et relevant fotografi til eksempelsliden for temaet ${tema}"
}

VIKTIG: Erstatt ALLE plassholdertekster med ekte, korrekt norsk innhold. Tilpass ALT til ${niva}-nivå. Kontroller at fasit stemmer med oppgavene.`;

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const svar = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          maxOutputTokens: 8000,
          temperature: 0.4,
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

    try {
      const innhold = JSON.parse(tekst);
      return res.status(200).json({ innhold });
    } catch (parseErr) {
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
