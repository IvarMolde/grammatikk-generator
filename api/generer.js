// api/generer.js
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ feil: "Kun POST er tillatt" });

  const { passord, system, prompt, maxTokens } = req.body;

  const riktigPassord = process.env.APP_PASSORD || "mbo2026";
  if (passord !== riktigPassord) {
    return res.status(401).json({ feil: "Feil passord" });
  }

  const apiKey = process.env.GOOGLE_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ feil: "Google API-nøkkel ikke konfigurert på serveren" });
  }

  if (!system || !prompt) {
    return res.status(400).json({ feil: "Mangler system eller prompt" });
  }

  try {
    // gemini-2.0-flash-lite er stabil og fungerer for alle brukere
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    const svar = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        system_instruction: { parts: [{ text: system }] },
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          maxOutputTokens: maxTokens || 8000,
          temperature: 0.7,
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

    return res.status(200).json({ tekst });

  } catch (feil) {
    console.error("Serverfeil:", feil);
    return res.status(500).json({ feil: "Serverfeil: " + feil.message });
  }
};
