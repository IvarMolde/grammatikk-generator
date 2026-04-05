// api/lag-pptx.js
// Vercel serverless function – genererer PowerPoint fra grammatikkdata

const pptxgen = require("pptxgenjs");

const C = {
  primary: "1F4E79", accent: "2E75B6", light: "D6E4F0",
  white: "FFFFFF", dark: "1A2744", green: "2C6E49",
  lightGreen: "D8F3DC", yellow: "FFF3CD", gray: "F5F5F5",
  textDark: "1A1A2E", textMid: "444466", amber: "E9A800",
};

const W = 10, H = 5.625;
const mk = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

async function lagPresentasjon(data) {
  const { tema, niva, forklaring, grammatikkForklaring, lesetekst, oppgaver } = data;
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${tema} – Nivå ${niva}`;
  pres.author = 'Molde voksenopplæringssenter';

  // ── Slide 1: Tittel ──
  const s1 = pres.addSlide();
  s1.background = { color: C.primary };
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText(tema, { x: 0.7, y: 1.1, w: 8.6, h: 1.8, fontSize: 42, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
  s1.addShape(pres.shapes.RECTANGLE, { x: 3.8, y: 3.0, w: 2.4, h: 0.65, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText(`Nivå ${niva}`, { x: 3.8, y: 3.0, w: 2.4, h: 0.65, fontSize: 20, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  s1.addText("Molde voksenopplæringssenter – MBO", { x: 0.7, y: 3.9, w: 8.6, h: 0.5, fontSize: 13, color: "A0B8D8", fontFace: "Calibri", align: "center" });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.08, w: W, h: 0.08, fill: { color: C.accent }, line: { color: C.accent } });
  s1.addNotes(`Tittelslide: ${tema} – Nivå ${niva}. Spør elevene hva de vet om dette grammatikktemaet.`);

  // ── Slide 2: Læringsmål ──
  const s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: H, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addText("Læringsmål", { x: 0.3, y: 0.2, w: 9.4, h: 0.7, fontSize: 28, bold: true, color: C.primary, fontFace: "Calibri" });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.3, h: 3.9, fill: { color: C.light }, line: { color: C.accent, pt: 1 }, shadow: mk() });
  s2.addText([
    { text: `Etter denne leksjonen kan jeg:`, options: { bold: true, breakLine: true, fontSize: 17 } },
    { text: " ", options: { breakLine: true } },
    { text: `  \u2022  Forklare hva ${tema} er`, options: { breakLine: true, fontSize: 15 } },
    { text: `  \u2022  Gjenkjenne ${tema} i tekster`, options: { breakLine: true, fontSize: 15 } },
    { text: `  \u2022  Lage egne setninger med riktig bruk`, options: { breakLine: true, fontSize: 15 } },
    { text: `  \u2022  Snakke om temaet med medelever`, options: { fontSize: 15 } },
  ], { x: 0.6, y: 1.25, w: 8.8, h: 3.6, color: C.textDark, fontFace: "Calibri", valign: "top" });
  s2.addNotes(`Gå gjennom læringsmålene. Spør elevene om de har sett eksempler på ${tema} i norsk. Gjenta målene på slutten av timen.`);

  // ── Slide 3: Forklaring ──
  const s3 = pres.addSlide();
  s3.background = { color: "F8FBFF" };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Grammatikk – forklaring", { x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.white, fontFace: "Calibri" });
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 4.0, fill: { color: C.white }, line: { color: C.light, pt: 1 }, shadow: mk() });

  const forkLinjer = (forklaring || '').split('\n').filter(l => l.trim()).slice(0, 9);
  const forkRuns = forkLinjer.map((l, i) => ({ text: l, options: { breakLine: i < forkLinjer.length - 1, fontSize: 14 } }));
  if (forkRuns.length > 0) {
    s3.addText(forkRuns, { x: 0.55, y: 1.25, w: 9.0, h: 3.7, color: C.textDark, fontFace: "Calibri", valign: "top" });
  }
  s3.addNotes(`Les forklaringen høyt. Pause ved hvert punkt og sjekk forståelse. Be elevene komme med egne eksempler.`);

  // ── Slide 4: Mønstre og eksempler ──
  const s4 = pres.addSlide();
  s4.background = { color: C.white };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.green }, line: { color: C.green } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: W, h: 0.92, fill: { color: C.lightGreen }, line: { color: C.lightGreen } });
  s4.addText("Mønstre og eksempler", { x: 0.4, y: 0.18, w: 9.2, h: 0.68, fontSize: 26, bold: true, color: C.green, fontFace: "Calibri" });

  const eksLinjer = (grammatikkForklaring || forklaring || '').split('\n').filter(l => l.trim()).slice(0, 5);
  eksLinjer.forEach((ex, i) => {
    const y = 1.15 + i * 0.85;
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y, w: 9.4, h: 0.72, fill: { color: i % 2 === 0 ? C.lightGreen : "F0FFF4" }, line: { color: C.green, pt: 0.5 } });
    s4.addText(ex.replace(/^[-\u2022*]\s*/, ''), { x: 0.5, y: y + 0.06, w: 9.0, h: 0.6, fontSize: 14, color: C.textDark, fontFace: "Calibri", valign: "middle" });
  });
  if (eksLinjer.length === 0) {
    s4.addText("Se eksempler i arbeidsarket", { x: 0.5, y: 2.0, w: 9.0, h: 1.0, fontSize: 15, color: C.green, fontFace: "Calibri", align: "center" });
  }
  s4.addNotes(`Gå gjennom eksemplene. Bruk tavlen til å skrive flere. La elevene lage egne lignende setninger.`);

  // ── Slide 5: Lesetekst ──
  const s5 = pres.addSlide();
  s5.background = { color: "FFFEF5" };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Lesetekst", { x: 0.4, y: 0.18, w: 9.2, h: 0.64, fontSize: 26, bold: true, color: C.dark, fontFace: "Calibri" });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 4.0, fill: { color: C.white }, line: { color: "DDCC88", pt: 1.5 }, shadow: mk() });

  const leseLinjer = (lesetekst || '').split('\n').filter(l => l.trim()).slice(0, 10);
  const leseRuns = leseLinjer.map((l, i) => ({ text: l, options: { breakLine: i < leseLinjer.length - 1, fontSize: 13 } }));
  if (leseRuns.length > 0) {
    s5.addText(leseRuns, { x: 0.55, y: 1.25, w: 9.0, h: 3.7, color: C.textDark, fontFace: "Calibri", valign: "top" });
  }
  s5.addNotes(`Les teksten høyt. Stopp ved setninger som illustrerer ${tema}. Spør: «Hva slags ord er dette?»`);

  // ── Slide 6: Oppgaver ──
  const s6 = pres.addSlide();
  s6.background = { color: C.white };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: W, h: 0.92, fill: { color: C.light }, line: { color: C.light } });
  s6.addText("Oppgaver – oversikt", { x: 0.4, y: 0.18, w: 9.2, h: 0.68, fontSize: 26, bold: true, color: C.primary, fontFace: "Calibri" });

  const brev = ['A', 'B', 'C', 'D', 'E'];
  (oppgaver || []).slice(0, 5).forEach((oppg, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const x = col === 0 ? 0.3 : 5.3;
    const y = 1.15 + row * 1.2;
    s6.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.7, h: 1.0, fill: { color: i % 2 === 0 ? C.light : "EEF5FF" }, line: { color: C.accent, pt: 0.5 }, shadow: mk() });
    s6.addText(`${brev[i]})  ${oppg.type}`, { x: x + 0.12, y: y + 0.07, w: 4.5, h: 0.38, fontSize: 13, bold: true, color: C.primary, fontFace: "Calibri" });
    const instruksjonsKort = (oppg.instruksjon || '').slice(0, 68) + ((oppg.instruksjon || '').length > 68 ? '…' : '');
    s6.addText(instruksjonsKort, { x: x + 0.12, y: y + 0.48, w: 4.5, h: 0.46, fontSize: 11, color: C.textMid, fontFace: "Calibri" });
  });
  s6.addNotes(`Presenter oppgaveoversikten. Forklar instruksjonene for hver oppgavetype. Elevene jobber i arbeidsarket.`);

  // ── Slide 7: Diskusjon ──
  const s7 = pres.addSlide();
  s7.background = { color: C.primary };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber } });
  s7.addText("Diskusjon og refleksjon", { x: 0.5, y: 0.4, w: 9.0, h: 0.9, fontSize: 30, bold: true, color: C.white, fontFace: "Calibri", align: "center" });

  const spørsmål = [
    `Forklar ${tema} til en klassekamerat med egne ord.`,
    `Lag en setning med ${tema} om noe fra din hverdag.`,
    `Når bruker vi dette på norsk? Gi et eksempel.`,
    `Hva er den vanligste feilen folk gjør med ${tema}?`
  ];
  spørsmål.forEach((spm, i) => {
    const y = 1.45 + i * 0.95;
    s7.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.2, h: 0.78, fill: { color: C.white, transparency: 88 }, line: { color: C.white, transparency: 60, pt: 0.5 } });
    s7.addText(`${i + 1}.  ${spm}`, { x: 0.6, y: y + 0.08, w: 8.8, h: 0.62, fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle" });
  });
  s7.addNotes(`Avsluttende diskusjon i par eller grupper. Oppsummer: Hva har vi lært om ${tema}? Knytt tilbake til læringsmålene.`);

  return pres;
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ feil: "Kun POST er tillatt" });

  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ feil: "Mangler data" });

    const pres = await lagPresentasjon(data);
    const filnavn = `${(data.tema || 'grammatikk').replace(/[^a-zA-ZæøåÆØÅ0-9]/g, '_')}_${data.niva || 'A1'}_presentasjon.pptx`;

    const buffer = await pres.write({ outputType: 'nodebuffer' });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filnavn}"`);
    res.send(buffer);

  } catch (e) {
    console.error(e);
    res.status(500).json({ feil: "Feil ved generering av PowerPoint: " + e.message });
  }
};
