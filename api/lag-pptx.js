// api/lag-pptx.js
const pptxgen = require("pptxgenjs");

const C = {
  primary: "1F4E79", accent: "2E75B6", light: "D6E4F0",
  white: "FFFFFF", dark: "1A2744", green: "2C6E49",
  lightGreen: "D8F3DC", yellow: "FFF3CD",
  textDark: "1A1A2E", textMid: "444466", amber: "E9A800",
};
const W = 10, H = 5.625;
const mk = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 });

// Hent bilde som base64 - bruker Unsplash API
async function hentBildeBase64(sokeord, apiKey) {
  try {
    // Hent bilde-URL fra Unsplash API
    const apiUrl = "https://api.unsplash.com/photos/random?query="
      + encodeURIComponent(sokeord)
      + "&orientation=landscape&client_id=" + apiKey;

    const ctrl = new AbortController();
    const timer = setTimeout(function() { ctrl.abort(); }, 7000);
    const resp = await fetch(apiUrl, { signal: ctrl.signal });
    clearTimeout(timer);

    if (!resp.ok) {
      console.error("Unsplash API feil " + resp.status + " for: " + sokeord);
      return null;
    }

    const json = await resp.json();
    // Bruk small-URL - bedre opplosning
    const url = json.urls && json.urls.small || json.urls.thumb;
    if (!url) { console.error("Ingen URL i svar"); return null; }

    // Last ned bildet
    const ctrl2 = new AbortController();
    const timer2 = setTimeout(function() { ctrl2.abort(); }, 10000);
    const imgResp = await fetch(url, { signal: ctrl2.signal });
    clearTimeout(timer2);

    if (!imgResp.ok) { console.error("Bilde-nedlasting feil: " + imgResp.status); return null; }

    const buf = await imgResp.arrayBuffer();
    const b64 = Buffer.from(buf).toString("base64");
    console.log("Bilde hentet OK for: " + sokeord + " (" + buf.byteLength + " bytes)");
    return "image/jpeg;base64," + b64;

  } catch (e) {
    console.error("hentBildeBase64 feil: " + e.message);
    return null;
  }
}

// Velg sokeord basert pa tema
function velgSokeord(tema) {
  const t = tema.toLowerCase();
  if (t.includes("verb"))        return "people working";
  if (t.includes("substantiv"))  return "objects table";
  if (t.includes("adjektiv"))    return "colorful nature";
  if (t.includes("setning"))     return "writing notebook";
  if (t.includes("pronomen"))    return "people talking";
  if (t.includes("preposisjon")) return "map direction";
  if (t.includes("konjunksjon")) return "bridge";
  if (t.includes("tall"))        return "numbers";
  if (t.includes("tid"))         return "clock";
  if (t.includes("passiv"))      return "classroom";
  if (t.includes("modal"))       return "crossroads choice";
  if (t.includes("imperativ"))   return "sign road";
  return "norway school learning";
}

async function lagPresentasjon(data, unsplashKey) {
  // Bruk Gemini-genererte sokeord hvis tilgjengelig
  var sokeord1 = data.bilde1 || velgSokeord(data.tema || "");
  var sokeord2 = data.bilde2 || "norway classroom students learning";
  const tema = data.tema || "";
  const niva = data.niva || "";
  const forklaring = data.forklaring || "";
  const grammatikkForklaring = data.grammatikkForklaring || "";
  const lesetekst = data.lesetekst || "";
  const oppgaver = data.oppgaver || [];

  // Hent bilder
  var bilde1 = null;
  var bilde2 = null;
  if (unsplashKey) {
    console.log("Henter bilder fra Unsplash...");
    var res = await Promise.all([
      hentBildeBase64(sokeord1, unsplashKey),
      hentBildeBase64(sokeord2, unsplashKey)
    ]);
    bilde1 = res[0];
    bilde2 = res[1];
    console.log("bilde1: " + (bilde1 ? "OK" : "FEILET"));
    console.log("bilde2: " + (bilde2 ? "OK" : "FEILET"));
  } else {
    console.log("Ingen UNSPLASH_ACCESS_KEY konfigurert");
  }

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = tema + " – Niva " + niva;

  // ── Slide 1: Tittel ──
  var s1 = pres.addSlide();

  if (bilde1) {
    s1.addImage({ data: bilde1, x: 0, y: 0, w: W * 0.5, h: H, sizing: { type: "cover", w: W * 0.5, h: H } });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W * 0.5, h: H, fill: { color: C.primary, transparency: 40 }, line: { color: C.primary, transparency: 40 } });
  } else {
    // Fallback: gradient-effekt med to rektangler
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W * 0.5, h: H, fill: { color: "1A5C9A" }, line: { color: "1A5C9A" } });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W * 0.5, h: H * 0.5, fill: { color: "2470B8", transparency: 30 }, line: { color: "2470B8", transparency: 30 } });
  }

  s1.addShape(pres.shapes.RECTANGLE, { x: W * 0.5, y: 0, w: W * 0.5, h: H, fill: { color: C.primary }, line: { color: C.primary } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.07, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.07, w: W, h: 0.07, fill: { color: C.accent }, line: { color: C.accent } });
  s1.addText(tema, { x: W * 0.5 + 0.3, y: 1.1, w: W * 0.46, h: 2.1, fontSize: 34, bold: true, color: C.white, fontFace: "Calibri", align: "left", valign: "middle" });
  s1.addShape(pres.shapes.RECTANGLE, { x: W * 0.5 + 0.3, y: 3.3, w: 1.9, h: 0.58, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText("Niva " + niva, { x: W * 0.5 + 0.3, y: 3.3, w: 1.9, h: 0.58, fontSize: 18, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  s1.addText("Molde voksenopplaringssenter - MBO", { x: W * 0.5 + 0.3, y: 4.1, w: W * 0.46, h: 0.4, fontSize: 11, color: "A0B8D8", fontFace: "Calibri" });
  s1.addNotes("Tittelslide: " + tema + " – Niva " + niva + ". Spor elevene hva de vet om dette grammatikktemaet.");

  // ── Slide 2: Laeringsmal ──
  var s2 = pres.addSlide();
  s2.background = { color: C.white };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: H, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addText("Laeringsmal", { x: 0.3, y: 0.18, w: 9.4, h: 0.7, fontSize: 28, bold: true, color: C.primary, fontFace: "Calibri" });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.3, h: 3.9, fill: { color: C.light }, line: { color: C.accent, pt: 1 }, shadow: mk() });
  s2.addText([
    { text: "Etter denne leksjonen kan jeg:", options: { bold: true, breakLine: true, fontSize: 17 } },
    { text: " ", options: { breakLine: true } },
    { text: "  *  Forklare hva " + tema + " er", options: { breakLine: true, fontSize: 15 } },
    { text: "  *  Gjenkjenne " + tema + " i tekster", options: { breakLine: true, fontSize: 15 } },
    { text: "  *  Lage egne setninger med riktig bruk", options: { breakLine: true, fontSize: 15 } },
    { text: "  *  Snakke om temaet med medelever", options: { fontSize: 15 } },
  ], { x: 0.6, y: 1.25, w: 8.8, h: 3.6, color: C.textDark, fontFace: "Calibri", valign: "top" });
  s2.addNotes("Ga gjennom laeringsmalene. Spor elevene om de har sett eksempler pa " + tema + ".");

  // ── Slide 3: Forklaring ──
  var s3 = pres.addSlide();
  s3.background = { color: "F8FBFF" };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Grammatikk - forklaring", { x: 0.4, y: 0.18, w: 9.2, h: 0.65, fontSize: 26, bold: true, color: C.white, fontFace: "Calibri" });
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 4.0, fill: { color: C.white }, line: { color: C.light, pt: 1 }, shadow: mk() });
  var forkLinjer = forklaring.split("\n").filter(function(l) { return l.trim(); }).slice(0, 9);
  if (forkLinjer.length > 0) {
    var forkRuns = forkLinjer.map(function(l, i) {
      return { text: l, options: { breakLine: i < forkLinjer.length - 1, fontSize: 14 } };
    });
    s3.addText(forkRuns, { x: 0.55, y: 1.25, w: 9.0, h: 3.7, color: C.textDark, fontFace: "Calibri", valign: "top" });
  }
  s3.addNotes("Les forklaringen hoyt. Pause og sjekk forstaelse ved hvert punkt.");

  // ── Slide 4: Monstre med bilde ──
  var s4 = pres.addSlide();
  s4.background = { color: C.white };
  var tw = bilde2 ? 5.3 : 9.4;
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.green }, line: { color: C.green } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: bilde2 ? 5.7 : W, h: 0.92, fill: { color: C.lightGreen }, line: { color: C.lightGreen } });
  s4.addText("Monstre og eksempler", { x: 0.4, y: 0.18, w: bilde2 ? 5.0 : 9.2, h: 0.68, fontSize: 24, bold: true, color: C.green, fontFace: "Calibri" });

  if (bilde2) {
    s4.addImage({ data: bilde2, x: 5.9, y: 1.0, w: 3.8, h: 4.3, sizing: { type: "cover", w: 3.8, h: 4.3 } });
  }

  var eksLinjer = grammatikkForklaring.split("\n").filter(function(l) { return l.trim(); }).slice(0, 5);
  eksLinjer.forEach(function(ex, i) {
    var y = 1.15 + i * 0.85;
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y, w: tw, h: 0.72, fill: { color: i % 2 === 0 ? C.lightGreen : "F0FFF4" }, line: { color: C.green, pt: 0.5 } });
    s4.addText(ex.replace(/^[-*]\s*/, ""), { x: 0.5, y: y + 0.06, w: tw - 0.2, h: 0.6, fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "middle" });
  });
  s4.addNotes("Ga gjennom eksemplene. La elevene lage egne lignende setninger.");

  // ── Slide 5: Lesetekst ──
  var s5 = pres.addSlide();
  s5.background = { color: "FFFEF5" };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Lesetekst", { x: 0.4, y: 0.18, w: 9.2, h: 0.65, fontSize: 26, bold: true, color: C.dark, fontFace: "Calibri" });
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 4.0, fill: { color: C.white }, line: { color: "DDCC88", pt: 1.5 }, shadow: mk() });
  var leseLinjer = lesetekst.split("\n").filter(function(l) { return l.trim(); }).slice(0, 10);
  if (leseLinjer.length > 0) {
    var leseRuns = leseLinjer.map(function(l, i) {
      return { text: l, options: { breakLine: i < leseLinjer.length - 1, fontSize: 13 } };
    });
    s5.addText(leseRuns, { x: 0.55, y: 1.25, w: 9.0, h: 3.7, color: C.textDark, fontFace: "Calibri", valign: "top" });
  }
  s5.addNotes("Les teksten hoyt. Stopp ved setninger som illustrerer " + tema + ".");

  // ── Slide 6: Oppgaver ──
  var s6 = pres.addSlide();
  s6.background = { color: C.white };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: W, h: 0.92, fill: { color: C.light }, line: { color: C.light } });
  s6.addText("Oppgaver - oversikt", { x: 0.4, y: 0.18, w: 9.2, h: 0.68, fontSize: 26, bold: true, color: C.primary, fontFace: "Calibri" });
  var brev = ["A","B","C","D","E"];
  oppgaver.slice(0, 5).forEach(function(oppg, i) {
    var col = i < 3 ? 0 : 1;
    var row = i < 3 ? i : i - 3;
    var x = col === 0 ? 0.3 : 5.3;
    var y = 1.15 + row * 1.2;
    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 4.7, h: 1.0, fill: { color: i % 2 === 0 ? C.light : "EEF5FF" }, line: { color: C.accent, pt: 0.5 }, shadow: mk() });
    s6.addText(brev[i] + ")  " + (oppg.type || ""), { x: x + 0.12, y: y + 0.07, w: 4.5, h: 0.38, fontSize: 13, bold: true, color: C.primary, fontFace: "Calibri" });
    var instr = (oppg.instruksjon || "").slice(0, 68) + ((oppg.instruksjon || "").length > 68 ? "..." : "");
    s6.addText(instr, { x: x + 0.12, y: y + 0.48, w: 4.5, h: 0.46, fontSize: 11, color: C.textMid, fontFace: "Calibri" });
  });
  s6.addNotes("Presenter oppgaveoversikten og forklar instruksjonene.");

  // ── Slide 7: Diskusjon ──
  var s7 = pres.addSlide();
  s7.background = { color: C.primary };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber } });
  s7.addText("Diskusjon og refleksjon", { x: 0.5, y: 0.4, w: 9.0, h: 0.9, fontSize: 30, bold: true, color: C.white, fontFace: "Calibri", align: "center" });
  var sporsmal = [
    "Forklar " + tema + " til en klassekamerat med egne ord.",
    "Lag en setning med " + tema + " om noe fra din hverdag.",
    "Nar bruker vi dette pa norsk? Gi et eksempel.",
    "Hva er den vanligste feilen folk gjor med " + tema + "?"
  ];
  sporsmal.forEach(function(spm, i) {
    var y = 1.45 + i * 0.95;
    s7.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: y, w: 9.2, h: 0.78, fill: { color: C.white, transparency: 88 }, line: { color: C.white, transparency: 60, pt: 0.5 } });
    s7.addText((i + 1) + ".  " + spm, { x: 0.6, y: y + 0.08, w: 8.8, h: 0.62, fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle" });
  });
  s7.addNotes("Avsluttende diskusjon. Oppsummer hva vi har laert om " + tema + ".");

  return pres;
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ feil: "Kun POST er tillatt" });

  try {
    var body = req.body;
    if (!body || !body.data) return res.status(400).json({ feil: "Mangler data" });

    var unsplashKey = process.env.UNSPLASH_ACCESS_KEY || null;
    console.log("UNSPLASH_ACCESS_KEY tilgjengelig: " + (unsplashKey ? "JA (" + unsplashKey.slice(0,6) + "...)" : "NEI"));

    var pres = await lagPresentasjon(body.data, unsplashKey);
    var filnavn = (body.data.tema || "grammatikk").replace(/[^a-zA-Z0-9]/g, "_") + "_" + (body.data.niva || "A1") + "_presentasjon.pptx";
    var buffer = await pres.write({ outputType: "nodebuffer" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", "attachment; filename=\"" + filnavn + "\"");
    res.send(buffer);

  } catch (e) {
    console.error("lag-pptx feil:", e);
    res.status(500).json({ feil: "Feil ved generering: " + e.message });
  }
};
