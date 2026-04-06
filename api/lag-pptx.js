// api/lag-pptx.js
const pptxgen = require("pptxgenjs");

const C = {
  primary: "1F4E79", accent: "2E75B6", light: "D6E4F0",
  white: "FFFFFF", dark: "1A2744", green: "2C6E49",
  lightGreen: "D8F3DC", amber: "E9A800", red: "C84B31",
  lightRed: "FFE4DE", lilla: "7B5EA7", lightLilla: "F0EBF8",
  textDark: "1A1A2E", textMid: "444466", gray: "F5F7FA"
};
const W = 10, H = 5.625;
const mk = () => ({ type: "outer", blur: 10, offset: 3, angle: 135, color: "000000", opacity: 0.10 });

// Hent bilde som base64
async function hentBilde(sokeord, apiKey) {
  try {
    const apiUrl = "https://api.unsplash.com/photos/random?query="
      + encodeURIComponent(sokeord)
      + "&orientation=landscape&client_id=" + apiKey;
    const c1 = new AbortController();
    const t1 = setTimeout(function() { c1.abort(); }, 7000);
    const r1 = await fetch(apiUrl, { signal: c1.signal });
    clearTimeout(t1);
    if (!r1.ok) { console.error("Unsplash feil " + r1.status); return null; }
    const json = await r1.json();
    const url = json.urls && (json.urls.regular || json.urls.small);
    if (!url) return null;
    const c2 = new AbortController();
    const t2 = setTimeout(function() { c2.abort(); }, 12000);
    const r2 = await fetch(url, { signal: c2.signal });
    clearTimeout(t2);
    if (!r2.ok) return null;
    const buf = await r2.arrayBuffer();
    console.log("Bilde OK: " + sokeord);
    return "image/jpeg;base64," + Buffer.from(buf).toString("base64");
  } catch (e) {
    console.error("Bilde feil: " + e.message);
    return null;
  }
}

// Hjelpefunksjoner
function bold(tekst, size, color) {
  return [{ text: tekst, options: { bold: true, fontSize: size || 18, color: color || C.textDark, fontFace: "Calibri" } }];
}

function kortTekst(tekst, maks) {
  if (!tekst) return "";
  return tekst.length > maks ? tekst.slice(0, maks - 1) + "…" : tekst;
}

function splittILinjer(tekst, maks) {
  if (!tekst) return [];
  // Del opp på punktum eller newline, ta maks N setninger
  var linjer = tekst.split(/\n|(?<=\.)\s+/).filter(function(l) { return l.trim().length > 3; });
  return linjer.slice(0, maks);
}

// Lag en ikonsirkel (farget sirkel med tall/bokstav)
function leggTilIkon(slide, x, y, r, farge, tekst) {
  slide.addShape(slide._pptx ? slide._pptx.shapes.OVAL : "ellipse", { x: x - r, y: y - r, w: r*2, h: r*2, fill: { color: farge }, line: { color: farge } });
  if (tekst) {
    slide.addText(tekst, { x: x - r, y: y - r * 1.1, w: r*2, h: r*2.2, fontSize: 14, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle" });
  }
}

async function lagPresentasjon(data, unsplashKey) {
  const tema = data.tema || "";
  const niva = data.niva || "";
  const forklaring = data.forklaring || "";
  const grammatikkForklaring = data.grammatikkForklaring || "";
  const lesetekst = data.lesetekst || "";
  const oppgaver = data.oppgaver || [];
  const sokeord1 = data.bilde1 || "language learning norway";
  const sokeord2 = data.bilde2 || "classroom students";

  var bilde1 = null, bilde2 = null;
  if (unsplashKey) {
    var res = await Promise.all([
      hentBilde(sokeord1, unsplashKey),
      hentBilde(sokeord2, unsplashKey)
    ]);
    bilde1 = res[0];
    bilde2 = res[1];
  }

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = tema + " – Nivå " + niva;
  pres.author = "Molde voksenopplæringssenter";

  // ══════════════════════════════════════
  // SLIDE 1: TITTEL – visuell impact
  // ══════════════════════════════════════
  var s1 = pres.addSlide();
  if (bilde1) {
    // Fullt bilde i bakgrunn med mørk overlay
    s1.addImage({ data: bilde1, x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
    s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 45 }, line: { color: "000000", transparency: 45 } });
  } else {
    s1.background = { color: C.primary };
  }
  // Gul toppstripe
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.1, fill: { color: C.amber }, line: { color: C.amber } });
  // Nivå-badge
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.55, y: 0.9, w: 1.4, h: 0.5, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText("Nivå " + niva, { x: 0.55, y: 0.9, w: 1.4, h: 0.5, fontSize: 16, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  // Tittel – stor og tydelig
  s1.addText(tema, {
    x: 0.5, y: 1.6, w: 9.0, h: 2.4,
    fontSize: 52, bold: true, color: C.white, fontFace: "Calibri",
    align: "left", valign: "middle"
  });
  // Institusjon
  s1.addText("Molde voksenopplæringssenter – MBO", {
    x: 0.5, y: 4.6, w: 9.0, h: 0.4,
    fontSize: 13, color: "CCDDEE", fontFace: "Calibri", align: "left"
  });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.1, w: W, h: 0.1, fill: { color: C.accent }, line: { color: C.accent } });
  s1.addNotes("Tittelslide: " + tema + " – Nivå " + niva + ".\nAktivering: Spør elevene – 'Hva vet dere om dette fra før?' La 2-3 elever svare. Skriv nøkkelord på tavlen.");

  // ══════════════════════════════════════
  // SLIDE 2: LÆRINGSMÅL – maks 3 mål, store og tydelige
  // ══════════════════════════════════════
  var s2 = pres.addSlide();
  s2.background = { color: C.white };
  // Toppbanner
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addText("Hva skal vi lære i dag?", { x: 0.5, y: 0.15, w: 9.0, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });
  // 3 store mål-kort
  var mal = [
    "Forstå og forklare: " + tema,
    "Gjenkjenne " + kortTekst(tema, 30) + " i tekster",
    "Bruke " + kortTekst(tema, 30) + " i egne setninger"
  ];
  var malFarger = [C.accent, C.green, "7B5EA7"];
  mal.forEach(function(m, i) {
    var x = 0.3 + i * 3.2;
    s2.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.3, w: 3.0, h: 3.8, fill: { color: malFarger[i] }, line: { color: malFarger[i] }, shadow: mk() });
    // Stort tall
    s2.addText((i + 1).toString(), { x: x, y: 1.5, w: 3.0, h: 1.2, fontSize: 52, bold: true, color: C.white, fontFace: "Calibri", align: "center" });
    // Mål-tekst
    s2.addText(m, { x: x + 0.15, y: 2.85, w: 2.7, h: 2.0, fontSize: 15, color: C.white, fontFace: "Calibri", align: "center", valign: "top" });
  });
  s2.addNotes("Læringsmål: Gå gjennom de tre målene. Si: 'På slutten av timen skal dere klare disse tre tingene.'\nTips: La elevene lese målene høyt – ett mål hver.");

  // ══════════════════════════════════════
  // SLIDE 3: GRAMMATIKKREGEL – én regel, stort og visuelt
  // ══════════════════════════════════════
  var s3 = pres.addSlide();
  s3.background = { color: C.gray };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Regelen", { x: 0.5, y: 0.1, w: 9.0, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  // Hent første 2 setninger fra forklaring = kjerneregelen
  var forkLinjer = splittILinjer(forklaring, 2);
  var kjerneRegel = forkLinjer.join(" ");

  // Stor regelkort i midten
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 1.9, fill: { color: C.white }, line: { color: C.accent, pt: 3 }, shadow: mk() });
  s3.addText(kortTekst(kjerneRegel, 200), {
    x: 0.6, y: 1.25, w: 8.8, h: 1.7,
    fontSize: 18, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "middle"
  });

  // Husk-boks
  var restForklaring = splittILinjer(forklaring, 5).slice(2).join(" ");
  if (restForklaring.length > 10) {
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.25, w: 9.2, h: 1.9, fill: { color: "FFF8E7" }, line: { color: C.amber, pt: 2 } });
    s3.addText("💡  " + kortTekst(restForklaring, 220), {
      x: 0.6, y: 3.35, w: 8.8, h: 1.6,
      fontSize: 15, color: "6B4F00", fontFace: "Calibri",
      align: "left", valign: "middle"
    });
  }
  s3.addNotes("Presentér kjerneregelen. Si regelen med egne ord – IKKE les fra sliden.\nSpør: 'Har noen sett dette i norsk før?' La elevene gi eksempler.\nSkriv regelen på tavlen med egne ord.");

  // ══════════════════════════════════════
  // SLIDE 4: EKSEMPEL – visuell before/after
  // ══════════════════════════════════════
  var s4 = pres.addSlide();
  s4.background = { color: C.white };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.green }, line: { color: C.green } });
  s4.addText("Se på eksemplene", { x: 0.5, y: 0.1, w: 9.0, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  // Hent eksempellinjer fra grammatikkForklaring
  var eksLinjer = grammatikkForklaring.split("\n").filter(function(l) { return l.trim().length > 3; }).slice(0, 4);

  if (bilde2 && eksLinjer.length > 0) {
    // To-kolonne: eksempler til venstre, bilde til høyre
    eksLinjer.forEach(function(ex, i) {
      var y = 1.1 + i * 1.1;
      var tekst = ex.replace(/^[-*•]\s*/, "").trim();
      s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y, w: 5.5, h: 0.9, fill: { color: i % 2 === 0 ? C.lightGreen : "F0FFF4" }, line: { color: C.green, pt: 1 } });
      s4.addText(kortTekst(tekst, 80), { x: 0.5, y: y + 0.05, w: 5.2, h: 0.8, fontSize: 14, color: C.textDark, fontFace: "Calibri", valign: "middle" });
    });
    s4.addImage({ data: bilde2, x: 6.1, y: 1.1, w: 3.6, h: 4.1, sizing: { type: "cover", w: 3.6, h: 4.1 } });
  } else {
    // Full bredde uten bilde
    eksLinjer.forEach(function(ex, i) {
      var y = 1.1 + i * 1.1;
      var tekst = ex.replace(/^[-*•]\s*/, "").trim();
      s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y, w: 9.4, h: 0.9, fill: { color: i % 2 === 0 ? C.lightGreen : "F0FFF4" }, line: { color: C.green, pt: 1 } });
      s4.addText(kortTekst(tekst, 120), { x: 0.5, y: y + 0.05, w: 9.1, h: 0.8, fontSize: 15, color: C.textDark, fontFace: "Calibri", valign: "middle" });
    });
  }
  s4.addNotes("Eksempler: Pek på hvert eksempel og les det høyt.\nSpør: 'Kan du forklare hvorfor dette er riktig?'\nBe elevene lage egne lignende setninger – enten muntlig eller på tavlen.");

  // ══════════════════════════════════════
  // SLIDE 5: LESETEKST – chunked, ikke vegg av tekst
  // ══════════════════════════════════════
  var s5 = pres.addSlide();
  s5.background = { color: "FFFEF5" };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Les teksten", { x: 0.5, y: 0.1, w: 6.0, h: 0.8, fontSize: 28, bold: true, color: C.dark, fontFace: "Calibri", valign: "middle" });
  // Instruksjonsboble
  s5.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: 0.15, w: 3.4, h: 0.7, fill: { color: "FFF3CD" }, line: { color: C.amber, pt: 1 } });
  s5.addText("Finn eksempler på " + kortTekst(tema, 25), { x: 6.4, y: 0.15, w: 3.2, h: 0.7, fontSize: 12, color: "7A5500", fontFace: "Calibri", valign: "middle" });

  // Teksten i én lesbar boks – ikke for mye
  var leseAvsnitt = lesetekst.split("\n").filter(function(l) { return l.trim().length > 3; }).slice(0, 6).join(" ");
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 4.1, fill: { color: C.white }, line: { color: "DDCC88", pt: 2 }, shadow: mk() });
  s5.addText(kortTekst(leseAvsnitt, 600), {
    x: 0.65, y: 1.3, w: 8.7, h: 3.8,
    fontSize: 15, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "top", paraSpaceAfter: 6
  });
  s5.addNotes("Lesetekst: Les teksten høyt – du eller en elev.\nOppgave: 'Finn alle eksemplene på " + tema + " i teksten. Understrek dem.'\nDiskuter: Hvilke ord/former la dere merke til?");

  // ══════════════════════════════════════
  // SLIDE 6: OPPGAVEOVERSIKT – visuell og oversiktlig
  // ══════════════════════════════════════
  var s6 = pres.addSlide();
  s6.background = { color: C.gray };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addText("Oppgavene dine", { x: 0.5, y: 0.1, w: 9.0, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var brev = ["A","B","C","D","E"];
  var oppgFarger = [C.accent, C.green, "7B5EA7", C.red, C.amber];
  var oppgLysFarger = [C.light, C.lightGreen, C.lightLilla, C.lightRed, "FFF8E7"];

  oppgaver.slice(0, 5).forEach(function(oppg, i) {
    var col = i < 3 ? 0 : 1;
    var row = i < 3 ? i : i - 3;
    var x = col === 0 ? 0.3 : 5.3;
    var y = 1.1 + row * 1.45;
    var w = col === 0 ? (i < 3 ? 4.6 : 4.6) : 4.6;

    // Farget venstre-stripe + lys bakgrunn
    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: 1.25, fill: { color: oppgLysFarger[i] }, line: { color: oppgFarger[i], pt: 1 } });
    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.45, h: 1.25, fill: { color: oppgFarger[i] }, line: { color: oppgFarger[i] } });
    // Bokstav i stripe
    s6.addText(brev[i], { x: x, y: y, w: 0.45, h: 1.25, fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    // Oppgavetype
    s6.addText(kortTekst(oppg.type || "", 30), { x: x + 0.55, y: y + 0.05, w: w - 0.65, h: 0.45, fontSize: 14, bold: true, color: oppgFarger[i], fontFace: "Calibri", valign: "middle" });
    // Instruksjon
    s6.addText(kortTekst(oppg.instruksjon || "", 65), { x: x + 0.55, y: y + 0.55, w: w - 0.65, h: 0.6, fontSize: 11, color: C.textMid, fontFace: "Calibri", valign: "top" });
  });
  s6.addNotes("Oppgaver: Gå gjennom instruksjonene for hver oppgave FØR elevene begynner.\nGjør ett eksempel på tavlen for oppgave A.\nElevene jobber i arbeidsarket – sirkuler og hjelp underveis.");

  // ══════════════════════════════════════
  // SLIDE 7: PAR-AKTIVITET – konkret og handlingsorientert
  // ══════════════════════════════════════
  var s7 = pres.addSlide();
  s7.background = { color: C.white };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.0, fill: { color: "7B5EA7" }, line: { color: "7B5EA7" } });
  s7.addText("🗣  Snakk med en partner", { x: 0.5, y: 0.1, w: 9.0, h: 0.8, fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var aktiviteter = [
    { ikon: "1️⃣", tittel: "Forklar regelen", tekst: "Forklar " + kortTekst(tema, 25) + " for partneren din med egne ord" },
    { ikon: "2️⃣", tittel: "Lag en setning", tekst: "Lag én setning om deg selv med " + kortTekst(tema, 25) },
    { ikon: "3️⃣", tittel: "Finn feilen", tekst: "Hva er galt? Rett setningen: " + lagFeilSetning(tema) }
  ];

  aktiviteter.forEach(function(akt, i) {
    var y = 1.15 + i * 1.4;
    s7.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y, w: 9.4, h: 1.2, fill: { color: C.lightLilla }, line: { color: "7B5EA7", pt: 1 }, shadow: mk() });
    s7.addText(akt.ikon, { x: 0.4, y: y + 0.05, w: 0.7, h: 1.1, fontSize: 28, fontFace: "Calibri", valign: "middle", align: "center" });
    s7.addText(akt.tittel, { x: 1.2, y: y + 0.08, w: 8.0, h: 0.45, fontSize: 16, bold: true, color: "5A3E88", fontFace: "Calibri", valign: "middle" });
    s7.addText(akt.tekst, { x: 1.2, y: y + 0.55, w: 8.0, h: 0.55, fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "middle" });
  });
  s7.addNotes("Par-aktivitet: Gi elevene 5-7 minutter.\nGå rundt og lytt – noter typiske feil for oppsummering.\nEtter par-arbeid: Ta 2-3 par som deler svaret med klassen.");

  // ══════════════════════════════════════
  // SLIDE 8: OPPSUMMERING – exit ticket
  // ══════════════════════════════════════
  var s8 = pres.addSlide();
  if (bilde1) {
    s8.addImage({ data: bilde1, x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
    s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 50 }, line: { color: "000000", transparency: 50 } });
  } else {
    s8.background = { color: C.primary };
  }
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.1, fill: { color: C.amber }, line: { color: C.amber } });
  s8.addText("Hva har du lært?", { x: 0.5, y: 0.3, w: 9.0, h: 0.9, fontSize: 34, bold: true, color: C.white, fontFace: "Calibri", align: "center" });

  var exitSporsmal = [
    "Hva er regelen for " + kortTekst(tema, 25) + "?",
    "Gi ett eksempel fra teksten vi leste.",
    "Hva var vanskeligst i dag?"
  ];
  exitSporsmal.forEach(function(spm, i) {
    var y = 1.35 + i * 1.3;
    s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9.0, h: 1.05, fill: { color: C.white, transparency: 15 }, line: { color: C.white, transparency: 40, pt: 1 } });
    s8.addText((i + 1) + ".", { x: 0.65, y: y + 0.05, w: 0.5, h: 0.95, fontSize: 22, bold: true, color: C.amber, fontFace: "Calibri", valign: "middle" });
    s8.addText(spm, { x: 1.25, y: y + 0.08, w: 8.0, h: 0.9, fontSize: 17, color: C.white, fontFace: "Calibri", valign: "middle" });
  });
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.1, w: W, h: 0.1, fill: { color: C.accent }, line: { color: C.accent } });
  s8.addNotes("Oppsummering / Exit ticket: Bruk disse spørsmålene til å sjekke forståelse.\nAlternativ: Be elevene skrive ned svaret på ett spørsmål på en lapp (exit ticket).\nKnytt tilbake til læringsmålene fra slide 2 – er alle målene nådd?");

  return pres;
}

// Lag en enkel feilsetning basert på tema
function lagFeilSetning(tema) {
  var t = tema.toLowerCase();
  if (t.includes("verb") || t.includes("presens")) return "Han ikke spiser frokost om morgenen.";
  if (t.includes("preteritum")) return "I går jeg gikk til jobben tidlig.";
  if (t.includes("perfektum")) return "Jeg har gå på skolen i tre år.";
  if (t.includes("adjektiv")) return "Hun har en rødt sykkel og en blåt bil.";
  if (t.includes("substantiv")) return "Jeg har to boka og en pennen.";
  if (t.includes("v2") || t.includes("inversjon")) return "I dag jeg jobber på kontoret.";
  if (t.includes("preposisjon")) return "Jeg bor på Norge og jobber i Oslo.";
  if (t.includes("pronomen")) return "Meg og han gikk til butikken.";
  return "Idag jeg er veldig glad for å lære norsk.";
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
