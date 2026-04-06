// api/lag-pptx.js – pedagogisk redesign
const pptxgen = require("pptxgenjs");

const C = {
  primary:    "1F4E79",
  accent:     "2E75B6",
  light:      "D6E4F0",
  white:      "FFFFFF",
  dark:       "1A2744",
  green:      "1E6B3C",
  lightGreen: "E8F5ED",
  amber:      "E9A800",
  lightAmber: "FFF8E1",
  lilla:      "6B4EA0",
  lightLilla: "F0EBF8",
  red:        "C0392B",
  lightRed:   "FDECEA",
  textDark:   "1A1A2E",
  textMid:    "4A4A6A",
  gray:       "F4F6F9",
  lightGray:  "FAFBFC"
};

const W = 10, H = 5.625;

function mk() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

function k(tekst, maks) {
  if (!tekst) return "";
  tekst = tekst.toString();
  return tekst.length > maks ? tekst.slice(0, maks - 1) + "…" : tekst;
}

// Hent bilde som base64 fra Unsplash
async function hentBilde(sokeord, apiKey) {
  try {
    var apiUrl = "https://api.unsplash.com/photos/random?query="
      + encodeURIComponent(sokeord)
      + "&orientation=landscape&client_id=" + apiKey;
    var c1 = new AbortController();
    var t1 = setTimeout(function() { c1.abort(); }, 7000);
    var r1 = await fetch(apiUrl, { signal: c1.signal });
    clearTimeout(t1);
    if (!r1.ok) { console.error("Unsplash " + r1.status); return null; }
    var json = await r1.json();
    var url = json.urls && (json.urls.regular || json.urls.small);
    if (!url) return null;
    var c2 = new AbortController();
    var t2 = setTimeout(function() { c2.abort(); }, 12000);
    var r2 = await fetch(url, { signal: c2.signal });
    clearTimeout(t2);
    if (!r2.ok) return null;
    var buf = await r2.arrayBuffer();
    // Hent bildebredde og høyde fra Unsplash-data for korrekt ratio
    var ratio = json.width && json.height ? json.width / json.height : 1.5;
    console.log("Bilde OK: " + sokeord + " ratio=" + ratio.toFixed(2));
    return { data: "image/jpeg;base64," + Buffer.from(buf).toString("base64"), ratio: ratio };
  } catch (e) {
    console.error("Bilde feil: " + e.message);
    return null;
  }
}

// Legg til bilde med korrekt proporsjoner innenfor en gitt boks
function leggTilBilde(slide, bildeObj, boksX, boksY, boksW, boksH) {
  if (!bildeObj) return;
  var ratio = bildeObj.ratio || 1.5;
  var imgW, imgH;
  // Beregn dimensjoner som passer innenfor boksen (contain, ikke crop)
  if (ratio > boksW / boksH) {
    imgW = boksW;
    imgH = boksW / ratio;
  } else {
    imgH = boksH;
    imgW = boksH * ratio;
  }
  var imgX = boksX + (boksW - imgW) / 2;
  var imgY = boksY + (boksH - imgH) / 2;
  slide.addImage({ data: bildeObj.data, x: imgX, y: imgY, w: imgW, h: imgH });
}

// Lag en feilsetning for «finn feilen»-aktivitet
function lagFeilSetning(tema) {
  var t = (tema || "").toLowerCase();
  if (t.includes("presens"))          return "Han ikke spiser frokost.";
  if (t.includes("preteritum"))       return "I går jeg gikk til jobben.";
  if (t.includes("perfektum"))        return "Jeg har gå på skolen i tre år.";
  if (t.includes("adjektiv"))         return "Hun har en rødt bil og et blåt hus.";
  if (t.includes("flertall"))         return "Jeg har to boka og mange penner.";
  if (t.includes("v2") || t.includes("inversjon")) return "I dag jeg jobber hjemme.";
  if (t.includes("preposisjon"))      return "Jeg bor på Norge og jobber i Oslo.";
  if (t.includes("pronomen"))         return "Meg og hun gikk til butikken.";
  if (t.includes("bestemt"))          return "Jeg liker hunden min, men bilen er gammel.";
  if (t.includes("modal"))            return "Hun kan ikke å snakke norsk ennå.";
  return "I dag jeg er veldig glad for å lære norsk.";
}

async function lagPresentasjon(data, unsplashKey) {
  var tema      = (data.tema || "").toString();
  var niva      = (data.niva || "").toString();
  var forklaring = (data.forklaring || "").toString();
  var gramForklaring = (data.grammatikkForklaring || "").toString();
  var lesetekst = (data.lesetekst || "").toString();
  var oppgaver  = Array.isArray(data.oppgaver) ? data.oppgaver : [];
  var sokeord1  = (data.bilde1 || "language learning norway").toString();
  var sokeord2  = (data.bilde2 || "classroom students").toString();

  // Hent bilder
  var bilde1 = null, bilde2 = null;
  if (unsplashKey) {
    var bildeRes = await Promise.all([
      hentBilde(sokeord1, unsplashKey),
      hentBilde(sokeord2, unsplashKey)
    ]);
    bilde1 = bildeRes[0];
    bilde2 = bildeRes[1];
  }

  // Del grammatikkforklaring i enkeltpunkter
  var gramLinjer = gramForklaring
    .split("\n")
    .map(function(l) { return l.replace(/^[-*•]\s*/, "").trim(); })
    .filter(function(l) { return l.length > 5; })
    .slice(0, 4);

  // Kjerneregel: første setning av forklaring
  var forklaringSetninger = forklaring
    .split(/(?<=[.!?])\s+/)
    .filter(function(s) { return s.trim().length > 10; });
  var kjerneregel = (forklaringSetninger[0] || forklaring).toString();
  var tilleggsregel = (forklaringSetninger[1] || "").toString();

  var pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = tema + " – Nivå " + niva;
  pres.author = "Molde voksenopplæringssenter";

  // ══════════════════════════════════════════
  // SLIDE 1: TITTEL
  // Full bakgrunn, stort tema-ord, nivå-badge
  // ══════════════════════════════════════════
  var s1 = pres.addSlide();
  s1.background = { color: C.primary };

  if (bilde1) {
    // Bilde fyller høyre halvdel – ikke strukket
    var b1 = bilde1;
    var ratio1 = b1.ratio || 1.5;
    var bildeH = H;
    var bildeW = bildeH * ratio1;
    if (bildeW < W * 0.5) bildeW = W * 0.5;
    slide1BildeX = W - bildeW;
    // Bilde bak overlay
    s1.addImage({ data: b1.data, x: W * 0.42, y: 0, w: W * 0.58, h: H, sizing: { type: "cover", w: W * 0.58, h: H } });
    // Gradient-overlay over bildet
    s1.addShape(pres.shapes.RECTANGLE, { x: W * 0.35, y: 0, w: W * 0.65, h: H, fill: { color: C.primary, transparency: 20 }, line: { color: C.primary, transparency: 20 } });
  }

  // Mørk venstre panel
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W * 0.48, h: H, fill: { color: C.primary }, line: { color: C.primary } });
  // Gul toppstripe
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: C.amber }, line: { color: C.amber } });
  // Blå bunnstripe
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.09, w: W, h: 0.09, fill: { color: C.accent }, line: { color: C.accent } });

  // Nivå-badge
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.85, w: 1.5, h: 0.52, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText("Nivå " + niva, { x: 0.5, y: 0.85, w: 1.5, h: 0.52, fontSize: 17, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

  // Tema – stort og tydelig
  s1.addText(tema, {
    x: 0.5, y: 1.55, w: W * 0.46, h: 2.6,
    fontSize: 46, bold: true, color: C.white, fontFace: "Calibri",
    align: "left", valign: "middle"
  });

  s1.addText("Molde voksenopplæringssenter – MBO", {
    x: 0.5, y: 4.6, w: W * 0.46, h: 0.4,
    fontSize: 12, color: "8AABCC", fontFace: "Calibri"
  });

  s1.addNotes("TITTELSLIDE: " + tema + " – Nivå " + niva + "\n\nAKTIVERING (2 min):\n– Spør klassen: «Hva vet dere om " + tema + " fra før?»\n– La 2–3 elever svare kort.\n– Skriv 1–2 nøkkelord fra svarene på tavlen.");

  // ══════════════════════════════════════════
  // SLIDE 2: LÆRINGSMÅL
  // 3 mål som store fargede kort – ikke punktliste
  // ══════════════════════════════════════════
  var s2 = pres.addSlide();
  s2.background = { color: C.gray };

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addText("Hva skal vi lære i dag?", {
    x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
  });

  var malData = [
    { farge: C.accent,  lys: "#EBF3FB", tall: "1", tekst: "Forstå hva\n" + k(tema, 28) + " er" },
    { farge: C.green,   lys: "#E8F5ED", tall: "2", tekst: "Gjenkjenne " + k(tema, 22) + "\ni norske tekster" },
    { farge: C.lilla,   lys: "#F0EBF8", tall: "3", tekst: "Bruke " + k(tema, 25) + "\ni egne setninger" }
  ];

  malData.forEach(function(m, i) {
    var x = 0.35 + i * 3.15;
    // Kort
    s2.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 2.95, h: 3.95, fill: { color: m.farge }, line: { color: m.farge }, shadow: mk() });
    // Stort nummer
    s2.addText(m.tall, {
      x: x, y: 1.35, w: 2.95, h: 1.5,
      fontSize: 64, bold: true, color: "FFFFFF", fontFace: "Calibri",
      align: "center", valign: "middle", transparency: 20
    });
    // Tekst
    s2.addText(m.tekst, {
      x: x + 0.15, y: 3.0, w: 2.65, h: 1.9,
      fontSize: 16, bold: false, color: C.white, fontFace: "Calibri",
      align: "center", valign: "top"
    });
  });

  s2.addNotes("LÆRINGSMÅL\n\nSi: «I dag skal vi jobbe med tre ting.»\nLes hvert mål høyt – be gjerne en elev lese.\n\nTips: Heng opp målene synlig og returner til dem på slutten av timen.");

  // ══════════════════════════════════════════
  // SLIDE 3: REGELEN
  // Én stor regelkort + ett eksempel – IKKE en vegg av tekst
  // ══════════════════════════════════════════
  var s3 = pres.addSlide();
  s3.background = { color: C.lightGray };

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Regelen – kort og tydelig", {
    x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
  });

  // Stor regelkort
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 1.85, fill: { color: C.white }, line: { color: C.accent, pt: 3 }, shadow: mk() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 0.18, h: 1.85, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText(kjerneregel, {
    x: 0.72, y: 1.22, w: 8.75, h: 1.7,
    fontSize: 18, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "middle", wrap: true
  });

  // Tilleggsregel / husk-boks
  if (tilleggsregel.length > 10) {
    s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.1, w: 9.2, h: 1.9, fill: { color: C.lightAmber }, line: { color: C.amber, pt: 2 } });
    s3.addText("💡", { x: 0.55, y: 3.15, w: 0.6, h: 1.8, fontSize: 28, fontFace: "Calibri", valign: "middle", align: "center" });
    s3.addText(tilleggsregel, {
      x: 1.2, y: 3.18, w: 8.2, h: 1.72,
      fontSize: 15, color: "6B4F00", fontFace: "Calibri",
      align: "left", valign: "top", wrap: true
    });
  }

  s3.addNotes("REGELEN\n\nVIKTIG: IKKE les fra sliden. Fortell regelen med egne ord.\n\nFremgangsmåte:\n1. Les regelkortet høyt\n2. Spør: «Kan noen si dette med andre ord?»\n3. Skriv regelen på tavlen – gjerne med et eget eksempel\n\nSjekk forståelse: «Tommel opp hvis du forstår, tommel ned hvis du trenger mer forklaring»");

  // ══════════════════════════════════════════
  // SLIDE 4: EKSEMPLER
  // Maks 4 eksempler – store og tydelige bokser
  // Bilde plassert korrekt (contain, ikke cover)
  // ══════════════════════════════════════════
  var s4 = pres.addSlide();
  s4.background = { color: C.white };

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.green }, line: { color: C.green } });
  s4.addText("Eksempler", {
    x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
  });

  var eksW = bilde2 ? 5.6 : 9.4;
  var eksFarger = [C.lightGreen, "#F0FFF4", C.lightGreen, "#F0FFF4"];
  var eksBorderFarger = [C.green, C.green, C.green, C.green];

  gramLinjer.forEach(function(ex, i) {
    var y = 1.1 + i * 1.18;
    // Del på pil hvis den finnes
    var deler = ex.split(/\s*->\s*/);
    var regel = deler[0] ? deler[0].trim() : "";
    var eks   = deler[1] ? deler[1].trim() : "";

    s4.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: y, w: eksW, h: 1.1,
      fill: { color: eksFarger[i] || C.lightGreen },
      line: { color: eksBorderFarger[i] || C.green, pt: 1.5 }
    });

    if (eks) {
      // Regel liten øverst, eksempel stort nederst
      s4.addText(regel, { x: 0.5, y: y + 0.04, w: eksW - 0.3, h: 0.42, fontSize: 12, color: C.textMid, fontFace: "Calibri", italic: true, wrap: true });
      s4.addText(eks,   { x: 0.5, y: y + 0.5,  w: eksW - 0.3, h: 0.52, fontSize: 15, bold: true, color: C.green, fontFace: "Calibri", wrap: true });
    } else {
      s4.addText(regel, { x: 0.5, y: y + 0.1, w: eksW - 0.3, h: 0.9, fontSize: 15, color: C.textDark, fontFace: "Calibri", valign: "middle", wrap: true });
    }
  });

  // Bilde – korrekt proporsjoner med contain
  if (bilde2) {
    var boksX = 6.2, boksY = 1.1, boksW = 3.6, boksH = 4.3;
    leggTilBilde(s4, bilde2, boksX, boksY, boksW, boksH);
  }

  s4.addNotes("EKSEMPLER\n\nFremgangsmåte:\n1. Pek på regel (kursiv) og si den høyt\n2. Les eksempelsetningen (fet) tydelig\n3. Spør: «Kan dere lage en lignende setning?\"\n4. La 2–3 elever gi muntlige eksempler\n\nTavleaktivitet: Skriv et nytt eleveksempel for hvert punkt");

  // ══════════════════════════════════════════
  // SLIDE 5: LESETEKST
  // Tekst i lesbar størrelse + aktiv oppgave synlig
  // ══════════════════════════════════════════
  var s5 = pres.addSlide();
  s5.background = { color: "#FFFEF8" };

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Les teksten", {
    x: 0.5, y: 0.1, w: 5.5, h: 0.85,
    fontSize: 26, bold: true, color: C.dark, fontFace: "Calibri", valign: "middle"
  });
  // Leseoppgave synlig i header
  s5.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 0.12, w: 3.7, h: 0.78, fill: { color: "FFF3CD" }, line: { color: C.dark, pt: 1 } });
  s5.addText("🔍 Finn alle " + k(tema, 20) + " i teksten", {
    x: 6.1, y: 0.12, w: 3.5, h: 0.78,
    fontSize: 13, bold: true, color: "6B4F00", fontFace: "Calibri", valign: "middle"
  });

  // Tekst i stor, lesbar font – IKKE for mye
  var kortLesetekst = lesetekst.replace(/\n+/g, " ").toString();
  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.15, w: 9.2, h: 4.1, fill: { color: C.white }, line: { color: "E0D080", pt: 2 }, shadow: mk() });
  s5.addText(kortLesetekst, {
    x: 0.65, y: 1.3, w: 8.7, h: 3.75,
    fontSize: 16, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "top", paraSpaceAfter: 8
  });

  s5.addNotes("LESETEKST\n\nFremgangsmåte:\n1. Lærer leser teksten høyt første gang (elevene lytter)\n2. Elevene leser stille og understreker eksempler på " + tema + "\n3. Diskuter: «Hvilke eksempler fant dere?»\n\nAlternativ: Høytlesning der elevene leser annenhver setning");

  // ══════════════════════════════════════════
  // SLIDE 6: OPPGAVEOVERSIKT
  // Fargede kort, ikke punktliste
  // ══════════════════════════════════════════
  var s6 = pres.addSlide();
  s6.background = { color: C.gray };

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addText("Jobbe med oppgavene", {
    x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
  });

  var oppgFarger = [C.accent, C.green, C.lilla, C.red, C.amber];
  var oppgLys    = [C.light,  C.lightGreen, C.lightLilla, C.lightRed, C.lightAmber];
  var brev       = ["A","B","C","D","E"];

  // Full bredde, 2-kolonne layout: A+B venstre, C+D høyre, E full bredde
  // Hver boks er høy nok til at instruksjonsteksten alltid får plass
  var oppgLayout = [
    { x: 0.3,  y: 1.15, w: 4.65, h: 1.9 },   // A
    { x: 0.3,  y: 3.15, w: 4.65, h: 1.9 },   // B – under A
    { x: 5.15, y: 1.15, w: 4.65, h: 1.9 },   // C
    { x: 5.15, y: 3.15, w: 4.65, h: 1.9 },   // D – under C
    { x: 0.3,  y: 1.15, w: 9.5,  h: 1.9 }    // E – brukes ikke som 5. kort, se under
  ];

  // Bruk 2x2 grid + eventuell femte full-bredde
  var antall = Math.min(oppgaver.length, 5);
  var layout4 = [
    { x: 0.3,  y: 1.15, w: 4.65 },
    { x: 0.3,  y: 3.18, w: 4.65 },
    { x: 5.15, y: 1.15, w: 4.65 },
    { x: 5.15, y: 3.18, w: 4.65 },
    { x: 0.3,  y: 1.15, w: 9.5  }
  ];
  // Beregn høyde basert på antall oppgaver
  var radH = antall <= 4 ? 2.05 : 1.9;

  oppgaver.slice(0, 4).forEach(function(oppg, i) {
    var pos = layout4[i];
    var x = pos.x, y = pos.y, w = pos.w, h = radH;

    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w, h: h, fill: { color: oppgLys[i] }, line: { color: oppgFarger[i], pt: 1.5 }, shadow: mk() });
    // Farget venstrestolpe med bokstav
    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.52, h: h, fill: { color: oppgFarger[i] }, line: { color: oppgFarger[i] } });
    s6.addText(brev[i], { x: x, y: y, w: 0.52, h: h, fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    // Oppgavetype – ingen kapping
    s6.addText((oppg.type || "").toString(), {
      x: x + 0.62, y: y + 0.1, w: w - 0.72, h: 0.48,
      fontSize: 14, bold: true, color: oppgFarger[i], fontFace: "Calibri", valign: "middle"
    });
    // Instruksjon – full tekst, ingen kapping, wrap automatisk
    s6.addText((oppg.instruksjon || "").toString(), {
      x: x + 0.62, y: y + 0.62, w: w - 0.72, h: h - 0.72,
      fontSize: 12, color: C.textMid, fontFace: "Calibri", valign: "top",
      wrap: true, paraSpaceAfter: 2
    });
  });

  // Hvis 5 oppgaver: legg E som egen rad under
  if (oppgaver.length >= 5) {
    var oppg5 = oppgaver[4];
    var y5 = 1.15 + radH * 2 + 0.08;
    // Sjekk om det er plass – hvis ikke, reduser høyde
    var h5 = Math.min(H - y5 - 0.1, 1.1);
    if (h5 > 0.5) {
      s6.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y5, w: 9.5, h: h5, fill: { color: oppgLys[4] }, line: { color: oppgFarger[4], pt: 1.5 } });
      s6.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y5, w: 0.52, h: h5, fill: { color: oppgFarger[4] }, line: { color: oppgFarger[4] } });
      s6.addText("E", { x: 0.3, y: y5, w: 0.52, h: h5, fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
      s6.addText((oppg5.type || "").toString(), {
        x: 0.92, y: y5 + 0.08, w: 4.0, h: 0.44,
        fontSize: 14, bold: true, color: oppgFarger[4], fontFace: "Calibri", valign: "middle"
      });
      s6.addText((oppg5.instruksjon || "").toString(), {
        x: 0.92, y: y5 + 0.55, w: 8.7, h: h5 - 0.6,
        fontSize: 12, color: C.textMid, fontFace: "Calibri", valign: "top", wrap: true
      });
    }
  }

  s6.addNotes("OPPGAVER\n\nFØR elevene begynner:\n– Gå gjennom instruksjonen for oppgave A høyt\n– Gjør det første eksemplet på tavlen\n– Sjekk at alle forstår\n\nUnderveis: Gå rundt, hjelp, noter typiske feil\nEtter: Gå gjennom fasit i plenum");

  // ══════════════════════════════════════════
  // SLIDE 7: PAR-AKTIVITET
  // 3 konkrete handlinger – aktiviserende
  // ══════════════════════════════════════════
  var s7 = pres.addSlide();
  s7.background = { color: C.white };

  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.05, fill: { color: C.lilla }, line: { color: C.lilla } });
  s7.addText("🗣  Snakk med en partner", {
    x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle"
  });

  var parAktiviteter = [
    { nr: "1", farge: C.accent, lys: C.light,      tittel: "Forklar regelen",   tekst: "Forklar «" + k(tema, 30) + "» for partneren din – bruk egne ord" },
    { nr: "2", farge: C.green,  lys: C.lightGreen,  tittel: "Lag en setning",    tekst: "Lag én setning om deg selv med " + k(tema, 25) + " – del den med partneren" },
    { nr: "3", farge: C.red,    lys: C.lightRed,    tittel: "Finn feilen",       tekst: "Hva er galt i denne setningen? Rett den: «" + lagFeilSetning(tema) + "»" }
  ];

  parAktiviteter.forEach(function(akt, i) {
    var y = 1.15 + i * 1.42;
    s7.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: y, w: 9.4, h: 1.25, fill: { color: akt.lys }, line: { color: akt.farge, pt: 1.5 }, shadow: mk() });
    // Nummerert sirkel
    s7.addShape(pres.shapes.OVAL, { x: 0.38, y: y + 0.18, w: 0.88, h: 0.88, fill: { color: akt.farge }, line: { color: akt.farge } });
    s7.addText(akt.nr, { x: 0.38, y: y + 0.18, w: 0.88, h: 0.88, fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    // Tittel
    s7.addText(akt.tittel, { x: 1.4, y: y + 0.08, w: 8.1, h: 0.44, fontSize: 15, bold: true, color: akt.farge, fontFace: "Calibri", valign: "middle" });
    // Instruksjon
    s7.addText(akt.tekst, { x: 1.4, y: y + 0.55, w: 8.1, h: 0.58, fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "middle" });
  });

  s7.addNotes("PAR-AKTIVITET (5–7 minutter)\n\nOrganisering: Snu deg mot naboen din\n\nGå rundt og lytt – noter typiske feil på et ark\n\nEtter par-arbeid:\n– Ta 2–3 par som deler svar på oppgave 3\n– Diskuter feilen i plenum\n– Bekreft/korriger riktig svar");

  // ══════════════════════════════════════════
  // SLIDE 8: EXIT TICKET / OPPSUMMERING
  // 3 enkle spørsmål – sjekk forståelse
  // ══════════════════════════════════════════
  var s8 = pres.addSlide();
  s8.background = { color: C.primary };

  if (bilde1) {
    // Bilde som bakgrunn med mørk overlay
    s8.addImage({ data: bilde1.data, x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
    s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 48 }, line: { color: "000000", transparency: 48 } });
  }

  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: C.amber }, line: { color: C.amber } });

  s8.addText("Hva har du lært?", {
    x: 0.5, y: 0.25, w: 9.0, h: 0.9,
    fontSize: 32, bold: true, color: C.white, fontFace: "Calibri", align: "center"
  });

  var exitSpm = [
    "Hva er regelen for " + k(tema, 28) + "?",
    "Gi ett eksempel fra leseteksten.",
    "Hva var vanskeligst i dag – og hva vil du øve mer på?"
  ];

  exitSpm.forEach(function(spm, i) {
    var y = 1.3 + i * 1.3;
    s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9.0, h: 1.1, fill: { color: C.white, transparency: 12 }, line: { color: C.white, transparency: 35, pt: 1.5 } });
    s8.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.12, w: 0.85, h: 0.85, fill: { color: C.amber }, line: { color: C.amber } });
    s8.addText((i + 1).toString(), { x: 0.62, y: y + 0.12, w: 0.85, h: 0.85, fontSize: 20, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s8.addText(spm, { x: 1.6, y: y + 0.1, w: 7.7, h: 0.9, fontSize: 17, color: C.white, fontFace: "Calibri", valign: "middle" });
  });

  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.09, w: W, h: 0.09, fill: { color: C.accent }, line: { color: C.accent } });

  s8.addNotes("OPPSUMMERING / EXIT TICKET\n\nAlternativ 1 – Muntlig:\nLa elevene svare på spørsmål 1 og 2 med partneren\n\nAlternativ 2 – Skriftlig exit ticket:\nBe elevene skrive svaret på spørsmål 1 på en lapp\nSaml inn lappene – bruk dem til å planlegge neste time\n\nKnytt tilbake til læringsmålene fra slide 2:\n– Mål 1: Forstår dere regelen? (spm 1)\n– Mål 2: Kan dere gjenkjenne den? (spm 2)\n– Mål 3: Kan dere bruke den? (par-aktiviteten)");

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
    var pres = await lagPresentasjon(body.data, unsplashKey);
    var filnavn = (body.data.tema || "grammatikk").replace(/[^a-zA-Z0-9]/g, "_")
      + "_" + (body.data.niva || "A1") + "_presentasjon.pptx";
    var buffer = await pres.write({ outputType: "nodebuffer" });
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", "attachment; filename=\"" + filnavn + "\"");
    res.send(buffer);
  } catch (e) {
    console.error("lag-pptx feil:", e);
    res.status(500).json({ feil: "Feil ved generering: " + e.message });
  }
};
