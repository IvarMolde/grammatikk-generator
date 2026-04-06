// api/lag-pptx.js – dynamiske bokser, ingen tekst kuttes
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
const HEADER_H = 1.05;
const MARGIN = 0.3;
const CONTENT_TOP = HEADER_H + MARGIN;
const CONTENT_H = H - CONTENT_TOP - 0.1;

function mk() {
  return { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.08 };
}

function str(v) { return (v || "").toString().trim(); }

// Estimer teksthøyde basert på tegn, bredde og skriftstørrelse
// Returnerer høyde i tommer
function estimerTekstH(tekst, breddeInch, fontSize) {
  if (!tekst) return 0.3;
  var t = str(tekst);
  // Antall tegn per linje (empirisk: ca 1.6 tegn per pt per tomme)
  var tegnPerLinje = Math.floor(breddeInch * (fontSize * 1.55));
  if (tegnPerLinje < 1) tegnPerLinje = 1;
  var linjer = Math.ceil(t.length / tegnPerLinje);
  // Legg til linjeskift i teksten
  var ekstraLinjer = (t.match(/\n/g) || []).length;
  linjer += ekstraLinjer;
  if (linjer < 1) linjer = 1;
  // Høyde per linje i tommer (fontSize i pt, 1 tomme = 72 pt, linjeavstand 1.3)
  var linjeh = (fontSize * 1.3) / 72;
  return linjer * linjeh + 0.15; // + padding
}

// Lag en feilsetning for «finn feilen»-aktivitet
function lagFeilSetning(tema) {
  var t = str(tema).toLowerCase();
  if (t.includes("presens"))          return "Han ikke spiser frokost om morgenen.";
  if (t.includes("preteritum"))       return "I går jeg gikk til jobben tidlig.";
  if (t.includes("perfektum"))        return "Jeg har gå på skolen i tre år.";
  if (t.includes("adjektiv"))         return "Hun har en rødt bil og et blåt hus.";
  if (t.includes("flertall"))         return "Jeg har to boka og mange penner.";
  if (t.includes("v2") || t.includes("inversjon")) return "I dag jeg jobber hjemme.";
  if (t.includes("preposisjon"))      return "Jeg bor på Norge og jobber i Oslo.";
  if (t.includes("pronomen"))         return "Meg og hun gikk til butikken.";
  if (t.includes("bestemt"))          return "Jeg ser en hunden og en katten.";
  if (t.includes("modal"))            return "Hun kan ikke å snakke norsk ennå.";
  if (t.includes("passiv"))           return "Brevet ble skrive av læreren.";
  return "I dag jeg er veldig glad for å lære norsk.";
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
    if (!r1.ok) return null;
    var json = await r1.json();
    var url = json.urls && (json.urls.regular || json.urls.small);
    if (!url) return null;
    var ratio = json.width && json.height ? json.width / json.height : 1.5;
    var c2 = new AbortController();
    var t2 = setTimeout(function() { c2.abort(); }, 12000);
    var r2 = await fetch(url, { signal: c2.signal });
    clearTimeout(t2);
    if (!r2.ok) return null;
    var buf = await r2.arrayBuffer();
    return { data: "image/jpeg;base64," + Buffer.from(buf).toString("base64"), ratio: ratio };
  } catch (e) {
    console.error("Bilde feil: " + e.message);
    return null;
  }
}

// Legg til bilde med korrekt proporsjoner (contain)
function leggTilBilde(slide, bildeObj, boksX, boksY, boksW, boksH) {
  if (!bildeObj) return;
  var ratio = bildeObj.ratio || 1.5;
  var imgW, imgH;
  if (ratio > boksW / boksH) {
    imgW = boksW; imgH = boksW / ratio;
  } else {
    imgH = boksH; imgW = boksH * ratio;
  }
  var imgX = boksX + (boksW - imgW) / 2;
  var imgY = boksY + (boksH - imgH) / 2;
  slide.addImage({ data: bildeObj.data, x: imgX, y: imgY, w: imgW, h: imgH });
}

// Legg til en tekstboks med farget venstrestolpe og dynamisk høyde
// Returnerer faktisk høyde brukt
function leggTilKortMedStolpe(slide, x, y, w, farge, lysFarge, stolpeTekst, tittel, innhold, titleSize, innholdSize) {
  var stolpeW = 0.52;
  var innerW = w - stolpeW - 0.15;
  var tH = estimerTekstH(tittel, innerW, titleSize || 14);
  var iH = estimerTekstH(innhold, innerW, innholdSize || 12);
  var h = Math.max(tH + iH + 0.35, 0.9);

  slide.addShape(slide._pptx ? "rect" : "rect", { x: x, y: y, w: w, h: h,
    fill: { color: lysFarge }, line: { color: farge, pt: 1.5 }, shadow: mk() });
  slide.addShape("rect", { x: x, y: y, w: stolpeW, h: h,
    fill: { color: farge }, line: { color: farge } });
  if (stolpeTekst) {
    slide.addText(stolpeTekst, { x: x, y: y, w: stolpeW, h: h,
      fontSize: 22, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0 });
  }
  if (tittel) {
    slide.addText(tittel, { x: x + stolpeW + 0.1, y: y + 0.08, w: innerW, h: tH,
      fontSize: titleSize || 14, bold: true, color: farge, fontFace: "Calibri",
      valign: "top", wrap: true });
  }
  if (innhold) {
    slide.addText(innhold, { x: x + stolpeW + 0.1, y: y + tH + 0.15, w: innerW, h: iH,
      fontSize: innholdSize || 12, color: C.textMid, fontFace: "Calibri",
      valign: "top", wrap: true });
  }
  return h;
}

async function lagPresentasjon(data, unsplashKey) {
  var tema       = str(data.tema);
  var niva       = str(data.niva);
  var forklaring = str(data.forklaring);
  var gramFork   = str(data.grammatikkForklaring);
  var lesetekst  = str(data.lesetekst);
  var oppgaver   = Array.isArray(data.oppgaver) ? data.oppgaver : [];
  var sokeord1   = str(data.bilde1) || "language learning norway";
  var sokeord2   = str(data.bilde2) || "classroom students";

  var bilde1 = null, bilde2 = null;
  if (unsplashKey) {
    var br = await Promise.all([hentBilde(sokeord1, unsplashKey), hentBilde(sokeord2, unsplashKey)]);
    bilde1 = br[0]; bilde2 = br[1];
  }

  // Del forklaring i setninger
  var forkSetn = forklaring.split(/(?<=[.!?])\s+/).filter(function(s) { return s.trim().length > 10; });
  var kjerneregel   = str(forkSetn[0] || forklaring);
  var tilleggsregel = str(forkSetn.slice(1).join(" "));

  // Del grammatikkforklaring i punkter
  var gramLinjer = gramFork.split("\n")
    .map(function(l) { return l.replace(/^[-*•]\s*/, "").trim(); })
    .filter(function(l) { return l.length > 5; })
    .slice(0, 5);

  var pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = tema + " – Nivå " + niva;
  pres.author = "Molde voksenopplæringssenter";

  // ══════════════════════════════════════════
  // SLIDE 1: TITTEL
  // ══════════════════════════════════════════
  var s1 = pres.addSlide();
  s1.background = { color: C.primary };
  if (bilde1) {
    s1.addImage({ data: bilde1.data, x: W * 0.42, y: 0, w: W * 0.58, h: H,
      sizing: { type: "cover", w: W * 0.58, h: H } });
    s1.addShape(pres.shapes.RECTANGLE, { x: W * 0.35, y: 0, w: W * 0.65, h: H,
      fill: { color: C.primary, transparency: 20 }, line: { color: C.primary, transparency: 20 } });
  }
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W * 0.5, h: H, fill: { color: C.primary }, line: { color: C.primary } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.09, w: W, h: 0.09, fill: { color: C.accent }, line: { color: C.accent } });
  s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.85, w: 1.5, h: 0.52, fill: { color: C.amber }, line: { color: C.amber } });
  s1.addText("Nivå " + niva, { x: 0.5, y: 0.85, w: 1.5, h: 0.52, fontSize: 17, bold: true,
    color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  s1.addText(tema, { x: 0.5, y: 1.55, w: W * 0.46, h: 2.6,
    fontSize: 44, bold: true, color: C.white, fontFace: "Calibri", align: "left", valign: "middle", wrap: true });
  s1.addText("Molde voksenopplæringssenter – MBO", { x: 0.5, y: 4.6, w: W * 0.46, h: 0.4,
    fontSize: 12, color: "8AABCC", fontFace: "Calibri" });
  s1.addNotes("TITTELSLIDE: " + tema + " – Nivå " + niva + "\n\nAKTIVERING (2 min):\n– Spør klassen: «Hva vet dere om " + tema + " fra før?»\n– La 2–3 elever svare. Skriv nøkkelord på tavlen.");

  // ══════════════════════════════════════════
  // SLIDE 2: LÆRINGSMÅL – 3 fargede kort
  // ══════════════════════════════════════════
  var s2 = pres.addSlide();
  s2.background = { color: C.gray };
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addText("Hva skal vi lære i dag?", { x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var malData = [
    { farge: C.accent, tall: "1", tekst: "Forstå hva\n" + tema + " er" },
    { farge: C.green,  tall: "2", tekst: "Gjenkjenne " + tema + "\ni norske tekster" },
    { farge: C.lilla,  tall: "3", tekst: "Bruke " + tema + "\ni egne setninger" }
  ];
  malData.forEach(function(m, i) {
    var x = 0.35 + i * 3.15;
    s2.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 2.95, h: 3.95,
      fill: { color: m.farge }, line: { color: m.farge }, shadow: mk() });
    s2.addText(m.tall, { x: x, y: 1.35, w: 2.95, h: 1.5,
      fontSize: 64, bold: true, color: "FFFFFF", fontFace: "Calibri", align: "center" });
    s2.addText(m.tekst, { x: x + 0.15, y: 3.0, w: 2.65, h: 1.9,
      fontSize: 16, color: C.white, fontFace: "Calibri", align: "center", valign: "top", wrap: true });
  });
  s2.addNotes("LÆRINGSMÅL\nGå gjennom de tre målene. Be gjerne en elev lese hvert mål høyt.\nReturner til målene på slutten av timen.");

  // ══════════════════════════════════════════
  // SLIDE 3: REGELEN – dynamiske bokser
  // ══════════════════════════════════════════
  var s3 = pres.addSlide();
  s3.background = { color: C.lightGray };
  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Regelen – kort og tydelig", { x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var regelTekstW = 8.75;
  var kjerneFontSize = 18;
  var kjH = estimerTekstH(kjerneregel, regelTekstW, kjerneFontSize) + 0.3;
  kjH = Math.max(kjH, 1.0);

  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: CONTENT_TOP, w: 9.2, h: kjH,
    fill: { color: C.white }, line: { color: C.accent, pt: 3 }, shadow: mk() });
  s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: CONTENT_TOP, w: 0.18, h: kjH,
    fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText(kjerneregel, { x: 0.72, y: CONTENT_TOP + 0.1, w: regelTekstW, h: kjH - 0.2,
    fontSize: kjerneFontSize, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "middle", wrap: true });

  // Husk-boks dynamisk under regelkort
  if (tilleggsregel.length > 15) {
    var huskTop = CONTENT_TOP + kjH + 0.15;
    var huskFontSize = 15;
    var huskH = estimerTekstH(tilleggsregel, 8.0, huskFontSize) + 0.3;
    huskH = Math.max(huskH, 0.8);
    // Ikke la husk-boks gå utenfor sliden
    if (huskTop + huskH <= H - 0.05) {
      s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: huskTop, w: 9.2, h: huskH,
        fill: { color: C.lightAmber }, line: { color: C.amber, pt: 2 } });
      s3.addText("💡", { x: 0.55, y: huskTop, w: 0.7, h: huskH,
        fontSize: 26, fontFace: "Calibri", valign: "middle", align: "center" });
      s3.addText(tilleggsregel, { x: 1.3, y: huskTop + 0.1, w: 8.1, h: huskH - 0.2,
        fontSize: huskFontSize, color: "6B4F00", fontFace: "Calibri",
        align: "left", valign: "top", wrap: true });
    }
  }
  s3.addNotes("REGELEN\nVIKTIG: IKKE les fra sliden – fortell regelen med egne ord.\n1. Les regelkortet høyt\n2. Spør: «Kan noen si dette med andre ord?»\n3. Skriv regelen på tavlen med et eget eksempel\nSjekk forståelse: «Tommel opp / tommel ned»");

  // ══════════════════════════════════════════
  // SLIDE 4: EKSEMPLER – dynamiske bokser
  // ══════════════════════════════════════════
  var s4 = pres.addSlide();
  s4.background = { color: C.white };
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.green }, line: { color: C.green } });
  s4.addText("Eksempler", { x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var bildeKol = bilde2 ? 3.6 : 0;
  var eksW = W - MARGIN * 2 - (bildeKol > 0 ? bildeKol + 0.2 : 0);
  var eksFarger  = [C.lightGreen, "#F0FFF4", C.lightGreen, "#F0FFF4", C.lightGreen];
  var regelFs = 12, eksFs = 15;
  var curY = CONTENT_TOP;

  gramLinjer.forEach(function(ex, i) {
    var deler = ex.split(/\s*->\s*/);
    var regel = deler[0] ? deler[0].trim() : "";
    var eks   = deler[1] ? deler[1].trim() : "";
    var innerW = eksW - 0.2;

    var rH = eks ? estimerTekstH(regel, innerW, regelFs) : 0;
    var eH = eks ? estimerTekstH(eks, innerW, eksFs) + 0.1 : estimerTekstH(regel, innerW, eksFs);
    var boksH = Math.max(rH + eH + 0.25, 0.8);

    // Ikke tegn utenfor sliden
    if (curY + boksH > H - 0.1) return;

    s4.addShape(pres.shapes.RECTANGLE, { x: MARGIN, y: curY, w: eksW, h: boksH,
      fill: { color: eksFarger[i % 2] }, line: { color: C.green, pt: 1.5 } });

    if (eks) {
      s4.addText(regel, { x: MARGIN + 0.12, y: curY + 0.06, w: innerW, h: rH,
        fontSize: regelFs, color: C.textMid, fontFace: "Calibri", italic: true, wrap: true, valign: "top" });
      s4.addText(eks, { x: MARGIN + 0.12, y: curY + rH + 0.1, w: innerW, h: eH,
        fontSize: eksFs, bold: true, color: C.green, fontFace: "Calibri", wrap: true, valign: "top" });
    } else {
      s4.addText(regel, { x: MARGIN + 0.12, y: curY + 0.08, w: innerW, h: boksH - 0.16,
        fontSize: eksFs, color: C.textDark, fontFace: "Calibri", valign: "middle", wrap: true });
    }
    curY += boksH + 0.1;
  });

  // Bilde til høyre, contain-proporsjoner
  if (bilde2) {
    leggTilBilde(s4, bilde2, W - bildeKol - MARGIN, CONTENT_TOP, bildeKol, CONTENT_H);
  }
  s4.addNotes("EKSEMPLER\n1. Pek på regelen (kursiv) og si den høyt\n2. Les eksempelsetningen (fet) tydelig\n3. Spør: «Kan dere lage en lignende setning?\"\n4. La 2–3 elever gi muntlige eksempler");

  // ══════════════════════════════════════════
  // SLIDE 5: LESETEKST – full tekst
  // ══════════════════════════════════════════
  var s5 = pres.addSlide();
  s5.background = { color: "#FFFEF8" };
  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Les teksten", { x: 0.5, y: 0.1, w: 5.5, h: 0.85,
    fontSize: 26, bold: true, color: C.dark, fontFace: "Calibri", valign: "middle" });
  s5.addShape(pres.shapes.RECTANGLE, { x: 6.0, y: 0.12, w: 3.7, h: 0.78,
    fill: { color: "FFF3CD" }, line: { color: C.dark, pt: 1 } });
  s5.addText("🔍 Finn alle eksempler på\n" + tema, { x: 6.1, y: 0.12, w: 3.5, h: 0.78,
    fontSize: 12, bold: true, color: "6B4F00", fontFace: "Calibri", valign: "middle", wrap: true });

  // Tekst-boks fyller hele innholdsfeltet
  var leseFs = 15;
  var leseW = 8.7;
  var lTekst = lesetekst.replace(/\n+/g, " ").trim();
  var leseH = estimerTekstH(lTekst, leseW, leseFs) + 0.3;
  leseH = Math.min(leseH, CONTENT_H); // ikke gå ut av sliden

  s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: CONTENT_TOP, w: 9.2, h: leseH,
    fill: { color: C.white }, line: { color: "E0D080", pt: 2 }, shadow: mk() });
  s5.addText(lTekst, { x: 0.65, y: CONTENT_TOP + 0.12, w: leseW, h: leseH - 0.24,
    fontSize: leseFs, color: C.textDark, fontFace: "Calibri",
    align: "left", valign: "top", paraSpaceAfter: 6, wrap: true });
  s5.addNotes("LESETEKST\n1. Lærer leser teksten høyt (elevene lytter)\n2. Elevene leser stille og understreker " + tema + "\n3. Diskuter: «Hvilke eksempler fant dere?»");

  // ══════════════════════════════════════════
  // SLIDE 6: OPPGAVEOVERSIKT – dynamiske bokser
  // ══════════════════════════════════════════
  var s6 = pres.addSlide();
  s6.background = { color: C.gray };
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addText("Jobbe med oppgavene", { x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var oppgFarger  = [C.accent, C.green, C.lilla, C.red, C.amber];
  var oppgLys     = [C.light,  C.lightGreen, C.lightLilla, C.lightRed, C.lightAmber];
  var brev        = ["A","B","C","D","E"];
  var antallOppg  = Math.min(oppgaver.length, 5);

  // Dynamisk 2-kolonne layout – beregn høyder først
  var kolW = (W - MARGIN * 3) / 2; // to kolonner med margin
  var titleFs6 = 14, innholdFs6 = 12;
  var innerW6 = kolW - 0.52 - 0.25;

  var hoyder = oppgaver.slice(0, antallOppg).map(function(oppg) {
    var tH = estimerTekstH(str(oppg.type), innerW6, titleFs6);
    var iH = estimerTekstH(str(oppg.instruksjon), innerW6, innholdFs6);
    return Math.max(tH + iH + 0.45, 1.0);
  });

  // Fordel på kolonner: A+B venstre, C+D høyre, E full bredde
  var yVenstre = CONTENT_TOP, yHoyre = CONTENT_TOP;
  oppgaver.slice(0, antallOppg).forEach(function(oppg, i) {
    var h = hoyder[i];
    var isFullBredde = (antallOppg === 5 && i === 4);
    var x, y, w6;

    if (isFullBredde) {
      x = MARGIN; y = Math.max(yVenstre, yHoyre) + 0.1;
      w6 = W - MARGIN * 2;
    } else if (i % 2 === 0) {
      // Venstre kolonne: A, C (evt E)
      x = MARGIN; y = yVenstre; w6 = kolW;
    } else {
      // Høyre kolonne: B, D
      x = MARGIN * 2 + kolW; y = yHoyre; w6 = kolW;
    }

    // Ikke tegn utenfor sliden
    if (y + h > H - 0.08) return;

    var innerW = w6 - 0.52 - 0.25;

    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: w6, h: h,
      fill: { color: oppgLys[i] }, line: { color: oppgFarger[i], pt: 1.5 }, shadow: mk() });
    s6.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.52, h: h,
      fill: { color: oppgFarger[i] }, line: { color: oppgFarger[i] } });
    s6.addText(brev[i], { x: x, y: y, w: 0.52, h: h,
      fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    var tH = estimerTekstH(str(oppg.type), innerW, titleFs6);
    s6.addText(str(oppg.type), { x: x + 0.62, y: y + 0.1, w: innerW, h: tH,
      fontSize: titleFs6, bold: true, color: oppgFarger[i], fontFace: "Calibri", valign: "top", wrap: true });
    s6.addText(str(oppg.instruksjon), { x: x + 0.62, y: y + tH + 0.18, w: innerW, h: h - tH - 0.28,
      fontSize: innholdFs6, color: C.textMid, fontFace: "Calibri", valign: "top", wrap: true });

    if (isFullBredde) {
      // ingenting
    } else if (i % 2 === 0) {
      yVenstre = y + h + 0.12;
    } else {
      yHoyre = y + h + 0.12;
    }
  });
  s6.addNotes("OPPGAVER\nFØR elevene begynner: Gå gjennom instruksjonen for oppgave A høyt og gjør første eksempel på tavlen.\nUnderveis: Gå rundt, hjelp, noter typiske feil.\nEtter: Gå gjennom fasit i plenum.");

  // ══════════════════════════════════════════
  // SLIDE 7: PAR-AKTIVITET
  // ══════════════════════════════════════════
  var s7 = pres.addSlide();
  s7.background = { color: C.white };
  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: HEADER_H, fill: { color: C.lilla }, line: { color: C.lilla } });
  s7.addText("🗣  Snakk med en partner", { x: 0.5, y: 0.12, w: 9.0, h: 0.8,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", valign: "middle" });

  var parData = [
    { nr: "1", farge: C.accent, lys: C.light,      tittel: "Forklar regelen",
      tekst: "Forklar «" + tema + "» for partneren din med egne ord" },
    { nr: "2", farge: C.green,  lys: C.lightGreen,  tittel: "Lag en setning",
      tekst: "Lag én setning om deg selv med " + tema + " – del den med partneren" },
    { nr: "3", farge: C.red,    lys: C.lightRed,    tittel: "Finn feilen",
      tekst: "Hva er galt i denne setningen?\n«" + lagFeilSetning(tema) + "»\nRett den og forklar hvorfor." }
  ];

  var curY7 = CONTENT_TOP;
  parData.forEach(function(akt) {
    var sirkelR = 0.44;
    var innerW7 = W - MARGIN * 2 - sirkelR * 2 - 0.4;
    var tH7 = estimerTekstH(akt.tittel, innerW7, 15);
    var iH7 = estimerTekstH(akt.tekst, innerW7, 13);
    var h7 = Math.max(tH7 + iH7 + 0.35, sirkelR * 2 + 0.2);

    if (curY7 + h7 > H - 0.05) return;

    s7.addShape(pres.shapes.RECTANGLE, { x: MARGIN, y: curY7, w: W - MARGIN * 2, h: h7,
      fill: { color: akt.lys }, line: { color: akt.farge, pt: 1.5 }, shadow: mk() });
    s7.addShape(pres.shapes.OVAL, { x: MARGIN + 0.1, y: curY7 + h7 / 2 - sirkelR, w: sirkelR * 2, h: sirkelR * 2,
      fill: { color: akt.farge }, line: { color: akt.farge } });
    s7.addText(akt.nr, { x: MARGIN + 0.1, y: curY7 + h7 / 2 - sirkelR, w: sirkelR * 2, h: sirkelR * 2,
      fontSize: 20, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    var tekX = MARGIN + sirkelR * 2 + 0.25;
    s7.addText(akt.tittel, { x: tekX, y: curY7 + 0.1, w: innerW7, h: tH7,
      fontSize: 15, bold: true, color: akt.farge, fontFace: "Calibri", valign: "top", wrap: true });
    s7.addText(akt.tekst, { x: tekX, y: curY7 + tH7 + 0.18, w: innerW7, h: iH7,
      fontSize: 13, color: C.textDark, fontFace: "Calibri", valign: "top", wrap: true });

    curY7 += h7 + 0.12;
  });
  s7.addNotes("PAR-AKTIVITET (5–7 minutter)\nOrganisering: Snu deg mot naboen\nGå rundt og lytt – noter typiske feil\nEtter: Ta 2–3 par som deler svar på oppgave 3\nDiskuter feilen i plenum");

  // ══════════════════════════════════════════
  // SLIDE 8: EXIT TICKET
  // ══════════════════════════════════════════
  var s8 = pres.addSlide();
  s8.background = { color: C.primary };
  if (bilde1) {
    s8.addImage({ data: bilde1.data, x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
    s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H,
      fill: { color: "000000", transparency: 48 }, line: { color: "000000", transparency: 48 } });
  }
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.09, fill: { color: C.amber }, line: { color: C.amber } });
  s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.09, w: W, h: 0.09, fill: { color: C.accent }, line: { color: C.accent } });

  s8.addText("Hva har du lært?", { x: 0.5, y: 0.2, w: 9.0, h: 0.9,
    fontSize: 32, bold: true, color: C.white, fontFace: "Calibri", align: "center" });

  var exitSpm = [
    "Hva er regelen for " + tema + "?",
    "Gi ett eksempel fra leseteksten.",
    "Hva var vanskeligst i dag – og hva vil du øve mer på?"
  ];
  var curY8 = 1.2;
  exitSpm.forEach(function(spm, i) {
    var spmH = estimerTekstH(spm, 7.7, 17) + 0.25;
    spmH = Math.max(spmH, 0.95);
    if (curY8 + spmH > H - 0.15) return;
    s8.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: curY8, w: 9.0, h: spmH,
      fill: { color: C.white, transparency: 12 }, line: { color: C.white, transparency: 35, pt: 1.5 } });
    s8.addShape(pres.shapes.OVAL, { x: 0.62, y: curY8 + spmH / 2 - 0.38, w: 0.76, h: 0.76,
      fill: { color: C.amber }, line: { color: C.amber } });
    s8.addText((i + 1).toString(), { x: 0.62, y: curY8 + spmH / 2 - 0.38, w: 0.76, h: 0.76,
      fontSize: 20, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s8.addText(spm, { x: 1.5, y: curY8 + 0.08, w: 7.8, h: spmH - 0.16,
      fontSize: 17, color: C.white, fontFace: "Calibri", valign: "middle", wrap: true });
    curY8 += spmH + 0.12;
  });
  s8.addNotes("OPPSUMMERING / EXIT TICKET\nAlternativ 1 – Muntlig: Svar på spm 1 og 2 med partner\nAlternativ 2 – Skriftlig: Skriv svaret på spm 1 på en lapp (samle inn)\nKnytt tilbake til læringsmålene fra slide 2");

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
