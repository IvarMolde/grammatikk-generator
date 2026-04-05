const pptxgen = require("pptxgenjs");
const fs = require('fs');

const data = JSON.parse(process.argv[2]);
const outputPath = process.argv[3];

const { tema, niva, tittel, forklaring, lesetekst, oppgaver, grammatikkForklaring } = data;

const C = {
  primary: "1F4E79",
  accent: "2E75B6",
  light: "D6E4F0",
  white: "FFFFFF",
  dark: "1A2744",
  green: "2C6E49",
  lightGreen: "D8F3DC",
  yellow: "FFF3CD",
  gray: "F5F5F5",
  textDark: "1A1A2E",
  textMid: "444466",
  sage: "4A7C59",
  amber: "E9A800",
};

const LAYOUT = {
  W: 10,
  H: 5.625,
};

function makeShadow() {
  return { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.12 };
}

async function lagPresentasjon() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${tema} – Nivå ${niva}`;
  pres.author = 'Molde voksenopplæringssenter';

  // ─── SLIDE 1: Tittelslide ───
  let s1 = pres.addSlide();
  s1.background = { color: C.primary };

  // Dekorativt rektangel øverst
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber } });

  // Stor tittel
  s1.addText(tema, {
    x: 0.7, y: 1.2, w: 8.6, h: 1.5,
    fontSize: 44, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle"
  });

  // Nivå-badge
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.0, y: 3.0, w: 2.0, h: 0.6, fill: { color: C.amber }, line: { color: C.amber }, rectRadius: 0.3 });
  s1.addText(`Nivå ${niva}`, { x: 4.0, y: 3.0, w: 2.0, h: 0.6, fontSize: 20, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

  // Undertittel
  s1.addText("Molde voksenopplæringssenter – MBO", {
    x: 0.7, y: 3.8, w: 8.6, h: 0.5,
    fontSize: 14, color: "A0B8D8", fontFace: "Calibri", align: "center"
  });

  // Dekorativ bunn
  s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: LAYOUT.H - 0.08, w: LAYOUT.W, h: 0.08, fill: { color: C.accent }, line: { color: C.accent } });

  s1.addNotes(`Tittelslide: ${tema} – Nivå ${niva}. Presenter temaet og nivå for klassen. Spør elevene hva de vet om dette grammatikktemaet fra før.`);

  // ─── SLIDE 2: Læringsmål ───
  let s2 = pres.addSlide();
  s2.background = { color: C.white };

  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: LAYOUT.H, fill: { color: C.primary }, line: { color: C.primary } });
  s2.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });

  s2.addText("Læringsmål", { x: 0.3, y: 0.2, w: 9.4, h: 0.7, fontSize: 28, bold: true, color: C.primary, fontFace: "Calibri" });

  // Mål-boks
  s2.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 9.3, h: 3.6,
    fill: { color: C.light }, line: { color: C.accent, pt: 1 },
    shadow: makeShadow()
  });

  s2.addText([
    { text: `Etter denne leksjonen kan jeg:`, options: { bold: true, breakLine: true, fontSize: 17 } },
    { text: " ", options: { breakLine: true } },
    { text: `• Forklare hva ${tema} er på norsk`, options: { bullet: false, breakLine: true, fontSize: 15 } },
    { text: `• Gjenkjenne ${tema} i tekster`, options: { bullet: false, breakLine: true, fontSize: 15 } },
    { text: `• Lage egne setninger med riktig bruk av ${tema}`, options: { bullet: false, breakLine: true, fontSize: 15 } },
    { text: `• Snakke om temaet med medelever`, options: { bullet: false, fontSize: 15 } },
  ], { x: 0.6, y: 1.3, w: 8.8, h: 3.0, color: C.textDark, fontFace: "Calibri", valign: "top" });

  s2.addNotes(`Gå gjennom læringsmålene med klassen. Spør elevene om de har sett eksempler på ${tema} i norsk før. Gjenta gjerne målene på slutten av timen.`);

  // ─── SLIDE 3: Grammatikk forklart ───
  let s3 = pres.addSlide();
  s3.background = { color: "F8FBFF" };

  s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 1.0, fill: { color: C.accent }, line: { color: C.accent } });
  s3.addText("Grammatikk – forklaring", { x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.white, fontFace: "Calibri" });

  // Forklaring-boks
  s3.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.15, w: 9.4, h: 3.8,
    fill: { color: C.white }, line: { color: C.light, pt: 1 },
    shadow: makeShadow()
  });

  const forklaringLinjer = forklaring.split('\n').filter(l => l.trim()).slice(0, 8);
  const forklaringRuns = forklaringLinjer.map((linje, i) => ({
    text: linje,
    options: { breakLine: i < forklaringLinjer.length - 1, fontSize: 14 }
  }));

  s3.addText(forklaringRuns, {
    x: 0.55, y: 1.3, w: 9.0, h: 3.5,
    color: C.textDark, fontFace: "Calibri", valign: "top"
  });

  s3.addNotes(`Forklaring på ${tema}. Les gjennom forklaringen høyt med klassen. Pause ved hvert punkt og sjekk forståelse. Be elevene komme med egne eksempler.`);

  // ─── SLIDE 4: Eksempler ───
  let s4 = pres.addSlide();
  s4.background = { color: C.white };

  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 0.08, fill: { color: C.green }, line: { color: C.green } });
  s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: LAYOUT.W, h: 0.9, fill: { color: C.lightGreen }, line: { color: C.lightGreen } });

  s4.addText("Eksempler", { x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.green, fontFace: "Calibri" });

  // Hent eksempel-setninger fra grammatikkForklaring
  const eksempelLinjer = (grammatikkForklaring || forklaring).split('\n')
    .filter(l => l.trim() && (l.includes('→') || l.includes(':') || l.match(/^[-•*]/)))
    .slice(0, 5);

  if (eksempelLinjer.length > 0) {
    eksempelLinjer.forEach((ex, i) => {
      const yPos = 1.15 + i * 0.82;
      s4.addShape(pres.shapes.RECTANGLE, {
        x: 0.3, y: yPos, w: 9.4, h: 0.68,
        fill: { color: i % 2 === 0 ? C.lightGreen : "F0FFF4" },
        line: { color: C.green, pt: 0.5 }
      });
      s4.addText(ex.replace(/^[-•*]\s*/, ''), {
        x: 0.5, y: yPos + 0.05, w: 9.0, h: 0.58,
        fontSize: 14, color: C.textDark, fontFace: "Calibri", valign: "middle"
      });
    });
  } else {
    s4.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.1, w: 9.4, h: 3.8, fill: { color: C.lightGreen }, line: { color: C.green, pt: 1 } });
    s4.addText("Se eksempler i arbeidsarket", { x: 0.5, y: 1.3, w: 9.0, h: 3.5, fontSize: 15, color: C.green, fontFace: "Calibri", align: "center", valign: "middle" });
  }

  s4.addNotes(`Gå gjennom eksemplene en etter en. Bruk gjerne tavlen til å skrive flere eksempler. La elevene lage egne lignende setninger.`);

  // ─── SLIDE 5: Lesetekst ───
  let s5 = pres.addSlide();
  s5.background = { color: "FFFEF5" };

  s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 1.0, fill: { color: C.amber }, line: { color: C.amber } });
  s5.addText("Lesetekst", { x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.dark, fontFace: "Calibri" });

  s5.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.1, w: 9.4, h: 4.0,
    fill: { color: C.white }, line: { color: "DDCC88", pt: 1.5 },
    shadow: makeShadow()
  });

  const leseLinjer = lesetekst.split('\n').filter(l => l.trim()).slice(0, 10);
  const leseRuns = leseLinjer.map((linje, i) => ({
    text: linje,
    options: { breakLine: i < leseLinjer.length - 1, fontSize: 13 }
  }));

  s5.addText(leseRuns, {
    x: 0.55, y: 1.25, w: 9.0, h: 3.7,
    color: C.textDark, fontFace: "Calibri", valign: "top"
  });

  s5.addNotes(`Les teksten høyt for klassen. Stopp ved setninger som illustrerer ${tema}. Spør: «Hva slags ord er dette? Hvorfor brukes denne formen?»`);

  // ─── SLIDE 6: Oppgave-oversikt ───
  let s6 = pres.addSlide();
  s6.background = { color: C.white };

  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.08, w: LAYOUT.W, h: 0.9, fill: { color: C.light }, line: { color: C.light } });
  s6.addText("Oppgaver – oversikt", { x: 0.4, y: 0.15, w: 9.2, h: 0.7, fontSize: 26, bold: true, color: C.primary, fontFace: "Calibri" });

  const bokstaver = ['A', 'B', 'C', 'D', 'E'];
  oppgaver.slice(0, 5).forEach((oppg, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const x = col === 0 ? 0.3 : 5.3;
    const y = 1.15 + row * 1.2;
    const w = 4.7;

    s6.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 1.0,
      fill: { color: i % 2 === 0 ? C.light : "EEF5FF" },
      line: { color: C.accent, pt: 0.5 },
      shadow: makeShadow()
    });

    s6.addText(`${bokstaver[i]})  ${oppg.type}`, {
      x: x + 0.12, y: y + 0.05, w: w - 0.2, h: 0.4,
      fontSize: 13, bold: true, color: C.primary, fontFace: "Calibri"
    });
    s6.addText(oppg.instruksjon.slice(0, 70) + (oppg.instruksjon.length > 70 ? '…' : ''), {
      x: x + 0.12, y: y + 0.46, w: w - 0.2, h: 0.46,
      fontSize: 11, color: C.textMid, fontFace: "Calibri"
    });
  });

  s6.addNotes(`Presenter oppgaveoversikten. Forklar at elevene skal jobbe med oppgavene i arbeidsarket. Gå gjennom instruksjonene for hver oppgave type.`);

  // ─── SLIDE 7: Diskusjon og refleksjon ───
  let s7 = pres.addSlide();
  s7.background = { color: C.primary };

  s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: LAYOUT.W, h: 0.08, fill: { color: C.amber }, line: { color: C.amber } });

  s7.addText("Diskusjon og refleksjon", {
    x: 0.5, y: 0.5, w: 9.0, h: 0.9,
    fontSize: 30, bold: true, color: C.white, fontFace: "Calibri", align: "center"
  });

  const diskSpørsmål = [
    `Kan du forklare ${tema} til en klassekamerat med egne ord?`,
    `Lag en setning med ${tema} om noe fra hverdagen din.`,
    `Når bruker vi dette på norsk? Gi et eksempel.`,
    `Hva er den vanligste feilen folk gjør med ${tema}?`
  ];

  diskSpørsmål.forEach((spm, i) => {
    const y = 1.5 + i * 0.95;
    s7.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.78,
      fill: { color: "FFFFFF", transparency: 88 },
      line: { color: "FFFFFF", transparency: 60, pt: 0.5 }
    });
    s7.addText(`${i + 1}.  ${spm}`, {
      x: 0.6, y: y + 0.08, w: 8.8, h: 0.62,
      fontSize: 14, color: C.white, fontFace: "Calibri", valign: "middle"
    });
  });

  s7.addNotes(`Avsluttende diskusjon. La elevene jobbe i par eller grupper med disse spørsmålene. Oppsummer timen: Hva har vi lært om ${tema} i dag? Knytt tilbake til læringsmålene fra slide 2.`);

  await pres.writeFile({ fileName: outputPath });
  console.log("OK:" + outputPath);
}

lagPresentasjon().catch(e => { console.error(e); process.exit(1); });
