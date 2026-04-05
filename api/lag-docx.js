// api/lag-docx.js
// Vercel serverless function – genererer Word-dokument fra grammatikkdata

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat, PageBreak
} = require('docx');

const PAGE_W = 11906;
const CONTENT_W = 9026;
const MARGIN = 1418;

const MBO_BLÅ = "1F4E79";
const LYS_BLÅ = "D6E4F0";
const LYS_GRØNN = "EAF4EA";
const LYS_GUL = "FFF3CD";
const LYS_GRÅ = "F5F5F5";
const HVIT = "FFFFFF";

function topptekst(tema, niva) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: MBO_BLÅ, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 100, left: 200, right: 200 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Molde voksenopplæringssenter – MBO", bold: true, color: HVIT, size: 28, font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Grammatikk: ${tema}   |   Nivå: ${niva}`, bold: true, color: HVIT, size: 24, font: "Arial" })] }),
        ]
      })]
    })]
  });
}

function seksjonsBoks(tittelTekst, farge, innhold) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: farge }, bottom: { style: BorderStyle.SINGLE, size: 4, color: farge }, left: { style: BorderStyle.SINGLE, size: 4, color: farge }, right: { style: BorderStyle.SINGLE, size: 4, color: farge } },
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: farge, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: [
          new Paragraph({ children: [new TextRun({ text: tittelTekst, bold: true, size: 26, font: "Arial", color: "1A1A1A" })] }),
          ...innhold
        ]
      })]
    })]
  });
}

function tomLinje(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function lagOppgaveBlokk(oppgave, index) {
  const bokstaver = ['A', 'B', 'C', 'D', 'E'];
  const brev = bokstaver[index] || String(index + 1);
  const rows = [];

  rows.push(new Paragraph({
    spacing: { before: 360, after: 100 },
    children: [new TextRun({ text: `Oppgave ${brev}) – ${oppgave.type}`, bold: true, size: 26, font: "Arial", color: MBO_BLÅ })]
  }));

  rows.push(new Paragraph({
    spacing: { before: 80, after: 100 },
    children: [new TextRun({ text: oppgave.instruksjon, size: 22, font: "Arial", color: "333333", italics: true })]
  }));

  (oppgave.innhold || []).forEach(linje => {
    rows.push(new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })]
    }));
  });

  if (oppgave.skrivelinje > 0) {
    for (let i = 0; i < oppgave.skrivelinje; i++) {
      rows.push(new Paragraph({
        spacing: { before: 160, after: 0 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
        children: [new TextRun({ text: " ", size: 22, font: "Arial" })]
      }));
    }
  }

  return rows;
}

async function lagDokument(data) {
  const { tema, niva, forklaring, grammatikkForklaring, lesetekst, oppgaver, fasit } = data;

  const children = [];

  children.push(topptekst(tema, niva));
  children.push(new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text: "Navn: ________________________________   Dato: ____________", size: 22, font: "Arial", color: "333333" })]
  }));
  children.push(tomLinje(120));

  // Læringsmål
  children.push(seksjonsBoks("Læringsmål", LYS_BLÅ, [
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: `Etter denne leksjonen kan jeg forstå og bruke: ${tema}`, size: 22, font: "Arial" })] })
  ]));
  children.push(tomLinje(180));

  // Grammatikkforklaring
  const forklaringLinjer = (forklaring || '').split('\n').filter(l => l.trim()).map(linje =>
    new Paragraph({ spacing: { before: 70, after: 70 }, children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })] })
  );
  children.push(seksjonsBoks("Grammatikk – forklaring", LYS_GRØNN, forklaringLinjer));
  children.push(tomLinje(180));

  // Mønstre og eksempler
  if (grammatikkForklaring) {
    const mønsterLinjer = grammatikkForklaring.split('\n').filter(l => l.trim()).map(linje =>
      new Paragraph({ spacing: { before: 70, after: 70 }, children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })] })
    );
    children.push(seksjonsBoks("Mønstre og eksempler", LYS_BLÅ, mønsterLinjer));
    children.push(tomLinje(180));
  }

  // Lesetekst
  const leseLinjer = (lesetekst || '').split('\n').filter(l => l.trim()).map(linje =>
    new Paragraph({ spacing: { before: 90, after: 90 }, children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })] })
  );
  children.push(seksjonsBoks("Lesetekst", "EBF5FA", leseLinjer));
  children.push(tomLinje(180));

  // Overskrift oppgaver
  children.push(new Paragraph({
    spacing: { before: 160, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MBO_BLÅ } },
    children: [new TextRun({ text: "Oppgaver", bold: true, size: 30, font: "Arial", color: MBO_BLÅ })]
  }));

  // Oppgaver a–e
  (oppgaver || []).forEach((oppgave, i) => {
    lagOppgaveBlokk(oppgave, i).forEach(p => children.push(p));
    children.push(tomLinje(160));
  });

  // Muntlig øvelse
  children.push(tomLinje(120));
  children.push(seksjonsBoks("🗣 Muntlig øvelse (par)", LYS_GUL, [
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Snakk med en partner:", bold: true, size: 22, font: "Arial" })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: `Bruk ${tema} i to setninger om deg selv.`, size: 22, font: "Arial" })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: `Spør hverandre: Kan du gi et eksempel på ${tema}?`, size: 22, font: "Arial" })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Forklar regelen til hverandre med egne ord.", size: 22, font: "Arial" })] }),
  ]));

  // Fasit på ny side
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: "888888" }, bottom: { style: BorderStyle.SINGLE, size: 2, color: "888888" }, left: { style: BorderStyle.SINGLE, size: 2, color: "888888" }, right: { style: BorderStyle.SINGLE, size: 2, color: "888888" } },
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill: LYS_GRÅ, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ children: [new TextRun({ text: "FASIT", bold: true, size: 32, font: "Arial", color: MBO_BLÅ })] }),
          new Paragraph({ children: [] }),
          ...(fasit || '').split('\n').filter(l => l.trim()).map(linje =>
            new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: linje, size: 20, font: "Arial", color: "333333" })] })
          )
        ]
      })]
    })]
  }));

  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 22 } } } },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: 16838 },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
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

    const buffer = await lagDokument(data);
    const filnavn = `${(data.tema || 'grammatikk').replace(/[^a-zA-ZæøåÆØÅ0-9]/g, '_')}_${data.niva || 'A1'}.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filnavn}"`);
    res.send(buffer);

  } catch (e) {
    console.error(e);
    res.status(500).json({ feil: "Feil ved generering av Word-dokument: " + e.message });
  }
};
