const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, VerticalAlign
} = require('docx');
const fs = require('fs');

const data = JSON.parse(process.argv[2]);
const outputPath = process.argv[3];

const {
  tema, niva, tittel, forklaring, lesetekst, oppgaver, grammatikkForklaring, fasit
} = data;

const PAGE_W = 11906;
const CONTENT_W = 9026;
const MARGIN = 1418;

const MBO_BLÅ = "1F4E79";
const LYS_BLÅ = "D6E4F0";
const LYS_GRØNN = "EAF4EA";
const LYS_GUL = "FFF3CD";
const LYS_GRÅ = "F5F5F5";
const HVIT = "FFFFFF";

function topptekst() {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: MBO_BLÅ, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 100, left: 200, right: 200 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "Molde voksenopplæringssenter – MBO", bold: true, color: HVIT, size: 28, font: "Arial" })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: `Grammatikk: ${tema}   |   Nivå: ${niva}`, bold: true, color: HVIT, size: 24, font: "Arial" })]
              }),
            ]
          })
        ]
      })
    ]
  });
}

function navnDatoLinje() {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [
      new TextRun({ text: "Navn: ________________________________   Dato: ____________", size: 22, font: "Arial", color: "333333" })
    ]
  });
}

function seksjonsBoks(tittelTekst, farge, innhold) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: farge },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: farge },
              left: { style: BorderStyle.SINGLE, size: 4, color: farge },
              right: { style: BorderStyle.SINGLE, size: 4, color: farge }
            },
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: farge, type: ShadingType.CLEAR },
            margins: { top: 140, bottom: 140, left: 200, right: 200 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: tittelTekst, bold: true, size: 26, font: "Arial", color: "1A1A1A" })]
              }),
              ...innhold
            ]
          })
        ]
      })
    ]
  });
}

function tomLinje(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function avsnitt(tekst, bold = false, size = 22) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, bold, size, font: "Arial", color: "1A1A1A" })]
  });
}

function lagFasit(fasitTekst) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: {
              top: { style: BorderStyle.SINGLE, size: 6, color: "888888" },
              bottom: { style: BorderStyle.SINGLE, size: 2, color: "888888" },
              left: { style: BorderStyle.SINGLE, size: 2, color: "888888" },
              right: { style: BorderStyle.SINGLE, size: 2, color: "888888" }
            },
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: LYS_GRÅ, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: "FASIT", bold: true, size: 28, font: "Arial", color: MBO_BLÅ })]
              }),
              ...fasitTekst.split('\n').filter(l => l.trim()).map(linje =>
                new Paragraph({
                  spacing: { before: 80, after: 40 },
                  children: [new TextRun({ text: linje, size: 20, font: "Arial", color: "333333" })]
                })
              )
            ]
          })
        ]
      })
    ]
  });
}

function lagOppgaveBlokk(oppgave, index) {
  const bokstaver = ['a', 'b', 'c', 'd', 'e'];
  const brev = bokstaver[index] || String(index + 1);

  const rows = [];

  rows.push(
    new Paragraph({
      spacing: { before: 320, after: 100 },
      children: [
        new TextRun({ text: `Oppgave ${brev.toUpperCase()}) – ${oppgave.type}`, bold: true, size: 24, font: "Arial", color: MBO_BLÅ })
      ]
    })
  );

  rows.push(
    new Paragraph({
      spacing: { before: 60, after: 80 },
      children: [new TextRun({ text: oppgave.instruksjon, size: 22, font: "Arial", color: "333333" })]
    })
  );

  oppgave.innhold.forEach(linje => {
    rows.push(
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })]
      })
    );
  });

  if (oppgave.skrivelinje) {
    for (let i = 0; i < oppgave.skrivelinje; i++) {
      rows.push(
        new Paragraph({
          spacing: { before: 120, after: 0 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
          children: [new TextRun({ text: " ", size: 22, font: "Arial" })]
        })
      );
    }
  }

  return rows;
}

async function lagDokument() {
  const numbering = {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  };

  const children = [];

  children.push(topptekst());
  children.push(navnDatoLinje());
  children.push(tomLinje(120));

  // Læringsmål
  const læringsmål = [
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: `Etter denne økten kan jeg forstå og bruke: ${tema}`, size: 22, font: "Arial" })] })
  ];
  children.push(seksjonsBoks("Læringsmål", LYS_BLÅ, læringsmål));
  children.push(tomLinje(160));

  // Grammatikkforklaring
  const forklaringLinjer = forklaring.split('\n').filter(l => l.trim()).map(linje =>
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })] })
  );
  children.push(seksjonsBoks("Grammatikk – forklaring", LYS_GRØNN, forklaringLinjer));
  children.push(tomLinje(160));

  // Lesetekst
  const leseLinjer = lesetekst.split('\n').filter(l => l.trim()).map(linje =>
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: linje, size: 22, font: "Arial", color: "222222" })] })
  );
  children.push(seksjonsBoks("Lesetekst", LYS_BLÅ, leseLinjer));
  children.push(tomLinje(160));

  // Overskrift oppgaver
  children.push(new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text: "Oppgaver", bold: true, size: 28, font: "Arial", color: MBO_BLÅ })]
  }));
  children.push(new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MBO_BLÅ } },
    children: []
  }));

  // Oppgaver a-e
  oppgaver.forEach((oppgave, i) => {
    lagOppgaveBlokk(oppgave, i).forEach(p => children.push(p));
    children.push(tomLinje(120));
  });

  // Muntlig øvelse
  const muntligInnhold = [
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Snakk med en partner:", bold: true, size: 22, font: "Arial" })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: `Bruk ${tema} i en setning om deg selv.`, size: 22, font: "Arial" })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Spør hverandre: Kan du gi et eksempel?", size: 22, font: "Arial" })] })
  ];
  children.push(tomLinje(160));
  children.push(seksjonsBoks("🗣 Muntlig øvelse (par)", LYS_GUL, muntligInnhold));
  children.push(tomLinje(240));

  // Sideskift + fasit
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(lagFasit(fasit));

  const doc = new Document({
    numbering,
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } }
    },
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

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
  console.log("OK:" + outputPath);
}

lagDokument().catch(e => { console.error(e); process.exit(1); });
