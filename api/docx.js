import { Document, Packer, Paragraph, TextRun, AlignmentType } from 'docx';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { text, filename } = req.body || {};
  if (!text) return res.status(400).json({ error: 'No text provided' });

  try {
    const doc = buildDocument(text);
    const buffer = await Packer.toBuffer(doc);

    const fname = (filename || 'minuta.docx').replace(/[^a-zA-Z0-9_\-\.]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}

// ── Helpers ────────────────────────────────────────────────────────────────────

const SEC_HEADERS = [
  'DO OBJETO:', 'DA SITUAÇÃO DO OBJETO:', 'DO VALOR DA VENDA:',
  'DAS CERTIDÕES:', 'DA IMISSÃO NA POSSE:', 'DAS DESPESAS',
  'DAS OBRIGAÇÕES:', 'DOS HONORÁRIOS DE CORRETAGEM:', 'DAS DISPOSIÇÕES FINAIS:',
  'EM TEMPO',
];

function isSecHeader(t) {
  return SEC_HEADERS.some(h => t.startsWith(h));
}

const FONT = 'Times New Roman';
const SIZE = 24; // 12pt em half-points

function inlineRuns(text) {
  const pat = /\b(OUTORGANTE[S]?|OUTORGADO[S]?|OUTORGADA[S]?|ANUENTE)\b/g;
  const runs = [];
  let last = 0, m;
  while ((m = pat.exec(text)) !== null) {
    if (m.index > last)
      runs.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: SIZE }));
    runs.push(new TextRun({ text: m[0], bold: true, font: FONT, size: SIZE }));
    last = m.index + m[0].length;
  }
  if (last < text.length)
    runs.push(new TextRun({ text: text.slice(last), font: FONT, size: SIZE }));
  return runs.length ? runs : [new TextRun({ text, font: FONT, size: SIZE })];
}

function buildRuns(line) {
  const t = line.trim();
  if (!t) return [new TextRun({ text: '', font: FONT, size: SIZE })];

  // Título ou cabeçalho de seção — linha inteira em negrito
  if (/^INSTRUMENTO PARTICULAR/.test(t) || isSecHeader(t)) {
    return [new TextRun({ text: line, bold: true, font: FONT, size: SIZE })];
  }

  // Título da parte (OUTORGANTE PROMITENTE VENDEDOR: ...) — negrito até o colon
  const partyM = line.match(/^((?:OUTORGANTE[S]?|OUTORGADO[S]?) PROMITENTE[S]?[^:]+:)([\s\S]*)/);
  if (partyM) {
    return [
      new TextRun({ text: partyM[1], bold: true, font: FONT, size: SIZE }),
      ...inlineRuns(partyM[2]),
    ];
  }

  // Numeração romana da cláusula — negrito só no número
  const romM = line.match(/^([IVXLCDM]+\)) ([\s\S]*)/);
  if (romM) {
    return [
      new TextRun({ text: romM[1] + ' ', bold: true, font: FONT, size: SIZE }),
      ...inlineRuns(romM[2]),
    ];
  }

  return inlineRuns(line);
}

function buildDocument(text) {
  const children = [];

  for (const para of text.split('\n\n')) {
    if (!para.trim()) {
      children.push(new Paragraph({
        children: [],
        spacing: { after: 80, line: 360, lineRule: 'auto' },
      }));
      continue;
    }

    for (const line of para.split('\n')) {
      if (!line.trim()) continue;
      const t = line.trim();
      const isSec = /^INSTRUMENTO PARTICULAR/.test(t) || isSecHeader(t);
      const isSig = /^_{5,}/.test(t);

      children.push(new Paragraph({
        children: buildRuns(line),
        alignment: isSig ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
        spacing: {
          before: isSec ? 280 : 0,
          after: 120,
          line: 360,
          lineRule: 'auto',
        },
      }));
    }
  }

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 }, // 2cm/2cm/2cm/3cm
        },
      },
      children,
    }],
  });
}
