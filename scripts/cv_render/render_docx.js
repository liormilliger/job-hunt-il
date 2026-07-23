/*
 * render_docx.js — job-hunt-il
 * ==========================================
 * Renders a CV or cover-letter/interview-prep package to .docx, matching the
 * visual style of the established HTML CV templates (CV and
 * Cover_Letter.html, one per filed application folder): Arial body text,
 * bold uppercase section titles with
 * a bottom border, role headers with the date right-aligned via a tab stop,
 * bullet lists, and (for cover letters) a shaded metadata block with a score
 * badge and bordered sections.
 *
 * Usage:  node render_docx.js <input.json> <output.docx>
 *
 * Input JSON shape — see generate_cvs_from_candidates.py for the producer.
 * Supports rtl:true for Hebrew (bidirectional paragraphs, right alignment).
 */

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  TabStopType, TabStopPosition, BorderStyle, ShadingType, HeadingLevel,
} = require("docx");

const [, , inputPath, outputPath] = process.argv;
if (!inputPath || !outputPath) {
  console.error("Usage: node render_docx.js <input.json> <output.docx>");
  process.exit(1);
}

const spec = JSON.parse(fs.readFileSync(inputPath, "utf-8"));
const RTL = !!spec.rtl;

// A4 in DXA, 0.75in margins to match the HTML template's `margin: 0.75in`
const PAGE = {
  size: { width: 11906, height: 16838 },
  margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
};

const COLOR = {
  text: "111111",
  gray: "555555",
  contactGray: "444444",
  border: "222222",
  green: "276221",
  greenBg: "C6EFCE",
  amber: "9C6500",
  amberBg: "FFEB9C",
  red: "9C0006",
  redBg: "FFC7CE",
  metaBorder: "1F4E79",
  metaBg: "F5F5F5",
  sectionDivider: "DDDDDD",
};

function para(children, opts = {}) {
  return new Paragraph({
    bidirectional: RTL,
    alignment: opts.alignment || (RTL ? AlignmentType.RIGHT : AlignmentType.LEFT),
    spacing: opts.spacing,
    border: opts.border,
    children,
  });
}

function run(text, opts = {}) {
  return new TextRun({
    text,
    font: "Arial",
    size: opts.size || 19, // 9.5pt
    bold: !!opts.bold,
    italics: !!opts.italics,
    color: opts.color || COLOR.text,
    rightToLeft: RTL,
  });
}

// ── Bullet list numbering config ────────────────────────────────────────────
const numbering = {
  config: [
    {
      reference: "cvBullets",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
          style: {
            paragraph: {
              indent: RTL
                ? { right: 720, hanging: 360 }
                : { left: 720, hanging: 360 },
            },
          },
        },
      ],
    },
  ],
};

function bullet(text) {
  return new Paragraph({
    bidirectional: RTL,
    alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
    numbering: { reference: "cvBullets", level: 0 },
    spacing: { after: 40 },
    children: [run(text, { size: 19 })],
  });
}

// ── CV renderer ──────────────────────────────────────────────────────────────
function buildCv(spec) {
  const children = [];

  // ── Name / title block ────────────────────────────────────────────────────
  if (RTL) {
    // Hebrew: "קורות חיים - [Name]" centered, bold, underlined (matches reference template)
    children.push(
      new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [run("קורות חיים - " + spec.name, { size: 40, bold: true, underline: {} })],
      })
    );
    // "פרטים אישיים:" section header before contact line
    children.push(
      new Paragraph({
        bidirectional: true,
        alignment: AlignmentType.RIGHT,
        spacing: { before: 60, after: 20 },
        children: [run("פרטים אישיים:", { size: 19, bold: true, underline: {} })],
      })
    );
  } else {
    // English: just the name
    children.push(
      new Paragraph({
        bidirectional: false,
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [run(spec.name, { size: 40, bold: true })],
      })
    );
  }

  // Contact line — right-aligned in RTL, centered in LTR
  children.push(
    new Paragraph({
      bidirectional: RTL,
      alignment: RTL ? AlignmentType.RIGHT : AlignmentType.CENTER,
      spacing: { after: 220 },
      children: [run(spec.contact, { size: 18, color: COLOR.contactGray })],
    })
  );

  // ── Section title helper ──────────────────────────────────────────────────
  // RTL: bold + text underline (reference style)
  // LTR: bold uppercase + full-width paragraph bottom border
  function sectionTitle(text) {
    if (RTL) {
      children.push(
        new Paragraph({
          bidirectional: true,
          alignment: AlignmentType.RIGHT,
          spacing: { before: 200, after: 80 },
          children: [run(text, { size: 20, bold: true, underline: {} })],
        })
      );
    } else {
      children.push(
        new Paragraph({
          bidirectional: false,
          alignment: AlignmentType.LEFT,
          spacing: { before: 200, after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR.border, space: 2 } },
          children: [run(text.toUpperCase(), { size: 20, bold: true })],
        })
      );
    }
  }

  // ── Professional summary ──────────────────────────────────────────────────
  // RTL: reference CV omits the section heading; summary flows after personal details
  if (!RTL) {
    sectionTitle(spec.summaryHeading || "Professional Summary");
  }
  children.push(
    new Paragraph({
      bidirectional: RTL,
      alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
      spacing: { after: 100, line: 270 },
      children: [run(spec.summary, { size: 19 })],
    })
  );

  // ── Work experience ───────────────────────────────────────────────────────
  sectionTitle(spec.experienceHeading || (RTL ? "ניסיון תעסוקתי:" : "Work Experience"));
  for (const role of spec.experience) {
    if (RTL) {
      // Split "Company — Hebrew Title" on em dash (template separator, not prose)
      const emDashIdx = role.title.indexOf(" — ");
      const company = emDashIdx >= 0 ? role.title.slice(0, emDashIdx).trim() : role.title;
      const jobTitle = emDashIdx >= 0 ? role.title.slice(emDashIdx + 3).trim() : "";

      // Line 1: Company name (bold, underlined) + dates in gray
      children.push(
        new Paragraph({
          bidirectional: true,
          alignment: AlignmentType.RIGHT,
          spacing: { before: 140, after: 0 },
          children: [
            run(company, { size: 19, bold: true, underline: {} }),
            run("  " + role.dates, { size: 18, color: COLOR.gray }),
          ],
        })
      );
      // Line 2: Hebrew job title in blue (matches reference color style)
      if (jobTitle) {
        children.push(
          new Paragraph({
            bidirectional: true,
            alignment: AlignmentType.RIGHT,
            spacing: { before: 0, after: 20 },
            children: [run(jobTitle, { size: 19, color: "0070C0" })],
          })
        );
      }
    } else {
      // LTR: company+title on one line, dates tab-stopped to the far right
      children.push(
        new Paragraph({
          bidirectional: false,
          spacing: { before: 140, after: 20 },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            run(role.title, { size: 20, bold: true }),
            run("\t" + role.dates, { size: 18, color: COLOR.gray }),
          ],
        })
      );
    }
    for (const b of role.bullets) {
      children.push(bullet(b));
    }
  }

  // ── Education ─────────────────────────────────────────────────────────────
  sectionTitle(spec.educationHeading || (RTL ? "השכלה:" : "Education"));
  for (const edu of spec.education) {
    if (RTL) {
      // RTL: dates first, then degree + school + honors (reference format)
      const honorsStr = edu.honors ? "  (" + edu.honors + ")" : "";
      children.push(
        new Paragraph({
          bidirectional: true,
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            run(edu.dates + ":  ", { size: 19 }),
            run(edu.degree, { bold: true, size: 19 }),
            run("  " + edu.school + honorsStr, { size: 19 }),
          ],
        })
      );
    } else {
      // LTR: degree · school · honors · dates
      children.push(
        new Paragraph({
          bidirectional: false,
          alignment: AlignmentType.LEFT,
          spacing: { after: 60 },
          children: [
            run(edu.degree, { bold: true, size: 19 }),
            run("  ·  " + edu.school + (edu.honors ? "  ·  " + edu.honors : "") + "  ·  " + edu.dates, { size: 19 }),
          ],
        })
      );
    }
  }

  // ── Skills ────────────────────────────────────────────────────────────────
  sectionTitle(spec.skillsHeading || (RTL ? "כישורים:" : "Skills"));
  for (const sk of spec.skills) {
    children.push(
      new Paragraph({
        bidirectional: RTL,
        alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [
          run(sk.label + "  ", { bold: true, size: 19 }),
          run(sk.items, { size: 19 }),
        ],
      })
    );
  }

  return children;
}

// ── Cover letter + interview prep renderer ──────────────────────────────────
function scoreColors(score) {
  if (score >= 80) return { bg: COLOR.greenBg, fg: COLOR.green };
  if (score >= 60) return { bg: COLOR.amberBg, fg: COLOR.amber };
  return { bg: COLOR.redBg, fg: COLOR.red };
}

function buildCover(spec) {
  const children = [];

  // Metadata block — shaded paragraphs with a left/right accent border,
  // mirroring the HTML .meta div (background + colored border-left).
  const metaRows = Object.entries(spec.meta || {});
  for (let i = 0; i < metaRows.length; i++) {
    const [label, value] = metaRows[i];
    const isFirst = i === 0;
    const isLast = i === metaRows.length - 1;
    const borderSide = RTL ? "right" : "left";
    const borderOpts = {
      [borderSide]: { style: BorderStyle.SINGLE, size: 24, color: COLOR.metaBorder, space: 8 },
    };
    children.push(
      new Paragraph({
        bidirectional: RTL,
        alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
        spacing: { before: isFirst ? 0 : 0, after: isLast ? 200 : 20 },
        shading: { fill: COLOR.metaBg, type: ShadingType.CLEAR },
        border: borderOpts,
        children: [
          run(label + ":  ", { bold: true, size: 18 }),
          run(String(value), { size: 18 }),
        ],
      })
    );
  }

  function h2(text) {
    children.push(
      new Paragraph({
        bidirectional: RTL,
        alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
        spacing: { before: 200, after: 100 },
        children: [run(text, { size: 24, bold: true })],
      })
    );
  }

  h2(spec.headingCover || "Cover Email Opening");
  children.push(
    new Paragraph({
      bidirectional: RTL,
      alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
      spacing: { after: 200, line: 300 },
      children: [run(spec.coverEmail, { size: 19 })],
    })
  );

  // Interview prep section — divider line above, like .section { border-top }
  children.push(
    new Paragraph({
      bidirectional: RTL,
      spacing: { before: 100 },
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: COLOR.sectionDivider, space: 8 } },
      children: [],
    })
  );
  h2(spec.headingPrep || "Interview Prep — Gaps and How to Frame Them");

  for (const gap of spec.gaps || []) {
    children.push(
      new Paragraph({
        bidirectional: RTL,
        alignment: RTL ? AlignmentType.RIGHT : AlignmentType.LEFT,
        spacing: { before: 140, after: 40 },
        children: [run(gap.gap, { bold: true, size: 19 })],
      })
    );
    if (gap.framing) children.push(bullet((gap.framingLabel || "Framing") + ": " + gap.framing));
    if (gap.angle) children.push(bullet((gap.angleLabel || "Angle") + ": " + gap.angle));
    if (gap.caveat) children.push(bullet((gap.caveatLabel || "Note") + ": " + gap.caveat));
  }

  return children;
}

// ── Assemble document ────────────────────────────────────────────────────────
const children = spec.type === "cv" ? buildCv(spec) : buildCover(spec);

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 19 } } },
  },
  numbering,
  sections: [
    {
      properties: { page: PAGE },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Wrote", outputPath);
});
