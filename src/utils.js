import { PDFDocument, StandardFonts, rgb, degrees } from 'pdf-lib';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// --- CONFIGURATION ---
const LAYOUT = {
  judge_y: 765,
  comp_y: 745,
  contest_y: 730,
  margin_left: 50,
  margin_right: 562,
  page_center: 306,
};

const FORMAT_MAPPING = {
  MUS: ["MUS_Long.pdf", "MUS_Short.pdf"],
  PER: ["PER_Long.pdf", "PER_Short.pdf"],
  SNG: ["SNG_Long.pdf", "SNG_Short.pdf"]
};

const CAT_FULL_NAMES = { MUS: "Musicality", PER: "Performance", SNG: "Singing" };

// --- DATA PROCESSING ---
export function balanceAndSortJudges(judges) {
  const categories = ['MUS', 'PER', 'SNG'];
  let maxCount = 0;
  
  categories.forEach(cat => {
    const count = judges.filter(j => j.Category === cat && j.Type === 'Official' && !j.Name.startsWith('Absent')).length;
    if (count > maxCount) maxCount = count;
  });

  let balanced = judges.filter(j => !j.Name.startsWith("Absent"));
  
  if (maxCount > 0) {
    categories.forEach(cat => {
      const currentCount = balanced.filter(j => j.Category === cat && j.Type === 'Official').length;
      for (let i = 0; i < maxCount - currentCount; i++) {
        balanced.push({ Name: `Absent ${cat} Judge`, Category: cat, Type: "Official", Print: false, Number: "" });
      }
    });
  }

  const catOrder = { MUS: 0, PER: 1, SNG: 2 };
  const typeOrder = { Official: 0, Practice: 1 };
  
  balanced.sort((a, b) => {
    if (catOrder[a.Category] !== catOrder[b.Category]) return catOrder[a.Category] - catOrder[b.Category];
    if (typeOrder[a.Type] !== typeOrder[b.Type]) return typeOrder[a.Type] - typeOrder[b.Type];
    const lastA = (a.Name || "").trim().split(' ').pop();
    const lastB = (b.Name || "").trim().split(' ').pop();
    return lastA.localeCompare(lastB);
  });

  let currentOfficial = 1;
  let currentPractice = 50;

  // 1. Assign standard numbers
  balanced.forEach(j => {
    if (!j.Name.startsWith("Absent")) {
      if (j.Type === 'Official') j.Number = currentOfficial++;
      else if (j.Type === 'Practice') j.Number = currentPractice++;
    }
  });

  // 2. Assign Absent numbers (Highest Official in their Category)
  balanced.forEach(j => {
    if (j.Name.startsWith("Absent")) {
      j.Print = false; // Always unselectable
      const officials = balanced.filter(x => x.Category === j.Category && x.Type === 'Official' && !x.Name.startsWith("Absent"));
      if (officials.length > 0) {
        j.Number = Math.max(...officials.map(o => o.Number || 0));
      } else {
        j.Number = "";
      }
    }
  });

  return balanced;
}

// --- PDF GENERATION ---
async function fetchTemplate(templateName) {
  const res = await fetch(`/templates/${templateName}`);
  if (!res.ok) throw new Error(`Template ${templateName} not found`);
  return await res.arrayBuffer();
}

async function drawOverlayText(page, font, boldFont, data, isShort, isRotated = false, applyMargin = false) {
  const { judge_name, judge_num, comp_name, comp_num, district, session, date, director } = data;

  // 0.25" Margin Transformation Math
  const s = 576 / 612; // Scale down to fit 0.25" margin on each side
  const tx = 18;       // Translate X inward by 18 points (0.25")
  const ty = (792 - (792 * s)) / 2; // Keep Y centered after scale

  const drawText = (text, origX, origY, size, f) => {
    let x = origX;
    let y = origY;

    if (isRotated) {
      x = 612 - x;
      y = 792 - y;
    }

    if (applyMargin) {
      x = (x * s) + tx;
      y = (y * s) + ty;
      page.drawText(text, { x, y, size: size * s, font: f, color: rgb(0,0,0), rotate: degrees(isRotated ? 180 : 0) });
    } else {
      page.drawText(text, { x, y, size: size, font: f, color: rgb(0,0,0), rotate: degrees(isRotated ? 180 : 0) });
    }
  };

  let nameText = String(judge_name);
  let numText = String(judge_num || ""); 
  
  if (isShort) {
    const nameWidth = boldFont.widthOfTextAtSize(nameText, 16);
    drawText(nameText, LAYOUT.margin_right - nameWidth, LAYOUT.judge_y, 16, boldFont);
    
    const numWidth = boldFont.widthOfTextAtSize(numText, 36);
    drawText(numText, LAYOUT.margin_right - nameWidth - 15 - numWidth, LAYOUT.judge_y, 36, boldFont);
  } else {
    const text = numText ? `${numText}. ${nameText}` : nameText;
    const textWidth = boldFont.widthOfTextAtSize(text, 16);
    drawText(text, LAYOUT.margin_right - textWidth, LAYOUT.judge_y, 16, boldFont);
  }

  drawText(`${comp_num}. ${comp_name}`, LAYOUT.margin_left, LAYOUT.comp_y, 12, font);

  if (!isShort && session.includes("Chorus") && director) {
    drawText(director, LAYOUT.margin_left, LAYOUT.comp_y - 14, 12, font);
  }

  const contestText = `${district} - ${session}, ${date}`;
  const contestWidth = font.widthOfTextAtSize(contestText, 10);
  drawText(contestText, LAYOUT.page_center - (contestWidth / 2), LAYOUT.contest_y, 10, font);
}

// 1. GENERATE BY CATEGORY (With 0.25" Margin Applied)
export async function generateCategoryPDFs(judges, competitors, context) {
  const zip = new JSZip();
  let filesGenerated = 0;

  for (const [cat, formats] of Object.entries(FORMAT_MAPPING)) {
    const catJudges = judges.filter(j => j.Category === cat && !j.Name.startsWith("Absent"));
    if (catJudges.length === 0) continue;

    for (const t_name of formats) {
      const templateBytes = await fetchTemplate(t_name).catch(() => null);
      if (!templateBytes) continue;

      const isShort = t_name.includes("Short");
      const outputDoc = await PDFDocument.create();
      const helvetica = await outputDoc.embedFont(StandardFonts.Helvetica);
      const helveticaBold = await outputDoc.embedFont(StandardFonts.HelveticaBold);

      for (const judge of catJudges) {
        if (isShort) {
          for (let i = 0; i < competitors.length; i += 2) {
            const comp1 = competitors[i];
            const comp2 = competitors[i + 1];
            const [copiedPage] = await outputDoc.copyPages(await PDFDocument.load(templateBytes), [0]);
            outputDoc.addPage(copiedPage);

            await drawOverlayText(copiedPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp1.Name, comp_num: comp1.Number }, true, false, true);
            if (comp2) await drawOverlayText(copiedPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp2.Name, comp_num: comp2.Number }, true, true, true);
          }
        } else {
           for (const comp of competitors) {
              const templateDoc = await PDFDocument.load(templateBytes);
              const copiedPages = await outputDoc.copyPages(templateDoc, templateDoc.getPageIndices());
              await drawOverlayText(copiedPages[0], helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp.Name, comp_num: comp.Number, director: comp.Director }, false, false, true);
              copiedPages.forEach(p => outputDoc.addPage(p));
           }
        }
      }
      const pdfBytes = await outputDoc.save();
      zip.file(`${context.session.replace(/[^a-z0-9]/gi, '_')}_${t_name.replace(".pdf", "")}_${context.date.replace(/\//g, "-")}.pdf`, pdfBytes);
      filesGenerated++;
    }
  }

  if (filesGenerated > 0) {
    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Category_Files.zip`);
  }
}

// 2. GENERATE BY JUDGE (Long Form Only, No Margin)
export async function generateJudgePDFs(judges, competitors, context) {
  const zip = new JSZip();
  let filesGenerated = 0;
  let singlePdfBytes = null;
  let singlePdfName = "";

  for (const judge of judges) {
    if (judge.Name.startsWith("Absent")) continue;

    const formats = FORMAT_MAPPING[judge.Category];
    if (!formats) continue;

    const outputDoc = await PDFDocument.create();
    const helvetica = await outputDoc.embedFont(StandardFonts.Helvetica);
    const helveticaBold = await outputDoc.embedFont(StandardFonts.HelveticaBold);
    let pagesAdded = 0;

    for (const t_name of formats) {
      if (!t_name.includes("Long")) continue; // Restrict to Long forms only

      const templateBytes = await fetchTemplate(t_name).catch(() => null);
      if (!templateBytes) continue;

      for (const comp of competitors) {
        const templateDoc = await PDFDocument.load(templateBytes);
        const copiedPages = await outputDoc.copyPages(templateDoc, templateDoc.getPageIndices());
        
        await drawOverlayText(copiedPages[0], helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp.Name, comp_num: comp.Number, director: comp.Director }, false, false, false);
        copiedPages.forEach(p => outputDoc.addPage(p));
        pagesAdded++;
      }
    }

    if (pagesAdded > 0) {
      const pdfBytes = await outputDoc.save();
      const safeJudge = judge.Name.replace(/[^a-z0-9]/gi, '_');
      const fname = `${context.session.replace(/[^a-z0-9]/gi, '_')}_${safeJudge}_${context.date.replace(/\//g, "-")}.pdf`;
      zip.file(fname, pdfBytes);
      filesGenerated++;
      
      singlePdfBytes = pdfBytes;
      singlePdfName = fname;
    }
  }

  if (filesGenerated === 1) {
    const blob = new Blob([singlePdfBytes], { type: "application/pdf" });
    saveAs(blob, singlePdfName);
  } else if (filesGenerated > 1) {
    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Judge_Packets.zip`);
  }
}

// --- RTF GENERATION ---
const escapeRTF = (text) => String(text || "").replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');

export function generateFolderLabelsRTF(judges, context) {
  let rtf = `{\\rtf1\\ansi\\deff0\\nouicompat\\viewkind4\\uc1{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}{\\colortbl ;\\red0\\green0\\blue0;}\\paperw12240\\paperh15840\\margl225\\margr225\\margt720\\margb720\\pard\\plain\\fs20\n`;
  const activeJudges = judges.filter(j => !j.Name.startsWith("Absent"));

  for (let i = 0; i < activeJudges.length; i += 2) {
    const j1 = activeJudges[i];
    const j2 = activeJudges[i + 1];
    
    rtf += `\\trowd\\trgaph108\\trleft0\\trrh2880\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx5760\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx6030\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx11790\n`;
    
    // Cell 1
    const cFull1 = CAT_FULL_NAMES[j1.Category] || j1.Category;
    rtf += `\\pard\\intbl\\qc\\sa0\\sb0\\b\\f0\\fs28 ${escapeRTF(j1.Name)}\\b0\\par ` +
           `\\fs22 ${escapeRTF(cFull1)} Category\\par ` +
           `\\fs20 ${escapeRTF(context.session)}\\par ` +
           `${escapeRTF(context.district)}\\par ` +
           `${escapeRTF(context.date)}\\cell\\pard\\intbl\\cell\n`;
    
    // Cell 2 (if exists)
    if (j2) {
      const cFull2 = CAT_FULL_NAMES[j2.Category] || j2.Category;
      rtf += `\\pard\\intbl\\qc\\sa0\\sb0\\b\\f0\\fs28 ${escapeRTF(j2.Name)}\\b0\\par ` +
             `\\fs22 ${escapeRTF(cFull2)} Category\\par ` +
             `\\fs20 ${escapeRTF(context.session)}\\par ` +
             `${escapeRTF(context.district)}\\par ` +
             `${escapeRTF(context.date)}\\cell\\row\n`;
    } else {
      rtf += `\\pard\\intbl\\cell\\row\n`;
    }
  }
  rtf += "}";
  const blob = new Blob([rtf], { type: "application/rtf" });
  saveAs(blob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Folder_Labels.rtf`);
}
