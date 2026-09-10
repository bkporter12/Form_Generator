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

// Helper to find the first capitalized word after the first name
function getLastNameSortKey(fullName) {
  const name = fullName || "";
  const parts = name.trim().split(/\s+/);
  
  if (parts.length <= 1) return name.toUpperCase();
  
  for (let i = 1; i < parts.length; i++) {
    if (/^[A-Z]/.test(parts[i])) {
      return parts.slice(i).join(' ').toUpperCase();
    }
  }
  
  return parts[parts.length - 1].toUpperCase();
}

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
    
    const isAbsentA = a.Name.startsWith("Absent");
    const isAbsentB = b.Name.startsWith("Absent");
    if (isAbsentA && !isAbsentB) return 1;
    if (!isAbsentA && isAbsentB) return -1;
    
    const lastA = getLastNameSortKey(a.Name);
    const lastB = getLastNameSortKey(b.Name);
    return lastA.localeCompare(lastB);
  });

  let currentOfficial = 1;
  let currentPractice = 51; 

  balanced.forEach(j => {
    if (j.Type === 'Official') {
      j.Number = currentOfficial++;
    } else if (j.Type === 'Practice') {
      j.Number = currentPractice++;
    }
    if (j.Name.startsWith("Absent")) j.Print = false;
  });

  return balanced;
}

// --- PDF GENERATION ---
async function fetchTemplate(templateName) {
  const res = await fetch(`/templates/${templateName}`);
  if (!res.ok) throw new Error(`Template ${templateName} not found`);
  return await res.arrayBuffer();
}

async function drawOverlayText(page, font, boldFont, data, isShort, isRotated = false, applyMargin = false, paperSize = "Letter") {
  const { judge_name, judge_num, comp_name, comp_num, district, session, date, director } = data;

  const PAGE_WIDTH = paperSize === 'A4' ? 595.28 : 612;
  const PAGE_HEIGHT = paperSize === 'A4' ? 841.89 : 792;
  const scaleX = PAGE_WIDTH / 612;
  const scaleY = PAGE_HEIGHT / 792;

  const s = 576 / 612; 
  const tx = 18;       
  const ty = (792 - (792 * s)) / 2; 

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
      size = size * s;
    }

    x = x * scaleX;
    y = y * scaleY;
    size = size * Math.min(scaleX, scaleY);

    page.drawText(text, { x, y, size, font: f, color: rgb(0,0,0), rotate: degrees(isRotated ? 180 : 0) });
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

  if (!isShort && director) {
    if (session.includes("Chorus")) {
      drawText(director, LAYOUT.margin_left, LAYOUT.comp_y - 14, 12, font);
    } else if (session.includes("Quartet")) {
      const parts = director.split(',').map(s => s.trim()).filter(s => s);
      let line1 = director;
      let line2 = "";
      
      if (parts.length >= 3) {
        line1 = parts.slice(0, 2).join(', ');
        line2 = parts.slice(2).join(', ');
      }
      
      drawText(line1, LAYOUT.margin_left, LAYOUT.comp_y - 12, 10, font);
      if (line2) {
        drawText(line2, LAYOUT.margin_left, LAYOUT.comp_y - 24, 10, font);
      }
    }
  }

  const contestText = `${district} - ${session}, ${date}`;
  const contestWidth = font.widthOfTextAtSize(contestText, 10);
  
  if (isShort) {
    drawText(contestText, LAYOUT.page_center - (contestWidth / 2), LAYOUT.contest_y, 10, font);
  } else {
    drawText(contestText, LAYOUT.margin_right - contestWidth, LAYOUT.contest_y, 10, font);
  }
}

// 1. GENERATE BY CATEGORY 
export async function generateCategoryPDFs(judges, competitors, context, paperSize) {
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
            
            const templateDoc = await PDFDocument.load(templateBytes);
            const [copiedPage] = await outputDoc.copyPages(templateDoc, [0]);
            
            let targetPage;
            if (paperSize === 'A4') {
              const embedded = await outputDoc.embedPage(copiedPage);
              targetPage = outputDoc.addPage([595.28, 841.89]);
              targetPage.drawPage(embedded, { width: 595.28, height: 841.89 });
            } else {
              targetPage = copiedPage;
              outputDoc.addPage(targetPage);
            }

            await drawOverlayText(targetPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp1.Name, comp_num: comp1.Number }, true, false, true, paperSize);
            if (comp2) await drawOverlayText(targetPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp2.Name, comp_num: comp2.Number }, true, true, true, paperSize);
          }
        } else {
           for (const comp of competitors) {
              const templateDoc = await PDFDocument.load(templateBytes);
              const copiedPages = await outputDoc.copyPages(templateDoc, templateDoc.getPageIndices());
              
              let firstTargetPage;
              for (let idx = 0; idx < copiedPages.length; idx++) {
                let targetPage;
                if (paperSize === 'A4') {
                  const embedded = await outputDoc.embedPage(copiedPages[idx]);
                  targetPage = outputDoc.addPage([595.28, 841.89]);
                  targetPage.drawPage(embedded, { width: 595.28, height: 841.89 });
                } else {
                  targetPage = copiedPages[idx];
                  outputDoc.addPage(targetPage);
                }
                if (idx === 0) firstTargetPage = targetPage;
              }

              await drawOverlayText(firstTargetPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp.Name, comp_num: comp.Number, director: comp.Director }, false, false, true, paperSize);
           }
        }
      }
      const pdfBytes = await outputDoc.save();
      const safeDate = context.date.replace(/[/]/g, "-");
      zip.file(`${context.session.replace(/[^a-z0-9]/gi, '_')}_${t_name.replace(".pdf", "")}_${safeDate}.pdf`, pdfBytes);
      filesGenerated++;
    }
  }

  if (filesGenerated > 0) {
    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Category_Files.zip`);
  }
}

// 2. GENERATE BY JUDGE 
export async function generateJudgePDFs(judges, competitors, context, paperSize) {
  const zip = new JSZip();
  let filesGenerated = 0;

  for (const judge of judges) {
    if (judge.Name.startsWith("Absent")) continue;

    const formats = FORMAT_MAPPING[judge.Category];
    if (!formats) continue;

    const outputDoc = await PDFDocument.create();
    const helvetica = await outputDoc.embedFont(StandardFonts.Helvetica);
    const helveticaBold = await outputDoc.embedFont(StandardFonts.HelveticaBold);
    let pagesAdded = 0;

    for (const t_name of formats) {
      if (!t_name.includes("Long")) continue;

      const templateBytes = await fetchTemplate(t_name).catch(() => null);
      if (!templateBytes) continue;

      for (const comp of competitors) {
        const templateDoc = await PDFDocument.load(templateBytes);
        const copiedPages = await outputDoc.copyPages(templateDoc, templateDoc.getPageIndices());
        
        let firstTargetPage;
        for (let idx = 0; idx < copiedPages.length; idx++) {
          let targetPage;
          if (paperSize === 'A4') {
            const embedded = await outputDoc.embedPage(copiedPages[idx]);
            targetPage = outputDoc.addPage([595.28, 841.89]);
            targetPage.drawPage(embedded, { width: 595.28, height: 841.89 });
          } else {
            targetPage = copiedPages[idx];
            outputDoc.addPage(targetPage);
          }
          if (idx === 0) firstTargetPage = targetPage;
        }
        
        await drawOverlayText(firstTargetPage, helvetica, helveticaBold, { ...context, judge_name: judge.Name, judge_num: judge.Number, comp_name: comp.Name, comp_num: comp.Number, director: comp.Director }, false, false, false, paperSize);
        pagesAdded++;
      }
    }

    if (pagesAdded > 0) {
      const pdfBytes = await outputDoc.save();
      const safeJudge = judge.Name.replace(/[^a-z0-9]/gi, '_');
      const safeDate = context.date.replace(/[/]/g, "-");
      const fname = `${context.session.replace(/[^a-z0-9]/gi, '_')}_${safeJudge}_${safeDate}.pdf`;
      zip.file(fname, pdfBytes);
      filesGenerated++;
    }
  }

  if (filesGenerated > 0) {
    const sortedComps = [...competitors].sort((a, b) => {
      const numA = a.Number ? parseInt(a.Number) : Infinity;
      const numB = b.Number ? parseInt(b.Number) : Infinity;
      if (numA !== numB) return numA - numB;
      return (a.Name || "").localeCompare(b.Name || "");
    });

    let txt = `${context.district} - ${context.session}\n${context.date}\n\n`;
    txt += `Order of Appearance:\n`;
    txt += `--------------------\n`;
    
    sortedComps.forEach(comp => {
      const oa = comp.Number || "TBD";
      let line = `${oa}. ${comp.Name}`;
      if (comp.Director) {
        line += ` (${comp.Director})`;
      }
      txt += line + `\n`;
    });

    zip.file(`${context.session.replace(/[^a-z0-9]/gi, '_')}_Competitor_List.txt`, txt);

    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Judge_Packets.zip`);
  }
}

// 3. GENERATE FOLDER LABELS
export function generateFolderLabelsRTF(judges, context, paperSize) {
  const pw = paperSize === 'A4' ? 11906 : 12240;
  const ph = paperSize === 'A4' ? 16838 : 15840;
  let rtf = `{\\rtf1\\ansi\\deff0\\nouicompat\\viewkind4\\uc1{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}{\\colortbl ;\\red0\\green0\\blue0;}\\paperw${pw}\\paperh${ph}\\margl225\\margr225\\margt720\\margb720\\pard\\plain\\fs20\n`;
  const activeJudges = judges.filter(j => !j.Name.startsWith("Absent"));

  for (let i = 0; i < activeJudges.length; i += 2) {
    const j1 = activeJudges[i];
    const j2 = activeJudges[i + 1];
    
    rtf += `\\trowd\\trgaph108\\trleft0\\trrh2880\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx5760\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx6030\\clvertalc\\brdrt\\brdrnil\\brdrl\\brdrnil\\brdrb\\brdrnil\\brdrr\\brdrnil\\cellx11790\n`;
    
    const cFull1 = CAT_FULL_NAMES[j1.Category] || j1.Category;
    rtf += `\\pard\\intbl\\qc\\sa0\\sb0\\b\\f0\\fs28 ${escapeRTF(j1.Name)}\\b0\\par ` +
           `\\fs22 ${escapeRTF(cFull1)} Category\\par ` +
           `\\fs20 ${escapeRTF(context.session)}\\par ` +
           `${escapeRTF(context.district)}\\par ` +
           `${escapeRTF(context.date)}\\cell\\pard\\intbl\\cell\n`;
    
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

// 4. GENERATE OVERLAYS ONLY (RTF)
export function generateOverlaysRTF(judges, competitors, context, paperSize) {
  const pw = paperSize === 'A4' ? 11906 : 12240;
  const ph = paperSize === 'A4' ? 16838 : 15840;
  
  let rtf = `{\\rtf1\\ansi\\deff0\\nouicompat\\viewkind4\\uc1{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}{\\colortbl ;\\red0\\green0\\blue0;}\\paperw${pw}\\paperh${ph}\\margl1000\\margr1000\\margt540\\margb1000\n`;
  
  const activeJudges = judges.filter(j => !j.Name.startsWith("Absent"));

  for (const judge of activeJudges) {
    for (const comp of competitors) {
      const judgeText = judge.Number ? `${judge.Number}. ${judge.Name}` : judge.Name;
      const contestText = `${context.district} - ${context.session}, ${context.date}`;
      
      rtf += `\\pard\\qc\\sa100\\f0\\fs20 ${escapeRTF(contestText)}\\par\n`;
      rtf += `\\pard\\qr\\sa100\\b\\fs32 ${escapeRTF(judgeText)}\\b0\\par\n`;
      rtf += `\\pard\\ql\\sa100\\fs24 ${escapeRTF(comp.Number + ". " + comp.Name)}\\par\n`;
      
      if (comp.Director) {
        if (context.session.includes("Chorus")) {
          rtf += `\\pard\\ql\\sa100\\fs24 ${escapeRTF(comp.Director)}\\par\n`;
        } else if (context.session.includes("Quartet")) {
          const parts = comp.Director.split(',').map(s => s.trim()).filter(s => s);
          let line1 = comp.Director;
          let line2 = "";
          if (parts.length >= 3) {
            line1 = parts.slice(0, 2).join(', ');
            line2 = parts.slice(2).join(', ');
          }
          rtf += `\\pard\\ql\\sa100\\fs20 ${escapeRTF(line1)}\\par\n`;
          if (line2) {
            rtf += `\\pard\\ql\\sa100\\fs20 ${escapeRTF(line2)}\\par\n`;
          }
        }
      }
      rtf += `\\page\n`;
    }
  }
  
  rtf += "}";
  const blob = new Blob([rtf], { type: "application/rtf" });
  saveAs(blob, `${context.session.replace(/[^a-z0-9]/gi, '_')}_Text_Overlays.rtf`);
}

// 5. GENERATE BLANK PDFs
export async function generateBlankPDFs(blankCounts, paperSize) {
  const zip = new JSZip();
  let filesGenerated = 0;
  let singlePdfBytes = null;
  let singlePdfName = "";

  const categories = [
    { id: 'MUS', name: 'Musicality' },
    { id: 'PER', name: 'Performance' },
    { id: 'SNG', name: 'Singing' }
  ];

  for (const cat of categories) {
    for (const type of ['Long', 'Short']) {
      const key = `${cat.id}_${type}`;
      const qty = parseInt(blankCounts[key]) || 0;
      if (qty <= 0) continue;

      const t_name = `${key}.pdf`; 
      const templateBytes = await fetchTemplate(t_name).catch(() => null);
      if (!templateBytes) continue;

      const outputDoc = await PDFDocument.create();
      const templateDoc = await PDFDocument.load(templateBytes);
      
      const copies = type === 'Short' ? Math.ceil(qty / 2) : qty;
      const pageIndices = templateDoc.getPageIndices();
      
      for (let i = 0; i < copies; i++) {
        const copiedPages = await outputDoc.copyPages(templateDoc, pageIndices);
        for (const copiedPage of copiedPages) {
          if (paperSize === 'A4') {
            const targetPage = outputDoc.addPage([595.28, 841.89]);
            const embedded = await outputDoc.embedPage(copiedPage);
            targetPage.drawPage(embedded, { width: 595.28, height: 841.89 });
          } else {
            outputDoc.addPage(copiedPage);
          }
        }
      }

      const pdfBytes = await outputDoc.save();
      const fname = `Blank_${cat.name}_${type}.pdf`;
      
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
    saveAs(zipBlob, `Blank_Forms.zip`);
  }
}
