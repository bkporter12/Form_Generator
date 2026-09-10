import React, { useState } from 'react';
import Papa from 'papaparse';
import { balanceAndSortJudges, generateCategoryPDFs, generateJudgePDFs, generateFolderLabelsRTF, generateOverlaysRTF, generateBlankPDFs } from './utils';

export default function App() {
  const [district, setDistrict] = useState("");
  const [date, setDate] = useState("");
  const [session, setSession] = useState("Quartet Semi-Finals");
  const [paperSize, setPaperSize] = useState("Letter");
  
  const [judges, setJudges] = useState([]);
  const [competitors, setCompetitors] = useState([]);
  
  const [blankCounts, setBlankCounts] = useState({
    MUS_Long: "", MUS_Short: "",
    PER_Long: "", PER_Short: "",
    SNG_Long: "", SNG_Short: ""
  });
  
  const [isGenerating, setIsGenerating] = useState(false);

  // --- PARSE UPLOADS ---
  const handleJudgeUpload = (e) => {
    if (!e.target.files.length) return;
    Papa.parse(e.target.files[0], {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        let clean = results.data
          .filter(r => r.Category && ['MUS', 'PER', 'SNG'].includes(r.Category.toUpperCase().trim()))
          .map(r => ({ Name: r.Name, Category: r.Category.toUpperCase().trim(), Type: r.Type, Print: true }));
        setJudges(balanceAndSortJudges(clean));
      }
    });
    e.target.value = null; 
  };

  const handleCompUpload = (e) => {
    if (!e.target.files.length) return;
    Papa.parse(e.target.files[0], {
      header: true, skipEmptyLines: true,
      complete: (results) => {
        let clean = results.data.map(r => ({ Number: r.OA, Name: r["Group Name"], Director: r["Director/Participant(s)"] || "", Print: true }));
        setCompetitors(clean);
      }
    });
    e.target.value = null; 
  };

  // --- JUDGE HANDLERS ---
  const toggleAllJudges = (checked) => setJudges(judges.map(j => j.Name.startsWith("Absent") ? j : { ...j, Print: checked }));
  const updateJudge = (index, field, value) => {
    const newJudges = [...judges];
    newJudges[index][field] = value;
    setJudges(newJudges);
  };
  const addJudge = () => setJudges([...judges, { Name: "", Category: "MUS", Type: "Official", Print: true, Number: "" }]);
  const removeJudge = (index) => setJudges(judges.filter((_, i) => i !== index));
  const clearJudges = () => setJudges([]);

  // --- COMPETITOR HANDLERS ---
  const toggleAllComps = (checked) => setCompetitors(competitors.map(c => ({ ...c, Print: checked })));
  const updateComp = (index, field, value) => {
    const newComps = [...competitors];
    newComps[index][field] = value;
    setCompetitors(newComps);
  };
  const addComp = () => setCompetitors([...competitors, { Number: "", Name: "", Director: "", Print: true }]);
  const removeComp = (index) => setCompetitors(competitors.filter((_, i) => i !== index));
  const clearComps = () => setCompetitors([]);
  
  const clearAllOAs = () => {
    setCompetitors(competitors.map(c => ({ ...c, Number: "" })));
  };

  const sortCompsByOA = () => {
    const seenNumbers = new Set();
    const duplicateNumbers = new Set();
    const validIntegers = [];
    
    competitors.forEach(c => {
      const numStr = String(c.Number).trim();
      if (numStr !== "") {
        if (seenNumbers.has(numStr)) {
          duplicateNumbers.add(numStr);
        } else {
          seenNumbers.add(numStr);
          const numInt = parseInt(numStr, 10);
          if (!isNaN(numInt)) {
            validIntegers.push(numInt);
          }
        }
      }
    });

    const skippedNumbers = [];
    if (validIntegers.length > 0) {
      const minNum = Math.min(...validIntegers);
      const maxNum = Math.max(...validIntegers);
      
      for (let i = minNum + 1; i < maxNum; i++) {
        if (!validIntegers.includes(i)) {
          skippedNumbers.push(i);
        }
      }
    }

    if (duplicateNumbers.size > 0 || skippedNumbers.length > 0) {
      let warningMsg = "⚠️ Warning regarding your Order of Appearance (OA) numbers:\n\n";
      if (duplicateNumbers.size > 0) warningMsg += `- Duplicated numbers: ${Array.from(duplicateNumbers).join(', ')}\n`;
      if (skippedNumbers.length > 0) warningMsg += `- Skipped numbers: ${skippedNumbers.join(', ')}\n`;
      warningMsg += "\nThe list has been sorted, but please verify your numbers.";
      alert(warningMsg);
    }

    const sorted = [...competitors].sort((a, b) => {
      const valA = String(a.Number).trim();
      const valB = String(b.Number).trim();
      
      const numA = valA !== "" ? parseInt(valA, 10) : Infinity;
      const numB = valB !== "" ? parseInt(valB, 10) : Infinity;
      
      if (numA !== numB) return numA - numB;
      return (a.Name || "").localeCompare(b.Name || "");
    });
    
    setCompetitors(sorted.map(c => ({ ...c })));
  };

  // --- BLANK FORMS HANDLER ---
  const handleBlankChange = (cat, type, val) => {
    setBlankCounts(prev => ({ ...prev, [`${cat}_${type}`]: val }));
  };

  const generateBlanks = async () => {
    setIsGenerating(true);
    try {
      await generateBlankPDFs(blankCounts, paperSize);
    } catch (err) {
      console.error(err);
      alert("Error generating blank forms. Check the browser console for details.");
    }
    setIsGenerating(false);
  };

  // --- GENERATION ACTIONS ---
  const generateByCategory = async () => {
    if (!district || !date) return alert("Please fill in District and Date");
    
    const activeJudges = judges.filter(j => j.Print);
    const activeComps = competitors.filter(c => c.Print);
    
    if (activeJudges.length === 0) return alert("Please select at least one Judge to print");
    if (activeComps.length === 0) return alert("Please select at least one Competitor to print");

    setIsGenerating(true);
    try {
      await generateCategoryPDFs(activeJudges, activeComps, { district, date, session }, paperSize);
    } catch (err) {
      console.error(err);
      alert("Error generating PDFs. Check the browser console for details.");
    }
    setIsGenerating(false);
  };

  const generateByJudge = async () => {
    if (!district || !date) return alert("Please fill in District and Date");
    
    const activeJudges = judges.filter(j => j.Print);
    const activeComps = competitors.filter(c => c.Print);
    
    if (activeJudges.length === 0) return alert("Please select at least one Judge to print");
    if (activeComps.length === 0) return alert("Please select at least one Competitor to print");

    setIsGenerating(true);
    try {
      await generateJudgePDFs(activeJudges, activeComps, { district, date, session }, paperSize);
    } catch (err) {
      console.error(err);
      alert("Error generating PDFs. Check the browser console for details.");
    }
    setIsGenerating(false);
  };

  const generateOverlays = () => {
    if (!district || !date) return alert("Please fill in District and Date");
    
    const activeJudges = judges.filter(j => j.Print);
    const activeComps = competitors.filter(c => c.Print);
    
    if (activeJudges.length === 0) return alert("Please select at least one Judge to print");
    if (activeComps.length === 0) return alert("Please select at least one Competitor to print");

    generateOverlaysRTF(activeJudges, activeComps, { district, date, session }, paperSize);
  };

  const generateLabels = () => {
    if (!district || !date) return alert("Please fill in District and Date");
    
    const activeJudges = judges.filter(j => j.Print);
    if (activeJudges.length === 0) return alert("Please select at least one Judge to print");

    generateFolderLabelsRTF(activeJudges, { district, date, session }, paperSize);
  };

  return (
    <div style={{ padding: '30px', fontFamily: 'system-ui, sans-serif', maxWidth: '1200px', margin: '0 auto', color: '#333' }}>
      <h1 style={{ borderBottom: '2px solid #eee', paddingBottom: '10px' }}>📝 Contest Form Generator</h1>
      
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr', gap: '15px', margin: '20px 0' }}>
        <input style={{ padding: '8px' }} placeholder="District" value={district} onChange={e => setDistrict(e.target.value)} />
        <input style={{ padding: '8px' }} type="date" value={date} onChange={e => setDate(e.target.value)} />
        <select style={{ padding: '8px' }} value={session} onChange={e => setSession(e.target.value)}>
          <option>Quartet Quarter-Finals</option>
          <option>Quartet Semi-Finals</option>
          <option>Chorus Finals</option>
          <option>Quartet Finals</option>
        </select>
        <select style={{ padding: '8px' }} value={paperSize} onChange={e => setPaperSize(e.target.value)}>
          <option value="Letter">Letter (8.5" x 11")</option>
          <option value="A4">A4 (210mm x 297mm)</option>
        </select>
      </div>

      <div style={{ display: 'flex', gap: '30px', marginTop: '30px', alignItems: 'flex-start' }}>
        
        {/* --- JUDGES PANEL --- */}
        <div style={{ flex: 1, background: '#f9f9f9', padding: '20px', borderRadius: '8px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h3 style={{ margin: 0 }}>🧑‍⚖️ Judges Upload</h3>
            <div style={{ display: 'flex', gap: '5px' }}>
              <button onClick={addJudge} style={{ padding: '4px 8px', cursor: 'pointer' }}>➕ Add</button>
              <button onClick={clearJudges} style={{ padding: '4px 8px', cursor: 'pointer', background: '#dc3545', color: 'white', border: 'none', borderRadius: '4px' }}>🗑️ Clear</button>
            </div>
          </div>
          
          <div style={{ marginBottom: '15px' }}>
            <input type="file" accept=".csv" onChange={handleJudgeUpload} />
            <div style={{ fontSize: '13px', color: '#555', marginTop: '4px' }}>
              <em>Note: Please upload the <strong>Assignments Report</strong> (.csv)</em>
            </div>
          </div>
          
          {judges.length > 0 && (
            <div style={{ marginBottom: '10px', paddingBottom: '10px', borderBottom: '1px solid #ddd' }}>
              <label style={{ fontWeight: 'bold', cursor: 'pointer' }}>
                <input 
                  type="checkbox" 
                  checked={judges.filter(j => !j.Name.startsWith("Absent")).every(j => j.Print)} 
                  onChange={(e) => toggleAllJudges(e.target.checked)} 
                  style={{ marginRight: '8px' }}
                />
                Select / Deselect All Judges
              </label>
            </div>
          )}

          <div style={{ maxHeight: '400px', overflowY: 'auto', fontSize: '14px', paddingRight: '5px' }}>
            {judges.map((j, i) => {
              const isAbsent = j.Name.startsWith("Absent");
              return (
                <div key={i} style={{ display: 'flex', gap: '5px', alignItems: 'center', marginBottom: '8px', opacity: isAbsent ? 0.6 : 1 }}>
                  <input 
                    type="checkbox" 
                    checked={j.Print} 
                    disabled={isAbsent}
                    onChange={e => updateJudge(i, 'Print', e.target.checked)} 
                    title={isAbsent ? "Absent judges cannot print" : "Print this judge?"}
                  />
                  <input 
                    type="number" 
                    value={j.Number} 
                    disabled={isAbsent}
                    onChange={e => updateJudge(i, 'Number', e.target.value)} 
                    style={{ width: '45px', padding: '4px' }}
                  />
                  <input 
                    type="text" 
                    value={j.Name} 
                    disabled={isAbsent}
                    onChange={e => updateJudge(i, 'Name', e.target.value)} 
                    style={{ flex: 1, padding: '4px' }}
                    placeholder="Judge Name"
                  />
                  <select value={j.Category} onChange={e => updateJudge(i, 'Category', e.target.value)} style={{ padding: '4px', width: '65px' }}>
                    <option>MUS</option><option>PER</option><option>SNG</option>
                  </select>
                  <select value={j.Type} onChange={e => updateJudge(i, 'Type', e.target.value)} style={{ padding: '4px', width: '85px' }}>
                    <option>Official</option><option>Practice</option>
                  </select>
                  <button onClick={() => removeJudge(i)} style={{ padding: '4px 8px', background: '#dc3545', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>🗑️</button>
                </div>
              );
            })}
          </div>
        </div>

        {/* --- COMPETITORS PANEL --- */}
        <div style={{ flex: 1, background: '#f9f9f9', padding: '20px', borderRadius: '8px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
            <h3 style={{ margin: 0 }}>🎤 Competitors</h3>
            <div style={{ display: 'flex', gap: '5px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
              <button onClick={clearAllOAs} style={{ padding: '4px 8px', cursor: 'pointer', background: '#ffc107', border: 'none', borderRadius: '4px' }} title="Clear all Order of Appearance numbers">🚫 Clear OAs</button>
              <button onClick={sortCompsByOA} style={{ padding: '4px 8px', cursor: 'pointer' }} title="Sort by Order of Appearance">🔄 Sort OA</button>
              <button onClick={addComp} style={{ padding: '4px 8px', cursor: 'pointer' }}>➕ Add</button>
              <button onClick={clearComps} style={{ padding: '4px 8px', cursor: 'pointer', background: '#dc3545', color: 'white', border: 'none', borderRadius: '4px' }}>🗑️ Clear</button>
            </div>
          </div>
          
          <div style={{ marginBottom: '15px' }}>
            <input type="file" accept=".csv" onChange={handleCompUpload} />
            <div style={{ fontSize: '13px', color: '#555', marginTop: '4px' }}>
              <em>Note: Please upload the <strong>DRCJ Report</strong> (.csv)</em>
            </div>
          </div>
          
          {competitors.length > 0 && (
            <div style={{ marginBottom: '10px', paddingBottom: '10px', borderBottom: '1px solid #ddd' }}>
              <label style={{ fontWeight: 'bold', cursor: 'pointer' }}>
                <input 
                  type="checkbox" 
                  checked={competitors.every(c => c.Print)} 
                  onChange={(e) => toggleAllComps(e.target.checked)} 
                  style={{ marginRight: '8px' }}
                />
                Select / Deselect All Competitors
              </label>
            </div>
          )}

          <div style={{ maxHeight: '400px', overflowY: 'auto', fontSize: '14px', paddingRight: '5px' }}>
            {competitors.map((c, i) => (
              <div key={i} style={{ display: 'flex', gap: '5px', alignItems: 'center', marginBottom: '8px' }}>
                <input 
                  type="checkbox" 
                  checked={c.Print} 
                  onChange={e => updateComp(i, 'Print', e.target.checked)} 
                  title="Print this competitor?"
                />
                <input 
                  type="number" 
                  value={c.Number} 
                  onChange={e => updateComp(i, 'Number', e.target.value)} 
                  placeholder="OA"
                  style={{ width: '50px', padding: '4px' }}
                  title="Order of Appearance"
                />
                <input 
                  type="text" 
                  value={c.Name} 
                  onChange={e => updateComp(i, 'Name', e.target.value)} 
                  style={{ flex: 1, padding: '4px' }}
                  placeholder="Competitor Name"
                />
                <input 
                  type="text" 
                  value={c.Director || ""} 
                  onChange={e => updateComp(i, 'Director', e.target.value)} 
                  style={{ flex: 1, padding: '4px' }}
                  placeholder={session.includes("Chorus") ? "Director Name" : "Members (comma separated)"}
                />
                <button onClick={() => removeComp(i)} style={{ padding: '4px 8px', background: '#dc3545', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>🗑️</button>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* --- BLANK FORMS SECTION --- */}
      <div style={{ marginTop: '30px', background: '#f9f9f9', padding: '20px', borderRadius: '8px' }}>
        <h3 style={{ margin: '0 0 15px 0' }}>📄 Print Blank Forms</h3>
        <table style={{ width: '100%', maxWidth: '600px', borderCollapse: 'collapse', textAlign: 'left' }}>
          <thead>
            <tr>
              <th style={{ padding: '8px', borderBottom: '2px solid #ddd' }}>Category</th>
              <th style={{ padding: '8px', borderBottom: '2px solid #ddd' }}>Long (Double-Sided)</th>
              <th style={{ padding: '8px', borderBottom: '2px solid #ddd' }}>Short (2 per page)</th>
            </tr>
          </thead>
          <tbody>
            {['MUS', 'PER', 'SNG'].map(cat => {
              const catName = cat === 'MUS' ? 'Musicality' : cat === 'PER' ? 'Performance' : 'Singing';
              return (
                <tr key={cat}>
                  <td style={{ padding: '8px', borderBottom: '1px solid #ddd', fontWeight: '500' }}>{catName}</td>
                  <td style={{ padding: '8px', borderBottom: '1px solid #ddd' }}>
                    <input 
                      type="number" 
                      min="0"
                      placeholder="0"
                      value={blankCounts[`${cat}_Long`]} 
                      onChange={(e) => handleBlankChange(cat, 'Long', e.target.value)}
                      style={{ width: '80px', padding: '6px' }}
                    />
                  </td>
                  <td style={{ padding: '8px', borderBottom: '1px solid #ddd' }}>
                    <input 
                      type="number" 
                      min="0"
                      placeholder="0"
                      value={blankCounts[`${cat}_Short`]} 
                      onChange={(e) => handleBlankChange(cat, 'Short', e.target.value)}
                      style={{ width: '80px', padding: '6px' }}
                    />
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
        <button 
          onClick={generateBlanks} 
          disabled={isGenerating || !Object.values(blankCounts).some(v => parseInt(v) > 0)}
          style={{ marginTop: '15px', padding: '12px 24px', background: '#6c757d', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '16px' }}
        >
          {isGenerating ? "⏳ Generating..." : "🖨️ Download Blank Forms"}
        </button>
      </div>

      <div style={{ marginTop: '30px', display: 'flex', gap: '15px', flexWrap: 'wrap' }}>
        <button onClick={generateByCategory} disabled={isGenerating || judges.length === 0} style={{ padding: '12px 24px', background: '#0066cc', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '16px' }}>
          {isGenerating ? "⏳ Generating..." : "📥 Generate PDFs by Category"}
        </button>
        <button onClick={generateByJudge} disabled={isGenerating || judges.length === 0} style={{ padding: '12px 24px', background: '#17a2b8', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '16px' }}>
          {isGenerating ? "⏳ Generating..." : "📥 Generate PDFs by Judge"}
        </button>
        <button onClick={generateOverlays} disabled={judges.length === 0 || competitors.length === 0} style={{ padding: '12px 24px', background: '#6f42c1', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '16px' }}>
          📄 Generate Overlays (RTF)
        </button>
        <button onClick={generateLabels} disabled={judges.length === 0} style={{ padding: '12px 24px', background: '#28a745', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '16px' }}>
          🏷️ Generate Folder Labels (RTF)
        </button>
      </div>
    </div>
  );
}
