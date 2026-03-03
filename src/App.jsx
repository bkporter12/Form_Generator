import React, { useState } from 'react';
import Papa from 'papaparse';
import { balanceAndSortJudges, generateCategoryPDFs, generateJudgePDFs, generateFolderLabelsRTF, generateOverlaysRTF } from './utils';

export default function App() {
  const [district, setDistrict] = useState("");
  const [date, setDate] = useState("");
  const [session, setSession] = useState("Quartet Semi-Finals");
  const [paperSize, setPaperSize] = useState("Letter");
  
  const [judges, setJudges] = useState([]);
  const [competitors, setCompetitors] = useState([]);
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

  // --- GENERATION ACTIONS ---
  const generateByCategory = async () => {
    if (!district || !date) return alert("Please fill in District and Date");
    setIsGenerating(true);
    try {
      await generateCategoryPDFs(judges.filter(j => j.Print), competitors.filter(c => c.Print), { district, date, session }, paperSize);
    } catch (err) {
      console.error(err);
      alert("Error generating PDFs. Check the browser console for details.");
    }
    setIsGenerating(false);
  };

  const generateByJudge = async () => {
    if (!district || !date) return alert("Please fill in District and Date");
    setIsGenerating(true);
    try {
      await generateJudgePDFs(judges.filter(j => j.Print), competitors.filter(c => c.Print), { district, date, session }, paperSize);
    } catch (err) {
      console.error(err);
      alert("Error generating PDFs. Check the browser console for details.");
    }
    setIsGenerating(false);
  };

  const generateLabels = () => {
    if (!district || !date) return alert("Please fill in District and Date");
    generateFolderLabelsRTF(judges.filter(j => j.Print), { district, date, session }, paperSize);
  };

  const generateOverlays = () => {
    if (!district || !date) return alert("Please fill in District and Date");
    generateOverlaysRTF(judges.filter(j => j.Print), competitors.filter(c => c.Print), { district, date, session }, paperSize);
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
            <div style={{ display: 'flex', gap: '5px' }}>
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
                {session.includes("Chorus") && (
                  <input 
                    type="text" 
                    value={c.Director || ""} 
                    onChange={e => updateComp(i, 'Director', e.target.value)} 
                    style={{ flex: 1, padding: '4px' }}
                    placeholder="Director Name"
                  />
                )}
                <button onClick={() => removeComp(i)} style={{ padding: '4px 8px', background: '#dc3545', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>🗑️</button>
              </div>
            ))}
          </div>
        </div>
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
