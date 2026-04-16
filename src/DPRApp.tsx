import React, { useState, useEffect } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';

// IMPORTANT: Unga Friend-oda Client ID-ya inga podunga
const CLIENT_ID = "683400126186-f3a9u3fbe6l50bv1vidci7oinq7socn6.apps.googleusercontent.com";
const SCOPES = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

export default function ProfessionalDPRSystem() {
  const [project, setProject] = useState("FLORA VILLA-75E");
  const [isAdmin, setIsAdmin] = useState(true); // Toggle for Testing
  const [isUploading, setIsUploading] = useState(false);
  const [accessToken, setAccessToken] = useState(() => sessionStorage.getItem('drive_token'));

  // DPR State: Labor, Work Progress, and Materials
  const [dprData, setDprData] = useState(() => {
    const saved = localStorage.getItem('dpr_v1_data');
    return saved ? JSON.parse(saved) : {
      date: new Date().toISOString().split('T')[0],
      labor: [
        { trade: 'Civil', skilled: 0, unskilled: 0 },
        { trade: 'Electrical', skilled: 0, unskilled: 0 },
        { trade: 'Plumbing', skilled: 0, unskilled: 0 }
      ],
      progress: [{ id: 1, desc: '', photos: [] }],
      materials: [{ item: 'Cement', used: 0, balance: 0 }]
    };
  });

  useEffect(() => {
    localStorage.setItem('dpr_v1_data', JSON.stringify(dprData));
    if (accessToken) sessionStorage.setItem('drive_token', accessToken);
    
    const gsiScript = document.createElement('script');
    gsiScript.src = "https://accounts.google.com/gsi/client";
    gsiScript.async = true; 
    document.body.appendChild(gsiScript);
  }, [dprData, accessToken]);

  const handleLogin = () => {
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID, scope: SCOPES,
      callback: (res) => { if (res.access_token) setAccessToken(res.access_token); },
    });
    client.requestAccessToken();
  };

  const uploadToDrive = async (rowId, files) => {
    if (!accessToken) return alert("Connect Drive First!");
    setIsUploading(true);
    for (const file of Array.from(files)) {
      const reader = new FileReader();
      reader.onloadend = async () => {
        const compressed = await compressImage(reader.result);
        const metadata = { name: `DPR_${project}_${Date.now()}.jpg`, mimeType: 'image/jpeg' };
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
        form.append('file', file);
        
        const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
          method: 'POST', headers: { 'Authorization': `Bearer ${accessToken}` }, body: form
        });
        const data = await res.json();
        setDprData(prev => ({
          ...prev,
          progress: prev.progress.map(p => p.id === rowId ? { ...p, photos: [...p.photos, { drive: data.webViewLink, preview: compressed }] } : p)
        }));
      };
      reader.readAsDataURL(file);
    }
    setIsUploading(false);
  };

  const compressImage = (base64) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.src = base64;
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const scale = 400 / img.width;
        canvas.width = 400; canvas.height = img.height * scale;
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', 0.7));
      };
    });
  };

  const generateExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Daily Progress Report');

    // Header Style
    sheet.mergeCells('A1:H1');
    const title = sheet.getCell('A1');
    title.value = `${project} - DAILY PROGRESS REPORT (${dprData.date})`;
    title.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1A1C1E' } };
    title.alignment = { horizontal: 'center' };

    // Labor Table
    sheet.addRow(['Labor Category', 'Skilled', 'Unskilled', 'Total']);
    dprData.labor.forEach(l => sheet.addRow([l.trade, l.skilled, l.unskilled, l.skilled + l.unskilled]));

    // Work Progress with Images (Replica Logic)
    sheet.addRow([]);
    sheet.addRow(['Work Description', 'Photos']);
    
    for (let i = 0; i < dprData.progress.length; i++) {
      const p = dprData.progress[i];
      const row = sheet.addRow([p.desc]);
      row.height = 80;
      
      if (p.photos.length > 0) {
        const imgId = workbook.addImage({ base64: p.photos[0].preview, extension: 'jpeg' });
        sheet.addImage(imgId, {
          tl: { col: 1, row: row.number - 1 },
          br: { col: 2, row: row.number },
          editAs: 'oneCell'
        });
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([buffer]));
    link.download = `DPR_${project}_${dprData.date}.xlsx`;
    link.click();
  };

  return (
    <div style={ui.container}>
      <header style={ui.nav}>
        <div style={{display:'flex', justifyContent:'space-between', alignItems:'center'}}>
          <h2 style={{color:'#fff', margin:0}}>DPR SYSTEM</h2>
          <button onClick={() => setIsAdmin(!isAdmin)} style={ui.toggleBtn}>
            {isAdmin ? "ADMIN MODE" : "SITE ENGINEER MODE"}
          </button>
        </div>
        <input value={project} onChange={e => setProject(e.target.value)} style={ui.headIn} disabled={!isAdmin} />
        {!accessToken && <button onClick={handleLogin} style={ui.authBtn}>Connect Google Drive</button>}
      </header>

      <main style={ui.main}>
        {/* Labor Strength Section */}
        <section style={ui.card}>
          <h3 style={ui.sectionTitle}>1. Labor Strength</h3>
          {dprData.labor.map((l, idx) => (
            <div key={idx} style={ui.inputGrid}>
              <span>{l.trade}</span>
              <input type="number" placeholder="Skilled" style={ui.field} 
                onChange={e => {
                  const newLabor = [...dprData.labor];
                  newLabor[idx].skilled = parseInt(e.target.value) || 0;
                  setDprData({...dprData, labor: newLabor});
                }} />
              <input type="number" placeholder="Unskilled" style={ui.field}
                onChange={e => {
                  const newLabor = [...dprData.labor];
                  newLabor[idx].unskilled = parseInt(e.target.value) || 0;
                  setDprData({...dprData, labor: newLabor});
                }} />
            </div>
          ))}
        </section>

        {/* Progress Section */}
        <section style={ui.card}>
          <h3 style={ui.sectionTitle}>2. Work Progress & Photos</h3>
          {dprData.progress.map((p, idx) => (
            <div key={p.id} style={{marginBottom:'15px', borderBottom:'1px solid #eee', paddingBottom:'10px'}}>
              <textarea placeholder="Work Description..." style={ui.area} value={p.desc}
                onChange={e => {
                  const newProg = [...dprData.progress];
                  newProg[idx].desc = e.target.value;
                  setDprData({...dprData, progress: newProg});
                }} />
              <label style={ui.upBtn}>
                📸 Upload Site Photos ({p.photos.length})
                <input type="file" multiple hidden onChange={e => uploadToDrive(p.id, e.target.files)} />
              </label>
              <div style={ui.thumbRow}>
                {p.photos.map((img, i) => <img key={i} src={img.preview} style={ui.thumb} />)}
              </div>
            </div>
          ))}
          <button onClick={() => setDprData({...dprData, progress: [...dprData.progress, {id: Date.now(), desc: '', photos: []}]})} style={ui.btnSmall}>+ Add Activity</button>
        </section>
      </main>

      <footer style={ui.footer}>
        <button onClick={generateExcel} style={ui.btnExcel}>GENERATE EXCEL (DPR)</button>
        {isAdmin && <button onClick={() => alert(`Share this link: ${window.location.href}?id=FLORA75E`)} style={ui.btnBlue}>SHARE LINK TO SITE</button>}
      </footer>
    </div>
  );
}

const ui = {
  container: { background: '#f0f2f5', minHeight: '100vh', paddingBottom: '100px', fontFamily: 'Segoe UI' },
  nav: { background: '#1a1c1e', padding: '15px', position: 'sticky', top: 0, zIndex: 10, display: 'flex', flexDirection: 'column', gap: '10px' },
  headIn: { background: 'transparent', border: '1px solid #555', color: '#fff', textAlign: 'center', padding: '8px', borderRadius: '4px', fontSize: '16px' },
  authBtn: { background: '#4285F4', color: '#fff', border: 'none', padding: '10px', borderRadius: '4px', fontWeight: 'bold', cursor: 'pointer' },
  toggleBtn: { background: '#ffd700', color: '#000', border: 'none', padding: '5px 10px', borderRadius: '20px', fontSize: '10px', fontWeight: 'bold' },
  main: { padding: '15px' },
  card: { background: '#fff', borderRadius: '8px', padding: '15px', marginBottom: '15px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' },
  sectionTitle: { fontSize: '14px', color: '#333', borderBottom: '2px solid #007bff', display: 'inline-block', marginBottom: '15px' },
  inputGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '10px', alignItems: 'center', marginBottom: '10px' },
  field: { padding: '8px', border: '1px solid #ddd', borderRadius: '4px' },
  area: { width: '100%', height: '60px', padding: '10px', borderRadius: '4px', border: '1px solid #ddd', marginBottom: '10px' },
  upBtn: { display: 'block', background: '#f8f9fa', border: '1px dashed #007bff', color: '#007bff', padding: '10px', textAlign: 'center', borderRadius: '4px', cursor: 'pointer' },
  thumbRow: { display: 'flex', gap: '5px', marginTop: '10px', flexWrap: 'wrap' },
  thumb: { width: '50px', height: '50px', objectFit: 'cover', borderRadius: '4px' },
  footer: { position: 'fixed', bottom: 0, width: '100%', background: '#fff', padding: '15px', display: 'flex', gap: '10px', boxShadow: '0 -2px 10px rgba(0,0,0,0.1)' },
  btnExcel: { flex: 2, background: '#217346', color: '#fff', border: 'none', padding: '15px', borderRadius: '8px', fontWeight: 'bold' },
  btnBlue: { flex: 1, background: '#007bff', color: '#fff', border: 'none', padding: '15px', borderRadius: '8px', fontWeight: 'bold' },
  btnSmall: { background: '#eee', border: 'none', padding: '5px 10px', borderRadius: '4px', fontSize: '12px' }
};
