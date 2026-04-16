import React, { useState, useEffect } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';
import { Share2, Download, Plus, HardDrive } from 'lucide-react';

// ✅ Updated with your official Client ID
const CLIENT_ID = "683400126186-f3a9u3fbe6l50bv1vidci7oinq7socn6.apps.googleusercontent.com";
const SCOPES = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

export default function MobileResponsiveDPR() {
  const [project, setProject] = useState("PROJECT");
  const [isAdmin, setIsAdmin] = useState(true);
  const [isUploading, setIsUploading] = useState(false);
  const [accessToken, setAccessToken] = useState(() => sessionStorage.getItem('drive_token'));

  const [dprData, setDprData] = useState(() => {
    const saved = localStorage.getItem('dpr_v2_mobile');
    return saved ? JSON.parse(saved) : {
      date: new Date().toISOString().split('T')[0],
      labor: [
        { trade: 'Civil', skilled: 0, unskilled: 0 },
        { trade: 'Electrical', skilled: 0, unskilled: 0 },
        { trade: 'Plumbing', skilled: 0, unskilled: 0 }
      ],
      progress: [{ id: 1, desc: '', photos: [] }]
    };
  });

  useEffect(() => {
    localStorage.setItem('dpr_v2_mobile', JSON.stringify(dprData));
    if (accessToken) sessionStorage.setItem('drive_token', accessToken);
  }, [dprData, accessToken]);

  // ✅ Account Picker: Shows all logged-in accounts to choose from
  const handleLogin = () => {
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      prompt: 'select_account', 
      callback: (res) => {
        if (res.access_token) {
          setAccessToken(res.access_token);
          sessionStorage.setItem('drive_token', res.access_token);
          alert("Google Account Connected Successfully!");
        }
      },
    });
    client.requestAccessToken();
  };

  // ✅ Auto Link Copy: Copies current URL to clipboard
  const handleShare = () => {
    const shareLink = window.location.href;
    navigator.clipboard.writeText(shareLink).then(() => {
      alert("Project Link Copied! Now you can paste it in WhatsApp.");
    });
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
        
        try {
          const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
            method: 'POST', headers: { 'Authorization': `Bearer ${accessToken}` }, body: form
          });
          const data = await res.json();
          setDprData(prev => ({
            ...prev,
            progress: prev.progress.map(p => p.id === rowId ? { ...p, photos: [...p.photos, { drive: data.webViewLink, preview: compressed }] } : p)
          }));
        } catch (e) { console.error("Upload Error", e); }
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
        const max = 600;
        let w = img.width, h = img.height;
        if (w > h) { if (w > max) { h *= max / w; w = max; } }
        else { if (h > max) { w *= max / h; h = max; } }
        canvas.width = w; canvas.height = h;
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', 0.6));
      };
    });
  };

  const generateExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Daily Progress');
    sheet.columns = [{ width: 15 }, { width: 15 }, { width: 15 }, { width: 35 }];

    sheet.addRow([`${project} - DAILY PROGRESS REPORT`]).font = { bold: true, size: 12 };
    sheet.addRow(['Date:', dprData.date]);
    sheet.addRow([]);
    sheet.addRow(['Trade', 'Skilled', 'Unskilled', 'Total']).font = { bold: true };
    
    dprData.labor.forEach(l => {
      sheet.addRow([l.trade, l.skilled, l.unskilled, l.skilled + l.unskilled]);
    });

    sheet.addRow([]);
    sheet.addRow(['Work Progress Details']).font = { bold: true };

    for (let i = 0; i < dprData.progress.length; i++) {
      const p = dprData.progress[i];
      const row = sheet.addRow([p.desc]);
      row.height = 95;
      row.alignment = { vertical: 'middle', wrapText: true };
      
      if (p.photos.length > 0) {
        const imgId = workbook.addImage({ base64: p.photos[0].preview, extension: 'jpeg' });
        sheet.addImage(imgId, {
          tl: { col: 1, row: row.number - 1 },
          ext: { width: 110, height: 110 },
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
        <div style={ui.flexBet}>
          <h2 style={ui.brand}>DPR TOOL</h2>
          <button onClick={() => setIsAdmin(!isAdmin)} style={ui.badge}>
            {isAdmin ? "ADMIN" : "SITE"} MODE
          </button>
        </div>
        <input value={project} onChange={e => setProject(e.target.value)} style={ui.inputDark} disabled={!isAdmin} />
        {!accessToken && (
          <button onClick={handleLogin} style={ui.btnDrive}>
            <HardDrive size={16} /> CONNECT DRIVE
          </button>
        )}
      </header>

      <main style={ui.main}>
        {/* Labor Grid - Finger Touch Optimized */}
        <section style={ui.card}>
          <h3 style={ui.sectionTitle}>LABOR STRENGTH</h3>
          {dprData.labor.map((l, idx) => (
            <div key={idx} style={ui.mobileRow}>
              <div style={{fontWeight:'bold', width:'75px', fontSize:'14px'}}>{l.trade}</div>
              <input type="number" placeholder="Skilled" style={ui.miniField} 
                onChange={e => {
                  const newL = [...dprData.labor];
                  newL[idx].skilled = parseInt(e.target.value) || 0;
                  setDprData({...dprData, labor: newL});
                }} />
              <input type="number" placeholder="Unskilled" style={ui.miniField}
                onChange={e => {
                  const newL = [...dprData.labor];
                  newL[idx].unskilled = parseInt(e.target.value) || 0;
                  setDprData({...dprData, labor: newL});
                }} />
            </div>
          ))}
        </section>

        {/* Work Progress Details */}
        <section style={ui.card}>
          <h3 style={ui.sectionTitle}>SITE ACTIVITIES</h3>
          {dprData.progress.map((p, idx) => (
            <div key={p.id} style={ui.progressItem}>
              <textarea placeholder="Work progress description..." style={ui.area} value={p.desc}
                onChange={e => {
                  const n = [...dprData.progress];
                  n[idx].desc = e.target.value;
                  setDprData({...dprData, progress: n});
                }} />
              <label style={ui.upBtn}>
                📸 ATTACH PHOTOS ({p.photos.length})
                <input type="file" multiple hidden onChange={e => uploadToDrive(p.id, e.target.files)} />
              </label>
              <div style={ui.thumbRow}>
                {p.photos.map((img, i) => <img key={i} src={img.preview} style={ui.thumb} />)}
              </div>
            </div>
          ))}
          <button onClick={() => setDprData({...dprData, progress: [...dprData.progress, {id: Date.now(), desc: '', photos: []}]})} style={ui.btnAdd}>
            <Plus size={16} /> NEW ACTIVITY
          </button>
        </section>
      </main>

      <footer style={ui.footer}>
        <button onClick={generateExcel} style={ui.btnExcel}>
          <Download size={18} /> GET EXCEL
        </button>
        {isAdmin && (
          <button onClick={handleShare} style={ui.btnShare}>
            <Share2 size={18} /> SHARE LINK
          </button>
        )}
      </footer>
    </div>
  );
}

const ui = {
  container: { background: '#f4f6f9', minHeight: '100vh', paddingBottom: '90px', fontFamily: 'Arial, sans-serif' },
  nav: { background: '#1a1c1e', padding: '15px', position: 'sticky', top: 0, zIndex: 10, display: 'flex', flexDirection: 'column', gap: '12px' },
  flexBet: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' },
  brand: { color: '#fff', fontSize: '18px', margin: 0, fontWeight: 'bold' },
  badge: { background: '#ffd700', border: 'none', padding: '5px 12px', borderRadius: '20px', fontSize: '11px', fontWeight: 'bold' },
  inputDark: { background: '#2c2e30', border: '1px solid #444', color: '#fff', padding: '12px', borderRadius: '8px', textAlign: 'center', fontSize: '15px' },
  btnDrive: { background: '#4285F4', color: '#fff', border: 'none', padding: '14px', borderRadius: '8px', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' },
  main: { padding: '12px' },
  card: { background: '#fff', borderRadius: '12px', padding: '15px', marginBottom: '15px', boxShadow: '0 2px 8px rgba(0,0,0,0.06)' },
  sectionTitle: { fontSize: '12px', color: '#888', marginBottom: '15px', fontWeight: 'bold', borderBottom: '1px solid #f0f0f0', paddingBottom: '8px' },
  mobileRow: { display: 'flex', gap: '10px', alignItems: 'center', marginBottom: '12px' },
  miniField: { flex: 1, padding: '14px', border: '1px solid #ddd', borderRadius: '10px', textAlign: 'center', fontSize: '15px' },
  progressItem: { marginBottom: '20px', borderBottom: '1px solid #eee', paddingBottom: '15px' },
  area: { width: '100%', height: '85px', padding: '12px', borderRadius: '10px', border: '1px solid #ddd', fontSize: '14px', marginBottom: '10px', boxSizing: 'border-box' },
  upBtn: { background: '#ebf5ff', color: '#007bff', padding: '14px', textAlign: 'center', borderRadius: '10px', border: '1px dashed #007bff', fontWeight: 'bold', fontSize: '13px', display: 'block' },
  thumbRow: { display: 'flex', gap: '8px', marginTop: '10px', flexWrap: 'wrap' },
  thumb: { width: '55px', height: '55px', objectFit: 'cover', borderRadius: '8px' },
  btnAdd: { width: '100%', background: '#fff', border: '1px solid #007bff', color: '#007bff', padding: '12px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px', fontSize: '13px', fontWeight: 'bold' },
  footer: { position: 'fixed', bottom: 0, width: '100%', background: '#fff', padding: '15px', display: 'flex', gap: '12px', boxShadow: '0 -5px 15px rgba(0,0,0,0.05)', boxSizing: 'border-box' },
  btnExcel: { flex: 1, background: '#217346', color: '#fff', border: 'none', padding: '16px', borderRadius: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' },
  btnShare: { flex: 1, background: '#007bff', color: '#fff', border: 'none', padding: '16px', borderRadius: '12px', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }
};
