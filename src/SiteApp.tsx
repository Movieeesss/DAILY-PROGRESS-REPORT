import React, { useState, useEffect } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';

const CLIENT_ID = "683400126186-f3a9u3fbe6l50bv1vidci7oinq7socn6.apps.googleusercontent.com";
const SCOPES = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

export default function FinalTenColumnRegister() {
  const [project, setProject] = useState("FLORA VILLA-75E");
  const [isUploading, setIsUploading] = useState(false);
  const [accessToken, setAccessToken] = useState(() => sessionStorage.getItem('drive_token'));

  const [rows, setRows] = useState(() => {
    const saved = localStorage.getItem('site_v20_final');
    return saved ? JSON.parse(saved) : [{ 
      id: 1, date: '2026-04-13', desc: '', loggedBy: '', 
      status: 'Open', closedDate: '', before: [], after: [], 
      owner: '', closedBy: '', remarks: '' 
    }];
  });

  useEffect(() => {
    const gsiScript = document.createElement('script');
    gsiScript.src = "https://accounts.google.com/gsi/client";
    gsiScript.async = true; gsiScript.defer = true;
    document.body.appendChild(gsiScript);
    localStorage.setItem('site_v20_final', JSON.stringify(rows));
    if (accessToken) sessionStorage.setItem('drive_token', accessToken);
    return () => { if (document.body.contains(gsiScript)) document.body.removeChild(gsiScript); };
  }, [rows, accessToken]);

  const handleLogin = () => {
    if (!window.google) return alert("Google Script ready aagala. Refresh!");
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID, scope: SCOPES,
      callback: (res) => { 
        if (res.access_token) {
          setAccessToken(res.access_token);
          sessionStorage.setItem('drive_token', res.access_token);
        }
      },
    });
    client.requestAccessToken();
  };

  const handleDisconnect = () => { setAccessToken(null); sessionStorage.removeItem('drive_token'); alert("Disconnected!"); };
  const formatDate = (d) => d ? d.split("-").reverse().join("-") : "";

  const uploadPhoto = async (id, type, files) => {
    if (!accessToken) return alert("Connect Drive First!");
    setIsUploading(true);
    for (const file of Array.from(files)) {
      const reader = new FileReader();
      reader.onloadend = async () => {
        const base64data = reader.result;
        const img = new Image();
        img.src = base64data;
        img.onload = async () => {
          const canvas = document.createElement('canvas');
          const ctx = canvas.getContext('2d');
          const maxWidth = 400;
          const scale = maxWidth / img.width;
          canvas.width = maxWidth; canvas.height = img.height * scale;
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          const compressed = canvas.toDataURL('image/jpeg', 0.7);
          const metadata = { name: `${project}_${type}_${Date.now()}.jpg`, mimeType: 'image/jpeg' };
          const form = new FormData();
          form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
          form.append('file', file);
          try {
            const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
              method: 'POST', headers: { 'Authorization': `Bearer ${accessToken}` }, body: form
            });
            const data = await res.json();
            setRows(prev => prev.map(r => r.id === id ? { ...r, [type]: [...r[type], { drive: data.webViewLink, preview: compressed }] } : r));
          } catch (e) { console.error(e); }
        };
      };
      reader.readAsDataURL(file);
    }
    setIsUploading(false);
  };

  const deletePhoto = (rowId, type, idx) => {
    setRows(prev => prev.map(r => r.id === rowId ? { ...r, [type]: r[type].filter((_, i) => i !== idx) } : r));
  };

  // ✅ UPDATED: Excel Export with Perfect Image Fitting & PDF Replica Logic
  const generateExcel = async () => {
    if (rows.length === 0) return alert("Data illai!");
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Action Register');

    // PDF replica logic: find max photos to create dynamic columns
    const maxBefore = Math.max(...rows.map(r => r.before.length), 1);
    const maxAfter = Math.max(...rows.map(r => r.after.length), 1);

    // 1. Setup Columns
    let columns = [
      { header: 'Slno', key: 'slno', width: 8 },
      { header: 'Date Logged', key: 'date', width: 15 },
      { header: 'Description', key: 'desc', width: 40 },
      { header: 'Logged By', key: 'loggedBy', width: 15 },
      { header: 'Status', key: 'status', width: 15 },
      { header: 'Actual Closed Date', key: 'closedDate', width: 18 },
    ];

    // Add dynamic Photo Before columns
    for (let i = 0; i < maxBefore; i++) {
      columns.push({ header: `P${i + 1}(B)`, key: `pB${i}`, width: 25 });
    }
    // Add dynamic Photo After columns
    for (let i = 0; i < maxAfter; i++) {
      columns.push({ header: `P${i + 1}(A)`, key: `pA${i}`, width: 25 });
    }

    columns.push(
      { header: 'Owner', key: 'owner', width: 15 },
      { header: 'Closed By', key: 'closedBy', width: 15 },
      { header: 'Remarks', key: 'remarks', width: 30 }
    );

    worksheet.columns = columns;

    // 2. Style Header
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '282C34' } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

    // 3. Add Rows
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rowIndex = i + 2;

      const rowData = {
        slno: i + 1,
        date: formatDate(r.date),
        desc: r.desc,
        loggedBy: r.loggedBy,
        status: r.status.toUpperCase(),
        closedDate: formatDate(r.closedDate),
        owner: r.owner,
        closedBy: r.closedBy,
        remarks: r.remarks.toUpperCase()
      };

      const row = worksheet.addRow(rowData);
      row.height = 100; // Increased height for better image fit
      row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

      // Status Styling
      const statusCell = worksheet.getCell(`E${rowIndex}`);
      statusCell.fill = {
        type: 'pattern', pattern: 'solid',
        fgColor: { argb: r.status.toUpperCase() === 'OPEN' ? 'FFFF0000' : 'FF92D050' }
      };
      statusCell.font = { color: { argb: 'FFFFFF' }, bold: true };

      // Photo Insertion Logic (Matching PDF Replica)
      const beforeStartCol = 7; // Column G
      const afterStartCol = beforeStartCol + maxBefore;

      // Insert Before Photos
      for (let j = 0; j < r.before.length; j++) {
        try {
          const imageId = workbook.addImage({
            base64: r.before[j].preview,
            extension: 'jpeg',
          });
          worksheet.addImage(imageId, {
            tl: { col: beforeStartCol + j - 1, row: rowIndex - 1 },
            br: { col: beforeStartCol + j, row: rowIndex },
            editAs: 'oneCell' // This forces image to stay within cell boundaries
          });
        } catch (e) { console.error(e); }
      }

      // Insert After Photos
      for (let k = 0; k < r.after.length; k++) {
        try {
          const imageId = workbook.addImage({
            base64: r.after[k].preview,
            extension: 'jpeg',
          });
          worksheet.addImage(imageId, {
            tl: { col: afterStartCol + k - 1, row: rowIndex - 1 },
            br: { col: afterStartCol + k, row: rowIndex },
            editAs: 'oneCell'
          });
        } catch (e) { console.error(e); }
      }
    }

    // Border for all cells
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' },
          bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${project}_Register.xlsx`;
    link.click();
  };

  const generatePDF = async () => {
    if (rows.length === 0) return;
    const batchSize = 5; 
    const totalBatches = Math.ceil(rows.length / batchSize);
    alert(`Total ${totalBatches} PDF download aagum...`);

    for (let i = 0; i < totalBatches; i++) {
      const start = i * batchSize;
      const batchRows = rows.slice(start, start + batchSize);
      const doc = new jsPDF('l', 'mm', 'a4');
      autoTable(doc, {
        body: [[project.toUpperCase(), 'ACTION REGISTER - NEW']],
        theme: 'grid',
        styles: { fontSize: 10, fontStyle: 'bold', halign: 'left', fillColor: [211, 211, 211] }
      });
      const maxBefore = Math.max(...batchRows.map(r => r.before.length), 1);
      const maxAfter = Math.max(...batchRows.map(r => r.after.length), 1);
      const head = [['Slno', 'Date Logged', 'Description', 'Logged by', 'Current Status', 'Actual Closed Date', ...Array.from({ length: maxBefore }, (_, idx) => `P${idx+1}(B)`), ...Array.from({ length: maxAfter }, (_, idx) => `P${idx+1}(A)`), 'Owner', 'Closed By', 'Remarks']];
      const body = batchRows.map((r, idx) => [start + idx + 1, formatDate(r.date), r.desc, r.loggedBy, r.status.toUpperCase(), formatDate(r.closedDate), ...Array(maxBefore + maxAfter).fill(''), r.owner, r.closedBy, r.remarks.toUpperCase()]);

      autoTable(doc, {
        startY: doc.lastAutoTable.finalY, head: head, body: body, theme: 'grid',
        styles: { fontSize: 4.5, halign: 'center', valign: 'middle', minCellHeight: 18 },
        headStyles: { fillColor: [40, 44, 52], textColor: [255, 255, 255] },
        didDrawCell: (data) => {
          if (data.section === 'head') return;
          const rowData = batchRows[data.row.index];
          if (data.column.index === 4) data.cell.styles.fillColor = rowData.status.toUpperCase() === 'OPEN' ? [255, 0, 0] : [146, 208, 80];
          const drawImg = (imgObj, cell) => { if (imgObj?.preview) doc.addImage(imgObj.preview, 'JPEG', cell.x + 0.5, cell.y + 0.5, cell.width - 1, cell.height - 1); };
          const photoStart = 6;
          if (data.column.index >= photoStart && data.column.index < photoStart + maxBefore) drawImg(rowData.before[data.column.index - photoStart], data.cell);
          if (data.column.index >= photoStart + maxBefore && data.column.index < photoStart + maxBefore + maxAfter) drawImg(rowData.after[data.column.index - (photoStart + maxBefore)], data.cell);
        }
      });
      doc.save(`${project}_Part_${i+1}.pdf`);
      await new Promise(res => setTimeout(res, 500));
    }
  };

  const addRow = () => {
    const newId = rows.length > 0 ? Math.max(...rows.map(r => r.id)) + 1 : 1;
    setRows([...rows, { id: newId, date: '2026-04-14', desc: '', loggedBy: '', status: 'Open', closedDate: '', before: [], after: [], owner: '', closedBy: '', remarks: '' }]);
  };

  const deleteRow = (id) => { if (window.confirm("Delete?")) setRows(rows.filter(r => r.id !== id)); };
  const clearAll = () => { if (window.confirm("Clear All?")) { setRows([{ id: Date.now(), date: '2026-04-14', desc: '', loggedBy: '', status: 'Open', closedDate: '', before: [], after: [], owner: '', closedBy: '', remarks: '' }]); localStorage.removeItem('site_v20_final'); } };

  return (
    <div style={ui.container}>
      <header style={ui.nav}>
        {!accessToken ? <button onClick={handleLogin} style={ui.authBtn}>Connect Drive</button> : <button onClick={handleDisconnect} style={ui.disBtn}>Disconnect Drive (✓)</button>}
        <input value={project} onChange={e => setProject(e.target.value)} style={ui.headIn} />
      </header>
      <div style={ui.main}>
        {rows.map((row, index) => (
          <div key={row.id} style={ui.card}>
            <div style={ui.topRow}>
              <div style={{display:'flex', gap:'10px', alignItems:'center'}}>
                <span style={ui.snoBadge}>Slno: {index + 1}</span>
                <button onClick={() => deleteRow(row.id)} style={ui.rowDelBtn}>Delete</button>
              </div>
              <input style={{...ui.badge, backgroundColor: row.status.toUpperCase()==='OPEN'?'#ff0000':'#92d050'}} value={row.status} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, status: e.target.value}:r))} />
            </div>
            <div style={ui.inputGrid}>
              <input type="date" style={ui.field} value={row.date} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, date: e.target.value}:r))} />
              <input placeholder="Logged By" style={ui.field} value={row.loggedBy} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, loggedBy: e.target.value}:r))} />
            </div>
            <textarea placeholder="Description" style={ui.area} value={row.desc} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, desc: e.target.value}:r))} />
            <div style={ui.inputGrid}>
               <div style={{fontSize:'10px'}}>Actual Closed Date: <input type="date" style={ui.field} value={row.closedDate} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, closedDate: e.target.value}:r))} /></div>
               <input placeholder="Owner" style={ui.field} value={row.owner} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, owner: e.target.value}:r))} />
            </div>
            <input placeholder="Closed By" style={{...ui.field, marginBottom:'10px'}} value={row.closedBy} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, closedBy: e.target.value}:r))} />
            <div style={ui.photoGrid}>
              <div style={ui.uploadSection}>
                <label style={ui.upBtn}>📸 Before ({row.before.length}) <input type="file" multiple hidden onChange={e => uploadPhoto(row.id, 'before', e.target.files)} /></label>
                <div style={ui.thumbnailRow}>{row.before.map((img, idx) => (<div key={idx} style={ui.thumbWrap}><img src={img.preview} style={ui.thumbImg} /><button onClick={() => deletePhoto(row.id, 'before', idx)} style={ui.delBtn}>×</button></div>))}</div>
              </div>
              <div style={ui.uploadSection}>
                <label style={ui.upBtn}>📸 After ({row.after.length}) <input type="file" multiple hidden onChange={e => uploadPhoto(row.id, 'after', e.target.files)} /></label>
                <div style={ui.thumbnailRow}>{row.after.map((img, idx) => (<div key={idx} style={ui.thumbWrap}><img src={img.preview} style={ui.thumbImg} /><button onClick={() => deletePhoto(row.id, 'after', idx)} style={ui.delBtn}>×</button></div>))}</div>
              </div>
            </div>
            <input placeholder="Remarks" style={ui.field} value={row.remarks} onChange={e => setRows(rows.map(r => r.id===row.id?{...r, remarks: e.target.value}:r))} />
          </div>
        ))}
      </div>
      <div style={ui.footer}>
        <button onClick={addRow} style={ui.btnGreen}>+ ROW</button>
        <button onClick={generatePDF} disabled={isUploading} style={ui.btnBlue}>GET PDF</button>
        <button onClick={generateExcel} style={ui.btnExcel}>GET EXCEL</button>
        <button onClick={clearAll} style={ui.btnRed}>CLEAR ALL</button>
      </div>
    </div>
  );
}

const ui = {
  container: { background: '#f4f7fa', minHeight: '100vh', paddingBottom: '110px', fontFamily: 'Arial' },
  nav: { background: '#1a1c1e', padding: '15px', position: 'sticky', top: 0, zIndex: 10, display: 'flex', flexDirection: 'column', gap: '8px' },
  authBtn: { background: '#4285F4', color: '#fff', border: 'none', padding: '12px', borderRadius: '4px', fontWeight: 'bold' },
  disBtn: { background: '#92d050', color: '#1a1c1e', border: 'none', padding: '12px', borderRadius: '4px', fontWeight: 'bold' },
  headIn: { width: '100%', background: 'transparent', border: '1px solid #fff', borderRadius: '4px', color: '#fff', textAlign: 'center', fontSize: '18px', fontWeight: 'bold' },
  main: { padding: '12px' },
  card: { background: '#fff', borderRadius: '10px', padding: '15px', marginBottom: '15px', border: '1px solid #e0e6ed' },
  topRow: { display: 'flex', justifyContent: 'space-between', marginBottom: '12px', alignItems: 'center' },
  snoBadge: { fontSize: '14px', fontWeight: 'bold', color: '#333' },
  rowDelBtn: { background: '#fff', border: '1px solid #dc3545', color: '#dc3545', borderRadius: '4px', padding: '2px 8px', fontSize: '11px' },
  badge: { border: 'none', borderRadius: '4px', color: '#fff', padding: '5px', width: '80px', textAlign: 'center', fontSize: '12px' },
  inputGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginBottom: '10px' },
  field: { width: '100%', padding: '10px', borderRadius: '6px', border: '1px solid #ced4da', boxSizing: 'border-box' },
  area: { width: '100%', height: '65px', borderRadius: '6px', border: '1px solid #ced4da', padding: '10px', boxSizing: 'border-box', marginBottom: '10px' },
  photoGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginBottom: '12px' },
  upBtn: { background: '#f8f9fa', padding: '10px', textAlign: 'center', borderRadius: '6px', border: '1px dashed #adb5bd', fontSize: '11px', cursor: 'pointer', display: 'block' },
  uploadSection: { display: 'flex', flexDirection: 'column', gap: '5px' },
  thumbnailRow: { display: 'flex', flexWrap: 'wrap', gap: '5px', marginTop: '5px' },
  thumbWrap: { position: 'relative', width: '40px', height: '40px' },
  thumbImg: { width: '100%', height: '100%', objectFit: 'cover', borderRadius: '4px' },
  delBtn: { position: 'absolute', top: '-5px', right: '-5px', background: '#ff3b30', color: 'white', border: 'none', borderRadius: '50%', width: '18px', height: '18px', fontSize: '12px' },
  footer: { position: 'fixed', bottom: 0, width: '100%', background: '#fff', padding: '15px', display: 'flex', gap: '10px', boxShadow: '0 -2px 10px rgba(0,0,0,0.1)', boxSizing: 'border-box', zIndex: 10 },
  btnGreen: { flex: 1, padding: '15px', background: '#28a745', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 'bold' },
  btnBlue: { flex: 1, padding: '15px', background: '#007bff', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 'bold' },
  btnExcel: { flex: 1, padding: '15px', background: '#217346', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 'bold' },
  btnRed: { padding: '15px', background: '#dc3545', color: '#fff', border: 'none', borderRadius: '8px', fontWeight: 'bold' }
};
