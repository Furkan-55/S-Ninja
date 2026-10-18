const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = 3000;

const EXCEL_PATH = path.join(__dirname, 'puanlar.xlsx');
const JSON_BACKUP_PATH = path.join(__dirname, 'puanlar.json');

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.use(express.json());
app.use(express.static(__dirname));
app.use('/mediapipe/hands', express.static(path.join(__dirname, 'node_modules', '@mediapipe', 'hands')));
app.use('/mediapipe/camera_utils', express.static(path.join(__dirname, 'node_modules', '@mediapipe', 'camera_utils')));

function normalizeEntry(r) {
  const rawName = r.isim || r.instagram || r['İsim'] || r['Instagram'] || '';
  const isim = String(rawName).startsWith('@') ? String(rawName) : '@' + String(rawName);
  const puan = Number(r.puan ?? r.score ?? r['Puan']) || 0;

  let tarih = r.tarih || r['Tarih'] || '';
  let saat = r.saat || r['Saat'] || '';

  if ((!tarih || !saat) && (r.dateFormatted || r['Tarih / Saat'])) {
    const parts = String(r.dateFormatted || r['Tarih / Saat']).trim().split(/\s+/);
    if (!tarih) tarih = parts[0] || '';
    if (!saat) saat = parts[1] || '';
  }

  if (!tarih || !saat) {
    const now = new Date();
    if (!tarih) {
      tarih = now.toLocaleDateString('tr-TR', { day: '2-digit', month: '2-digit', year: 'numeric' });
    }
    if (!saat) {
      saat = now.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    }
  }

  return { isim, puan, tarih, saat };
}

function sortRowsDesc(rows) {
  return rows.sort((a, b) => (Number(b.puan) || 0) - (Number(a.puan) || 0));
}

function readRows() {
  try {
    if (fs.existsSync(JSON_BACKUP_PATH)) {
      const data = JSON.parse(fs.readFileSync(JSON_BACKUP_PATH, 'utf8'));
      if (Array.isArray(data)) {
        return sortRowsDesc(
          data
            .filter(r => r && (r.isim || r.instagram))
            .map(normalizeEntry)
        );
      }
    }
  } catch (_) {}

  try {
    if (fs.existsSync(EXCEL_PATH)) {
      const wb = XLSX.readFile(EXCEL_PATH);
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(firstSheet);
      return sortRowsDesc(
        rawRows
          .filter(r => (r['İsim'] && r['İsim'] !== '-') || (r['Instagram'] && r['Instagram'] !== '-'))
          .map(normalizeEntry)
      );
    }
  } catch (_) {}

  return [];
}

function writeExcel(rows) {
  const cleanSorted = sortRowsDesc(rows.map(normalizeEntry));
  fs.writeFileSync(JSON_BACKUP_PATH, JSON.stringify(cleanSorted, null, 2), 'utf8');

  // Excel'de puana göre çoktan aza sıralı liste (Sadece İsim, Puan, Tarih, Saat)
  const sheetData = cleanSorted.map((item) => ({
    'İsim': item.isim,
    'Puan': item.puan,
    'Tarih': item.tarih,
    'Saat': item.saat
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(sheetData, {
    header: ['İsim', 'Puan', 'Tarih', 'Saat']
  });

  ws['!cols'] = [
    { wch: 28 },
    { wch: 12 },
    { wch: 16 },
    { wch: 14 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Puanlar');

  try {
    XLSX.writeFile(wb, EXCEL_PATH);
    return true;
  } catch (err) {
    console.warn('puanlar.xlsx şu an Excel programında açık olabilir:', err.message);
    return false;
  }
}

if (!fs.existsSync(EXCEL_PATH)) {
  writeExcel(readRows());
}

app.get('/api/ping', (req, res) => {
  res.json({ ok: true });
});

app.get('/api/scores', (req, res) => {
  const rows = readRows();
  res.json({ ok: true, scores: rows, excelPath: EXCEL_PATH });
});

app.post('/api/scores', (req, res) => {
  const body = req.body || {};
  const rawName = body.isim || body.instagram;
  if (!rawName) {
    return res.status(400).json({ ok: false, error: 'İsim gerekli' });
  }
  const rows = readRows();
  const entry = normalizeEntry(body);
  rows.push(entry);
  sortRowsDesc(rows);
  const saved = writeExcel(rows);
  res.json({ ok: true, saved, total: rows.length, scores: rows });
});

app.delete('/api/scores', (req, res) => {
  writeExcel([]);
  res.json({ ok: true, scores: [] });
});

app.listen(PORT, '127.0.0.1', () => {
  console.log(`SÜDEV Fruit Ninja Uygulama Servisi Aktif: ${EXCEL_PATH}`);
});
