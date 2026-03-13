// ═══════════════════════════════════════════
//   ESTADO GLOBAL
// ═══════════════════════════════════════════
let planData  = null;
let visitData = null;
let allRows   = [];

// ═══════════════════════════════════════════
//   EVENTOS DE CARGA DE ARCHIVOS
// ═══════════════════════════════════════════
document.getElementById('file-planeador').addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('name-planeador').textContent = '⏳ Procesando...';
  await parsePlaneador(file);
  document.getElementById('name-planeador').textContent = '✓ ' + file.name;
  document.getElementById('card-planeador').classList.add('loaded');
  tryEnableControls();
});

document.getElementById('file-visitas').addEventListener('change', async e => {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('name-visitas').textContent = '⏳ Procesando...';
  await parseVisitas(file);
  document.getElementById('name-visitas').textContent = '✓ ' + file.name;
  document.getElementById('card-visitas').classList.add('loaded');
  tryEnableControls();
});

// ═══════════════════════════════════════════
//   UTILIDAD: limpiar texto de celdas
// ═══════════════════════════════════════════
function cleanText(val) {
  if (val == null) return '';
  return String(val).replace(/[\t\n\r]/g, ' ').trim().replace(/\s+/g, ' ');
}

// Correcciones de nombres con errores en el Excel (F6)
const NOMBRE_CORRECCIONES = {
  '71385266': 'JULIAN BETANCUR LOPEZ',  // Excel dice "JULIANBETANCUR LOPEZ"
};

// ═══════════════════════════════════════════
//   DETECTAR SI UNA CELDA TIENE FONDO VERDE
//
//   Contempla todos los casos del planeador:
//   - rgb FF92D050 / FF8ED965  → verde explícito
//   - theme 9 (con o sin tint) → verde tema Office
//   - La celda puede tener X, texto, o estar vacía
// ═══════════════════════════════════════════
function isCeldaVerde(cell) {
  if (!cell || !cell.s || !cell.s.fgColor) return false;
  const fc = cell.s.fgColor;
  if (fc.type === 'rgb'   && ['FF92D050','FF8ED965'].includes((fc.rgb||'').toUpperCase())) return true;
  if (fc.type === 'theme' && fc.theme === 9) return true;
  return false;
}

// ═══════════════════════════════════════════
//   PARSEAR PLANEADOR EXCEL
//
//   F6  = Nombre persona
//   F7  = Cédula
//   Fila 17, cols G+ = días del mes (1–31)
//   Col E = código de zona  ← cruce con CSV.RUTA
//   Col F = actividad
//
//   Una celda cuenta como visita planeada si:
//   → tiene fondo VERDE  (con o sin X)
// ═══════════════════════════════════════════
async function parsePlaneador(file) {
  const data = await file.arrayBuffer();
  const wb   = XLSX.read(data, { type: 'array', cellStyles: true });
  planData   = {};

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];

    const nombreCell = ws['F6'];
    const cedulaCell = ws['F7'];
    if (!nombreCell || !cedulaCell) continue;

    const nombreRaw = cleanText(nombreCell.v);
    const cedula = String(Math.round(Number(cedulaCell.v)) || '').trim();
    if (!nombreRaw || !cedula || cedula === '0' || cedula === 'NaN') continue;
    const nombre = NOMBRE_CORRECCIONES[cedula] || nombreRaw;

    // Fila 17 → mapa columna:día (solo días reales 1–31, sin duplicados)
    const dateMap = {};
    const diasUsados = new Set();
    for (let col = 7; col <= 55; col++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 16, c: col - 1 })];
      if (cell && cell.v != null) {
        const num = Math.round(Number(cell.v));
        if (num >= 1 && num <= 31 && !diasUsados.has(num)) {
          dateMap[col] = num;
          diasUsados.add(num);
        }
      }
    }

    // Filas 18+ → buscar celdas con fondo VERDE (con o sin X)
    const rows  = [];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');

    for (let r = 17; r <= range.e.r; r++) {
      const zonaCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];
      const actCell  = ws[XLSX.utils.encode_cell({ r, c: 5 })];
      if (!zonaCell || !zonaCell.v) continue;

      const zona      = cleanText(zonaCell.v).toUpperCase();
      const actividad = cleanText(actCell ? actCell.v : '');
      if (!zona || zona.length < 3) continue;

      for (let col = 7; col <= 55; col++) {
        const dia = dateMap[col];
        if (!dia) continue;

        const cell = ws[XLSX.utils.encode_cell({ r, c: col - 1 })];

        // Contar si la celda tiene fondo verde (independiente del valor)
        if (isCeldaVerde(cell)) {
          rows.push({ zona, actividad, dia });
        }
      }
    }

    planData[sheetName] = { nombre, cedula, rows };
  }
}

// ═══════════════════════════════════════════
//   PARSEAR CSV DE VISITAS R_VIS
//
//   FECHA   = "2026-03-02 12:03:41"
//   RUTA    = "R2ENVTEA0206"  ← igual a col E planeador
//   CEDULA  = "71766474"      ← clave de cruce
//
//   Clave: CEDULA|RUTA|DIA
// ═══════════════════════════════════════════
async function parseVisitas(file) {
  visitData = new Map();

  const text  = await file.text();
  const lines = text.replace(/^\uFEFF/, '').split('\n').filter(l => l.trim());
  if (lines.length < 2) return;

  const sep     = lines[0].includes(';') ? ';' : ',';
  const headers = lines[0].split(sep).map(h => h.trim().replace(/^"|"$/g, ''));

  const iFecha  = headers.indexOf('FECHA');
  const iRuta   = headers.indexOf('RUTA');
  const iCedula = headers.indexOf('CEDULA');

  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(sep).map(c => c.trim().replace(/^"|"$/g, ''));
    if (cols.length < 9) continue;

    const m = (cols[iFecha] || '').match(/\d{4}-\d{2}-(\d{2})/);
    if (!m) continue;
    const dia = parseInt(m[1], 10);

    const ruta   = (cols[iRuta]   || '').trim().toUpperCase();
    const cedula = (cols[iCedula] || '').trim();
    if (!ruta || !cedula) continue;

    const key = `${cedula}|${ruta}|${dia}`;
    visitData.set(key, (visitData.get(key) || 0) + 1);
  }
}

// ═══════════════════════════════════════════
//   HABILITAR CONTROLES
// ═══════════════════════════════════════════
function tryEnableControls() {
  if (!planData || !visitData) return;

  const sel = document.getElementById('sel-persona');
  sel.innerHTML = '<option value="">— Selecciona una persona —</option>';

  const sorted = Object.entries(planData).sort((a, b) =>
    a[1].nombre.localeCompare(b[1].nombre, 'es')
  );

  for (const [sheet, info] of sorted) {
    const opt       = document.createElement('option');
    opt.value       = sheet;
    opt.textContent = `${info.nombre} (CC ${info.cedula})`;
    sel.appendChild(opt);
  }

  sel.disabled = false;
  document.getElementById('sel-mes').disabled  = false;
  document.getElementById('sel-mes').innerHTML = '<option value="3">Marzo 2026</option>';
  document.getElementById('btn-analizar').disabled = false;
}

// ═══════════════════════════════════════════
//   ANALIZAR: planeación vs visitas reales
// ═══════════════════════════════════════════
function analizarPersona() {
  const sheet = document.getElementById('sel-persona').value;
  if (!sheet || !planData[sheet]) return;

  const persona = planData[sheet];

  if (persona.rows.length === 0) {
    const initials = persona.nombre.split(' ').slice(0,2).map(w=>w[0]).join('');
    document.getElementById('avatar').textContent         = initials;
    document.getElementById('persona-nombre').textContent = persona.nombre;
    document.getElementById('persona-meta').textContent   = `CC ${persona.cedula} · Sin visitas de terreno planeadas`;
    document.getElementById('stat-total').textContent    = 0;
    document.getElementById('stat-cumple').textContent   = 0;
    document.getElementById('stat-nocumple').textContent = 0;
    document.getElementById('stat-pct').textContent      = '—';
    document.getElementById('progress-bar').style.width  = '0%';
    document.getElementById('table-title').textContent   = `Detalle · ${persona.nombre}`;
    renderTabla([]);
    document.getElementById('results').style.display     = 'block';
    document.getElementById('empty-state').style.display = 'none';
    return;
  }

  allRows = persona.rows.map(row => {
    const key   = `${persona.cedula}|${row.zona}|${row.dia}`;
    const count = visitData.get(key) || 0;
    return { ...row, visitas: count, cumple: count > 0 };
  }).sort((a, b) => a.dia - b.dia);

  const total    = allRows.length;
  const cumple   = allRows.filter(r => r.cumple).length;
  const noCumple = total - cumple;
  const pct      = total > 0 ? Math.round(cumple / total * 100) : 0;

  const initials = persona.nombre.split(' ').slice(0, 2).map(w => w[0]).join('');
  document.getElementById('avatar').textContent         = initials;
  document.getElementById('persona-nombre').textContent = persona.nombre;
  document.getElementById('persona-meta').textContent   =
    `CC ${persona.cedula} · ${total} visitas planeadas · Marzo 2026`;

  document.getElementById('stat-total').textContent    = total;
  document.getElementById('stat-cumple').textContent   = cumple;
  document.getElementById('stat-nocumple').textContent = noCumple;
  document.getElementById('stat-pct').textContent      = pct + '%';

  const bar = document.getElementById('progress-bar');
  bar.style.width      = pct + '%';
  bar.style.background =
    pct >= 70 ? 'var(--green)' :
    pct >= 40 ? 'var(--amber)' : 'var(--red)';

  document.getElementById('table-title').textContent = `Detalle · ${persona.nombre}`;
  renderTabla(allRows);

  document.getElementById('results').style.display     = 'block';
  document.getElementById('empty-state').style.display = 'none';

  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  document.querySelector('.filter-btn').classList.add('active');
}

// ═══════════════════════════════════════════
//   RENDERIZAR TABLA
// ═══════════════════════════════════════════
function renderTabla(rows) {
  const tbody = document.getElementById('tabla-body');
  tbody.innerHTML = '';

  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:2rem;color:var(--muted)">
      Sin visitas de terreno planeadas para este período
    </td></tr>`;
    return;
  }

  for (const row of rows) {
    const tr         = document.createElement('tr');
    const badgeClass = row.cumple ? 'cumple'             : 'nocumple';
    const badgeText  = row.cumple ? '✅ Cumplió'         : '❌ No cumplió';
    const countStyle = row.cumple ? 'color:var(--green)' : 'color:var(--red)';

    tr.innerHTML = `
      <td><span class="day-cell">${row.dia}</span></td>
      <td><span class="zona-cell">${row.zona}</span></td>
      <td style="color:var(--muted);font-size:0.82rem">${row.actividad || '—'}</td>
      <td><span class="visits-count" style="${countStyle}">${row.visitas}</span></td>
      <td><span class="badge ${badgeClass}">${badgeText}</span></td>
    `;
    tbody.appendChild(tr);
  }
}

// ═══════════════════════════════════════════
//   FILTRAR TABLA
// ═══════════════════════════════════════════
function filtrarTabla(tipo, btn) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');

  const filtered =
    tipo === 'cumple'   ? allRows.filter(r => r.cumple)  :
    tipo === 'nocumple' ? allRows.filter(r => !r.cumple) :
    allRows;

  renderTabla(filtered);
}
