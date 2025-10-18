// === CONFIGURAÇÕES ===

// Google Sheets com as seleções
const SPREADSHEET_HII = '1G6tZyr3zY0rjZCSCxAN-W72YUH_ZHw8Aka30C8Y2H1I';

// Base de dados de indicadores
const INDICATORS_HII = '1jW-Q1dz-q5t_3pHiWG1HSwAm4k-fsaWaiJstxM56EF8';

// Colunas a exportar
const EXPORT_COLUMNS = [
  'Indicator Number',
  'Indicators',
  'Associated Sub-Indicators/Sub-Items',
  'Associated Scales',
  'Notes',
  'Link of the Articles'
];

// Linha sentinela
const SENTINEL_ID = '-1';

// === FUNÇÕES AUXILIARES ===
function openSheet_(id, name) {
  const ss = SpreadsheetApp.openById(id);
  return name ? ss.getSheetByName(name) : ss.getSheets()[0];
}

function getAll_(sh) {
  return sh.getDataRange().getValues();
}

function headerMap_(row) {
  const m = {};
  row.forEach((v, i) => (m[String(v).trim()] = i));
  return m;
}

// === MENSAGEM HTML BONITA ===
function htmlMessage_(msg, emoji = '✅') {
  return HtmlService.createHtmlOutput(`
    <html>
      <body style="font-family:system-ui; background:#f7faff; height:100vh; display:flex; align-items:center; justify-content:center;">
        <div style="text-align:center; padding:30px 40px; background:white; border-radius:14px; box-shadow:0 4px 14px rgba(0,0,0,0.1); max-width:600px;">
          <div style="font-size:48px;">${emoji}</div>
          <div style="font-size:20px; line-height:1.5; color:#333; margin-top:10px;">${msg}</div>
          <p style="margin-top:25px; color:#888; font-size:14px;">You can now return to Innovation Vitals tool.</p>
        </div>
      </body>
    </html>
  `);
}

// === FUNÇÃO PRINCIPAL ===
function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_HII);
  const sh = ss.getSheets()[0];

  const p = (e && e.parameter) ? e.parameter : {};
  const project = (p.name_of_the_project || '').toString().trim();
  const ind = (p.id || '').toString().trim();
  const op = (p.op || 'add').toString().trim(); // add | remove | clear | export_full

  // ⚠️ Verificar se o nome do projeto foi indicado
  if (!project) {
    return htmlMessage_(
      '⚠️ No project name provided.<br><b>Please enter a project name before using this feature.</b>',
      '⚠️'
    );
  }

  // 🟩 ADICIONAR INDICADOR (todas as linhas com o mesmo número)
  if (op === 'add' && ind) {
    const indSheet = SpreadsheetApp.openById(INDICATORS_HII).getSheets()[0];
    const indVals = indSheet.getDataRange().getValues();
    const indHdr = indVals[0].map(String);
    const indMap = headerMap_(indHdr);
    const numCol = indMap['Indicator Number'];

    let added = 0;

    for (let r = 1; r < indVals.length; r++) {
      const row = indVals[r];
      if ((row[numCol] + '').trim() === ind) {
        sh.appendRow([new Date(), ind, project]);
        added++;
      }
    }

    if (added > 0) {
      return htmlMessage_(
        `The indicator was successfully added to the list of indicators selected to evaluate the project <b>${project}</b>.`
      );
    } else {
      return htmlMessage_(
        `No indicators found in the main database with the number <b>${ind}</b>.`,
        '⚠️'
      );
    }
  }

  const vals = sh.getDataRange().getValues();

  // 🗑 REMOVER TODOS OS REGISTOS COM O MESMO NÚMERO
  if (op === 'remove' && ind) {
    let deleted = 0;
    for (let row = vals.length; row >= 2; row--) {
      const rowNum = (vals[row - 1][1] + '').trim();
      const rowProj = (vals[row - 1][2] + '').trim();
      if (rowNum === ind && rowProj === project) {
        sh.deleteRow(row);
        deleted++;
      }
    }

    if (deleted > 0) {
      return htmlMessage_(
        `The indicator was successfully removed from the list of indicators selected to evaluate the project <b>${project}</b>.`,
        '🗑️'
      );
    } else {
      return htmlMessage_(
        `The indicator <b>${ind}</b> was not found in the list of indicators selected to evaluate the project <b>${project}</b>.`,
        '⚠️'
      );
    }
  }

  // 🧹 REMOVER TODOS OS INDICADORES DO PROJETO
  if (op === 'clear') {
    let deleted = 0;
    for (let row = vals.length; row >= 2; row--) {
      if ((vals[row - 1][2] + '').trim() === project) {
        sh.deleteRow(row);
        deleted++;
      }
    }
    return htmlMessage_(
      `All the indicators were successfully removed from the list of indicators selected to evaluate the project <b>${project}</b>.`,
      '🧹'
    );
  }

  // 📤 EXPORTAÇÃO
  if (op === 'export_full') return exportFull_(project);

  return htmlMessage_('✅ Action completed successfully.');
}

// === EXPORTAÇÃO COMPLETA ===
function exportFull_(project) {
  const sel = openSheet_(SPREADSHEET_HII);
  const sVals = getAll_(sel);

  if (sVals.length < 2)
    return HtmlService.createHtmlOutput('<p>No indicators selected.</p>');

  const idCol = 1;
  const projectCol = 2;
  const ids = new Set();

  for (let r = 2; r <= sVals.length; r++) {
    const row = sVals[r - 1];
    if ((row[projectCol] + '').trim() === project) {
      const id = String(row[idCol] ?? '').trim();
      if (id && id !== SENTINEL_ID) ids.add(id);
    }
  }

  if (!ids.size) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:system-ui;text-align:center;padding:50px;">
        <h2>📭 No indicators found for project <b>${project}</b>.</h2>
        <p>You can now return to the report and start selecting indicators.</p>
      </body></html>`);
  }

  // Carregar base de dados
  const ind = openSheet_(INDICATORS_HII);
  const iVals = getAll_(ind);
  const iHdr = iVals[0].map(String);
  const iMap = headerMap_(iHdr);
  const idxNum = iMap['Indicator Number'];

  // Mapear todos os registos por número
  const rowById = new Map();
  for (let r = 2; r <= iVals.length; r++) {
    const row = iVals[r - 1];
    const id = String(row[idxNum] ?? '').trim();
    if (!id) continue;
    if (!rowById.has(id)) rowById.set(id, []);
    rowById.get(id).push(row);
  }

  // Função de escape CSV (UTF-8 friendly)
  const escapeCSV = (text) => {
    if (text == null) return '';
    const s = String(text).replace(/"/g, '""');
    return /[",\n]/.test(s) ? `"${s}"` : s;
  };

  let html = `
  <html><head><meta charset='UTF-8'>
  <style>
    body { font-family: system-ui; padding: 20px; background:#f9fafc; }
    h2 { color: #1a73e8; margin-bottom: 20px; text-align:center; }
    table { border-collapse: collapse; width: 100%; margin-top: 10px; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }
    th { background-color: #e8f0fe; }
    tr:nth-child(even) { background: #f9f9f9; }
    .btn { background: #1a73e8; color: white; border: none; border-radius: 6px;
           padding: 10px 18px; margin-top: 15px; font-size: 15px; cursor: pointer; display:block; margin-left:auto; margin-right:auto; }
    .btn:hover { background: #1558b0; }
  </style>
  </head><body>
  <h2>Indicators selected for evaluating the project ${project}</h2>
  <table><tr>${EXPORT_COLUMNS.map(c => `<th>${c}</th>`).join('')}</tr>`;

  const exportRows = [];
  Array.from(ids).sort((a, b) => Number(a) - Number(b)).forEach(id => {
    const rows = rowById.get(id);
    if (!rows) return;
    rows.forEach(row => {
      html += `<tr>${EXPORT_COLUMNS.map(c => `<td>${row[iMap[c]] ?? ''}</td>`).join('')}</tr>`;
      exportRows.push(EXPORT_COLUMNS.map(c => escapeCSV(row[iMap[c]])).join(','));
    });
  });

  html += `</table>
  <button class='btn' onclick='downloadCSV()'>⬇️ Download CSV</button>
  <script>
    function downloadCSV() {
      const header = ${JSON.stringify(EXPORT_COLUMNS)};
      const rows = ${JSON.stringify(exportRows)};
      const csv = "\\uFEFF" + header.join(",") + "\\n" + rows.join("\\n"); // UTF-8 BOM
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "Indicators_${project.replace(/\\s+/g, '_')}.csv";
      a.click();
    }
  </script></body></html>`;

  return HtmlService.createHtmlOutput(html);
}
