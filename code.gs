function doGet(e) {
  const log = (msg) => Logger.log("🧩 " + msg);

  log("Function started");

  if (!e.parameter.id) {
    log("No ID provided — showing HTML template form");
    return HtmlService.createTemplateFromFile("Form").evaluate();
  }

  const id = e.parameter.id.trim().toUpperCase();
  log("ID received: " + id);

  const ss = SpreadsheetApp.openById("1gn9XbNxrLqlnR229G79IKQnj5LIJG-QuXj2YAuvBlqw");
  const sheet = ss.getSheetByName("Master");
  if (!sheet) return HtmlService.createHtmlOutput("❌ Sheet not found");

  const data = sheet.getDataRange().getValues();
  log("Rows loaded: " + data.length);

  let match = null;

  for (let i = 1; i < data.length; i++) {
    const rowID = data[i][15];
    if (rowID && rowID.toString().trim().toUpperCase() === id) {
      match = data[i];
      break;
    }
  }

  if (!match) {
    return HtmlService.createHtmlOutput(`<h3>No se encontró la guía con ID: ${id}</h3>`);
  }

  const label = `
    <html>
      <head>
        <title>Etiqueta</title>
        <style>
          @media print { body { margin: 0; } }
          body { font-family: Arial; width: 56mm; margin: 0; padding: 0; }
          .label {
            width: 56mm; padding: 4mm;
            border: 1px dashed #000;
            font-size: 10pt;
          }
          .section { margin-bottom: 2mm; }
          .bold { font-weight: bold; }
          button { margin: 10px auto; display: block; }
        </style>
      </head>
      <body>
        <div class="label">
          <div class="section"><span class="bold">Cliente:</span> ${match[1]}</div>
          <div class="section"><span class="bold">Destinatario:</span> ${match[7]}</div>
          <div class="section"><span class="bold">Dirección:</span> ${match[9]}, ${match[10]}, ${match[11]}</div>
          <div class="section"><span class="bold">Guía:</span> ${match[15]}</div>
        </div>
        <button onclick="window.print()">🖨️ Imprimir</button>
      </body>
    </html>
  `;

  log("✅ Label generated");
  return HtmlService.createHtmlOutput(label);
}

