function doGet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
    var lastRow = sheet.getLastRow();
    
    Logger.log("📊 Last Row: " + lastRow);

    if (lastRow < 2) { // Ensure there's data (assuming row 1 has headers)
        Logger.log("🚨 No valid data found!");
        return HtmlService.createHtmlOutput("<h2>No ID Found! 🚨</h2>");
    }

    var uniqueID = sheet.getRange(lastRow, 16).getValue();
    Logger.log("✅ Unique ID Retrieved: " + uniqueID);

var html = HtmlService.createHtmlOutput(`
    <html>
        <head>
            <title>Guía de Rastreo</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    text-align: center;
                    padding: 40px;
                    background-color: #f4f4f9;
                }
                h2 {
                    color: #333;
                }
                h1 {
                    color: #007bff;
                    font-size: 32px;
                }
                .btn {
                    display: inline-block;
                    margin: 20px;
                    padding: 12px 24px;
                    font-size: 16px;
                    font-weight: bold;
                    color: white;
                    background: #007bff;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    text-decoration: none;
                    transition: background 0.3s ease;
                }
                .btn:hover {
                    background: #0056b3;
                }
                .btn.secondary {
                    background: #28a745;
                }
                .btn.secondary:hover {
                    background: #1e7e34;
                }
            </style>
        </head>
        <body>
            <h2>Agrega la siguiente guía a tu paquete:</h2>
            <h1>${uniqueID}</h1>
            <a href="https://forms.gle/p4Euoy2TRspRoRbU8" class="btn">Crear guia</a>
            <a href="https://www.google.com/maps/d/edit?mid=1nwfEhqaxXNW0bTzCg-dW5ZPqGOtPRvg&usp=sharing" class="btn secondary">Zona de cobertura</a>
        </body>
    </html>
`);
return html;

}

function onFormSubmit(e) {
  if (!e || !e.range) {
    Logger.log("⚠️ Invalid event object!");
    return;
  }

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const masterSheet = e.source.getSheetByName("Master");

  const lastRow = sheet.getLastRow();

  // Get raw row from the submitting sheet
  const rowData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Define mappings for each form
  let client, cphone, rname, raddress, rcol, rcity, dname, dphone, daddress, dcol, dcity, parcel, notes;

  if (sheetName === "Form Responses 1") {
    client = rowData[1];
    cphone = rowData[2].toString();
    rname = rowData[3];
    raddress = rowData[4];
    rcol = rowData[5];
    rcity = rowData[6];
    dname = rowData[7];
    dphone = rowData[8].toString();
    daddress = rowData[9];
    dcol = rowData[10];
    dcity = rowData[11];
    notes = rowData[12];
  } else if (sheetName === "Form Responses 2") {
    client = "Electrycom";
    cphone = "3331194835";
    dname = rowData[2];
    dphone = rowData[3].toString();
    daddress = rowData[4];
    dcol = rowData[5];
    dcity = rowData[6];
    parcel = rowData[7];
    notes = rowData[8];
  } else {
    Logger.log("❌ Not a target sheet.");
    return;
  }

  // Validate and format dphone
  if (!dphone || typeof dphone !== "string") {
    dphone = "error";
  } else {
    dphone = dphone.replace(/\s+/g, '');
    if (dphone.length === 10) {
      dphone = "52" + dphone;
    } else if (dphone.length === 12 && dphone.startsWith("52")) {
      // Already formatted
    } else {
      dphone = "error";
    }
  }

  // Validate and format cphone
  if (!cphone || typeof cphone !== "string") {
    cphone = "error";
  } else {
    cphone = cphone.replace(/\s+/g, '');
    if (cphone.length === 10) {
      cphone = "52" + cphone;
    } else if (cphone.length === 12 && cphone.startsWith("52")) {
      // Already formatted
    } else {
      cphone = "error";
    }
  }

  // Add timestamp and form source
  const timestamp = new Date();
  const source = sheetName;

  // Generate unique ID with random letter
  var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  var randomLetter = alphabet[Math.floor(Math.random() * alphabet.length)];
  var uniqueID = "ZMG-" + new Date().getTime().toString().slice(-6) + randomLetter;

  /////////////////// WhatsApp message generation
  var message = `Hola ${dname}, te notificamos que tienes un paquete por parte de ${client} con el número de guía ${uniqueID}, el mismo se entregará en un lapso no mayor de 2 horas. ¿Tienes alguna duda? No dudes en escribir.`;
  var encoderMessage = encodeURIComponent(message);

  var whatsappURL = dphone !== "error"
    ? `https://wa.me/${dphone}?text=${encoderMessage}`
    : "error";

  var message2 = `Hola ${client}, el número de guía ${uniqueID}, ha sido asignado a tu envío, el mismo se entregará al destinatario ${dname} en un lapso no mayor de 2 horas. ¿Tienes alguna duda? No dudes en escribir.`;
  var encoderMessage2 = encodeURIComponent(message2);

  var whatsappURL2 = cphone !== "error"
    ? `https://wa.me/${cphone}?text=${encoderMessage2}`
    : "error";
  
  var message3 = `Hola ${client}, el número de guía ${uniqueID}, ha sido entregado con éxito al destinatario ${dname}, DLC Logística agradece tu confianza.`;
  var encoderMessage3 = encodeURIComponent(message3);

  var whatsappURL3 = cphone !== "error"
    ? `https://wa.me/${cphone}?text=${encoderMessage3}`
    : "error";
  // Append to Master sheet
  masterSheet.appendRow([
    timestamp,
    client,
    cphone,
    rname,
    raddress,
    rcol,
    rcity,
    dname,
    dphone,
    daddress,
    dcol,
    dcity,
    parcel,
    notes,
    source,
    uniqueID,
    whatsappURL2,
    whatsappURL,
    whatsappURL3
  ]);

  Logger.log(`✅ Data from ${sheetName} mapped and added to Master`);
}



