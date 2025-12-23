const { PDFDocument, StandardFonts } = PDFLib;

// Helper: ubah angka jadi string dengan titik ribuan
function cleanNumber(n) {
  if (typeof n === "number") return Math.round(n).toLocaleString("id-ID");
  if (!n) return "";
  const parsed = Number(n);
  return isNaN(parsed) ? String(n) : Math.round(parsed).toLocaleString("id-ID");
}

// Parse Excel sesuai struktur GUEST / ORDERS / TOTALS
function loadInvoiceFromUploadedExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils
    .sheet_to_json(sheet, { header: 1, raw: false })
    .flat();

  let section = null;
  let guestData = [];
  let ordersData = [];
  let totalsData = [];

  for (let i = 0; i < rows.length; i++) {
    const cell = rows[i];
    if (cell === "GUEST") {
      section = "guest";
      continue;
    }
    if (cell === "ORDERS") {
      section = "orders";
      continue;
    }
    if (cell === "TOTALS") {
      section = "totals";
      continue;
    }

    if (section === "guest") guestData.push(cell);
    if (section === "orders") ordersData.push(cell);
    if (section === "totals") totalsData.push(cell);
  }

  // --- GUEST ---
  const invoiceData = {
    name: guestData[6],
    invNumber: guestData[7],
    dateOrder: guestData[8],
    checkIn: guestData[9],
    checkOut: guestData[10],
    room: guestData[11],
  };

  // --- ORDERS ---
  const orders = [];
  for (let i = 5; i < ordersData.length; i += 5) {
    const chunk = ordersData.slice(i, i + 5);
    if (chunk.length < 5) continue; // skip kalau tidak lengkap

    const row = {
      nights: cleanNumber(chunk[0]) + " Malam",
      rate: cleanNumber(chunk[1]),
      isDiscount: chunk[2],
      totalRate: cleanNumber(chunk[3]),
      description: chunk[4],
    };

    // skip placeholder
    if (
      String(row.nights).includes("REQUIRED!!") ||
      String(row.rate).includes("ISI DISINI") ||
      String(row.totalRate).includes("ISI DISINI")
    )
      continue;

    orders.push(row);
  }
  invoiceData.orders = orders;

  // --- TOTALS ---
  invoiceData.subtotal = cleanNumber(totalsData[4]);
  invoiceData.pajak = cleanNumber(totalsData[5]);
  invoiceData.total = cleanNumber(totalsData[6]);
  invoiceData.dp = "= " + cleanNumber(totalsData[7]);

  return invoiceData;
}

// Helper: rata tengah terhadap titik acuan
function drawCenteredAtPoint(page, font, text, fontSize, anchorX, anchorY) {
  if (!text) return;
  const textWidth = font.widthOfTextAtSize(text, fontSize);
  const x = anchorX - textWidth / 2;
  page.drawText(text, { x, y: anchorY, size: fontSize, font });
}

// Helper: tulis satu baris order
function drawOrderRow(page, font, order, fontSize, startX, startY) {
  drawCenteredAtPoint(page, font, order.nights, fontSize, startX, startY);
  drawCenteredAtPoint(
    page,
    font,
    order.description,
    fontSize,
    startX + 130,
    startY
  );
  drawCenteredAtPoint(page, font, order.rate, fontSize, startX + 280, startY);
  drawCenteredAtPoint(
    page,
    font,
    order.totalRate,
    fontSize,
    startX + 390,
    startY
  );
}

// Generate invoice PDF
async function fillInvoice(pdfArrayBuffer, data) {
  const pdfDoc = await PDFDocument.load(pdfArrayBuffer);
  const page = pdfDoc.getPages()[0];

  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  // --- Guest Info ---
  page.drawText(data.name || "", { x: 130, y: 653, size: 18, font: fontBold });
  drawCenteredAtPoint(page, font, data.invNumber || "", 12, 490, 630);
  page.drawText(data.dateOrder || "", { x: 203, y: 620, size: 12, font });
  page.drawText(data.checkIn || "", { x: 203, y: 600, size: 12, font });
  page.drawText(data.checkOut || "", { x: 203, y: 582, size: 12, font });
  page.drawText(data.room || "", { x: 203, y: 566, size: 12, font });

  // --- Orders (1–10 baris fleksibel) ---
  let startY = 500;
  const lineGap = 20;
  for (let i = 0; i < data.orders.length; i++) {
    drawOrderRow(page, font, data.orders[i], 12, 85, startY - i * lineGap);
  }

  // --- DP ---
  page.drawText("Minimal DP 50%", { x: 123, y: 285, size: 12, font });
  page.drawText(data.dp || "", { x: 130, y: 267, size: 12, font: fontBold });

  // --- Totals ---
  drawCenteredAtPoint(page, font, data.subtotal || "", 12, 475, 233);
  drawCenteredAtPoint(page, font, data.pajak || "", 12, 475, 205);
  drawCenteredAtPoint(page, fontBold, data.total || "", 12, 475, 177);

  return await pdfDoc.save();
}

// FileReader helper (mobile friendly)
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    // fallback: gunakan binary string di mobile
    if (/Android|iPhone|iPad/i.test(navigator.userAgent)) {
      reader.onload = () => {
        const binary = reader.result;
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        resolve(bytes.buffer);
      };
      reader.onerror = (e) => {
        console.error("FileReader error:", e);
        reject(new Error("Gagal membaca file di mobile browser"));
      };
      reader.readAsBinaryString(file);
    } else {
      reader.onload = () => resolve(reader.result);
      reader.onerror = (e) => {
        console.error("FileReader error:", e);
        reject(
          new Error("Gagal membaca file. Pastikan format dan ukuran sesuai.")
        );
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

// Download helper (desktop + Android friendly)
function downloadBytes(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);

  const isMobile = /Android|iPhone|iPad/i.test(navigator.userAgent);

  if (isMobile) {
    // buka di tab baru agar user bisa download manual
    window.open(url, "_blank");
  } else {
    // desktop: langsung download
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
}

// Event handler
document.getElementById("generateBtn").addEventListener("click", async () => {
  const pdfFile = document.getElementById("pdfInput").files[0];
  const xlsxFile = document.getElementById("xlsxInput").files[0];
  if (!pdfFile || !xlsxFile) {
    alert("Upload template PDF dan data Excel terlebih dahulu.");
    return;
  }

  try {
    const [pdfBuf, xlsxBuf] = await Promise.all([
      readFileAsArrayBuffer(pdfFile),
      readFileAsArrayBuffer(xlsxFile),
    ]);

    const data = loadInvoiceFromUploadedExcel(xlsxBuf);
    const outBytes = await fillInvoice(pdfBuf, data);
    const outputName = `Invoice_${data.name || "Guest"}_${
      data.invNumber || "INV"
    }.pdf`;
    downloadBytes(outBytes, outputName);
  } catch (err) {
    console.error(err);
    alert("Terjadi kesalahan saat memproses file. Lihat console untuk detail.");
  }
});
