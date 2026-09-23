/**
 * Rudra Enterprises Invoice Generator — Core Application Logic
 * Direct Vector PDF Engine (100% Genuine Excel Base, Zero HTML)
 */

/* ===== Buyer Configuration ===== */
const BUYERS = {
  kaka: {
    id: 'kaka',
    name: 'Kaka Tile Studio',
    location: 'Hyderabad, Telengana',
    icon: '🏪',
    address: '104 Sagar Enclave, Hyderabad, Telengana 500035',
    gst: '36AAWFK9520H1ZS',
    storageKey: 'lastInvoiceNum_kaka'
  },
  my: {
    id: 'my',
    name: 'My Tile Studio',
    location: 'Hyderabad, Telengana',
    icon: '🏬',
    address: 'MCH No. \u2013 8-16-61/PLOT NO 6\nSagar Road\nHyderabad\nTelengana\n500079',
    gst: '36BBCPK7265H1ZI',
    storageKey: 'lastInvoiceNum_my'
  }
};

/* ===== App State ===== */
let selectedBuyer = null;
let currentPDFBlob = null;
let currentXLSXBlob = null;
let currentInvoiceNo = '';
let currentPDFUrl = '';
let currentXLSXUrl = '';

/* ===== DOM Elements ===== */
const screens = {
  select: document.getElementById('screen-select'),
  form: document.getElementById('screen-form'),
  preview: document.getElementById('screen-preview'),
  result: document.getElementById('screen-result')
};

const elements = {
  // Form
  buyerBadgeName: document.getElementById('buyer-badge-name'),
  buyerBadgeIcon: document.getElementById('buyer-badge-icon'),
  invoiceNoInput: document.getElementById('invoice-no'),
  invoiceDateInput: document.getElementById('invoice-date'),
  quantityInput: document.getElementById('quantity'),
  invoiceNoHint: document.getElementById('invoice-no-hint'),

  // Preview
  previewBuyer: document.getElementById('preview-buyer'),
  previewInvoiceNo: document.getElementById('preview-invoice-no'),
  previewDate: document.getElementById('preview-date'),
  previewQty: document.getElementById('preview-qty'),
  previewAmount: document.getElementById('preview-amount'),
  previewTaxable: document.getElementById('preview-taxable'),
  previewIGST: document.getElementById('preview-igst'),
  previewTotal: document.getElementById('preview-total'),
  previewWords: document.getElementById('preview-words'),

  // Result
  resultFilename: document.getElementById('result-filename'),
  resultPreviewCanvas: document.getElementById('result-preview-canvas'),
  resultPreviewImg: document.getElementById('result-preview-img'),
  btnShare: document.getElementById('btn-share'),
  btnWhatsApp: document.getElementById('btn-whatsapp'),
  btnDownload: document.getElementById('btn-download'),
  btnDownloadExcel: document.getElementById('btn-download-excel'),

  // Loading
  loadingOverlay: document.getElementById('loading-overlay'),
  loadingText: document.getElementById('loading-text')
};

/* ===== Screen Navigation ===== */
function showScreen(screenId) {
  Object.values(screens).forEach(s => s.classList.remove('active'));
  screens[screenId].classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ===== Buyer Selection ===== */
function selectBuyer(buyerId) {
  selectedBuyer = BUYERS[buyerId];

  elements.buyerBadgeIcon.textContent = selectedBuyer.icon;
  elements.buyerBadgeName.textContent = selectedBuyer.name;

  // Auto-suggest next invoice number
  const lastNum = localStorage.getItem(selectedBuyer.storageKey);
  if (lastNum) {
    const parsed = parseInt(lastNum, 10);
    const suggested = !isNaN(parsed) ? parsed + 1 : '';
    elements.invoiceNoInput.value = suggested;
    elements.invoiceNoHint.textContent = `Last used: #${lastNum}`;
  } else {
    elements.invoiceNoInput.value = '';
    elements.invoiceNoHint.textContent = 'No previous invoice recorded';
  }

  // Default to today's date in dd/mm/yyyy
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy = today.getFullYear();
  elements.invoiceDateInput.value = `${dd}/${mm}/${yyyy}`;

  elements.quantityInput.value = '';
  showScreen('form');
}

/* ===== Form Submission & Calculations ===== */
function handleFormSubmit(e) {
  e.preventDefault();

  const invoiceNo = elements.invoiceNoInput.value.trim();
  const dateVal = elements.invoiceDateInput.value.trim();
  const quantity = parseInt(elements.quantityInput.value, 10);

  if (!invoiceNo) {
    elements.invoiceNoInput.focus();
    shakeElement(elements.invoiceNoInput);
    return;
  }
  if (!dateVal || !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateVal)) {
    elements.invoiceDateInput.focus();
    shakeElement(elements.invoiceDateInput);
    return;
  }
  if (!quantity || quantity <= 0) {
    elements.quantityInput.focus();
    shakeElement(elements.quantityInput);
    return;
  }

  // Calculate pricing & taxes
  const amount = quantity * 1000;
  const taxableValue = amount;
  const igst = Math.round(taxableValue * 0.18 * 100) / 100;
  const total = taxableValue + igst;
  const words = numberToWordsIndian(total);

  currentInvoiceNo = invoiceNo;
  window._invoiceData = {
    buyer: selectedBuyer.id,
    invoiceNo,
    invoiceDate: dateVal,
    quantity,
    amount,
    taxableValue,
    igst,
    total,
    amountInWords: words
  };

  // Populate preview screen
  elements.previewBuyer.textContent = selectedBuyer.name;
  elements.previewInvoiceNo.textContent = `#${invoiceNo}`;
  elements.previewDate.textContent = dateVal;
  elements.previewQty.textContent = `${quantity} Box`;
  elements.previewAmount.textContent = `₹${amount.toLocaleString('en-IN')}`;
  elements.previewTaxable.textContent = `₹${taxableValue.toLocaleString('en-IN')}`;
  elements.previewIGST.textContent = `₹${igst.toLocaleString('en-IN')}`;
  elements.previewTotal.textContent = `₹${total.toLocaleString('en-IN')}`;
  elements.previewWords.textContent = `Amount charged in words : ${words}`;

  showScreen('preview');
}

/* ===== Direct Vector PDF Generation (100% Genuine Excel Base) ===== */
async function generateInvoiceDirect() {
  showLoading(true, 'Generating authentic Excel invoice...');

  try {
    const data = window._invoiceData;
    const { PDFDocument, StandardFonts, rgb } = PDFLib;

    // 1. Load authentic base PDF exported directly by Microsoft Excel
    const baseB64 = data.buyer === 'kaka' ? window.KAKA_BASE_PDF : window.MY_BASE_PDF;
    if (!baseB64) {
      throw new Error('Base template not loaded. Please refresh the page.');
    }

    const binaryString = atob(baseB64);
    const pdfBytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      pdfBytes[i] = binaryString.charCodeAt(i);
    }

    const pdfDoc = await PDFDocument.load(pdfBytes);
    const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

    const page = pdfDoc.getPages()[0];
    const { width, height } = page.getSize(); // 612 x 792 (Letter standard)

    // Formatted currency strings
    const fmt = num => Number(num).toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });

    const formattedAmount = fmt(data.amount);
    const formattedTaxable = fmt(data.taxableValue);
    const formattedIGST = fmt(data.igst);
    const formattedTotal = fmt(data.total);
    const wordsText = `Amount charged in words : ${data.amountInWords}`;

    function drawRight(text, rightX, y, font, size) {
      const textWidth = font.widthOfTextAtSize(text, size);
      page.drawText(text, {
        x: rightX - textWidth,
        y: y,
        size: size,
        font: font,
        color: rgb(0, 0, 0)
      });
    }

    function drawCenter(text, y, font, size) {
      const textWidth = font.widthOfTextAtSize(text, size);
      page.drawText(text, {
        x: (width - textWidth) / 2,
        y: y,
        size: size,
        font: font,
        color: rgb(0, 0, 0)
      });
    }

    if (data.buyer === 'kaka') {
      // Kaka Tile Studio exact vector positions
      page.drawText(String(data.invoiceNo), { x: 318.3, y: 660.0, size: 8.5, font: helvetica });
      page.drawText(String(data.invoiceDate), { x: 318.3, y: 638.5, size: 8.5, font: helvetica });
      page.drawText(String(data.quantity), { x: 350.0, y: 457.0, size: 8.5, font: helveticaBold });

      drawRight(formattedAmount, 530.0, 457.0, helvetica, 8.5);
      drawRight(formattedTaxable, 530.0, 792.0 - 409.0, helveticaBold, 8.5);
      drawRight(formattedIGST, 530.0, 792.0 - 447.0, helveticaBold, 8.5);
      drawRight(formattedTotal, 530.0, 792.0 - 459.5, helveticaBold, 8.5);

      drawCenter(wordsText, 792.0 - 480.0, helvetica, 10.0);
    } else {
      // My Tile Studio exact vector positions (Image 3)
      page.drawText(String(data.invoiceNo), { x: 319.1, y: 688.0, size: 8.5, font: helvetica });
      page.drawText(String(data.invoiceDate), { x: 319.1, y: 678.0, size: 8.5, font: helvetica });
      page.drawText(String(data.quantity), { x: 356.0, y: 467.0, size: 8.5, font: helveticaBold });

      drawRight(formattedAmount, 558.0, 467.0, helvetica, 8.5);
      drawRight(formattedTaxable, 558.0, 792.0 - 394.5, helveticaBold, 8.5);
      drawRight(formattedIGST, 558.0, 792.0 - 429.5, helveticaBold, 8.5);
      drawRight(formattedTotal, 558.0, 792.0 - 441.5, helveticaBold, 8.5);

      drawCenter(wordsText, 792.0 - 461.0, helvetica, 10.0);
    }

    // 2. Export vector PDF bytes
    const finalPdfBytes = await pdfDoc.save();
    if (currentPDFUrl) URL.revokeObjectURL(currentPDFUrl);
    currentPDFBlob = new Blob([finalPdfBytes], { type: 'application/pdf' });
    currentPDFUrl = URL.createObjectURL(currentPDFBlob);

    // 3. Generate updated .xlsx Excel workbook in memory via SheetJS
    try {
      const xlsxB64 = data.buyer === 'kaka' ? window.KAKA_BASE_XLSX : window.MY_BASE_XLSX;
      if (xlsxB64 && window.XLSX) {
        const wb = XLSX.read(xlsxB64, { type: 'base64' });
        if (data.buyer === 'kaka') {
          const ws = wb.Sheets['Table 1'];
          if (ws) {
            ws['D5'] = { t: 's', v: String(data.invoiceNo) };
            ws['D7'] = { t: 's', v: String(data.invoiceDate) };
            ws['D18'] = { t: 'n', v: data.quantity };
            ws['G18'] = { t: 'n', v: data.amount };
            ws['G19'] = { t: 'n', v: data.taxableValue };
            ws['G22'] = { t: 'n', v: data.igst };
            ws['G23'] = { t: 'n', v: data.total };
            ws['A24'] = { t: 's', v: wordsText };
          }
        } else {
          const ws = wb.Sheets['My tile'];
          if (ws) {
            const rawD3 = ws['D3'] ? (ws['D3'].v || '') : '';
            const lines = rawD3.split('\n');
            const rest = lines.slice(2).join('\n');
            ws['D3'] = { t: 's', v: `${data.invoiceNo}\n${data.invoiceDate}\n${rest}` };
            ws['D10'] = { t: 'n', v: data.quantity };
            ws['G10'] = { t: 'n', v: data.amount };
            ws['G11'] = { t: 'n', v: data.taxableValue };
            ws['G14'] = { t: 'n', v: data.igst };
            ws['G15'] = { t: 'n', v: data.total };
            ws['A16'] = { t: 's', v: wordsText };
          }
        }
        const xlsxArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        if (currentXLSXUrl) URL.revokeObjectURL(currentXLSXUrl);
        currentXLSXBlob = new Blob([xlsxArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        currentXLSXUrl = URL.createObjectURL(currentXLSXBlob);
      }
    } catch (xlsxErr) {
      console.warn('XLSX generation warning:', xlsxErr);
    }

    // 4. Render crisp 2x retina preview canvas using PDF.js
    if (window.pdfjsLib && elements.resultPreviewCanvas) {
      try {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
        const loadingTask = pdfjsLib.getDocument({ data: finalPdfBytes });
        const pdf = await loadingTask.promise;
        const pageObj = await pdf.getPage(1);
        const viewport = pageObj.getViewport({ scale: 2.0 });

        const canvas = elements.resultPreviewCanvas;
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.display = 'block';
        if (elements.resultPreviewImg) elements.resultPreviewImg.style.display = 'none';

        const ctx = canvas.getContext('2d');
        await pageObj.render({ canvasContext: ctx, viewport }).promise;
      } catch (renderErr) {
        console.warn('PDF.js preview render warning:', renderErr);
      }
    }

    // 5. Update UI & Save last invoice number
    elements.resultFilename.textContent = `${currentInvoiceNo}.pdf`;
    localStorage.setItem(selectedBuyer.storageKey, currentInvoiceNo);

    // Share buttons are permanently visible
    if (elements.btnShare) elements.btnShare.style.display = 'flex';
    if (elements.btnWhatsApp) elements.btnWhatsApp.style.display = 'flex';

    showScreen('result');
  } catch (err) {
    console.error('Invoice generation failed:', err);
    alert(`Generation failed: ${err.message || err}. Please try again.`);
  } finally {
    showLoading(false);
  }
}

/* ===== Share & Download Actions ===== */
async function sharePDF() {
  const filename = `${currentInvoiceNo}.pdf`;
  const data = window._invoiceData || {};
  const buyerName = selectedBuyer ? selectedBuyer.name : 'Customer';
  const totalStr = data.total ? `₹${data.total.toLocaleString('en-IN')}` : '';
  const shareText = `Tax Invoice #${currentInvoiceNo} for ${buyerName} (Total: ${totalStr}) from Rudra Enterprises.`;

  // Native Web Share API (mobile Safari / Chrome on HTTPS)
  if (navigator.share) {
    try {
      if (currentPDFBlob && typeof File !== 'undefined') {
        const file = new File([currentPDFBlob], filename, { type: 'application/pdf' });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          await navigator.share({
            files: [file],
            title: `Tax Invoice ${currentInvoiceNo}`,
            text: shareText
          });
          return;
        }
      }

      await navigator.share({
        title: `Tax Invoice ${currentInvoiceNo}`,
        text: shareText,
        url: window.location.href
      });
      return;
    } catch (err) {
      if (err.name === 'AbortError') return;
      console.warn('Native share failed, falling back:', err);
    }
  }

  // Fallback for desktop / plain HTTP
  downloadPDF();
  const waConfirm = confirm(`Invoice ${filename} has been downloaded to your device!\n\nWould you like to open WhatsApp to share invoice details?`);
  if (waConfirm) {
    shareWhatsApp();
  }
}

function shareWhatsApp() {
  const data = window._invoiceData || {};
  const buyerName = selectedBuyer ? selectedBuyer.name : 'Customer';
  const totalStr = data.total ? `₹${data.total.toLocaleString('en-IN')}` : '';
  const dateStr = data.invoiceDate || '';
  const qtyStr = data.quantity ? `${data.quantity} Boxes` : '';

  // Download PDF automatically so user can attach it
  downloadPDF();

  const msg = `*RUDRA ENTERPRISES — TAX INVOICE*\n` +
    `─────────────────────────\n` +
    `*Invoice No:* #${currentInvoiceNo}\n` +
    `*Date:* ${dateStr}\n` +
    `*Buyer:* ${buyerName}\n` +
    `*Quantity:* ${qtyStr}\n` +
    `*Total Amount:* ${totalStr}\n` +
    `─────────────────────────\n` +
    `_Tax invoice PDF has been downloaded to your device._`;

  const waUrl = `https://api.whatsapp.com/send?text=${encodeURIComponent(msg)}`;
  window.open(waUrl, '_blank');
}

function downloadPDF() {
  const filename = `${currentInvoiceNo}.pdf`;
  if (!currentPDFBlob) return;
  const url = currentPDFUrl || URL.createObjectURL(currentPDFBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try { document.body.removeChild(a); } catch (e) {}
  }, 1000);
}

function downloadXLSX() {
  const filename = `${currentInvoiceNo}.xlsx`;
  if (!currentXLSXBlob) return;
  const url = currentXLSXUrl || URL.createObjectURL(currentXLSXBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try { document.body.removeChild(a); } catch (e) {}
  }, 1000);
}

/* ===== Navigation ===== */
function goBackToSelect() {
  showScreen('select');
  selectedBuyer = null;
}

function goBackToForm() {
  showScreen('form');
}

function startNewInvoice() {
  showScreen('select');
  selectedBuyer = null;
  currentPDFBlob = null;
  currentXLSXBlob = null;
  currentInvoiceNo = '';
  if (currentPDFUrl) URL.revokeObjectURL(currentPDFUrl);
  if (currentXLSXUrl) URL.revokeObjectURL(currentXLSXUrl);
  currentPDFUrl = '';
  currentXLSXUrl = '';
}

/* ===== Utilities ===== */
function showLoading(show, message) {
  if (show) {
    if (message) elements.loadingText.textContent = message;
    elements.loadingOverlay.classList.add('active');
  } else {
    elements.loadingOverlay.classList.remove('active');
  }
}

function shakeElement(el) {
  el.style.animation = 'none';
  el.offsetHeight; // reflow
  el.style.animation = 'shake 0.4s ease';
  setTimeout(() => { el.style.animation = ''; }, 400);
}

const shakeStyle = document.createElement('style');
shakeStyle.textContent = `
  @keyframes shake {
    0%, 100% { transform: translateX(0); }
    20% { transform: translateX(-6px); }
    40% { transform: translateX(6px); }
    60% { transform: translateX(-4px); }
    80% { transform: translateX(4px); }
  }
`;
document.head.appendChild(shakeStyle);
