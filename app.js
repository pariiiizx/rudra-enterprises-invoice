/**
 * Rudra Enterprises Invoice Generator — Core Application Logic
 * 100% Client-Side Generation with Authentic Excel Page Margins
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
  resultPreviewImg: document.getElementById('result-preview-img'),
  btnShare: document.getElementById('btn-share'),
  btnDownload: document.getElementById('btn-download'),
  btnDownloadExcel: document.getElementById('btn-download-excel'),

  // Loading
  loadingOverlay: document.getElementById('loading-overlay'),
  loadingText: document.getElementById('loading-text')
};

/* ===== Screen Management ===== */
function showScreen(screenId) {
  Object.values(screens).forEach(s => s.classList.remove('active'));
  screens[screenId].classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ===== Buyer Selection ===== */
function selectBuyer(buyerId) {
  selectedBuyer = BUYERS[buyerId];

  // Update form badge
  elements.buyerBadgeIcon.textContent = selectedBuyer.icon;
  elements.buyerBadgeName.textContent = selectedBuyer.name;

  // Auto-suggest invoice number
  const lastNum = localStorage.getItem(selectedBuyer.storageKey);
  if (lastNum) {
    const parsed = parseInt(lastNum, 10);
    const suggested = !isNaN(parsed) ? parsed + 1 : '';
    elements.invoiceNoInput.value = suggested;
    elements.invoiceNoHint.textContent = `Last used: ${lastNum}`;
  } else {
    elements.invoiceNoInput.value = '';
    elements.invoiceNoHint.textContent = 'No previous invoice found';
  }

  // Set today's date as default in dd/mm/yyyy
  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yyyy = today.getFullYear();
  elements.invoiceDateInput.value = `${dd}/${mm}/${yyyy}`;

  // Clear quantity
  elements.quantityInput.value = '';

  showScreen('form');
}

/* ===== Form Submission ===== */
function handleFormSubmit(e) {
  e.preventDefault();

  const invoiceNo = elements.invoiceNoInput.value.trim();
  const dateVal = elements.invoiceDateInput.value.trim();
  const quantity = parseInt(elements.quantityInput.value, 10);

  // Validation
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

  // Calculations
  const amount = quantity * 1000;
  const taxableValue = amount;
  const igst = taxableValue * 0.18;
  const total = taxableValue + igst;
  const words = numberToWordsIndian(total); // already contains '... Only'
  const amountInWords = words;

  // Store computed data
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
    amountInWords
  };

  // Populate preview screen
  elements.previewBuyer.textContent = selectedBuyer.name;
  elements.previewInvoiceNo.textContent = invoiceNo;
  elements.previewDate.textContent = dateVal;
  elements.previewQty.textContent = `${quantity} Box`;
  elements.previewAmount.textContent = `₹${amount.toLocaleString('en-IN')}`;
  elements.previewTaxable.textContent = `₹${taxableValue.toLocaleString('en-IN')}`;
  elements.previewIGST.textContent = `₹${igst.toLocaleString('en-IN')}`;
  elements.previewTotal.textContent = `₹${total.toLocaleString('en-IN')}`;
  elements.previewWords.textContent = `Amount charged in words : ${amountInWords}`;

  showScreen('preview');
}

/* ===== 100% Client-Side Invoice Generation with Proper Page Margins ===== */
async function generateInvoiceDirect() {
  showLoading(true, 'Rendering high-resolution invoice...');

  try {
    const data = window._invoiceData;
    const renderContainer = document.getElementById('invoice-render-container');
    if (!renderContainer) throw new Error('Render container missing');

    // 1. Generate HTML template
    const invoiceHTML = selectedBuyer.id === 'kaka' 
      ? generateKakaInvoiceHTML(data) 
      : generateMyTilesInvoiceHTML(data);

    renderContainer.innerHTML = invoiceHTML;

    // Small delay to ensure DOM and fonts settle
    await new Promise(r => setTimeout(r, 60));

    const invoiceElem = document.getElementById('invoice-doc') || renderContainer.firstElementChild;

    // 2. Render to high-DPI canvas via html2canvas
    const canvas = await html2canvas(invoiceElem, {
      scale: 2.5, // Crisp print resolution
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff'
    });

    const previewDataUrl = canvas.toDataURL('image/png');

    // 3. Generate A4 PDF via jsPDF with authentic page margins
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4',
      compress: true
    });

    // A4 sheet: 210mm wide x 297mm high
    // Authentic Excel margins: 16mm left/right, 18mm top
    const marginX = 16; // mm
    const marginY = 18; // mm
    const printWidth = 210 - (marginX * 2); // 178mm printable width
    const printHeight = (canvas.height * printWidth) / canvas.width;

    pdf.addImage(previewDataUrl, 'PNG', marginX, marginY, printWidth, printHeight);

    currentPDFBlob = pdf.output('blob');
    if (currentPDFUrl) URL.revokeObjectURL(currentPDFUrl);
    currentPDFUrl = URL.createObjectURL(currentPDFBlob);

    // 4. Generate modified Excel (.xlsx) file in browser via SheetJS
    try {
      const templateFilename = selectedBuyer.id === 'kaka' ? 'kaka tiles.xlsx' : 'My tiles.xlsx';
      const xlsxResponse = await fetch(templateFilename);
      if (xlsxResponse.ok) {
        const arrayBuf = await xlsxResponse.arrayBuffer();
        const wb = XLSX.read(arrayBuf, { type: 'array' });
        const wordsFormatted = `Amount charged in words : ${data.amountInWords}`;

        if (selectedBuyer.id === 'kaka') {
          const ws = wb.Sheets['Table 1'];
          if (ws) {
            ws['D5'] = { t: 's', v: String(data.invoiceNo) };
            ws['D7'] = { t: 's', v: String(data.invoiceDate) };
            ws['B18'] = { t: 's', v: 'Cermaic Tile' };
            ws['D18'] = { t: 'n', v: data.quantity };
            ws['G18'] = { t: 'n', v: data.amount };
            ws['G19'] = { t: 'n', v: data.taxableValue };
            ws['G22'] = { t: 'n', v: data.igst };
            ws['G23'] = { t: 'n', v: data.total };
            ws['A24'] = { t: 's', v: wordsFormatted };
          }
        } else {
          const ws = wb.Sheets['My tile'];
          if (ws) {
            const rawD3 = ws['D3'] ? String(ws['D3'].v || '') : '';
            const lines = rawD3.split('\n');
            if (lines.length >= 2) {
              lines[0] = String(data.invoiceNo);
              lines[1] = String(data.invoiceDate);
              ws['D3'] = { t: 's', v: lines.join('\n') };
            } else {
              ws['D3'] = { t: 's', v: `${data.invoiceNo}\n${data.invoiceDate}` };
            }
            ws['B10'] = { t: 's', v: 'Cermaic Tile' };
            ws['D10'] = { t: 'n', v: data.quantity };
            ws['G10'] = { t: 'n', v: data.amount };
            ws['G11'] = { t: 'n', v: data.taxableValue };
            ws['G14'] = { t: 'n', v: data.igst };
            ws['G15'] = { t: 'n', v: data.total };
            ws['A16'] = { t: 's', v: wordsFormatted };
          }
        }

        const outArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        currentXLSXBlob = new Blob([outArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        if (currentXLSXUrl) URL.revokeObjectURL(currentXLSXUrl);
        currentXLSXUrl = URL.createObjectURL(currentXLSXBlob);
      }
    } catch (xlsxErr) {
      console.warn('SheetJS in-browser Excel generation note:', xlsxErr);
    }

    // 5. Save last invoice number to localStorage
    localStorage.setItem(selectedBuyer.storageKey, currentInvoiceNo);

    // 6. Update Result Screen
    elements.resultFilename.textContent = `${currentInvoiceNo}.pdf`;
    elements.resultPreviewImg.src = previewDataUrl;

    // 7. Check Web Share API capability
    elements.btnShare.style.display = 'none';
    try {
      if (navigator.share) {
        if (currentPDFBlob && navigator.canShare && typeof File !== 'undefined') {
          const testFile = new File([currentPDFBlob], `${currentInvoiceNo}.pdf`, { type: 'application/pdf' });
          if (navigator.canShare({ files: [testFile] })) {
            elements.btnShare.style.display = 'flex';
          }
        } else {
          elements.btnShare.style.display = 'flex';
        }
      }
    } catch (shareCheckErr) {
      console.warn('Share check:', shareCheckErr);
    }

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
  try {
    if (navigator.share) {
      if (currentPDFBlob && typeof File !== 'undefined') {
        const file = new File([currentPDFBlob], filename, { type: 'application/pdf' });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          await navigator.share({
            files: [file],
            title: `Tax Invoice ${currentInvoiceNo}`,
            text: `Tax Invoice ${currentInvoiceNo} for ${selectedBuyer ? selectedBuyer.name : ''} from Rudra Enterprises`
          });
          return;
        }
      }
      // Fallback share URL
      await navigator.share({
        title: `Tax Invoice ${currentInvoiceNo}`,
        text: `Tax Invoice ${currentInvoiceNo} from Rudra Enterprises`,
        url: window.location.href
      });
    } else {
      downloadPDF();
    }
  } catch (err) {
    if (err.name !== 'AbortError') {
      console.error('Share failed:', err);
      downloadPDF();
    }
  }
}

function downloadPDF() {
  const filename = `${currentInvoiceNo}.pdf`;
  if (!currentPDFUrl && currentPDFBlob) {
    currentPDFUrl = URL.createObjectURL(currentPDFBlob);
  }
  const a = document.createElement('a');
  a.href = currentPDFUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    try { document.body.removeChild(a); } catch (e) {}
  }, 1000);
}

function downloadXLSX() {
  const filename = `${currentInvoiceNo}.xlsx`;
  if (!currentXLSXUrl && currentXLSXBlob) {
    currentXLSXUrl = URL.createObjectURL(currentXLSXBlob);
  }
  if (!currentXLSXUrl) {
    alert('Excel template could not be loaded.');
    return;
  }
  const a = document.createElement('a');
  a.href = currentXLSXUrl;
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
  el.offsetHeight; // force reflow
  el.style.animation = 'shake 0.4s ease';
  setTimeout(() => { el.style.animation = ''; }, 400);
}

// Keyframes for validation feedback
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
