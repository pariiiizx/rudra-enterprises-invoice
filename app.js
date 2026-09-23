/**
 * Rudra Enterprises Invoice Generator — Core Application Logic
 * Native Excel Automation with In-Browser Fallback
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
  const words = numberToWordsIndian(total); // contains '... Only'
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

/* ===== Authentic Invoice Generation (Excel Backend with Client-Side Fallback) ===== */
async function generateInvoiceDirect() {
  showLoading(true, 'Updating Excel workbook & Exporting PDF...');

  try {
    const data = window._invoiceData;

    // Try native Excel server first
    let generatedViaServer = false;
    try {
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          buyer: data.buyer,
          invoiceNumber: data.invoiceNo,
          invoiceDate: data.invoiceDate,
          quantity: data.quantity
        })
      });

      if (response.ok) {
        const result = await response.json();
        if (result.success) {
          // Convert base64 PDF into Blob for sharing
          try {
            const byteCharacters = atob(result.pdfBase64);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
              byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            const byteArray = new Uint8Array(byteNumbers);
            currentPDFBlob = new Blob([byteArray], { type: 'application/pdf' });
          } catch (b64Err) {
            console.warn('Base64 decode warning:', b64Err);
            currentPDFBlob = null;
          }

          currentPDFUrl = result.pdfUrl;
          currentXLSXUrl = result.xlsxUrl;

          elements.resultFilename.textContent = result.fileName;
          elements.resultPreviewImg.src = `${result.previewUrl}?t=${Date.now()}`;
          generatedViaServer = true;
        }
      }
    } catch (serverErr) {
      console.log('Native Excel server not reachable, using in-browser engine:', serverErr);
    }

    // Fallback: In-browser generation if server is offline
    if (!generatedViaServer) {
      const renderContainer = document.getElementById('invoice-render-container');
      const invoiceHTML = selectedBuyer.id === 'kaka' 
        ? generateKakaInvoiceHTML(data) 
        : generateMyTilesInvoiceHTML(data);

      renderContainer.innerHTML = invoiceHTML;
      await new Promise(r => setTimeout(r, 60));

      const invoiceElem = document.getElementById('invoice-doc') || renderContainer.firstElementChild;
      const canvas = await html2canvas(invoiceElem, {
        scale: 2.5,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });

      const previewDataUrl = canvas.toDataURL('image/png');
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4',
        compress: true
      });

      const marginX = 16;
      const marginY = 18;
      const printWidth = 210 - (marginX * 2);
      const printHeight = (canvas.height * printWidth) / canvas.width;
      pdf.addImage(previewDataUrl, 'PNG', marginX, marginY, printWidth, printHeight);

      currentPDFBlob = pdf.output('blob');
      if (currentPDFUrl) URL.revokeObjectURL(currentPDFUrl);
      currentPDFUrl = URL.createObjectURL(currentPDFBlob);

      elements.resultFilename.textContent = `${currentInvoiceNo}.pdf`;
      elements.resultPreviewImg.src = previewDataUrl;
    }

    // Save last invoice number to localStorage
    localStorage.setItem(selectedBuyer.storageKey, currentInvoiceNo);

    // Check Web Share API capability
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
  const url = currentPDFUrl || `/generated/${filename}`;
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
  const url = currentXLSXUrl || `/generated/${filename}`;
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
