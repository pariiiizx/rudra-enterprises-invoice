/**
 * Rudra Enterprises Invoice Generator — Template Engine
 * 1:1 Cell-by-Cell Excel Table Replication
 */

function formatIndianNumber(num) {
  if (isNaN(num) || num === null || num === undefined) return '0.00';
  return Number(num).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Generates the Kaka Tile Studio invoice HTML.
 * Matches kaka tiles.xlsx cell-by-cell with exact column widths and vertical alignment.
 */
function generateKakaInvoiceHTML(data) {
  const formattedAmount = formatIndianNumber(data.amount);
  const formattedTaxable = formatIndianNumber(data.taxableValue);
  const formattedIGST = formatIndianNumber(data.igst);
  const formattedTotal = formatIndianNumber(data.total);

  return `
    <div id="invoice-doc" class="invoice-sheet" style="
      width: 680px;
      margin: 0 auto;
      background: #ffffff;
      color: #000000;
      font-family: Arial, sans-serif;
      box-sizing: border-box;
      padding: 0;
      -webkit-font-smoothing: antialiased;
    ">
      <!-- Master Excel Table (Cols A to G) -->
      <table style="
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        border: 1.5px solid #000;
        box-sizing: border-box;
        background: #ffffff;
      ">
        <colgroup>
          <col style="width: 10.5%;"> <!-- A: Sl. No -->
          <col style="width: 28.5%;"> <!-- B: Description -->
          <col style="width: 15.5%;"> <!-- C: HSN & Labels -->
          <col style="width: 15.0%;"> <!-- D: Qty & Values -->
          <col style="width: 9.0%;">  <!-- E: Rate -->
          <col style="width: 10.0%;"> <!-- F: Per & Tax Labels -->
          <col style="width: 11.5%;"> <!-- G: Amount & Tax Values -->
        </colgroup>

        <tbody>
          <!-- Row 1: TAX INVOICE Header -->
          <tr style="height: 28px; border-bottom: 1px solid #000;">
            <td colspan="7" style="padding: 4px 10px; font-family: Arial, sans-serif; font-size: 10.5pt; font-weight: normal; vertical-align: middle;">
              TAX INVOICE
              <span style="float: right; font-weight: bold; font-size: 9.5pt;">1/ 2/ 3</span>
            </td>
          </tr>

          <!-- Row 2: Blank Spacer -->
          <tr style="height: 26px; border-bottom: 1px solid #000;">
            <td colspan="7"></td>
          </tr>

          <!-- Row 3: Seller Name Label -->
          <tr style="height: 20px;">
            <td colspan="2" style="font-size: 11.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; vertical-align: top;">
              Seller Name
            </td>
            <td colspan="5" style="border-bottom: 1px solid #000;"></td>
          </tr>

          <!-- Row 4: RUDRA ENTERPRISES -->
          <tr style="height: 18px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 2px 6px; border-right: 1px solid #000; vertical-align: top;">
              RUDRA ENTERPRISES
            </td>
            <td colspan="5"></td>
          </tr>

          <!-- Row 5: Invoice No -->
          <tr style="height: 18px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              AC-4C, Shalimar Bagh, Delhi-110088
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Invoice No.
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              ${data.invoiceNo}
            </td>
          </tr>

          <!-- Row 6: Spacer -->
          <tr style="height: 15px;">
            <td colspan="2" style="border-right: 1px solid #000;"></td>
            <td style="border-right: 1px solid #000;"></td>
            <td colspan="4"></td>
          </tr>

          <!-- Row 7: Invoice Date -->
          <tr style="height: 16px;">
            <td colspan="2" style="border-right: 1px solid #000;"></td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Invoice date
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              ${data.invoiceDate}
            </td>
          </tr>

          <!-- Row 8: Transporter Name -->
          <tr style="height: 18px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Contact No: 9971008585
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Transporter Name
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              Okara Parcel
            </td>
          </tr>

          <!-- Row 9: Payment Term -->
          <tr style="height: 18px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              GST IN: 07AAJPK3334Q1ZA
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Payment Term
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              45 Days
            </td>
          </tr>

          <!-- Row 10: Destination -->
          <tr style="height: 18px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              State Name: Delhi
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Destination
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;"></td>
          </tr>

          <!-- Row 11: Final Destination -->
          <tr style="height: 20px;">
            <td colspan="2" style="font-size: 10.5pt; padding: 1px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; vertical-align: top;">
              Email ID: naresh_kalra@yahoo.com
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Final Destination
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              Telengana
            </td>
          </tr>

          <!-- Row 12: Country of Origin -->
          <tr style="height: 18px;">
            <td colspan="2" rowspan="2" style="font-size: 11.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; vertical-align: top;">
              Buyer Name
            </td>
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Country of origin
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              India
            </td>
          </tr>

          <!-- Row 13: Loading Point -->
          <tr style="height: 18px;">
            <td style="font-size: 8.5pt; font-weight: bold; padding: 1px 6px; border-right: 1px solid #000; vertical-align: top;">
              Loading Point
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 1px 6px; vertical-align: top;">
              Delhi
            </td>
          </tr>

          <!-- Row 14: Kaka Tile Studio Address Block -->
          <tr style="height: 72px;">
            <td colspan="2" style="font-size: 10.5pt; font-weight: bold; padding: 3px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; vertical-align: top; line-height: 1.45;">
              Kaka Tile Studio<br>104 Sagar Enclave, Hyderabad, Telengana 500035
            </td>
            <td colspan="5" style="border-bottom: 1px solid #000;"></td>
          </tr>

          <!-- Row 15 & 16: GST NO & Inco Term -->
          <tr style="height: 20px;">
            <td style="font-size: 10.5pt; font-weight: bold; padding: 2px 6px; border-bottom: 1px solid #000; white-space: nowrap;">
              GST NO.
            </td>
            <td style="font-size: 10.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000;">
              36AAWFK9520H1ZS
            </td>
            <td style="font-size: 9.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000;">
              Inco Term
            </td>
            <td colspan="4" style="font-size: 8.5pt; padding: 2px 6px; border-bottom: 1px solid #000;">
              FOB
            </td>
          </tr>

          <!-- Row 17: Table Headers -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 20px; border-bottom: 1px solid #000;">
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">Sl. No</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">Description of Goods</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">HSN/SAC</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">Quantity</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Rate</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Per</td>
            <td style="padding: 2px 4px; text-align: left;">Amount</td>
          </tr>

          <!-- Row 18: Line Item with natural tall whitespace below text -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 110px; vertical-align: top; border-bottom: 1px solid #000;">
            <td style="padding: 3px 4px; border-right: 1px solid #000; text-align: center; font-weight: normal;">1</td>
            <td style="padding: 3px 4px; border-right: 1px solid #000;">Ceramic Tile</td>
            <td style="padding: 3px 4px; border-right: 1px solid #000;">69072300</td>
            <td style="padding: 3px 4px; border-right: 1px solid #000; text-align: center;">${data.quantity}</td>
            <td style="padding: 3px 4px; border-right: 1px solid #000; text-align: center;">1000</td>
            <td style="padding: 3px 4px; border-right: 1px solid #000; text-align: center;">Box</td>
            <td style="padding: 3px 4px; text-align: right;">${formattedAmount}</td>
          </tr>

          <!-- Row 19: Taxable Value -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 20px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center; white-space: nowrap;">Taxable Value</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedTaxable}</td>
          </tr>

          <!-- Row 20: CGST -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 20px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">CGST</td>
            <td style="padding: 2px 4px; text-align: right;">Nil</td>
          </tr>

          <!-- Row 21: SGST -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 20px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">SGST</td>
            <td style="padding: 2px 4px; text-align: right;">Nil</td>
          </tr>

          <!-- Row 22: IGST (18%) -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 20px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center; white-space: nowrap;">IGST (18%)</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedIGST}</td>
          </tr>

          <!-- Row 23: Total Amount & Total -->
          <tr style="font-size: 8.5pt; font-weight: bold; height: 28px; border-bottom: 1px solid #000;">
            <td style="border-right: 1px solid #000;"></td>
            <td style="padding: 3px 4px; border-right: 1px solid #000; font-size: 9.5pt;">Total Amount</td>
            <td colspan="3" style="border-right: 1px solid #000;"></td>
            <td style="padding: 3px 4px; border-right: 1px solid #000; text-align: right;">Total</td>
            <td style="padding: 3px 4px; text-align: right;">${formattedTotal}</td>
          </tr>

          <!-- Row 24: Amount in Words -->
          <tr style="border-bottom: 1px solid #000; height: 26px;">
            <td colspan="7" style="padding: 3px 6px; font-size: 11pt; text-align: center; font-weight: normal;">
              Amount charged in words : ${data.amountInWords}
            </td>
          </tr>

          <!-- Row 25 & 26: Bank Details & Declaration (Left), Signatory (Right) -->
          <tr>
            <td colspan="3" style="font-size: 9.5pt; padding: 4px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; height: 50px; vertical-align: top; line-height: 1.45;">
              Bank Name  : HDFC<br>
              Bank Account No.  50200004017434 RTGS/IFS Code :HDFC0000331
            </td>
            <td colspan="4" rowspan="2" style="font-size: 9.5pt; padding: 6px 12px; vertical-align: top; line-height: 1.4;">
              For Rudra Enterprises<br><br><br>
              Authorised Signatory
            </td>
          </tr>
          <tr>
            <td colspan="3" style="font-size: 9.5pt; padding: 4px 6px; border-right: 1px solid #000; height: 44px; vertical-align: top; line-height: 1.35;">
              Declaration<br>
              We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct
            </td>
          </tr>

        </tbody>
      </table>
    </div>
  `;
}

/**
 * Generates the My Tile Studio invoice HTML.
 * Matches My tiles.xlsx cell-by-cell with consolidated header layout.
 */
function generateMyTilesInvoiceHTML(data) {
  const formattedAmount = formatIndianNumber(data.amount);
  const formattedTaxable = formatIndianNumber(data.taxableValue);
  const formattedIGST = formatIndianNumber(data.igst);
  const formattedTotal = formatIndianNumber(data.total);

  return `
    <div id="invoice-doc" class="invoice-sheet" style="
      width: 680px;
      margin: 0 auto;
      background: #ffffff;
      color: #000000;
      font-family: Arial, sans-serif;
      box-sizing: border-box;
      padding: 0;
      -webkit-font-smoothing: antialiased;
    ">
      <!-- Master Excel Table (Cols A to G) -->
      <table style="
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        border: 1px solid #000;
        box-sizing: border-box;
        background: #ffffff;
      ">
        <colgroup>
          <col style="width: 11.0%;"> <!-- A: Sl. No -->
          <col style="width: 26.5%;"> <!-- B: Description -->
          <col style="width: 15.0%;"> <!-- C: Labels -->
          <col style="width: 16.0%;"> <!-- D: Values -->
          <col style="width: 9.5%;">  <!-- E: Rate -->
          <col style="width: 10.5%;"> <!-- F: Per & Tax Labels -->
          <col style="width: 11.5%;"> <!-- G: Amount & Tax Values -->
        </colgroup>

        <tbody>
          <!-- Row 1: TAX INVOICE Header -->
          <tr style="height: 22px; border-bottom: 1px solid #000;">
            <td colspan="7" style="padding: 2px 8px; font-family: 'Times New Roman', serif; font-size: 9.5pt; font-weight: normal; vertical-align: middle;">
              TAX INVOICE
              <span style="float: right; font-weight: bold; font-size: 9pt;">1/ 2/ 3</span>
            </td>
          </tr>

          <!-- Row 2: Blank Spacer -->
          <tr style="height: 20px; border-bottom: 1px solid #000;">
            <td colspan="7"></td>
          </tr>

          <!-- Row 3 & 4: Seller Info (Left) + Consolidated Details (Right) -->
          <tr>
            <td colspan="2" style="font-size: 8.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; vertical-align: top;">
              Seller Name
            </td>
            <td style="font-size: 8pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000; vertical-align: top; line-height: 1.6;" rowspan="2">
              Invoice No.<br>
              Invoice date Transporter Name<br>
              Payment Term<br>
              Final Destination<br>
              Country of origin Loading Point
            </td>
            <td colspan="4" style="font-size: 8pt; padding: 2px 6px; vertical-align: top; line-height: 1.6;" rowspan="2">
              ${data.invoiceNo}<br>
              ${data.invoiceDate}<br>
              Okara Parcel &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 45 Days<br>
              <br>
              Telengana &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; India<br>
              Delhi
            </td>
          </tr>
          <tr style="border-bottom: 1px solid #000;">
            <td colspan="2" style="font-size: 9.5pt; padding: 2px 6px; border-right: 1px solid #000; vertical-align: top; line-height: 1.45;">
              RUDRA ENTERPRISES<br>
              AC-4C, Shalimar Bagh, Delhi-110088 Contact No: 9971008585<br>
              GST IN: 07AAJPK3334Q1ZA<br>
              State Name: Delhi<br>
              Email ID: naresh_kalra@yahoo.com
            </td>
          </tr>

          <!-- Row 5: Buyer Name Label -->
          <tr style="height: 18px; border-bottom: 1px solid #000;">
            <td colspan="2" style="font-size: 8.5pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000;">
              Buyer Name
            </td>
            <td colspan="5" rowspan="2" style="border-bottom: 1px solid #000;"></td>
          </tr>

          <!-- Row 6: Buyer Details Block -->
          <tr style="height: 70px; border-bottom: 1px solid #000;">
            <td colspan="2" style="font-size: 9pt; padding: 2px 6px; border-right: 1px solid #000; vertical-align: top; line-height: 1.35;">
              My Tile Studio<br>
              MCH No. – 8-16-61/PLOT NO 6<br>
              Sagar Road<br>
              Hyderabad<br>
              Telengana<br>
              500079
            </td>
          </tr>

          <!-- Row 7 & 8: GST NO & Inco Term -->
          <tr style="height: 18px; border-bottom: 1px solid #000;">
            <td colspan="2" style="font-size: 8pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000;">
              GST NO. &nbsp;36BBCPK7265H1ZI
            </td>
            <td style="font-size: 7pt; font-weight: bold; padding: 2px 6px; border-right: 1px solid #000;">
              Inco Term
            </td>
            <td colspan="4" style="font-size: 6.5pt; padding: 2px 6px;">
              FOB
            </td>
          </tr>

          <!-- Row 9: Table Headers -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 18px; border-bottom: 1px solid #000;">
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">Sl. No</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">Description of Goods</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: right;">HSN/SAC</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Quantity</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Rate</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Per</td>
            <td style="padding: 2px 4px; text-align: left;">Amount</td>
          </tr>

          <!-- Row 10: Line Item -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 80px; vertical-align: top; border-bottom: 1px solid #000;">
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center; font-weight: normal;">1</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000;">Ceramic Tile</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: right;">69072300</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">${data.quantity}</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">1000</td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: center;">Box</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedAmount}</td>
          </tr>

          <!-- Row 11: Taxable Value -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 16px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: right; white-space: nowrap;">Taxable Value</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedTaxable}</td>
          </tr>

          <!-- Row 12: CGST -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 16px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">CGST</td>
            <td style="padding: 2px 4px; text-align: right;">Nil</td>
          </tr>

          <!-- Row 13: SGST -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 16px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left;">SGST</td>
            <td style="padding: 2px 4px; text-align: right;">Nil</td>
          </tr>

          <!-- Row 14: IGST (18%) -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 16px; border-bottom: 1px solid #000;">
            <td colspan="5" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: left; white-space: nowrap;">IGST (18%)</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedIGST}</td>
          </tr>

          <!-- Row 15: Total Amount & Total -->
          <tr style="font-size: 6.5pt; font-weight: bold; height: 22px; border-bottom: 1px solid #000;">
            <td style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; font-size: 7pt;">Total Amount</td>
            <td colspan="3" style="border-right: 1px solid #000;"></td>
            <td style="padding: 2px 4px; border-right: 1px solid #000; text-align: right;">Total</td>
            <td style="padding: 2px 4px; text-align: right;">${formattedTotal}</td>
          </tr>

          <!-- Row 16: Words -->
          <tr style="border-bottom: 1px solid #000; height: 22px;">
            <td colspan="7" style="padding: 3px 6px; font-size: 9.5pt; text-align: center; font-weight: normal;">
              Amount charged in words : ${data.amountInWords}
            </td>
          </tr>

          <!-- Row 17 & 18: Bank Details, Signatory, Declaration -->
          <tr>
            <td colspan="3" style="font-size: 9.5pt; padding: 3px 6px; border-right: 1px solid #000; border-bottom: 1px solid #000; height: 44px; vertical-align: top; line-height: 1.35;">
              Bank Name  : HDFC<br>
              Bank Account No.  50200004017434 RTGS/IFS Code :HDFC0000331
            </td>
            <td colspan="4" rowspan="2" style="font-size: 9.5pt; padding: 4px 8px; vertical-align: top; line-height: 1.35;">
              For Rudra Enterprises<br><br><br>
              Authorised Signatory
            </td>
          </tr>
          <tr>
            <td colspan="3" style="font-size: 9.5pt; padding: 3px 6px; border-right: 1px solid #000; height: 36px; vertical-align: top; line-height: 1.3;">
              Declaration<br>
              We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct
            </td>
          </tr>

        </tbody>
      </table>
    </div>
  `;
}
