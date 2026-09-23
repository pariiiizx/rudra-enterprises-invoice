/**
 * Rudra Enterprises Invoice Generator — Template Engine
 * Produces pixel-perfect HTML/CSS representations matching the original Excel invoice templates.
 */

function formatIndianNumber(num) {
  if (isNaN(num) || num === null || num === undefined) return '0.00';
  return Number(num).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/**
 * Generates the Kaka Tile Studio invoice HTML matching kaka tiles.xlsx exactly.
 */
function generateKakaInvoiceHTML(data) {
  const formattedAmount = formatIndianNumber(data.amount);
  const formattedTaxable = formatIndianNumber(data.taxableValue);
  const formattedIGST = formatIndianNumber(data.igst);
  const formattedTotal = formatIndianNumber(data.total);

  return `
    <div id="invoice-doc" class="invoice-sheet kaka-theme" style="
      width: 780px;
      margin: 0 auto;
      background: #ffffff;
      color: #000000;
      font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
      font-size: 12px;
      line-height: 1.25;
      box-sizing: border-box;
      padding: 0;
      -webkit-font-smoothing: antialiased;
    ">
      <!-- Main Outer Border Box -->
      <div style="border: 2px solid #000000; box-sizing: border-box;">
        
        <!-- Row 1: Header -->
        <div style="border-bottom: 1.5px solid #000; padding: 5px 12px; text-align: center; position: relative;">
          <span style="font-size: 15px; font-weight: bold; letter-spacing: 0.5px;">TAX INVOICE</span>
          <span style="position: absolute; right: 14px; top: 6px; font-size: 11px; font-weight: bold;">1/ 2/ 3</span>
        </div>

        <!-- Row 2: Spacer -->
        <div style="height: 12px; border-bottom: 1.5px solid #000; background: #fff;"></div>

        <!-- Rows 3-14: Seller & Buyer (Left) + Invoice Fields (Right) -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          
          <!-- Left Column: Seller & Buyer Info (48%) -->
          <div style="flex: 0 0 48%; border-right: 1.5px solid #000; box-sizing: border-box;">
            
            <!-- Seller Name Label -->
            <div style="padding: 3px 8px; font-size: 13px; font-weight: bold; border-bottom: 1px solid #000;">
              Seller Name
            </div>
            
            <!-- Seller Details -->
            <div style="padding: 4px 8px; font-size: 12px; line-height: 1.45; border-bottom: 1px solid #000;">
              <div style="font-weight: bold; font-size: 12px;">RUDRA ENTERPRISES</div>
              <div>AC-4C, Shalimar Bagh, Delhi-110088</div>
              <div>Contact No: 9971008585</div>
              <div>GST IN: 07AAJPK3334Q1ZA</div>
              <div>State Name: Delhi</div>
              <div>Email ID: naresh_kalra@yahoo.com</div>
            </div>

            <!-- Buyer Name Label -->
            <div style="padding: 3px 8px; font-size: 13px; font-weight: bold; border-bottom: 1px solid #000;">
              Buyer Name
            </div>

            <!-- Buyer Details -->
            <div style="padding: 4px 8px; font-size: 12px; line-height: 1.45; min-height: 58px;">
              <div style="font-weight: bold;">Kaka Tile Studio</div>
              <div>104 Sagar Enclave, Hyderabad, Telengana 500035</div>
            </div>
          </div>

          <!-- Right Column: Meta Fields (52%) -->
          <div style="flex: 1; display: flex; flex-direction: column;">
            
            <!-- Empty Row aligning with Seller Name -->
            <div style="height: 24px; border-bottom: 1px solid #000;"></div>

            <!-- Invoice No -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Invoice No.</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">${data.invoiceNo}</div>
            </div>

            <!-- Spacer aligning with seller address -->
            <div style="height: 16px; border-bottom: 1px solid #000;"></div>

            <!-- Invoice Date -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Invoice date</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">${data.invoiceDate}</div>
            </div>

            <!-- Transporter Name -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Transporter Name</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">Okara Parcel</div>
            </div>

            <!-- Payment Term -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Payment Term</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">45 Days</div>
            </div>

            <!-- Destination -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Destination</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;"></div>
            </div>

            <!-- Final Destination -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Final Destination</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">Telengana</div>
            </div>

            <!-- Country of origin -->
            <div style="display: flex; border-bottom: 1px solid #000;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Country of origin</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">India</div>
            </div>

            <!-- Loading Point -->
            <div style="display: flex;">
              <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; border-right: 1.5px solid #000; font-size: 11px;">Loading Point</div>
              <div style="flex: 1; padding: 3px 8px; font-size: 11px;">Delhi</div>
            </div>

          </div>
        </div>

        <!-- Row 15-16: GST NO & Inco Term -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          <div style="flex: 0 0 48%; padding: 3px 8px; font-weight: bold; font-size: 12px; border-right: 1.5px solid #000; box-sizing: border-box;">
            GST NO.&nbsp;&nbsp;&nbsp;36AAWFK9520H1ZS
          </div>
          <div style="flex: 1; display: flex;">
            <div style="flex: 0 0 45%; padding: 3px 8px; font-weight: bold; font-size: 11px; border-right: 1.5px solid #000;">Inco Term</div>
            <div style="flex: 1; padding: 3px 8px; font-size: 11px;">FOB</div>
          </div>
        </div>

        <!-- Table: Items & Pricing -->
        <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 11px;">
          <thead>
            <tr style="border-bottom: 1.5px solid #000; font-weight: bold;">
              <th style="width: 7%; padding: 4px 6px; border-right: 1px solid #000; text-align: left;">Sl. No</th>
              <th style="width: 25%; padding: 4px 6px; border-right: 1px solid #000; text-align: left;">Description of Goods</th>
              <th style="width: 14%; padding: 4px 6px; border-right: 1px solid #000; text-align: center;">HSN/SAC</th>
              <th style="width: 14%; padding: 4px 6px; border-right: 1px solid #000; text-align: center;">Quantity</th>
              <th style="width: 10%; padding: 4px 6px; border-right: 1px solid #000; text-align: center;">Rate</th>
              <th style="width: 10%; padding: 4px 6px; border-right: 1px solid #000; text-align: center;">Per</th>
              <th style="width: 20%; padding: 4px 6px; text-align: right;">Amount</th>
            </tr>
          </thead>
          <tbody>
            <!-- Line Item Row -->
            <tr style="font-weight: bold;">
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">1</td>
              <td style="padding: 4px 6px; border-right: 1px solid #000;">Ceramic Tile</td>
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">69072300</td>
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">${data.quantity}</td>
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">1000</td>
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">Box</td>
              <td style="padding: 4px 6px; text-align: right;">${formattedAmount}</td>
            </tr>

            <!-- Empty Vertical Space preserving table lines -->
            <tr style="height: 100px;">
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td></td>
            </tr>

            <!-- Taxable Value -->
            <tr style="border-top: 1px solid #000; border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">Taxable Value</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">${formattedTaxable}</td>
            </tr>

            <!-- CGST -->
            <tr style="border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">CGST</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">Nil</td>
            </tr>

            <!-- SGST -->
            <tr style="border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">SGST</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">Nil</td>
            </tr>

            <!-- IGST -->
            <tr style="border-bottom: 1.5px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">IGST (18%)</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">${formattedIGST}</td>
            </tr>

            <!-- Total -->
            <tr style="border-bottom: 1.5px solid #000; font-weight: bold;">
              <td style="border-right: 1px solid #000;"></td>
              <td style="padding: 4px 6px; border-right: 1px solid #000;">Total Amount</td>
              <td colspan="3" style="border-right: 1px solid #000;"></td>
              <td style="padding: 4px 6px; border-right: 1px solid #000; text-align: center;">Total</td>
              <td style="padding: 4px 6px; text-align: right;">${formattedTotal}</td>
            </tr>
          </tbody>
        </table>

        <!-- Amount charged in words -->
        <div style="border-bottom: 1.5px solid #000; padding: 6px 12px; font-size: 13px; font-weight: bold; text-align: center;">
          Amount charged in words : ${data.amountInWords}
        </div>

        <!-- Bank Details & Signatory -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          <div style="flex: 0 0 52%; padding: 6px 8px; font-size: 11px; border-right: 1.5px solid #000; box-sizing: border-box; min-height: 48px;">
            <div>Bank Name  : HDFC</div>
            <div>Bank Account No.  50200004017434 RTGS/IFS Code :HDFC0000331</div>
          </div>
          <div style="flex: 1; padding: 6px 10px; font-size: 11px;">
            <div style="font-weight: bold;">For Rudra Enterprises</div>
            <div style="height: 28px;"></div>
            <div>Authorised Signatory</div>
          </div>
        </div>

        <!-- Declaration -->
        <div style="padding: 6px 8px; font-size: 11px; line-height: 1.35;">
          <div style="font-weight: bold;">Declaration</div>
          <div>We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct</div>
        </div>

      </div>

      <!-- Footer Outside Border -->
      <div style="padding: 4px 0 2px 0; font-size: 11px; font-weight: bold; text-align: center;">
        This is a computer generated invoice
      </div>
    </div>
  `;
}

/**
 * Generates the My Tile Studio invoice HTML matching My tiles.xlsx exactly.
 */
function generateMyTilesInvoiceHTML(data) {
  const formattedAmount = formatIndianNumber(data.amount);
  const formattedTaxable = formatIndianNumber(data.taxableValue);
  const formattedIGST = formatIndianNumber(data.igst);
  const formattedTotal = formatIndianNumber(data.total);

  return `
    <div id="invoice-doc" class="invoice-sheet my-theme" style="
      width: 780px;
      margin: 0 auto;
      background: #ffffff;
      color: #000000;
      font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
      font-size: 12px;
      line-height: 1.25;
      box-sizing: border-box;
      padding: 0;
      -webkit-font-smoothing: antialiased;
    ">
      <!-- Main Outer Border Box -->
      <div style="border: 2px solid #000000; box-sizing: border-box;">
        
        <!-- Row 1: Header -->
        <div style="border-bottom: 1.5px solid #000; padding: 5px 12px; text-align: center; position: relative;">
          <span style="font-size: 15px; font-weight: bold; letter-spacing: 0.5px;">TAX INVOICE</span>
          <span style="position: absolute; right: 14px; top: 6px; font-size: 11px; font-weight: bold;">1/ 2/ 3</span>
        </div>

        <!-- Row 2: Spacer -->
        <div style="height: 10px; border-bottom: 1.5px solid #000; background: #fff;"></div>

        <!-- Rows 3-7: Seller & Buyer (Left) + Consolidated Details (Right) -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          
          <!-- Left Column (38%) -->
          <div style="flex: 0 0 38%; border-right: 1.5px solid #000; box-sizing: border-box;">
            
            <!-- Seller Name Label -->
            <div style="padding: 2px 6px; font-size: 12px; font-weight: bold; border-bottom: 1px solid #000;">
              Seller Name
            </div>
            
            <!-- Seller Details Block -->
            <div style="padding: 4px 6px; font-size: 11px; line-height: 1.45; border-bottom: 1px solid #000;">
              <div style="font-weight: bold;">RUDRA ENTERPRISES</div>
              <div>AC-4C, Shalimar Bagh, Delhi-110088 Contact No: 9971008585</div>
              <div>GST IN: 07AAJPK3334Q1ZA</div>
              <div>State Name: Delhi</div>
              <div>Email ID: naresh_kalra@yahoo.com</div>
            </div>

            <!-- Buyer Name Label -->
            <div style="padding: 2px 6px; font-size: 12px; font-weight: bold; border-bottom: 1px solid #000;">
              Buyer Name
            </div>

            <!-- Buyer Details Block -->
            <div style="padding: 4px 6px; font-size: 11px; line-height: 1.35; min-height: 60px;">
              <div style="font-weight: bold;">My Tile Studio</div>
              <div>MCH No. – 8-16-61/PLOT NO 6</div>
              <div>Sagar Road</div>
              <div>Hyderabad</div>
              <div>Telengana</div>
              <div>500079</div>
            </div>

          </div>

          <!-- Right Column (62%): Labels & Values -->
          <div style="flex: 1; display: flex; flex-direction: column;">
            
            <!-- Details Grid -->
            <div style="display: flex; flex: 1;">
              
              <!-- Labels Column -->
              <div style="flex: 0 0 35%; padding: 4px 8px; font-size: 11px; font-weight: bold; line-height: 1.7; border-right: 1.5px solid #000;">
                <div>Invoice No.</div>
                <div>Invoice date</div>
                <div>Transporter Name</div>
                <div>Payment Term</div>
                <div>Final Destination</div>
                <div>Country of origin</div>
                <div>Loading Point</div>
              </div>

              <!-- Values Column -->
              <div style="flex: 1; padding: 4px 8px; font-size: 11px; line-height: 1.7;">
                <div>${data.invoiceNo}</div>
                <div>${data.invoiceDate}</div>
                <div>Okara Parcel</div>
                <div>45 Days</div>
                <div>Telengana</div>
                <div>India</div>
                <div>Delhi</div>
              </div>

            </div>

          </div>

        </div>

        <!-- Row 8: GST NO & Inco Term -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          <div style="flex: 0 0 38%; padding: 3px 6px; font-weight: bold; font-size: 11px; border-right: 1.5px solid #000; box-sizing: border-box;">
            GST NO.&nbsp;&nbsp;36BBCPK7265H1ZI
          </div>
          <div style="flex: 1; display: flex;">
            <div style="flex: 0 0 35%; padding: 3px 8px; font-weight: bold; font-size: 11px; border-right: 1.5px solid #000;">Inco Term</div>
            <div style="flex: 1; padding: 3px 8px; font-size: 11px;">FOB</div>
          </div>
        </div>

        <!-- Table: Items & Pricing -->
        <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 11px;">
          <thead>
            <tr style="border-bottom: 1.5px solid #000; font-weight: bold;">
              <th style="width: 7%; padding: 3px 6px; border-right: 1px solid #000; text-align: left;">Sl. No</th>
              <th style="width: 25%; padding: 3px 6px; border-right: 1px solid #000; text-align: left;">Description of Goods</th>
              <th style="width: 14%; padding: 3px 6px; border-right: 1px solid #000; text-align: center;">HSN/SAC</th>
              <th style="width: 14%; padding: 3px 6px; border-right: 1px solid #000; text-align: center;">Quantity</th>
              <th style="width: 10%; padding: 3px 6px; border-right: 1px solid #000; text-align: center;">Rate</th>
              <th style="width: 10%; padding: 3px 6px; border-right: 1px solid #000; text-align: center;">Per</th>
              <th style="width: 20%; padding: 3px 6px; text-align: right;">Amount</th>
            </tr>
          </thead>
          <tbody>
            <!-- Line Item Row -->
            <tr style="font-weight: bold;">
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">1</td>
              <td style="padding: 3px 6px; border-right: 1px solid #000;">Ceramic Tile</td>
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">69072300</td>
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">${data.quantity}</td>
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">1000</td>
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">Box</td>
              <td style="padding: 3px 6px; text-align: right;">${formattedAmount}</td>
            </tr>

            <!-- Empty Vertical Space -->
            <tr style="height: 95px;">
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td style="border-right: 1px solid #000;"></td>
              <td></td>
            </tr>

            <!-- Taxable Value -->
            <tr style="border-top: 1px solid #000; border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">Taxable Value</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">${formattedTaxable}</td>
            </tr>

            <!-- CGST -->
            <tr style="border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">CGST</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">Nil</td>
            </tr>

            <!-- SGST -->
            <tr style="border-bottom: 1px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">SGST</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">Nil</td>
            </tr>

            <!-- IGST -->
            <tr style="border-bottom: 1.5px solid #000;">
              <td colspan="5" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; font-weight: bold; border-right: 1px solid #000; text-align: center;">IGST (18%)</td>
              <td style="padding: 3px 6px; font-weight: bold; text-align: right;">${formattedIGST}</td>
            </tr>

            <!-- Total -->
            <tr style="border-bottom: 1.5px solid #000; font-weight: bold;">
              <td style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; border-right: 1px solid #000;">Total Amount</td>
              <td colspan="3" style="border-right: 1px solid #000;"></td>
              <td style="padding: 3px 6px; border-right: 1px solid #000; text-align: center;">Total</td>
              <td style="padding: 3px 6px; text-align: right;">${formattedTotal}</td>
            </tr>
          </tbody>
        </table>

        <!-- Amount charged in words -->
        <div style="border-bottom: 1.5px solid #000; padding: 6px 12px; font-size: 13px; font-weight: bold; text-align: center;">
          Amount charged in words : ${data.amountInWords}
        </div>

        <!-- Bank Details & Signatory -->
        <div style="display: flex; border-bottom: 1.5px solid #000;">
          <div style="flex: 0 0 52%; padding: 6px 8px; font-size: 11px; border-right: 1.5px solid #000; box-sizing: border-box; min-height: 48px;">
            <div>Bank Name  : HDFC</div>
            <div>Bank Account No.  50200004017434 RTGS/IFS Code :HDFC0000331</div>
          </div>
          <div style="flex: 1; padding: 6px 10px; font-size: 11px;">
            <div style="font-weight: bold;">For Rudra Enterprises</div>
            <div style="height: 28px;"></div>
            <div>Authorised Signatory</div>
          </div>
        </div>

        <!-- Declaration -->
        <div style="padding: 6px 8px; font-size: 11px; line-height: 1.35;">
          <div style="font-weight: bold;">Declaration</div>
          <div>We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct</div>
        </div>

      </div>

      <!-- Footer Outside Border -->
      <div style="padding: 4px 0 2px 0; font-size: 11px; font-weight: bold; text-align: center;">
        This is a computer generated invoice
      </div>
    </div>
  `;
}
