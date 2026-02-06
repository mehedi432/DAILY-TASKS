# DAILY-TASKS
A REPO FOR COPYING HOME SERVER DATA TO OF SERVER
Scrpt for automatically fetch and arrange production report with MD Sir & CM Sir Format.
```
function main(workbook: ExcelScript.Workbook) {
  let sourceSheet = workbook.getWorksheet("Today,Total & Balance");
  let dailySummary = workbook.getWorksheet("Daily Summary");

  if (!sourceSheet) return;
  if (!dailySummary) {
    dailySummary = workbook.addWorksheet("Daily Summary");
  } else {
    dailySummary.getUsedRange()?.clear();
  }

  // Set solid white background
  dailySummary.getRange("A1:Z2000").getFormat().getFill().setColor("FFFFFF");

  let sourceValues: (string | number | boolean)[][] = sourceSheet.getUsedRange().getValues();
  let buyers: string[] = ["PAAYRA", "PEPCO"];
  let finalReportRows: (string | number | boolean)[][] = [];
  let slNo: number = 1;

  // Global Totals for Grand Total
  let gQty = 0, gKnit = 0, gLink = 0, gMend = 0, gWash = 0, gPoly = 0;

  // 1. DATA GATHERING
  buyers.forEach((buyer: string) => {
    let buyerData = sourceValues.filter(row =>
      String(row[0]).toUpperCase().indexOf(buyer.toUpperCase()) !== -1
    );

    if (buyerData.length > 0) {
      let bQty = 0, bKnit = 0, bLink = 0, bMend = 0, bWash = 0, bPoly = 0;

      buyerData.forEach((row) => {
        let qty = Number(row[3]) || 0;
        let knit = Number(row[7]) || 0;
        let link = Number(row[11]) || 0;
        let mend = Number(row[19]) || 0;
        let wash = Number(row[23]) || 0;
        let poly = Number(row[27]) || 0;

        finalReportRows.push([
          slNo++, "", String(row[2]), String(row[1]), String(row[4]), "50% COT/ACR",
          qty, "29.03.26", "APPROVED", String(row[5]), knit, link, mend, wash, poly,
          "YES", "DONE", qty > 0 ? Math.round((poly / qty) * 100) : 0
        ]);

        bQty += qty; bKnit += knit; bLink += link; bMend += mend; bWash += wash; bPoly += poly;
      });

      // ALIGNED BUYER TOTAL ROW
      let tRow: (string | number | boolean)[] = new Array(18).fill("");
      tRow[0] = buyer + " TOTAL";
      tRow[6] = bQty; tRow[10] = bKnit; tRow[11] = bLink; tRow[12] = bMend; tRow[13] = bWash; tRow[14] = bPoly;
      tRow[17] = "BAL: " + (bQty - bPoly);
      finalReportRows.push(tRow);

      // Add to Grand Totals
      gQty += bQty; gKnit += bKnit; gLink += bLink; gMend += bMend; gWash += bWash; gPoly += bPoly;
    }
  });

  // ADD GRAND TOTAL ROW AT THE VERY END
  let grandTotalRow: (string | number | boolean)[] = new Array(18).fill("");
  grandTotalRow[0] = "GRAND TOTAL (ALL BUYERS)";
  grandTotalRow[6] = gQty; grandTotalRow[10] = gKnit; grandTotalRow[11] = gLink; grandTotalRow[12] = gMend; grandTotalRow[13] = gWash; grandTotalRow[14] = gPoly;
  grandTotalRow[17] = "TOTAL BAL: " + (gQty - gPoly);
  finalReportRows.push(grandTotalRow);

  // 2. DASHBOARD TOP SECTION
  let titleRange = dailySummary.getRange("A1");
  titleRange.setValue("EXECUTIVE PRODUCTION DASHBOARD");
  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setSize(16);

  dailySummary.getRange("A2").setValue("TOTAL PACKED:");
  let packVal = dailySummary.getRange("B2");
  packVal.setValue(gQty > 0 ? gPoly / gQty : 0);
  packVal.setNumberFormat("0.0%");

  // 3. MAIN TABLE HEADERS
  const headers: string[][] = [["SL No", "IMAGE", "FACTORY", "STYLE", "GG", "COMP", "ORDER QTY", "EX-FTY", "PPS", "M/C", "KNIT", "LINK", "MEND", "WASH", "POLY", "L/C", "DYEING", "PROGRESS %"]];
  let headerRange = dailySummary.getRange("A4:R4");
  headerRange.setValues(headers);
  headerRange.getFormat().getFill().setColor("#2D2D2D");
  headerRange.getFormat().getFont().setColor("FFFFFF");
  headerRange.getFormat().getFont().setBold(true);

  // 4. DATA INSERTION AND VISIBLE LINES
  if (finalReportRows.length > 0) {
    let dataRange = dailySummary.getRange("A5").getResizedRange(finalReportRows.length - 1, 17);
    dataRange.setValues(finalReportRows);

    for (let i = 0; i < finalReportRows.length; i++) {
      let rIdx = i + 5;
      let rowRange = dailySummary.getRange(`A${rIdx}:R${rIdx}`);
      
      // Apply Grid Borders (Inside Horizontal and Vertical)
      let hBorders = rowRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal);
      hBorders.setStyle(ExcelScript.BorderLineStyle.continuous);
      hBorders.setColor("#BDBDBD");
      
      let vBorders = rowRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical);
      vBorders.setStyle(ExcelScript.BorderLineStyle.continuous);
      vBorders.setColor("#BDBDBD");

      // Shade Sub-Totals and Grand Total
      let firstCell = String(finalReportRows[i][0]);
      if (firstCell.indexOf("TOTAL") !== -1) {
        rowRange.getFormat().getFill().setColor("#F2F2F2");
        rowRange.getFormat().getFont().setBold(true);
      }
      if (firstCell.indexOf("GRAND TOTAL") !== -1) {
        rowRange.getFormat().getFill().setColor("#D9EAD3"); // Light Green for Grand Total
      }
    }
    
    dailySummary.getRange("R5").getResizedRange(finalReportRows.length - 1, 0).setNumberFormat("0'%'");
  }

  // --- THE CORRECT FIX FOR COLUMN WIDTH ---
  dailySummary.getUsedRange().getFormat().autofitColumns();
  dailySummary.getRange("B:B").getFormat().setColumnWidth(70); // Access .getFormat() first
  
  // Set Alignment
  let fullTable = dailySummary.getRange("A4").getSurroundingRegion();
  fullTable.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}
```

# Sample Yarn Tag
```
<style>
@page {
  size: A4;
}
body {
  -webkit-print-color-adjust: exact;
}
</style>

<div style="width:100%; font-family:'Inter','Segoe UI',Arial,sans-serif; padding:13px;">

  <div style="max-width:100%; margin:auto; border:1px solid #111; border-radius:13px; padding:13px; background:#fff;">

    <!-- HEADER -->
<div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:18px;">

  <!-- LOGO (LEFT) -->
  <div>
    <img src="/files/logo_meek-01.png" style="height:55px; max-width:89px;">
  </div>

  <!-- TITLE (CENTER) -->
  <div style="text-align:center; flex:1;">
    <div style="font-size:34px; font-weight:800; letter-spacing:4px;">
      YARN TAG
    </div>
    <div style="width:34%; height:1.6px; background:#111; margin:7px auto 0;"></div>
  </div>

  <!-- QR (RIGHT) -->
    <div>
      <img
        src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=http://192.168.0.102:8003/app/daily-sample-request/{{ doc.name }}"
        style="width:55px; height:55px;"
        alt="QR">
    </div>


</div>


    <!-- INFO -->
    <table style="width:100%; border-collapse:collapse;">
        
      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555; width:30%;">BUYER</td>
        <td style="font-size:21px; padding:10px 6px; font-weight:800;">
          {{ doc.buyer or "" }}
        </td>
      </tr>
      
      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555; width:30%;">STYLE</td>
        <td style="font-size:21px; padding:10px 6px; font-weight:800;">
          {{ doc.style or "" }}
        </td>
      </tr>

      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555;">COMPOSITION</td>
        <td style="font-size:21px; padding:14px 6px;">
          {% for y in doc.raw_material %}
            {{ y.yarn_composition }} - {{ y.yarn_count }}{% if not loop.last %}, {% endif %}
          {% endfor %}
        </td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">SDO NO</td><br/>
        <td style="font-size:21px; padding:22px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">DYE HOUSE</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">LOT NO</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>
      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">QTY.</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

    </table>

    <!-- FOOT -->
    <div style="margin-top:18px; display:flex; justify-content:space-between; font-size:9px; color:#888;">
      <div>RECEIVE DATE: {{ frappe.utils.formatdate(frappe.utils.nowdate(),"dd MMM yyyy") }}</div>
      <div>{{ doc.name }}</div>
    </div>

  </div>

</div>

<div style="width:100%; font-family:'Inter','Segoe UI',Arial,sans-serif; padding:13px;">

  <div style="max-width:100%; margin:auto; border:1px solid #111; border-radius:13px; padding:13px; background:#fff;">

    <!-- HEADER -->
    <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:18px;">
    
      <!-- LOGO (LEFT) -->
      <div>
        <img src="/files/logo_meek-01.png" style="height:55px; max-width:89px;">
      </div>
    
      <!-- TITLE (CENTER) -->
      <div style="text-align:center; flex:1;">
        <div style="font-size:34px; font-weight:800; letter-spacing:4px;">
          YARN TAG
        </div>
        <div style="width:34%; height:1.6px; background:#111; margin:7px auto 0;"></div>
      </div>
    
      <!-- QR (RIGHT) -->
        <div>
          <img
            src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=http://192.168.0.102:8003/app/daily-sample-request/{{ doc.name }}"
            style="width:55px; height:55px;"
            alt="QR">
        </div>
    
    
    </div>


    <!-- INFO -->
    <table style="width:100%; border-collapse:collapse;">
        
      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555; width:30%;">BUYER</td>
        <td style="font-size:21px; padding:10px 6px; font-weight:800;">
          {{ doc.buyer or "" }}
        </td>
      </tr>
      
      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555; width:30%;">STYLE</td>
        <td style="font-size:21px; padding:10px 6px; font-weight:800;">
          {{ doc.style or "" }}
        </td>
      </tr>

      <tr>
        <td style="font-size:13px; padding:10px 6px; font-weight:700; color:#555;">COMPOSITION</td>
        <td style="font-size:21px; padding:14px 6px;">
          {% for y in doc.raw_material %}
            {{ y.yarn_composition }} - {{ y.yarn_count }}{% if not loop.last %}, {% endif %}
          {% endfor %}
        </td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">SDO NO</td><br/>
        <td style="font-size:21px; padding:22px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">DYE HOUSE</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">LOT NO</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>
      <tr style="margin-top: 21px;">
        <td style="font-size:21px; padding:12px 6px; font-weight:700; color:#555;">QTY.</td>
        <td style="font-size:21px; padding:23px 6px; border-bottom:1px solid #bbb;"></td>
      </tr>

    </table>

    <!-- FOOT -->
    <div style="margin-top:18px; display:flex; justify-content:space-between; font-size:9px; color:#888;">
      <div>RECEIVE DATE: {{ frappe.utils.formatdate(frappe.utils.nowdate(),"dd MMM yyyy") }}</div>
      <div>{{ doc.name }}</div>
    </div>

  </div>

</div>
```
ASSET TRANSFER ACKNOWLEDGEMENT -
```
<style>
    @page {
        size: A4 portrait;
        margin: 20mm;
    }

    .letter-wrapper {
        font-family: 'Inter', 'Hind Siliguri', sans-serif;
        color: #000;
        line-height: 1.5;
    }

    /* RESTORED: Your Preferred Header Styles */
    .header-table {
        width: 100%;
        border-bottom: 2.5px solid #000;
        padding-bottom: 21px;
        margin-bottom: 34px;
        border-collapse: collapse;
    }

    .company-identity {
        display: flex;
        align-items: center;
        gap: 21px;
    }

    .logo-container {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 55px;
        width: auto;
    }

    .logo-img {
        max-height: 55px;
        width: auto;
        display: block;
    }

    .company-title {
        font-size: 21px;
        font-weight: 800;
        letter-spacing: -0.8px;
        text-transform: uppercase;
        line-height: 1.1;
    }

    .document-label {
        text-align: right;
        font-size: 11px;
        letter-spacing: 1.5px;
        font-weight: 700;
        text-transform: uppercase;
        line-height: 1.4;
    }

    /* Meta Information Bar */
    .meta-bar {
        display: flex;
        justify-content: space-between;
        margin-bottom: 40px;
        font-size: 12px;
    }

    .section-heading {
        font-size: 9px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 10px;
        display: block;
        border-bottom: 1.5px solid #000;
        padding-bottom: 3px;
    }

    .undertaking-text {
        font-size: 13px;
        text-align: justify;
        font-weight: 400;
        color: #000;
    }

    /* NEW: Billion-Dollar Professional Table Styling */
    .specs-table {
        width: 100%;
        border: 2px solid #555;
        border-collapse: collapse;
        margin-top: 21px;
        margin-bottom: 34px;
        table-layout: fixed;
    }

    .specs-table thead th {
        background: #888;
        color: #fff;
        text-align: center;
        padding: 13px;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 2px;
        border: 1px solid #000;
    }

    .specs-table td {
        padding: 13px 15px;
        border: 1px solid #000;
        font-size: 13px;
        vertical-align: middle;
    }

    .label-cell {
        font-weight: 800;
        width: 34%;
        background: #f2f2f2;
        text-transform: uppercase;
        font-size: 10px !important;
        color: #000;
        letter-spacing: 0.5px;
    }

    .value-cell {
        font-weight: 500;
        color: #000;
        background: #ffffff;
    }

    .desc-label {
        font-size: 8px;
        font-weight: 800;
        text-transform: uppercase;
        color: #000;
        margin-bottom: 8px;
        display: block;
        text-decoration: underline;
    }

    /* Signatures */
    .signature-row {
        margin-top: 110px;
        display: flex;
        justify-content: space-between;
    }

    .sig-block {
        width: 220px;
        text-align: center;
        border-top: 2px solid #000;
        padding-top: 10px;
        font-size: 11px;
        font-weight: 700;
    }

    .footer-reference {
        margin-top: 70px;
        font-size: 8px;
        color: #666;
        text-align: center;
        border-top: 1px solid #eee;
        padding-top: 10px;
    }
</style>

{# --- DYNAMIC LOGIC --- #}
{% set item_desc = frappe.db.get_value("Item", doc.item_code, "description") if doc.item_code else "" %}
{% set emp_name = frappe.db.get_value("Employee", doc.custodian, "employee_name") if doc.custodian else "N/A" %}
{% set emp_des = frappe.db.get_value("Employee", doc.custodian, "designation") if doc.custodian else "" %}
{% set base_url = frappe.utils.get_url() %}

<div class="letter-wrapper">
    <table class="header-table">
        <tr>
            <td style="width: 70%;">
                <div class="company-identity">
                    <div class="logo-container">
                        <img src="{{ base_url }}/files/logo_meek-01.png" class="logo-img" alt="LOGO">
                    </div>
                    <div>
                        <div class="company-title">{{ doc.company }}</div>
                        <div style="font-size: 10px; font-weight: 500; color: #444;">CORPORATE ASSET GOVERNANCE</div>
                    </div>
                </div>
            </td>
            <td class="document-label">
                সম্পদ গ্রহণ ও অঙ্গীকারনামা<br>
                <span style="font-size: 8px; font-weight: 400; color: #555;">ASSET HANDOVER & UNDERTAKING</span>
            </td>
        </tr>
    </table>

    <div class="meta-bar">
        <div style="width: 48%;">
            <span class="section-heading">কর্মচারীর তথ্য (Employee Details)</span>
            <div style="line-height: 1.8;">
                <b>নাম:</b> {{ emp_name }}<br>
                <b>আইডি:</b> {{ doc.custodian or "---" }}<br>
                <b>পদবী:</b> {{ emp_des }}<br>
                <b>বিভাগ:</b> {{ doc.department or "---" }}
            </div>
        </div>
        <div style="width: 48%; text-align: right;">
            <span class="section-heading">নথি তথ্য (Document Reference)</span>
            <div style="line-height: 1.8;">
                <b>রেফারেন্স নং:</b> {{ doc.name }}<br>
                <b>তারিখ:</b> {{ frappe.utils.formatdate(frappe.utils.nowdate(), "dd MMMM, yyyy") }}<br>
                <b>অবস্থান:</b> {{ doc.location or "Office Premises" }}
            </div>
        </div>
    </div>

    <div class="undertaking-section">
        <span class="section-heading">অঙ্গীকারনামা (Terms of Undertaking)</span>
        <div class="undertaking-text">
            আমি নিম্নস্বাক্ষরকারী অঙ্গীকার করছি যে, অদ্য তারিখে আমার দাপ্তরিক কাজের সুবিধার্থে <b>{{ doc.company }}</b> এর পক্ষ হতে বর্ণিত সম্পদটি বুঝে নিলাম। উক্ত সম্পদ ব্যবহারকালীন সময়ে আমি নিম্নোক্ত শর্তসমূহ মেনে চলতে বাধ্য থাকব:
            <ul style="margin-top: 12px; padding-left: 25px; line-height: 1.7;">
                <li>উক্ত সম্পদটি শুধুমাত্র কোম্পানির দাপ্তরিক কাজে ব্যবহারের জন্য অনুমোদিত।</li>
                <li>সম্পদটির পূর্ণ নিরাপত্তা এবং যথাযথ রক্ষণাবেক্ষণের দায়ভার ব্যক্তিগতভাবে আমার ওপর ন্যস্ত।</li>
                <li>কর্তৃপক্ষের পূর্বানুমতি ব্যতীত হার্ডওয়্যার পরিবর্তন বা কোনো সফটওয়্যার কনফিগারেশন পরিবর্তন করা নিষিদ্ধ।</li>
                <li>সম্পদটি হারিয়ে গেলে বা অবহেলাজনিত কারণে ক্ষতিগ্রস্ত হলে, আমি বর্তমান বাজার মূল্য অনুযায়ী পূর্ণ ক্ষতিপূরণ প্রদানে বাধ্য থাকব।</li>
                <li>চাকরি ইস্তফা বা বদলির ক্ষেত্রে, আমি দায়বদ্ধতার সাথে সম্পদটি অ্যাডমিন বিভাগকে বুঝিয়ে দিয়ে ছাড়পত্র গ্রহণ করব।</li>
            </ul>
        </div>
    </div>

    <table class="specs-table">
        <thead>
            <tr>
                <th colspan="2">সম্পদের কারিগরি বিবরণ (Technical Specifications)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td class="label-cell">অ্যাসেট নাম (Asset Name)</td>
                <td class="value-cell">{{ doc.asset_name | upper }}</td>
            </tr>
            <tr>
                <td class="label-cell">সিরিয়াল নম্বর (Serial No)</td>
                <td class="value-cell" style="font-family: monospace;">{{ doc.serial_number or "NOT REGISTERED" }}</td>
            </tr>
            <tr>
                <td class="label-cell">আইটেম কোড (Item Code)</td>
                <td class="value-cell">{{ doc.item_code }}</td>
            </tr>
            <tr>
                <td class="label-cell">ইনভয়েস তারিখ (Invoice Date)</td>
                <td class="value-cell">{{ frappe.utils.formatdate(doc.purchase_date, "dd MMM yyyy") if doc.purchase_date else "N/A" }}</td>
            </tr>
            <tr>
                <td class="label-cell">ব্যবহারযোগ্য তারিখ (Available Date)</td>
                <td class="value-cell">{{ frappe.utils.formatdate(doc.available_for_use_date, "dd MMM yyyy") if doc.available_for_use_date else "N/A" }}</td>
            </tr>
            <tr>
                <td class="label-cell">বর্তমান অবস্থা (Status)</td>
                <td class="value-cell" style="font-weight: 800;">{{ doc.status | upper }}</td>
            </tr>
            {% if item_desc %}
            <tr class="desc-row">
                <td colspan="2" style="border-top: 2px solid #000;">
                    <span class="desc-label">বিস্তারিত বিবরণ (Technical Description from Master)</span>
                    <div style="line-height: 1.7; font-size: 12.5px; color: #1a1a1a;">
                        {{ item_desc }}
                    </div>
                </td>
            </tr>
            {% endif %}
        </tbody>
    </table>

    <div class="signature-row">
        <div class="sig-block">
            প্রদানকারীর স্বাক্ষর ও সীল<br>
            <span style="font-size: 8px; font-weight: 400;">(Issuing Authority)</span>
        </div>
        <div class="sig-block">
            গ্রহীতার স্বাক্ষর ও তারিখ<br>
            <span style="font-size: 8px; font-weight: 400;">(Receiver's Acknowledgment)</span>
        </div>
    </div>

    <div class="footer-reference">
        This document serves as an official handover record. System generated on {{ frappe.utils.now_datetime().strftime('%d-%m-%Y %H:%M') }} by {{ frappe.session.user }}.
    </div>
    <div style="display:flex; justify-content:space-between; align-items:center;">
            <img class="qr" src="https://api.qrserver.com/v1/create-qr-code/?size=150x150&data={{ frappe.utils.get_url() }}/app/asset/{{ doc.name }}" style="width:55px; height:55px;">
    </div>
</div>
```
