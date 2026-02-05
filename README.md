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
