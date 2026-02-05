# DAILY-TASKS
A REPO FOR COPYING HOME SERVER DATA TO OF SERVER
Scrpt for automatically fetch and arrange production report with MD Sir & CM Sir Format.
```
function main(workbook: ExcelScript.Workbook) {
    let sourceSheet = workbook.getWorksheet("Today,Total & Balance");
    let dailySummary = workbook.getWorksheet("Daily Summary");
    let mdfpSheet = workbook.getWorksheet("MDFP-WC6-2ND");

    if (!sourceSheet) return;

    if (!dailySummary) {
        dailySummary = workbook.addWorksheet("Daily Summary");
    } else {
        dailySummary.getUsedRange()?.clear();
    }

    // Clean Background for that Premium Look
    dailySummary.getRange("A1:Z2000").getFormat().getFill().setColor("FFFFFF");

    let sourceValues: (string | number | boolean)[][] = sourceSheet.getUsedRange().getValues();
    
    // Exact Header Sequence from your Image
    const headers: string[][] = [["SL No", "IMAGE", "FACTORY", "STYLE", "GG", "COMPOSITION", "ORDER QTY", "EX-FTY DATE", "PPS", "M/C", "KNIT", "LINK", "MEND", "WASH", "POLY", "BTB L/C", "DYEING ORDER", "REMARKS"]];

    let finalReportRows: (string | number | boolean)[][] = [];
    let buyers: string[] = ["PAAYRA", "PEPCO"];
    let slNo: number = 1;

    // 1. BUYER-DRIVEN DATA FETCHING
    buyers.forEach((buyer: string) => {
        // Collect ALL rows for this buyer
        let buyerData = sourceValues.filter(row => 
            String(row[0]).toUpperCase().indexOf(buyer.toUpperCase()) !== -1
        );

        if (buyerData.length > 0) {
            let bQty = 0, bKnit = 0, bLink = 0, bMend = 0, bWash = 0, bPoly = 0;
            let startedRows: (string | number | boolean)[][] = [];
            let pendingRows: (string | number | boolean)[][] = [];

            // Sort by status (Started vs Not Started)
            buyerData.forEach((row) => {
                let knit = Number(row[7]) || 0;
                let poly = Number(row[27]) || 0;
                if (knit === 0 && poly === 0) {
                    pendingRows.push(row);
                } else {
                    startedRows.push(row);
                }
            });

            // Add Styles with Production Activity
            startedRows.forEach((row) => {
                let qty = Number(row[3]) || 0;
                let knit = Number(row[7]) || 0;
                let poly = Number(row[27]) || 0;

                finalReportRows.push([
                    slNo++, "", String(row[2]), String(row[1]), String(row[4]), 
                    "50% COTTON 50% ACRYLIC", qty, "29.03.26", "APPROVED", 
                    String(row[5]), knit, Number(row[11]) || 0, Number(row[19]) || 0, 
                    Number(row[23]) || 0, poly, "YARN IN HOUSE", "FRI-22.12.25", "IN PRODUCTION"
                ]);
                bQty += qty; bKnit += knit; bPoly += poly;
                // Add Link/Mend/Wash to totals if needed
                bLink += (Number(row[11]) || 0);
                bMend += (Number(row[19]) || 0);
                bWash += (Number(row[23]) || 0);
            });

            // Add "Pending" Separator and Rows
            if (pendingRows.length > 0) {
                finalReportRows.push(["--- " + buyer + " PENDING ORDERS ---", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
                pendingRows.forEach((row) => {
                    let qty = Number(row[3]) || 0;
                    finalReportRows.push([slNo++, "", String(row[2]), String(row[1]), String(row[4]), "50% COTTON 50% ACRYLIC", qty, "29.03.26", "APPROVED", String(row[5]), 0, 0, 0, 0, 0, "YARN IN HOUSE", "FRI-22.12.25", "WAITING"]);
                    bQty += qty;
                });
            }

            // Buyer Sub-Total Row
            finalReportRows.push([buyer + " TOTAL", "", "", bQty, "", "", bKnit, bLink, bMend, bWash, bPoly, (bQty - bPoly), "", "", "", "", "", ""]);
        }
    });

    // 2. OUTPUT & BILLION DOLLAR STYLING
    dailySummary.getRange("A4:R4").setValues(headers);
    let headerRange = dailySummary.getRange("A4:R4");
    headerRange.getFormat().getFill().setColor("#2D2D2D"); // Slate Gray
    headerRange.getFormat().getFont().setColor("#FFFFFF");
    headerRange.getFormat().getFont().setBold(true);

    if (finalReportRows.length > 0) {
        let dataRange = dailySummary.getRange("A5").getResizedRange(finalReportRows.length - 1, 17);
        dataRange.setValues(finalReportRows);
        dataRange.getFormat().getFont().setName("Segoe UI");

        finalReportRows.forEach((row, idx) => {
            let rIdx = idx + 5;
            let rowRange = dailySummary.getRange(`A${rIdx}:R${rIdx}`);
            
            // Zebra Striping
            if (idx % 2 === 0) rowRange.getFormat().getFill().setColor("#F9F9F9");

            // Total Row Style
            if (String(row[0]).indexOf("TOTAL") !== -1) {
                rowRange.getFormat().getFill().setColor("#E8E8E8");
                rowRange.getFormat().getFont().setBold(true);
                rowRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
            }

            // Pending Separator Style
            if (String(row[0]).indexOf("PENDING") !== -1) {
                rowRange.getFormat().getFill().setColor("#FFF9E6");
                rowRange.getFormat().getFont().setItalic(true);
                rowRange.getFormat().getHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
            }
        });
    }

    // 3. AUTO-SYNC MDFP SHEET (Design Safe)
    if (mdfpSheet) {
        let mdfpRange = mdfpSheet.getUsedRange();
        let mdfpValues = mdfpRange.getValues();
        for (let i = 0; i < mdfpValues.length; i++) {
            let mStyle = String(mdfpValues[i][3]).replace(/\s/g, "").toUpperCase();
            if (mStyle.length > 4) {
                let match = sourceValues.find(row => String(row[1]).replace(/\s/g, "").toUpperCase().indexOf(mStyle) !== -1);
                if (match) {
                    mdfpSheet.getRangeByIndexes(i, 9, 1, 6).setValues([[match[5], match[7], match[11], match[19], match[23], match[27]]]);
                }
            }
        }
    }

    // Finishing Touches
    let tableRange = dailySummary.getRange("A4").getSurroundingRegion();
    tableRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.continuous);
    tableRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setColor("#E0E0E0");
    
    dailySummary.getUsedRange().getFormat().autofitColumns();
    dailySummary.getRange("B:B").getFormat().setColumnWidth(75); // For Images
    
    console.log("Complete Buyer-Driven Report Generated. No Styles Missed.");
}
```

# Sample Yarn Tag
<code>
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
</code>
