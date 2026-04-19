PROMPT — A&G Physiotherapy Quarterly Report Generator
You are a senior healthcare data analyst and report designer for A & G Physiotherapy Inc.
I am attaching an Excel file with physiotherapy data for a long-term care facility. Your job is to generate a professional Word document (.docx) quarterly report by reading every number directly from the Excel — do not fabricate, estimate, or assume any values.

EXCEL FILE STRUCTURE
The file has these sheets:

Resident Flow — row 0 is a sub-header, row 1 is the data. Columns: Quarter, Start Residents, Admissions, Discharged(Total) [= Deaths/Palliative], Unnamed: 4 [= Moved Out], non compliant, goal achieved, final(Current residents)
Therapy Minutes — 1:1 Minutes, Evaluation Minutes, Group Sessions (count/notes)
PT Programs — Ambulation, Weight Bearing/Pre-Gait, AROM/PROM, Strengthening + ROM
Referals — Total Referrals, Jan, Feb, Mar (monthly breakdown)
Assesments — Total Assessments, Jan, Feb, Mar (monthly breakdown)
Staffing — PTA Hours (Total), PTA 1:1 Hours, PTA Group Hours, PT Hours
Summary Metrics — % Residents on 1:1 PT, Total Residents
total beds — rows: January, February, March with resident counts in column 2. Total beds in facility = 128


REQUIRED LIBRARIES
pip install pandas openpyxl python-docx matplotlib pillow

REPORT DESIGN SPECIFICATIONS
Page Setup: US Letter, margins 0.45" top/bottom, 0.55" left/right. Font: Arial throughout.
Colour Palette:

Navy: #1B3A6B — primary headers, main chart bars
Teal: #1A7A8A — accents, dividers, section borders
White: #FFFFFF
Dark: #2C3E50 — body text
Mid: #7F8C8D — captions, notes
Light Blue tile: #EBF5FB
Light Green tile: #EAFAF1
Light Yellow tile: #FEFAED
Light Red tile: #FDEDEC
Amber badge: #B7770D
Green badge: #1E8449
Red badge: #922B21
Chart BG: #F4F8FC


REPORT SECTIONS — IN ORDER
1. HEADER BANNER (top of page)
Two-column table spanning full content width (9360 DXA):

Left cell (3000 DXA): White background. Embed logo_clean.png at 1.85" × 0.44"
Right cell (6360 DXA): Navy background. Text: "QUARTERLY PHYSIOTHERAPY REPORT" bold 13pt white. Second line: "Burton Manor, Brampton | {quarter} | Confidential" in 8.5pt light blue

2. KEY PERFORMANCE INDICATORS
Section header: navy filled bar, white uppercase text 10pt.
Two rows of KPI tiles (4 tiles in row 1, 5 tiles in row 2). Each tile: coloured background, large bold value (20pt navy), bold label (8pt dark), italic note (7.5pt grey), optional bold coloured badge below note.
Row 1 (4 tiles):
ValueLabelNoteTile colourBadgeflow["end"]Current Census"59 of 128 beds occupied"Light Blue"↓ X from opening" grey{pct_1on1}%On Active 1:1 Rx"Of {total_res} current residents"Light Green"Strong coverage" greenflow["deceased"]Deaths / Palliative"{mortality_pct}% of opening census"Light Red"High-acuity quarter" redreferrals["total"]Total Referrals"Q{quarter}"Light Yellow"Completed this quarter" amber
Row 2 (5 tiles):
ValueLabelNoteTile colourBadgeassessments["total"]Total Assessments"Q{quarter}"Light Green"Completed this quarter" green{one2one:,}1:1 Therapy Min."Direct therapy delivered"Light Bluenone{evaluation:,}Evaluation Min."Assessment time"Light Bluenone{pta_total}hPTA Hours / Week"2 PTAs (1 FT + 1 PT)"Light Bluenone{pt_hours}hPT Hours / Week"Mon, Tue, Thu"Light Yellow"3 days / week" amber
3. SERVICE HIGHLIGHTS BANNER
Immediately after KPIs. Two-column table, light blue fill, teal border on all sides.

Left cell: ● "A & G Physiotherapy Inc. provides ADP (Assistive Devices Program) services to eligible residents."
Right cell: ● "PT assesses all new residents on admission and completes Quarterly Assessments. The team participates in Daily Huddles, Falls Meetings, MDS, PT + NRCC Meetings and Care Conferences, and follows up on all referrals as required."

4. RESIDENT FLOW — {Quarter}
Section header navy bar.
Side-by-side layout:

Left (5400 DXA): Waterfall bar chart titled "Quarterly Resident Flow". 5 bars: Opening Census (navy), Admissions (green, stacked on top), Deaths/Palliative (red, descending), Moved Out (amber, descending), Closing Census (teal, from zero). Each bar labelled with exact number.
Right (3960 DXA): Bar chart titled "Monthly Census". 3 bars (Jan/Feb/Mar) in navy, each labelled with exact count from the total beds sheet.

Caption below: plain text 8.5pt grey describing the quarter movement using bold for key numbers.
5. PT REFERRALS, ASSESSMENTS AND FALLS
Section header navy bar.
Side-by-side layout:

Left (4680 DXA): Grouped bar chart titled "Monthly Referrals and Assessments". Two bars per month (Jan/Feb/Mar): navy = referrals, teal = assessments. Each bar labelled. No "gap" language. Legend top-left.
Right (4680 DXA): Nested breakdown table with columns: [blank], Jan, Feb, Mar, Total. Header row navy. Total column header teal. Two data rows: "PT Referrals" (light blue row, navy text) and "Assessments" (light green row, green text). Total column bold.

6. PROGRAM MIX, THERAPY DELIVERY AND WORKFORCE
Section header navy bar.
Two-column layout (4500 DXA left, 4860 DXA right):

Left: Donut chart of PT Programs. Segments: Ambulation (navy), Wt. Bearing & Pre-Gait (teal), Strengthening + ROM (blue #2980B9), AAROM/PROM (light blue #AED6F1). Percentages in white bold inside wedges. Legend below in 2 columns.
Right: Two charts stacked vertically with a gap between:

Top: Horizontal bar chart "Therapy Minutes" — 1:1 Therapy (navy) and Evaluations (teal), each labelled with exact minutes
Bottom: Stacked bar chart "Workforce Hours / Week" — PTA bar (teal = group hours stacked on navy = 1:1 hours) and PT bar (blue = direct hours). Each bar labelled with exact total h/wk. Legend shows exact hours per category.



Three-column caption row below all charts (teal left-border dividers on columns 2 and 3), 8.5pt grey text with bold highlights for key numbers.
7. FOOTER
Centred italic 8pt grey text: "End of Report | {Quarter} | Burton Manor Brampton | A & G Physiotherapy Inc."

CHART STYLING (all charts)

Background: #F4F8FC
No top or right spine
Left and bottom spines: #CCCCCC
Y-axis gridlines: #DDDDDD
All labels: DejaVu Sans, values bold
dpi=160, bbox_inches="tight"
Return as io.BytesIO buffer


WORD DOC TABLE RULES

All table widths set explicitly in DXA using w:tblW
All cell widths set with w:tcW
All cell shading uses w:shd val="clear"
Borders set via w:tcBorders with w:val="nil" for no-border cells
Cell margins always set explicitly
No \n inside runs — use separate paragraphs
Paragraph spacing always set explicitly with w:spacing before/after
Never use WidthType.PERCENTAGE


OUTPUT
Save as: PT_{Quarter}_Report.docx in the same folder as the script.
Print: "Generating charts..." then "Building document..." then "Saved: {path}"
IMPORTANT RULES

Every single number in the report must come directly from the Excel file
Do not fabricate, estimate, or calculate any number not in the Excel
The only derived values allowed are: mortality_pct = deceased/start*100 (rounded 1dp), hrs_1on1 = one2one/60 (rounded 1dp), hrs_eval = evaluation/60 (rounded 1dp), active_res = pct_1on1/100 * total_res (rounded to integer), prog_total = sum of all 4 program values
Do not show any "gap" between referrals and assessments — they are independent metrics
Do not mention PT coverage gaps or anything with negative framing
Do not include a Risks section or Recommendations section


LOGO
The file logo_clean.png must be in the same folder as the script. It is the A&G Physiotherapy Inc. logo on a white background. It is used in the header banner on a white cell background.

Save this prompt. Next quarter, paste it into Claude or ChatGPT, attach your updated Excel file alongside it, and you will get the exact same report generated fresh with the new data.Sonnet 4.6Adaptive