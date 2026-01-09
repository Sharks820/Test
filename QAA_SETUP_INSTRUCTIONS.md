# QAA Quote System v12.0 - Setup Instructions

## Quality Aircraft Accessories - A Hartzell Engine Technologies Company

---

## DELIVERABLES INCLUDED

| File | Description |
|------|-------------|
| `QAA_Quote_System.xlsx` | Complete Excel workbook with all sheets, formulas, and embedded button images |
| `QAA_VBA_MODULE.bas` | VBA code module to install for macros |
| `btn_*.png` | Professional button icons (already embedded in workbook) |
| `qaa_logo_header.png` | Logo for PDF headers |

---

## CRITICAL FIXES IN THIS VERSION

### 1. 500HR Pricing - FIXED
- **Now uses `5H` suffix (NOT `-5H`)**
- Example: `10-382320` + `5H` = `10-3823205H`
- Each magneto type has its own 500HR price
- Formulas: `=VLOOKUP(B10&"5H",'Parts Master'!$A:$H,4,FALSE)`

### 2. Button Icons - INCLUDED
- Real PNG button images embedded in all sheets
- Professional navy blue with white text
- Rounded corners, consistent styling

### 3. Email Logging - COMPLETE
Logs to **BOTH** Quote Log and Email Log:
- Quote Number
- Date Sent
- Time Sent
- Customer Email Address
- Email Subject
- **FULL Email Body Text**
- Quote Type
- Units/Items Count
- Total Amount
- Status

---

## SETUP STEPS

### Step 1: Open and Save as Macro-Enabled
1. Open `QAA_Quote_System.xlsx` in Microsoft Excel
2. Click **File > Save As**
3. Change type to **Excel Macro-Enabled Workbook (.xlsm)**
4. Save as `QAA_Quote_System.xlsm`

### Step 2: Install VBA Code
1. Press **Alt+F11** to open VBA Editor
2. In the Project Explorer, right-click on your workbook
3. Select **Insert > Module**
4. Open `QAA_VBA_MODULE.bas` in a text editor
5. Copy ALL the code
6. Paste into the new module
7. Press **Ctrl+S** to save
8. Close VBA Editor (Alt+Q)

### Step 3: Assign Macros to Buttons
For each button image in the workbook:
1. Right-click the button image
2. Select **Assign Macro...**
3. Select the appropriate macro (see table below)
4. Click OK

### Macro Assignments

| Button | Macro Name | Sheet |
|--------|------------|-------|
| 500HR SERVICE | `Go500HR` | Menu |
| CUSTOMER PROPERTY | `GoCustProp` | Menu |
| EXCHANGE / SALE | `GoExchange` | Menu |
| GENERATE PDF | `GenPDF500HR` | 500HR |
| EMAIL QUOTE | `Email500HR` | 500HR |
| CLEAR FORM | `Clear500HR` | 500HR |
| MENU | `GoMenu` | 500HR |
| GENERATE PDF | `GenPDFCustProp` | Cust Property |
| EMAIL QUOTE | `EmailCustProp` | Cust Property |
| CLEAR FORM | `ClearCustProp` | Cust Property |
| MENU | `GoMenu` | Cust Property |
| UPDATE DATE | `UpdateCustPropDate` | Cust Property |
| GENERATE PDF | `GenPDFExchange` | Exchange |
| EMAIL QUOTE | `EmailExchange` | Exchange |
| CLEAR FORM | `ClearExchange` | Exchange |
| MENU | `GoMenu` | Exchange |

---

## SHEET STRUCTURE

| Sheet | Purpose |
|-------|---------|
| Menu | Navigation hub with buttons to each quote type |
| 500HR | 500HR Magneto Inspection Service quotes |
| Cust Property | Customer Property Overhaul quotes |
| Exchange | Exchange/Inventory Sale quotes |
| Parts Master | Product database with pricing |
| BOM | Bill of Materials (determines In-House vs Vendor) |
| Config | System settings (quote prefix, fees, tier discounts) |
| Quote Log | Summary of all sent quotes |
| Email Log | Complete email body records |
| Inventory | Stock data placeholder |

---

## PARTS MASTER - 5H SUFFIX EXAMPLES

| Part Number | Description | Price |
|-------------|-------------|-------|
| 10-3823205H | S-20/200 Magneto - 500HR Inspection | $575 |
| 10-3206805H | S-1200 Magneto - 500HR Inspection | $495 |
| 10-3910885H | D-2000/D-3000 - 500HR Inspection | $695 |
| 10-1630205H | Bendix S-20 - 500HR Inspection | $625 |

**Note:** The suffix is `5H` (no hyphen before it), NOT `-5H`

---

## CONFIG SETTINGS

| Setting | Value | Notes |
|---------|-------|-------|
| Quote Prefix | QAA | Quote numbers: QAA-00001 |
| Next Quote # | 1 | Auto-increments after email |
| Quote Valid Days | 30 | Shown in emails |
| 500HR Default Fee | $550 | Fallback if 5H part not found |
| 500HR Findings Add-On | $110 | Per unit with ANY findings |

### Tier Discounts
| Tier | Discount |
|------|----------|
| Standard | 0% |
| Tier 1 | 7% |
| Tier 2 | 10% |
| Tier 3 | 13% |

---

## TESTING CHECKLIST

- [ ] Enter magneto P/N (e.g., 10-382320) → 5H price ($575) appears
- [ ] Click checkbox in findings grid → changes ☐ to ☑
- [ ] Mark any finding → $110 add-on calculated
- [ ] Mark Ovhl Req = Yes → CO part# and price appear
- [ ] Click GENERATE PDF → PDF created with logo
- [ ] Click EMAIL QUOTE → Outlook opens with PDF attached
- [ ] After email sent → logged to Quote Log AND Email Log
- [ ] Quote number increments after email
- [ ] Tier discount applies correctly to Exchange prices
- [ ] Core In Hand = Yes → Core charges show $0

---

## TROUBLESHOOTING

### "Macros disabled" warning
- Click **Enable Content** when prompted
- Or: File > Options > Trust Center > Trust Center Settings > Enable all macros

### Buttons don't work
- Make sure you assigned macros to each button (Step 3)
- Check that VBA module is installed correctly

### Email error
- Ensure Microsoft Outlook is installed and configured
- Check that you entered a valid email address

### PDF not creating
- Make sure print area is set correctly
- Check that the folder path is valid

---

## CONTACT

Quality Aircraft Accessories
A Hartzell Engine Technologies Company
5746 E Apache St, Tulsa, OK 74115
Phone: 918-835-6948
Email: sales@qaainc.com
Web: qaa.com

---

*Generated by QAA Quote System Builder v12.0*
