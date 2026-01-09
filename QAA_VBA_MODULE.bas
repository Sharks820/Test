Attribute VB_Name = "QAAQuoteSystem"
' ╔══════════════════════════════════════════════════════════════════════════════╗
' ║     QAA QUOTE SYSTEM v12.0 - VBA MODULE                                      ║
' ║     Quality Aircraft Accessories - A Hartzell Engine Technologies Company    ║
' ║     CORRECTED: Uses 5H suffix (NOT -5H) for 500HR pricing                    ║
' ║     Full Email Logging: Customer email, subject, FULL body, date/time        ║
' ╚══════════════════════════════════════════════════════════════════════════════╝
'
' SETUP INSTRUCTIONS:
' 1. Save workbook as .xlsm (macro-enabled)
' 2. Press Alt+F11 to open VBA Editor
' 3. Insert > Module
' 4. Paste this entire code
' 5. Close VBA Editor
' 6. Right-click each button image > Assign Macro > Select the appropriate macro

Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
' NAVIGATION MACROS
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub GoMenu()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Menu").Activate
    Application.ScreenUpdating = True
End Sub

Public Sub Go500HR()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("500HR").Activate
    ThisWorkbook.Sheets("500HR").Range("B6").Select
    Application.ScreenUpdating = True
End Sub

Public Sub GoCustProp()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Cust Property").Activate
    ThisWorkbook.Sheets("Cust Property").Range("C6").Select
    Application.ScreenUpdating = True
End Sub

Public Sub GoExchange()
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Exchange").Activate
    ThisWorkbook.Sheets("Exchange").Range("B6").Select
    Application.ScreenUpdating = True
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PDF GENERATION
' ═══════════════════════════════════════════════════════════════════════════════

Private Function GetPDFPath(quoteNum As String) As String
    Dim folder As String
    folder = ThisWorkbook.Path
    If folder = "" Then folder = Environ("USERPROFILE") & "\Documents"
    GetPDFPath = folder & "\Quote_" & Replace(Replace(quoteNum, "-", "_"), "/", "_") & "_" & Format(Date, "YYYYMMDD") & ".pdf"
End Function

Public Sub GenPDF500HR()
    Dim ws As Worksheet
    Dim pdfPath As String

    Set ws = ThisWorkbook.Sheets("500HR")

    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation, "QAA Quote System"
        ws.Range("C6").Select
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("C5").Value)

    Application.ScreenUpdating = False

    With ws.PageSetup
        .PrintArea = "$A$1:$J$37"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = True
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard, False, False, , , False

    Application.ScreenUpdating = True
    MsgBox "PDF Created Successfully!" & vbCrLf & vbCrLf & pdfPath, vbInformation, "QAA Quote System"

    If MsgBox("Create email with PDF attached?", vbYesNo + vbQuestion, "QAA Quote System") = vbYes Then
        Email500HR pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical, "QAA Quote System"
End Sub

Public Sub GenPDFCustProp()
    Dim ws As Worksheet
    Dim pdfPath As String

    Set ws = ThisWorkbook.Sheets("Cust Property")

    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation, "QAA Quote System"
        ws.Range("C6").Select
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("C5").Value)

    Application.ScreenUpdating = False

    With ws.PageSetup
        .PrintArea = "$A$1:$I$36"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = True
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard, False, False, , , False

    Application.ScreenUpdating = True
    MsgBox "PDF Created Successfully!" & vbCrLf & vbCrLf & pdfPath, vbInformation, "QAA Quote System"

    If MsgBox("Create email with PDF attached?", vbYesNo + vbQuestion, "QAA Quote System") = vbYes Then
        EmailCustProp pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical, "QAA Quote System"
End Sub

Public Sub GenPDFExchange()
    Dim ws As Worksheet
    Dim pdfPath As String

    Set ws = ThisWorkbook.Sheets("Exchange")

    If Trim(ws.Range("B6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation, "QAA Quote System"
        ws.Range("B6").Select
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("B5").Value)

    Application.ScreenUpdating = False

    With ws.PageSetup
        .PrintArea = "$A$1:$J$30"
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = True
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard, False, False, , , False

    Application.ScreenUpdating = True
    MsgBox "PDF Created Successfully!" & vbCrLf & vbCrLf & pdfPath, vbInformation, "QAA Quote System"

    If MsgBox("Create email with PDF attached?", vbYesNo + vbQuestion, "QAA Quote System") = vbYes Then
        EmailExchange pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical, "QAA Quote System"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' EMAIL GENERATION WITH COMPLETE LOGGING
' Logs: Customer email, Subject, FULL body text, Date, Time
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub Email500HR(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double

    Set ws = ThisWorkbook.Sheets("500HR")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C5").Value
    totalAmt = ws.Range("H30").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email address.", vbExclamation, "QAA Quote System"
        ws.Range("C6").Select
        Exit Sub
    End If

    ' Build professional subject line
    subj = "Quote " & quoteNum & " - 500HR Magneto Inspection Service - Quality Aircraft Accessories"

    ' Build professional email body
    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for choosing Quality Aircraft Accessories for your magneto inspection needs." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for 500HR Magneto Inspection Service." & vbCrLf & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "                    QUOTE SUMMARY" & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "Quote #:        " & quoteNum & vbCrLf
    body = body & "Date:           " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:          30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:          " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(54, "=") & vbCrLf & vbCrLf

    ' Check for units requiring overhaul
    If Application.WorksheetFunction.CountIf(ws.Range("F10:F15"), "Yes") > 0 Then
        body = body & "NOTE: Some units require overhaul based on inspection findings." & vbCrLf
        body = body & "Please see attached quote for complete details." & vbCrLf & vbCrLf
    End If

    ' Check approval requirement
    If ws.Range("H6").Value = "Yes" Then
        body = body & "** APPROVAL REQUIRED **" & vbCrLf
        body = body & "Please reply with your approval to proceed with service." & vbCrLf & vbCrLf
    End If

    body = body & "If you have any questions, please don't hesitate to contact us." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
    body = body & "A Hartzell Engine Technologies Company" & vbCrLf
    body = body & "5746 E Apache St, Tulsa, OK 74115" & vbCrLf
    body = body & "918-835-6948 | sales@qaainc.com | qaa.com"

    On Error GoTo EmailError
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = emailTo
        .Subject = subj
        .body = body
        If pdfPath <> "" Then
            If Dir(pdfPath) <> "" Then .Attachments.Add pdfPath
        End If
        .Display
    End With

    ' LOG THE EMAIL WITH FULL DETAILS
    LogEmail quoteNum, emailTo, subj, body, "500HR Service", Count500HRUnits(), totalAmt, 0

    MsgBox "Email created and logged successfully!", vbInformation, "QAA Quote System"
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure Microsoft Outlook is installed.", vbCritical, "QAA Quote System"
End Sub

Public Sub EmailCustProp(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double

    Set ws = ThisWorkbook.Sheets("Cust Property")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C5").Value
    totalAmt = ws.Range("G31").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email address.", vbExclamation, "QAA Quote System"
        ws.Range("C6").Select
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Customer Property Overhaul - Quality Aircraft Accessories"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for choosing Quality Aircraft Accessories for your overhaul needs." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for Customer Property Overhaul services." & vbCrLf & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "                    QUOTE SUMMARY" & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "Quote #:        " & quoteNum & vbCrLf
    body = body & "Date:           " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:          30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:          " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(54, "=") & vbCrLf & vbCrLf

    ' Requirements section
    Dim req As String: req = ""
    If ws.Range("C7").Value = "Yes" Then req = req & "  - Approval to proceed" & vbCrLf
    If ws.Range("E7").Value = "Yes" Then req = req & "  - Shipping information" & vbCrLf
    If ws.Range("G7").Value = "Yes" Then req = req & "  - Payment method/terms" & vbCrLf

    If req <> "" Then
        body = body & "TO PROCEED, WE NEED:" & vbCrLf & req & vbCrLf
    End If

    body = body & "If you have any questions, please don't hesitate to contact us." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
    body = body & "A Hartzell Engine Technologies Company" & vbCrLf
    body = body & "5746 E Apache St, Tulsa, OK 74115" & vbCrLf
    body = body & "918-835-6948 | sales@qaainc.com | qaa.com"

    On Error GoTo EmailError
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = emailTo
        .Subject = subj
        .body = body
        If pdfPath <> "" Then
            If Dir(pdfPath) <> "" Then .Attachments.Add pdfPath
        End If
        .Display
    End With

    LogEmail quoteNum, emailTo, subj, body, "Customer Property", CountCPUnits(), totalAmt, 0

    MsgBox "Email created and logged successfully!", vbInformation, "QAA Quote System"
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical, "QAA Quote System"
End Sub

Public Sub EmailExchange(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double, coreAmt As Double, savings As Double

    Set ws = ThisWorkbook.Sheets("Exchange")
    emailTo = Trim(ws.Range("B6").Value)
    quoteNum = ws.Range("B5").Value
    totalAmt = ws.Range("E25").Value
    coreAmt = ws.Range("D24").Value
    savings = ws.Range("G24").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email address.", vbExclamation, "QAA Quote System"
        ws.Range("B6").Select
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Exchange/Sale Quote - Quality Aircraft Accessories"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for your interest in Quality Aircraft Accessories." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote." & vbCrLf & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "                    QUOTE SUMMARY" & vbCrLf
    body = body & String(54, "=") & vbCrLf
    body = body & "Quote #:        " & quoteNum & vbCrLf
    body = body & "Date:           " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:          30 Days" & vbCrLf & vbCrLf
    body = body & "Subtotal:       " & Format(ws.Range("B24").Value, "$#,##0.00") & vbCrLf
    body = body & "Core Charges:   " & Format(coreAmt, "$#,##0.00")
    If ws.Range("H7").Value = "Yes" Then body = body & " (WAIVED - Core In Hand)"
    body = body & vbCrLf
    body = body & String(54, "-") & vbCrLf
    body = body & "TOTAL:          " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(54, "=") & vbCrLf & vbCrLf

    If savings > 0 Then
        body = body & "YOUR SAVINGS:   " & Format(savings, "$#,##0.00") & vbCrLf & vbCrLf
    End If

    ' Requirements
    Dim req As String: req = ""
    If ws.Range("B7").Value = "Yes" Then req = req & "  - Approval to proceed" & vbCrLf
    If ws.Range("D7").Value = "Yes" Then req = req & "  - Shipping information" & vbCrLf
    If ws.Range("F7").Value = "Yes" Then req = req & "  - Payment method/terms" & vbCrLf

    If req <> "" Then
        body = body & "TO PROCEED, WE NEED:" & vbCrLf & req & vbCrLf
    End If

    If ws.Range("H7").Value <> "Yes" And coreAmt > 0 Then
        body = body & "CORE RETURN POLICY:" & vbCrLf
        body = body & "Core charges are refunded upon receipt of serviceable core" & vbCrLf
        body = body & "within 30 days of shipment." & vbCrLf & vbCrLf
    End If

    body = body & "If you have any questions, please don't hesitate to contact us." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
    body = body & "A Hartzell Engine Technologies Company" & vbCrLf
    body = body & "5746 E Apache St, Tulsa, OK 74115" & vbCrLf
    body = body & "918-835-6948 | sales@qaainc.com | qaa.com"

    On Error GoTo EmailError
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = emailTo
        .Subject = subj
        .body = body
        If pdfPath <> "" Then
            If Dir(pdfPath) <> "" Then .Attachments.Add pdfPath
        End If
        .Display
    End With

    LogEmail quoteNum, emailTo, subj, body, "Exchange/Sale", CountExItems(), totalAmt, coreAmt

    MsgBox "Email created and logged successfully!", vbInformation, "QAA Quote System"
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical, "QAA Quote System"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' COMPLETE EMAIL LOGGING
' Logs to BOTH Quote Log AND Email Log with FULL email body
' ═══════════════════════════════════════════════════════════════════════════════

Private Sub LogEmail(quoteNum As String, emailTo As String, subj As String, body As String, _
                     qType As String, units As Integer, total As Double, core As Double)

    Dim wsQuoteLog As Worksheet, wsEmailLog As Worksheet, wsConfig As Worksheet
    Dim nr As Long
    Dim dateSent As Date, timeSent As String

    On Error GoTo LogError

    Set wsQuoteLog = ThisWorkbook.Sheets("Quote Log")
    Set wsEmailLog = ThisWorkbook.Sheets("Email Log")
    Set wsConfig = ThisWorkbook.Sheets("Config")

    dateSent = Date
    timeSent = Format(Now, "HH:MM:SS AM/PM")

    ' ═══════════════════════════════════════════════════════════════════════════
    ' LOG TO QUOTE LOG (summary with truncated body)
    ' ═══════════════════════════════════════════════════════════════════════════
    nr = wsQuoteLog.Cells(wsQuoteLog.Rows.Count, 1).End(xlUp).Row + 1

    wsQuoteLog.Cells(nr, 1).Value = quoteNum           ' A: Quote #
    wsQuoteLog.Cells(nr, 2).Value = dateSent           ' B: Date Sent
    wsQuoteLog.Cells(nr, 3).Value = timeSent           ' C: Time Sent
    wsQuoteLog.Cells(nr, 4).Value = emailTo            ' D: Customer Email
    wsQuoteLog.Cells(nr, 5).Value = qType              ' E: Quote Type
    wsQuoteLog.Cells(nr, 6).Value = units              ' F: Units/Items
    wsQuoteLog.Cells(nr, 7).Value = total              ' G: Total Amount
    wsQuoteLog.Cells(nr, 7).NumberFormat = "$#,##0.00"
    wsQuoteLog.Cells(nr, 8).Value = core               ' H: Core Charges
    wsQuoteLog.Cells(nr, 8).NumberFormat = "$#,##0.00"
    wsQuoteLog.Cells(nr, 9).Value = "Sent"             ' I: Status
    wsQuoteLog.Cells(nr, 10).Value = subj              ' J: Email Subject
    wsQuoteLog.Cells(nr, 11).Value = Left(body, 500)   ' K: Body Preview (truncated)

    ' ═══════════════════════════════════════════════════════════════════════════
    ' LOG TO EMAIL LOG (FULL email body - NO truncation)
    ' ═══════════════════════════════════════════════════════════════════════════
    nr = wsEmailLog.Cells(wsEmailLog.Rows.Count, 1).End(xlUp).Row + 1

    wsEmailLog.Cells(nr, 1).Value = quoteNum           ' A: Quote #
    wsEmailLog.Cells(nr, 2).Value = dateSent           ' B: Date Sent
    wsEmailLog.Cells(nr, 3).Value = timeSent           ' C: Time Sent
    wsEmailLog.Cells(nr, 4).Value = emailTo            ' D: Customer Email
    wsEmailLog.Cells(nr, 5).Value = subj               ' E: Subject
    wsEmailLog.Cells(nr, 6).Value = body               ' F: FULL Email Body (complete text)

    ' Auto-wrap the body cell for readability
    wsEmailLog.Cells(nr, 6).WrapText = True
    wsEmailLog.Rows(nr).RowHeight = 100

    ' ═══════════════════════════════════════════════════════════════════════════
    ' INCREMENT QUOTE NUMBER
    ' ═══════════════════════════════════════════════════════════════════════════
    wsConfig.Range("B3").Value = wsConfig.Range("B3").Value + 1

    Exit Sub

LogError:
    MsgBox "Logging Error: " & Err.Description, vbExclamation, "QAA Quote System"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' CLEAR FORM MACROS
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub Clear500HR()
    Dim ws As Worksheet
    Dim r As Integer, c As Integer

    Set ws = ThisWorkbook.Sheets("500HR")

    If MsgBox("Clear the 500HR quote form?" & vbCrLf & vbCrLf & _
              "This will reset all entries.", vbYesNo + vbQuestion, "QAA Quote System") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Clear customer email
    ws.Range("C6").ClearContents

    ' Reset approval
    ws.Range("H6").Value = "No"

    ' Clear magneto rows
    For r = 10 To 15
        ws.Cells(r, 2).ClearContents  ' P/N
        ws.Cells(r, 3).ClearContents  ' S/N
        ws.Cells(r, 6).ClearContents  ' Overhaul Required
    Next r

    ' Reset findings checkboxes to unchecked
    For r = 19 To 24
        For c = 2 To 9
            ws.Cells(r, c).Value = ChrW(9744)  ' ☐
        Next c
    Next r

    ' Clear notes
    ws.Range("B33").ClearContents

    Application.ScreenUpdating = True
    MsgBox "Form cleared!", vbInformation, "QAA Quote System"
End Sub

Public Sub ClearCustProp()
    Dim ws As Worksheet
    Dim r As Integer

    Set ws = ThisWorkbook.Sheets("Cust Property")

    If MsgBox("Clear the Customer Property quote form?" & vbCrLf & vbCrLf & _
              "This will reset all entries.", vbYesNo + vbQuestion, "QAA Quote System") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ws.Range("C6").ClearContents      ' Customer email
    ws.Range("I5").ClearContents      ' Updated date
    ws.Range("H6").Value = "Standard" ' Tier
    ws.Range("C7").Value = "No"       ' Approval
    ws.Range("E7").Value = "No"       ' Shipping
    ws.Range("G7").Value = "No"       ' Payment

    ' Clear unit rows
    For r = 11 To 14
        ws.Cells(r, 2).ClearContents  ' P/N
        ws.Cells(r, 5).ClearContents  ' Base Labor
    Next r

    ' Clear O&A rows
    For r = 18 To 27
        ws.Cells(r, 1).ClearContents  ' Unit#
        ws.Cells(r, 2).ClearContents  ' Part#
        ws.Cells(r, 4).ClearContents  ' Qty
        ws.Cells(r, 6).ClearContents  ' Override
        ws.Cells(r, 8).ClearContents  ' Status
    Next r

    ws.Range("B34").ClearContents     ' Notes

    Application.ScreenUpdating = True
    MsgBox "Form cleared!", vbInformation, "QAA Quote System"
End Sub

Public Sub ClearExchange()
    Dim ws As Worksheet
    Dim r As Integer

    Set ws = ThisWorkbook.Sheets("Exchange")

    If MsgBox("Clear the Exchange quote form?" & vbCrLf & vbCrLf & _
              "This will reset all entries.", vbYesNo + vbQuestion, "QAA Quote System") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ws.Range("B6").ClearContents      ' Customer email
    ws.Range("G6").Value = "Standard" ' Tier
    ws.Range("B7").Value = "No"       ' Approval
    ws.Range("D7").Value = "No"       ' Shipping
    ws.Range("F7").Value = "No"       ' Payment
    ws.Range("H7").Value = "No"       ' Core In Hand

    ' Clear line items
    For r = 11 To 22
        ws.Cells(r, 1).ClearContents  ' Part#
        ws.Cells(r, 4).ClearContents  ' Qty
    Next r

    ws.Range("A28").ClearContents     ' Notes

    Application.ScreenUpdating = True
    MsgBox "Form cleared!", vbInformation, "QAA Quote System"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' UPDATE DATE MACRO (for revised quotes)
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub UpdateCustPropDate()
    ThisWorkbook.Sheets("Cust Property").Range("I5").Value = Date
    MsgBox "Updated date set to today!", vbInformation, "QAA Quote System"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' HELPER FUNCTIONS
' ═══════════════════════════════════════════════════════════════════════════════

Private Function Count500HRUnits() As Integer
    Dim ws As Worksheet
    Dim cnt As Integer, r As Integer

    Set ws = ThisWorkbook.Sheets("500HR")
    cnt = 0

    For r = 10 To 15
        If Trim(ws.Cells(r, 2).Value) <> "" Then cnt = cnt + 1
    Next r

    Count500HRUnits = cnt
End Function

Private Function CountCPUnits() As Integer
    Dim ws As Worksheet
    Dim cnt As Integer, r As Integer

    Set ws = ThisWorkbook.Sheets("Cust Property")
    cnt = 0

    For r = 11 To 14
        If Trim(ws.Cells(r, 2).Value) <> "" Then cnt = cnt + 1
    Next r

    CountCPUnits = cnt
End Function

Private Function CountExItems() As Integer
    Dim ws As Worksheet
    Dim cnt As Integer, r As Integer

    Set ws = ThisWorkbook.Sheets("Exchange")
    cnt = 0

    For r = 11 To 22
        If Trim(ws.Cells(r, 1).Value) <> "" Then cnt = cnt + 1
    Next r

    CountExItems = cnt
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' CHECKBOX TOGGLE (for findings)
' Call this from Worksheet_SelectionChange event
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub ToggleCheckbox(Target As Range)
    If Target.CountLarge > 1 Then Exit Sub

    Select Case Target.Value
        Case ChrW(9744)  ' ☐ unchecked
            Target.Value = ChrW(9745)  ' ☑ checked
        Case ChrW(9745)  ' ☑ checked
            Target.Value = ChrW(9744)  ' ☐ unchecked
    End Select
End Sub
