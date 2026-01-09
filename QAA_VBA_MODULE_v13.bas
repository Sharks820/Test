Attribute VB_Name = "QAAQuoteSystem"
' ╔══════════════════════════════════════════════════════════════════════════════╗
' ║     QAA QUOTE SYSTEM v13.0 - PROFESSIONAL EDITION                            ║
' ║     Quality Aircraft Accessories - A Hartzell Engine Technologies Company    ║
' ║                                                                              ║
' ║     FEATURES:                                                                ║
' ║     - Uses 5H suffix (NOT -5H) for 500HR pricing                            ║
' ║     - Single Quote Tracker with follow-up workflow                          ║
' ║     - Reminder counting and status management                               ║
' ║     - Professional PDF generation                                           ║
' ╚══════════════════════════════════════════════════════════════════════════════╝

Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
' NAVIGATION
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub GoMenu()
    ThisWorkbook.Sheets("Menu").Activate
End Sub

Public Sub Go500HR()
    ThisWorkbook.Sheets("500HR").Activate
    ThisWorkbook.Sheets("500HR").Range("C6").Select
End Sub

Public Sub GoCustProp()
    ThisWorkbook.Sheets("Cust Property").Activate
    ThisWorkbook.Sheets("Cust Property").Range("C6").Select
End Sub

Public Sub GoExchange()
    ThisWorkbook.Sheets("Exchange").Activate
    ThisWorkbook.Sheets("Exchange").Range("B6").Select
End Sub

Public Sub GoTracker()
    ThisWorkbook.Sheets("Quote Tracker").Activate
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PDF GENERATION - Professional Layout
' ═══════════════════════════════════════════════════════════════════════════════

Private Function GetPDFPath(quoteNum As String) As String
    Dim folder As String
    folder = ThisWorkbook.Path
    If folder = "" Then folder = Environ("USERPROFILE") & "\Documents"
    GetPDFPath = folder & "\QAA_Quote_" & Replace(Replace(quoteNum, "-", "_"), "/", "_") & "_" & Format(Date, "YYYYMMDD") & ".pdf"
End Function

Public Sub GenPDF500HR()
    Dim ws As Worksheet, pdfPath As String
    Set ws = ThisWorkbook.Sheets("500HR")

    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation, "QAA Quote System"
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
        .PrintGridlines = False
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard
    Application.ScreenUpdating = True

    MsgBox "PDF Created:" & vbCrLf & vbCrLf & pdfPath, vbInformation, "QAA Quote System"

    If MsgBox("Send email with PDF attached?", vbYesNo + vbQuestion) = vbYes Then
        Email500HR pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

Public Sub GenPDFCustProp()
    Dim ws As Worksheet, pdfPath As String
    Set ws = ThisWorkbook.Sheets("Cust Property")

    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("C5").Value)

    Application.ScreenUpdating = False
    With ws.PageSetup
        .PrintArea = "$A$1:$I$36"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard
    Application.ScreenUpdating = True

    MsgBox "PDF Created:" & vbCrLf & vbCrLf & pdfPath, vbInformation

    If MsgBox("Send email with PDF attached?", vbYesNo + vbQuestion) = vbYes Then
        EmailCustProp pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

Public Sub GenPDFExchange()
    Dim ws As Worksheet, pdfPath As String
    Set ws = ThisWorkbook.Sheets("Exchange")

    If Trim(ws.Range("B6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("B5").Value)

    Application.ScreenUpdating = False
    With ws.PageSetup
        .PrintArea = "$A$1:$J$31"
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard
    Application.ScreenUpdating = True

    MsgBox "PDF Created:" & vbCrLf & vbCrLf & pdfPath, vbInformation

    If MsgBox("Send email with PDF attached?", vbYesNo + vbQuestion) = vbYes Then
        EmailExchange pdfPath
    End If
    Exit Sub

PDFError:
    Application.ScreenUpdating = True
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' EMAIL GENERATION WITH QUOTE TRACKER LOGGING
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub Email500HR(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double
    Dim needsApproval As String

    Set ws = ThisWorkbook.Sheets("500HR")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C5").Value
    totalAmt = ws.Range("H29").Value
    needsApproval = ws.Range("H6").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - 500HR Magneto Inspection - Quality Aircraft Accessories"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for choosing Quality Aircraft Accessories for your magneto service needs." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for 500HR Magneto Inspection Service." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "QUOTE SUMMARY" & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:       30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

    If needsApproval = "Yes" Then
        body = body & "** APPROVAL REQUIRED **" & vbCrLf
        body = body & "Please reply with your approval to proceed." & vbCrLf & vbCrLf
    End If

    body = body & "Questions? Contact us anytime." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
    body = body & "A Hartzell Engine Technologies Company" & vbCrLf
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

    ' Log to Quote Tracker
    LogToTracker quoteNum, emailTo, "500HR Service", totalAmt, needsApproval, "No", "No", "Initial quote sent"

    MsgBox "Email created and logged to Quote Tracker!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

Public Sub EmailCustProp(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double
    Dim needsAppr As String, needsShip As String, needsPay As String

    Set ws = ThisWorkbook.Sheets("Cust Property")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C5").Value
    totalAmt = ws.Range("G30").Value
    needsAppr = ws.Range("D7").Value
    needsShip = ws.Range("F7").Value
    needsPay = ws.Range("H7").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Customer Property Overhaul - Quality Aircraft Accessories"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for choosing Quality Aircraft Accessories." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for Customer Property Overhaul services." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "QUOTE SUMMARY" & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:       30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

    Dim req As String: req = ""
    If needsAppr = "Yes" Then req = req & "  - Approval to proceed" & vbCrLf
    If needsShip = "Yes" Then req = req & "  - Shipping information" & vbCrLf
    If needsPay = "Yes" Then req = req & "  - Payment method" & vbCrLf
    If req <> "" Then body = body & "TO PROCEED, WE NEED:" & vbCrLf & req & vbCrLf

    body = body & "Questions? Contact us anytime." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
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

    LogToTracker quoteNum, emailTo, "Cust Property", totalAmt, needsAppr, needsShip, needsPay, "Initial quote sent"

    MsgBox "Email created and logged to Quote Tracker!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

Public Sub EmailExchange(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double
    Dim needsAppr As String, needsShip As String, needsPay As String

    Set ws = ThisWorkbook.Sheets("Exchange")
    emailTo = Trim(ws.Range("B6").Value)
    quoteNum = ws.Range("B5").Value
    totalAmt = ws.Range("F24").Value
    needsAppr = ws.Range("B7").Value
    needsShip = ws.Range("D7").Value
    needsPay = ws.Range("F7").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Exchange/Sale - Quality Aircraft Accessories"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Thank you for your interest in Quality Aircraft Accessories." & vbCrLf & vbCrLf
    body = body & "Please find attached your quote." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "QUOTE SUMMARY" & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:       30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

    Dim req As String: req = ""
    If needsAppr = "Yes" Then req = req & "  - Approval" & vbCrLf
    If needsShip = "Yes" Then req = req & "  - Shipping info" & vbCrLf
    If needsPay = "Yes" Then req = req & "  - Payment method" & vbCrLf
    If req <> "" Then body = body & "TO PROCEED:" & vbCrLf & req & vbCrLf

    body = body & "Questions? Contact us anytime." & vbCrLf & vbCrLf
    body = body & "Best regards," & vbCrLf
    body = body & "Quality Aircraft Accessories" & vbCrLf
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

    LogToTracker quoteNum, emailTo, "Exchange/Sale", totalAmt, needsAppr, needsShip, needsPay, "Initial quote sent"

    MsgBox "Email created and logged to Quote Tracker!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' QUOTE TRACKER LOGGING
' ═══════════════════════════════════════════════════════════════════════════════

Private Sub LogToTracker(quoteNum As String, emailTo As String, qType As String, _
                         total As Double, needsAppr As String, needsShip As String, _
                         needsPay As String, notes As String)

    Dim wsTracker As Worksheet, wsConfig As Worksheet
    Dim nr As Long
    Dim existingRow As Long

    On Error GoTo LogError

    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    Set wsConfig = ThisWorkbook.Sheets("Config")

    ' Check if quote already exists (for follow-ups)
    existingRow = FindQuoteRow(quoteNum)

    If existingRow > 0 Then
        ' Update existing entry (increment reminder count)
        wsTracker.Cells(existingRow, 9).Value = wsTracker.Cells(existingRow, 9).Value + 1  ' Reminders
        wsTracker.Cells(existingRow, 10).Value = Date  ' Last Contact
        wsTracker.Cells(existingRow, 13).Value = wsTracker.Cells(existingRow, 13).Value & vbCrLf & _
            Format(Now, "MM/DD HH:MM") & " - " & notes
    Else
        ' New entry
        nr = wsTracker.Cells(wsTracker.Rows.Count, 1).End(xlUp).Row + 1
        If nr < 4 Then nr = 4  ' Start after headers

        wsTracker.Cells(nr, 1).Value = quoteNum                     ' Quote #
        wsTracker.Cells(nr, 2).Value = Date                          ' Created
        wsTracker.Cells(nr, 3).Value = emailTo                       ' Customer Email
        wsTracker.Cells(nr, 4).Value = qType                         ' Type
        wsTracker.Cells(nr, 5).Value = total                         ' Amount
        wsTracker.Cells(nr, 6).Value = IIf(needsAppr = "Yes", "Yes", "No")  ' Approval
        wsTracker.Cells(nr, 7).Value = IIf(needsShip = "Yes", "Yes", "No")  ' Shipping
        wsTracker.Cells(nr, 8).Value = IIf(needsPay = "Yes", "Yes", "No")   ' Payment
        wsTracker.Cells(nr, 9).Value = 1                             ' Reminders (first contact)
        wsTracker.Cells(nr, 10).Value = Date                         ' Last Contact
        wsTracker.Cells(nr, 12).Value = "OPEN"                       ' Status
        wsTracker.Cells(nr, 13).Value = Format(Now, "MM/DD HH:MM") & " - " & notes  ' Notes

        ' Increment quote number
        wsConfig.Range("B3").Value = wsConfig.Range("B3").Value + 1
    End If

    Exit Sub

LogError:
    MsgBox "Tracker Log Error: " & Err.Description, vbExclamation
End Sub

Private Function FindQuoteRow(quoteNum As String) As Long
    Dim wsTracker As Worksheet
    Dim r As Long

    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    FindQuoteRow = 0

    For r = 4 To wsTracker.Cells(wsTracker.Rows.Count, 1).End(xlUp).Row
        If wsTracker.Cells(r, 1).Value = quoteNum Then
            FindQuoteRow = r
            Exit Function
        End If
    Next r
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' REMINDER & STATUS MANAGEMENT
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub SendReminder()
    ' Send a reminder for the currently selected quote in tracker
    Dim wsTracker As Worksheet
    Dim r As Long
    Dim quoteNum As String, emailTo As String

    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    r = ActiveCell.Row

    If r < 4 Then
        MsgBox "Please select a quote row first.", vbExclamation
        Exit Sub
    End If

    quoteNum = wsTracker.Cells(r, 1).Value
    emailTo = wsTracker.Cells(r, 3).Value

    If quoteNum = "" Then
        MsgBox "No quote selected.", vbExclamation
        Exit Sub
    End If

    ' Update tracker
    wsTracker.Cells(r, 9).Value = wsTracker.Cells(r, 9).Value + 1  ' Increment reminders
    wsTracker.Cells(r, 10).Value = Date  ' Update last contact
    wsTracker.Cells(r, 13).Value = wsTracker.Cells(r, 13).Value & vbCrLf & _
        Format(Now, "MM/DD HH:MM") & " - Reminder sent"

    MsgBox "Reminder logged for " & quoteNum & vbCrLf & _
           "Total reminders: " & wsTracker.Cells(r, 9).Value, vbInformation
End Sub

Public Sub CloseQuote()
    ' Mark the selected quote as CLOSED
    Dim wsTracker As Worksheet
    Dim r As Long

    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    r = ActiveCell.Row

    If r < 4 Then
        MsgBox "Please select a quote row first.", vbExclamation
        Exit Sub
    End If

    wsTracker.Cells(r, 12).Value = "CLOSED"
    wsTracker.Cells(r, 13).Value = wsTracker.Cells(r, 13).Value & vbCrLf & _
        Format(Now, "MM/DD HH:MM") & " - Quote CLOSED"

    MsgBox "Quote marked as CLOSED.", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' CLEAR FORMS
' ═══════════════════════════════════════════════════════════════════════════════

Public Sub Clear500HR()
    Dim ws As Worksheet, r As Integer, c As Integer
    Set ws = ThisWorkbook.Sheets("500HR")

    If MsgBox("Clear 500HR form?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ws.Range("C6").ClearContents
    ws.Range("H6").Value = "No"

    For r = 9 To 14
        ws.Cells(r, 2).ClearContents
        ws.Cells(r, 3).ClearContents
        ws.Cells(r, 6).ClearContents
    Next r

    For r = 18 To 23
        For c = 2 To 9
            ws.Cells(r, c).Value = ChrW(9744)
        Next c
    Next r

    ws.Range("B32").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub ClearCustProp()
    Dim ws As Worksheet, r As Integer
    Set ws = ThisWorkbook.Sheets("Cust Property")

    If MsgBox("Clear Cust Property form?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ws.Range("C6").ClearContents
    ws.Range("H5").ClearContents
    ws.Range("G6").Value = "Standard"
    ws.Range("D7").Value = "No"
    ws.Range("F7").Value = "No"
    ws.Range("H7").Value = "No"

    For r = 10 To 13
        ws.Cells(r, 2).ClearContents
        ws.Cells(r, 5).ClearContents
    Next r

    For r = 17 To 26
        ws.Cells(r, 1).ClearContents
        ws.Cells(r, 2).ClearContents
        ws.Cells(r, 4).ClearContents
        ws.Cells(r, 6).ClearContents
        ws.Cells(r, 8).ClearContents
    Next r

    ws.Range("B33").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub ClearExchange()
    Dim ws As Worksheet, r As Integer
    Set ws = ThisWorkbook.Sheets("Exchange")

    If MsgBox("Clear Exchange form?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ws.Range("B6").ClearContents
    ws.Range("G6").Value = "Standard"
    ws.Range("B7").Value = "No"
    ws.Range("D7").Value = "No"
    ws.Range("F7").Value = "No"
    ws.Range("H7").Value = "No"

    For r = 10 To 21
        ws.Cells(r, 1).ClearContents
        ws.Cells(r, 4).ClearContents
    Next r

    ws.Range("A27").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub UpdateDate()
    ThisWorkbook.Sheets("Cust Property").Range("H5").Value = Date
    MsgBox "Date updated!", vbInformation
End Sub
