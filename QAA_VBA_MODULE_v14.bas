Attribute VB_Name = "QAAQuoteSystem"
' ============================================================================
' QAA QUOTE SYSTEM v14.0 - VBA MODULE
' Quality Aircraft Accessories - A Hartzell Engine Technologies Company
'
' FEATURES:
' - 5H suffix for 500HR pricing (NOT -5H)
' - Order # and PO # support
' - Single Quote Tracker with follow-up workflow
' - PO Tracking integration
' ============================================================================

Option Explicit

' ============================================================================
' NAVIGATION
' ============================================================================

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
    ThisWorkbook.Sheets("Exchange").Range("B5").Select
End Sub

Public Sub GoTracker()
    ThisWorkbook.Sheets("Quote Tracker").Activate
End Sub

Public Sub GoPOTracking()
    ThisWorkbook.Sheets("PO Tracking").Activate
End Sub

Public Sub GoInforXA()
    ThisWorkbook.Sheets("Infor XA Data").Activate
End Sub

Public Sub GoQuickRef()
    ThisWorkbook.Sheets("Quick Reference").Activate
End Sub

' ============================================================================
' PDF GENERATION
' ============================================================================

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
        MsgBox "Please enter customer email first.", vbExclamation
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("C4").Value)

    With ws.PageSetup
        .PrintArea = "$A$1:$J$34"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard

    MsgBox "PDF Created:" & vbCrLf & pdfPath, vbInformation

    If MsgBox("Send email with PDF?", vbYesNo) = vbYes Then
        Email500HR pdfPath
    End If
    Exit Sub

PDFError:
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

Public Sub GenPDFCustProp()
    Dim ws As Worksheet, pdfPath As String
    Set ws = ThisWorkbook.Sheets("Cust Property")

    If Trim(ws.Range("C6").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("C4").Value)

    With ws.PageSetup
        .PrintArea = "$A$1:$H$34"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard

    MsgBox "PDF Created:" & vbCrLf & pdfPath, vbInformation

    If MsgBox("Send email with PDF?", vbYesNo) = vbYes Then
        EmailCustProp pdfPath
    End If
    Exit Sub

PDFError:
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

Public Sub GenPDFExchange()
    Dim ws As Worksheet, pdfPath As String
    Set ws = ThisWorkbook.Sheets("Exchange")

    If Trim(ws.Range("B5").Value) = "" Then
        MsgBox "Please enter customer email first.", vbExclamation
        Exit Sub
    End If

    pdfPath = GetPDFPath(ws.Range("B4").Value)

    With ws.PageSetup
        .PrintArea = "$A$1:$J$27"
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    On Error GoTo PDFError
    ws.ExportAsFixedFormat xlTypePDF, pdfPath, xlQualityStandard

    MsgBox "PDF Created:" & vbCrLf & pdfPath, vbInformation

    If MsgBox("Send email with PDF?", vbYesNo) = vbYes Then
        EmailExchange pdfPath
    End If
    Exit Sub

PDFError:
    MsgBox "PDF Error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' EMAIL WITH QUOTE TRACKER LOGGING
' ============================================================================

Public Sub Email500HR(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, orderNum As String, poNum As String, totalAmt As Double

    Set ws = ThisWorkbook.Sheets("500HR")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C4").Value
    orderNum = ws.Range("C5").Value
    poNum = ws.Range("F5").Value
    totalAmt = ws.Range("H28").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - 500HR Magneto Service - QAA"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for 500HR Magneto Inspection Service." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "QUOTE SUMMARY" & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    If orderNum <> "" Then body = body & "Order #:     " & orderNum & vbCrLf
    If poNum <> "" Then body = body & "Your PO #:   " & poNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:       30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

    If ws.Range("G6").Value = "Yes" Then
        body = body & "** APPROVAL REQUIRED - Please reply to proceed **" & vbCrLf & vbCrLf
    End If

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

    LogToTracker quoteNum, emailTo, "500HR", totalAmt, ws.Range("G6").Value, "No", "No"

    MsgBox "Email created and logged!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

Public Sub EmailCustProp(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, orderNum As String, poNum As String, totalAmt As Double

    Set ws = ThisWorkbook.Sheets("Cust Property")
    emailTo = Trim(ws.Range("C6").Value)
    quoteNum = ws.Range("C4").Value
    orderNum = ws.Range("C5").Value
    poNum = ws.Range("F5").Value
    totalAmt = ws.Range("F28").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Customer Property Overhaul - QAA"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Please find attached your quote for Customer Property Overhaul." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "QUOTE SUMMARY" & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    If orderNum <> "" Then body = body & "Order #:     " & orderNum & vbCrLf
    If poNum <> "" Then body = body & "Your PO #:   " & poNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "Valid:       30 Days" & vbCrLf & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

    Dim req As String: req = ""
    If ws.Range("D7").Value = "Yes" Then req = req & "  - Approval" & vbCrLf
    If ws.Range("F7").Value = "Yes" Then req = req & "  - Shipping info" & vbCrLf
    If ws.Range("H7").Value = "Yes" Then req = req & "  - Payment method" & vbCrLf
    If req <> "" Then body = body & "TO PROCEED, WE NEED:" & vbCrLf & req & vbCrLf

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

    LogToTracker quoteNum, emailTo, "Cust Prop", totalAmt, ws.Range("D7").Value, ws.Range("F7").Value, ws.Range("H7").Value

    MsgBox "Email created and logged!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

Public Sub EmailExchange(Optional pdfPath As String = "")
    Dim OutApp As Object, OutMail As Object
    Dim ws As Worksheet
    Dim emailTo As String, subj As String, body As String
    Dim quoteNum As String, totalAmt As Double

    Set ws = ThisWorkbook.Sheets("Exchange")
    emailTo = Trim(ws.Range("B5").Value)
    quoteNum = ws.Range("B4").Value
    totalAmt = ws.Range("F21").Value

    If emailTo = "" Then
        MsgBox "Please enter customer email.", vbExclamation
        Exit Sub
    End If

    subj = "Quote " & quoteNum & " - Exchange/Sale - QAA"

    body = "Dear Valued Customer," & vbCrLf & vbCrLf
    body = body & "Please find attached your quote." & vbCrLf & vbCrLf
    body = body & String(50, "=") & vbCrLf
    body = body & "Quote #:     " & quoteNum & vbCrLf
    body = body & "Date:        " & Format(Date, "MMMM d, yyyy") & vbCrLf
    body = body & "TOTAL:       " & Format(totalAmt, "$#,##0.00") & vbCrLf
    body = body & String(50, "=") & vbCrLf & vbCrLf

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

    LogToTracker quoteNum, emailTo, "Exchange", totalAmt, ws.Range("B6").Value, ws.Range("D6").Value, ws.Range("F6").Value

    MsgBox "Email created and logged!", vbInformation
    Exit Sub

EmailError:
    MsgBox "Email Error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' QUOTE TRACKER LOGGING
' ============================================================================

Private Sub LogToTracker(quoteNum As String, emailTo As String, qType As String, _
                         total As Double, needsAppr As String, needsShip As String, needsPay As String)

    Dim wsTracker As Worksheet, wsConfig As Worksheet
    Dim nr As Long, existingRow As Long

    On Error GoTo LogError

    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    Set wsConfig = ThisWorkbook.Sheets("Config")

    existingRow = FindQuoteRow(quoteNum)

    If existingRow > 0 Then
        ' Update existing - increment reminders
        wsTracker.Cells(existingRow, 9).Value = wsTracker.Cells(existingRow, 9).Value + 1
        wsTracker.Cells(existingRow, 10).Value = Date
        wsTracker.Cells(existingRow, 13).Value = wsTracker.Cells(existingRow, 13).Value & " | " & Format(Now, "MM/DD HH:MM") & " follow-up"
    Else
        ' New entry
        nr = wsTracker.Cells(wsTracker.Rows.Count, 1).End(xlUp).Row + 1
        If nr < 4 Then nr = 4

        wsTracker.Cells(nr, 1).Value = quoteNum
        wsTracker.Cells(nr, 2).Value = Date
        wsTracker.Cells(nr, 3).Value = emailTo
        wsTracker.Cells(nr, 4).Value = qType
        wsTracker.Cells(nr, 5).Value = total
        wsTracker.Cells(nr, 6).Value = IIf(needsAppr = "Yes", "Yes", "No")
        wsTracker.Cells(nr, 7).Value = IIf(needsShip = "Yes", "Yes", "No")
        wsTracker.Cells(nr, 8).Value = IIf(needsPay = "Yes", "Yes", "No")
        wsTracker.Cells(nr, 9).Value = 1
        wsTracker.Cells(nr, 10).Value = Date
        wsTracker.Cells(nr, 12).Value = "OPEN"
        wsTracker.Cells(nr, 13).Value = Format(Now, "MM/DD HH:MM") & " - Initial quote"

        wsConfig.Range("B3").Value = wsConfig.Range("B3").Value + 1
    End If

    Exit Sub

LogError:
    MsgBox "Log Error: " & Err.Description, vbExclamation
End Sub

Private Function FindQuoteRow(quoteNum As String) As Long
    Dim wsTracker As Worksheet, r As Long
    Set wsTracker = ThisWorkbook.Sheets("Quote Tracker")
    FindQuoteRow = 0

    For r = 4 To wsTracker.Cells(wsTracker.Rows.Count, 1).End(xlUp).Row
        If wsTracker.Cells(r, 1).Value = quoteNum Then
            FindQuoteRow = r
            Exit Function
        End If
    Next r
End Function

' ============================================================================
' TRACKER MANAGEMENT
' ============================================================================

Public Sub SendReminder()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Quote Tracker")
    r = ActiveCell.Row

    If r < 4 Or ws.Cells(r, 1).Value = "" Then
        MsgBox "Select a quote row first.", vbExclamation
        Exit Sub
    End If

    ws.Cells(r, 9).Value = ws.Cells(r, 9).Value + 1
    ws.Cells(r, 10).Value = Date
    ws.Cells(r, 13).Value = ws.Cells(r, 13).Value & " | " & Format(Now, "MM/DD") & " reminder"

    MsgBox "Reminder logged! Total: " & ws.Cells(r, 9).Value, vbInformation
End Sub

Public Sub CloseQuote()
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets("Quote Tracker")
    r = ActiveCell.Row

    If r < 4 Or ws.Cells(r, 1).Value = "" Then
        MsgBox "Select a quote row first.", vbExclamation
        Exit Sub
    End If

    ws.Cells(r, 12).Value = "CLOSED"
    ws.Cells(r, 13).Value = ws.Cells(r, 13).Value & " | " & Format(Now, "MM/DD") & " CLOSED"

    MsgBox "Quote marked CLOSED.", vbInformation
End Sub

' ============================================================================
' CLEAR FORMS
' ============================================================================

Public Sub Clear500HR()
    Dim ws As Worksheet, r As Integer, c As Integer
    Set ws = ThisWorkbook.Sheets("500HR")

    If MsgBox("Clear form?", vbYesNo) = vbNo Then Exit Sub

    ws.Range("C5").ClearContents
    ws.Range("F5").ClearContents
    ws.Range("C6").ClearContents
    ws.Range("G6").Value = "No"

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

    ws.Range("B31").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub ClearCustProp()
    Dim ws As Worksheet, r As Integer
    Set ws = ThisWorkbook.Sheets("Cust Property")

    If MsgBox("Clear form?", vbYesNo) = vbNo Then Exit Sub

    ws.Range("C5").ClearContents
    ws.Range("F5").ClearContents
    ws.Range("C6").ClearContents
    ws.Range("H4").ClearContents
    ws.Range("G6").Value = "Standard"
    ws.Range("D7").Value = "No"
    ws.Range("F7").Value = "No"
    ws.Range("H7").Value = "No"

    For r = 10 To 13
        ws.Cells(r, 2).ClearContents
        ws.Cells(r, 5).ClearContents
    Next r

    For r = 17 To 24
        ws.Cells(r, 1).ClearContents
        ws.Cells(r, 2).ClearContents
        ws.Cells(r, 4).ClearContents
        ws.Cells(r, 7).ClearContents
    Next r

    ws.Range("B31").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub ClearExchange()
    Dim ws As Worksheet, r As Integer
    Set ws = ThisWorkbook.Sheets("Exchange")

    If MsgBox("Clear form?", vbYesNo) = vbNo Then Exit Sub

    ws.Range("B5").ClearContents
    ws.Range("G5").Value = "Standard"
    ws.Range("B6").Value = "No"
    ws.Range("D6").Value = "No"
    ws.Range("F6").Value = "No"
    ws.Range("H6").Value = "No"

    For r = 9 To 18
        ws.Cells(r, 1).ClearContents
        ws.Cells(r, 4).ClearContents
    Next r

    ws.Range("A24").ClearContents
    MsgBox "Cleared!", vbInformation
End Sub

Public Sub UpdateInforXADate()
    ThisWorkbook.Sheets("Infor XA Data").Range("B6").Value = Now
    MsgBox "Timestamp updated!", vbInformation
End Sub
