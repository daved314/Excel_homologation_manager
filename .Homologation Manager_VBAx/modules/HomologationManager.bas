Attribute VB_Name = "HomologationManager"
Option Explicit

Private Const HM_DATA_SHEET As String = "Projects"
Private Const HM_LIST_SHEET As String = "Lists"

Private Const HM_COL_PROJECT_NAME As Long = 1
Private Const HM_COL_START_DATE As Long = 2
Private Const HM_COL_HOMOLOGATION_TYPE As Long = 3
Private Const HM_COL_HOMOLOGATION_SPEC As Long = 4
Private Const HM_COL_APPLICATION_NO As Long = 5
Private Const HM_COL_PO_NO As Long = 6
Private Const HM_COL_INVOICE_NO As Long = 7
Private Const HM_COL_CERTIFICATE_NO As Long = 8
Private Const HM_COL_CLOSE_DATE As Long = 9
Private Const HM_COL_COMMENT As Long = 10
Private Const HM_COL_STATUS As Long = 11
Private Const HM_COL_LAST_UPDATED As Long = 12

Public Sub HM_OpenManager()
    Dim frm As Object

    On Error GoTo CleanFail

    Set frm = CreateUserFormInstance()
    If frm Is Nothing Then
        Err.Raise vbObjectError + 8101, "HM_OpenManager", "Could not create runtime UserForm."
    End If

    HM_BuildRuntimeUi frm
    HM_PopulateDropdowns frm
    HM_PopulateList frm
    HM_SetNewRecordState frm

    frm.caption = "Homologation Manager"
    frm.Width = 900
    frm.Height = 520

    frm.Show vbModeless
    HM_RunFormLoop frm

CleanExit:
    On Error Resume Next
    Unload frm
    Set frm = Nothing
    Exit Sub

CleanFail:
    MsgBox "Unable to open Homologation Manager form." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub HM_ReportOpenProjects()
    HM_CreateStatusReport "Open"
End Sub

Public Sub HM_ReportClosedProjects()
    HM_CreateStatusReport "Closed"
End Sub

Private Function CreateUserFormInstance() As Object
    Dim frm As Object
    Dim vbComp As Object

    On Error GoTo CreateFail

    On Error Resume Next
    Set frm = VBA.UserForms.Add("UserForm1")
    On Error GoTo CreateFail
    If Not frm Is Nothing Then
        Set CreateUserFormInstance = frm
        Exit Function
    End If

    Set vbComp = EnsureUserFormComponent("UserForm1")
    If vbComp Is Nothing Then GoTo CreateFail

    On Error Resume Next
    Set frm = VBA.UserForms.Add(vbComp.Name)
    On Error GoTo CreateFail
    If Not frm Is Nothing Then
        Set CreateUserFormInstance = frm
        Exit Function
    End If

    On Error Resume Next
    Set frm = vbComp.Designer
    On Error GoTo CreateFail
    If Not frm Is Nothing Then
        Set CreateUserFormInstance = frm
        Exit Function
    End If

CreateFail:
    Set CreateUserFormInstance = Nothing
End Function

Private Function EnsureUserFormComponent(ByVal formName As String) As Object
    Const VBEXT_CT_MSFORM As Long = 3
    Dim vbComp As Object

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        On Error GoTo EnsureFail
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(VBEXT_CT_MSFORM)
        On Error Resume Next
        vbComp.Name = formName
        On Error GoTo EnsureFail
    End If

    ClearVBComponentCode vbComp
    Set EnsureUserFormComponent = vbComp
    Exit Function

EnsureFail:
    Set EnsureUserFormComponent = Nothing
End Function

Private Sub ClearVBComponentCode(ByVal vbComp As Object)
    Dim cm As Object
    Dim lineCount As Long

    On Error Resume Next
    Set cm = vbComp.CodeModule
    If cm Is Nothing Then Exit Sub

    lineCount = cm.CountOfLines
    If lineCount > 0 Then cm.DeleteLines 1, lineCount
End Sub

Private Sub HM_BuildRuntimeUi(ByVal frm As Object)
    Dim lbl As Object
    Dim i As Long

    RemoveIfExists frm, "lstRecords"
    RemoveIfExists frm, "lblListHeader"
    RemoveIfExists frm, "lblFieldsHeader"
    RemoveIfExists frm, "txtRecordRow"
    RemoveIfExists frm, "btnSave"
    RemoveIfExists frm, "btnNew"
    RemoveIfExists frm, "btnReportOpen"
    RemoveIfExists frm, "btnReportClosed"
    RemoveIfExists frm, "btnClose"

    For i = 1 To 8
        RemoveIfExists frm, "lblField" & CStr(i)
    Next i
    RemoveIfExists frm, "lblComment"
    RemoveIfExists frm, "lblCloseDate"

    RemoveIfExists frm, "txtProjectName"
    RemoveIfExists frm, "txtStartDate"
    RemoveIfExists frm, "cboType"
    RemoveIfExists frm, "cboSpec"
    RemoveIfExists frm, "txtApplicationNo"
    RemoveIfExists frm, "txtPONo"
    RemoveIfExists frm, "txtInvoiceNo"
    RemoveIfExists frm, "txtCertificateNo"
    RemoveIfExists frm, "txtCloseDate"
    RemoveIfExists frm, "txtComment"
    RemoveIfExists frm, "txtStatus"

    Set lbl = EnsureControl(frm, "lblListHeader", "Forms.Label.1")
    With lbl
        .caption = "Projects"
        .Left = 18
        .Top = 16
        .Width = 240
        .Height = 18
        .Font.Bold = True
        .Font.Size = 11
    End With

    Set lbl = EnsureControl(frm, "lblFieldsHeader", "Forms.Label.1")
    With lbl
        .caption = "Project Details"
        .Left = 322
        .Top = 16
        .Width = 220
        .Height = 18
        .Font.Bold = True
        .Font.Size = 11
    End With

    With EnsureControl(frm, "lstRecords", "Forms.ListBox.1")
        .Left = 18
        .Top = 40
        .Width = 290
        .Height = 398
        .ColumnCount = 3
        .ColumnWidths = "0 pt;170 pt;90 pt"
        .BoundColumn = 1
    End With

    HM_AddLabel frm, "lblField1", "Project Name", 322, 44
    With EnsureControl(frm, "txtProjectName", "Forms.TextBox.1")
        .Left = 502
        .Top = 40
        .Width = 360
        .Height = 18
    End With

    HM_AddLabel frm, "lblField2", "Project Start Date", 322, 72
    With EnsureControl(frm, "txtStartDate", "Forms.TextBox.1")
        .Left = 502
        .Top = 68
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblField3", "Homologation Type", 322, 100
    With EnsureControl(frm, "cboType", "Forms.ComboBox.1")
        .Left = 502
        .Top = 96
        .Width = 160
        .Height = 18
        .Style = 0
        .MatchRequired = False
    End With

    HM_AddLabel frm, "lblField4", "Homologation Spec", 322, 128
    With EnsureControl(frm, "cboSpec", "Forms.ComboBox.1")
        .Left = 502
        .Top = 124
        .Width = 160
        .Height = 18
        .Style = 0
        .MatchRequired = False
    End With

    HM_AddLabel frm, "lblField5", "Application Number", 322, 156
    With EnsureControl(frm, "txtApplicationNo", "Forms.TextBox.1")
        .Left = 502
        .Top = 152
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblField6", "PO #", 322, 184
    With EnsureControl(frm, "txtPONo", "Forms.TextBox.1")
        .Left = 502
        .Top = 180
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblField7", "Invoice #", 322, 212
    With EnsureControl(frm, "txtInvoiceNo", "Forms.TextBox.1")
        .Left = 502
        .Top = 208
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblField8", "Certificate #", 322, 240
    With EnsureControl(frm, "txtCertificateNo", "Forms.TextBox.1")
        .Left = 502
        .Top = 236
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblCloseDate", "Close Date", 322, 268
    With EnsureControl(frm, "txtCloseDate", "Forms.TextBox.1")
        .Left = 502
        .Top = 264
        .Width = 160
        .Height = 18
    End With

    HM_AddLabel frm, "lblComment", "Comment", 322, 296
    With EnsureControl(frm, "txtComment", "Forms.TextBox.1")
        .Left = 502
        .Top = 292
        .Width = 360
        .Height = 78
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = True
    End With

    HM_AddLabel frm, "lblStatus", "Status", 322, 382
    With EnsureControl(frm, "txtStatus", "Forms.TextBox.1")
        .Left = 502
        .Top = 378
        .Width = 160
        .Height = 18
        .Locked = True
        .Enabled = False
    End With

    With EnsureControl(frm, "txtRecordRow", "Forms.TextBox.1")
        .Visible = False
        .Text = ""
    End With

    With EnsureControl(frm, "btnSave", "Forms.ToggleButton.1")
        .caption = "Save"
        .Left = 322
        .Top = 414
        .Width = 90
        .Height = 26
        .Value = False
    End With

    With EnsureControl(frm, "btnNew", "Forms.ToggleButton.1")
        .caption = "New"
        .Left = 420
        .Top = 414
        .Width = 90
        .Height = 26
        .Value = False
    End With

    With EnsureControl(frm, "btnReportOpen", "Forms.ToggleButton.1")
        .caption = "Report Open"
        .Left = 518
        .Top = 414
        .Width = 110
        .Height = 26
        .Value = False
    End With

    With EnsureControl(frm, "btnReportClosed", "Forms.ToggleButton.1")
        .caption = "Report Closed"
        .Left = 636
        .Top = 414
        .Width = 112
        .Height = 26
        .Value = False
    End With

    With EnsureControl(frm, "btnClose", "Forms.ToggleButton.1")
        .caption = "Close"
        .Left = 756
        .Top = 414
        .Width = 90
        .Height = 26
        .Value = False
    End With
End Sub

Private Sub HM_AddLabel(ByVal frm As Object, ByVal controlName As String, ByVal caption As String, ByVal leftPos As Long, ByVal topPos As Long)
    With EnsureControl(frm, controlName, "Forms.Label.1")
        .caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = 170
        .Height = 16
    End With
End Sub

Private Function EnsureControl(ByVal frm As Object, ByVal controlName As String, ByVal progId As String) As Object
    Dim ctl As Object

    On Error Resume Next
    Set ctl = frm.Controls(controlName)
    On Error GoTo 0

    If ctl Is Nothing Then
        On Error Resume Next
        Set ctl = frm.Controls.Add(progId, controlName, True)
        On Error GoTo 0
    End If

    If ctl Is Nothing Then
        Err.Raise vbObjectError + 8102, "EnsureControl", _
                  "Could not create control '" & controlName & "' (" & progId & ")."
    End If

    Set EnsureControl = ctl
End Function

Private Sub RemoveIfExists(ByVal frm As Object, ByVal controlName As String)
    On Error Resume Next
    frm.Controls.Remove controlName
End Sub

Private Sub HM_RunFormLoop(ByVal frm As Object)
    Dim lastListIndex As Long
    Dim currentListIndex As Long

    On Error GoTo LoopFail

    lastListIndex = -2

    Do
        DoEvents
        If Not IsFormAlive(frm) Then Exit Do
        If Not frm.Visible Then Exit Do

        currentListIndex = frm.Controls("lstRecords").ListIndex
        If currentListIndex <> lastListIndex Then
            If currentListIndex >= 0 Then HM_LoadSelectedRecord frm
            lastListIndex = currentListIndex
        End If

        If CBool(frm.Controls("btnSave").Value) Then
            frm.Controls("btnSave").Value = False
            HM_SaveRecord frm
            HM_PopulateDropdowns frm
            HM_PopulateList frm
            lastListIndex = frm.Controls("lstRecords").ListIndex
        End If

        If CBool(frm.Controls("btnNew").Value) Then
            frm.Controls("btnNew").Value = False
            HM_SetNewRecordState frm
            lastListIndex = frm.Controls("lstRecords").ListIndex
        End If

        If CBool(frm.Controls("btnReportOpen").Value) Then
            frm.Controls("btnReportOpen").Value = False
            HM_CreateStatusReport "Open"
        End If

        If CBool(frm.Controls("btnReportClosed").Value) Then
            frm.Controls("btnReportClosed").Value = False
            HM_CreateStatusReport "Closed"
        End If

        If CBool(frm.Controls("btnClose").Value) Then
            frm.Controls("btnClose").Value = False
            frm.Hide
            Exit Do
        End If
    Loop
    Exit Sub

LoopFail:
    MsgBox "Form loop failed. Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function IsFormAlive(ByVal frm As Object) As Boolean
    On Error GoTo NotAlive
    Dim v As Boolean
    v = frm.Visible
    IsFormAlive = True
    Exit Function
NotAlive:
    IsFormAlive = False
End Function

Private Sub HM_SetNewRecordState(ByVal frm As Object)
    frm.Controls("txtRecordRow").Text = ""
    frm.Controls("txtProjectName").Text = ""
    frm.Controls("txtStartDate").Text = ""
    frm.Controls("cboType").Text = ""
    frm.Controls("cboSpec").Text = ""
    frm.Controls("txtApplicationNo").Text = ""
    frm.Controls("txtPONo").Text = ""
    frm.Controls("txtInvoiceNo").Text = ""
    frm.Controls("txtCertificateNo").Text = ""
    frm.Controls("txtCloseDate").Text = ""
    frm.Controls("txtComment").Text = ""
    frm.Controls("txtStatus").Text = "Open"
    frm.Controls("lstRecords").ListIndex = -1
End Sub

Private Sub HM_SaveRecord(ByVal frm As Object)
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim closeDate As String
    Dim statusValue As String
    Dim rowToken As String
    If Trim$(frm.Controls("txtProjectName").Text) = "" Then
        MsgBox "Project Name is required.", vbExclamation
        Exit Sub
    End If

    Set ws = HM_EnsureDataSheet()

    rowToken = Trim$(frm.Controls("txtRecordRow").Text)
    If rowToken = "" Then
        rowNum = ws.Cells(ws.Rows.Count, HM_COL_PROJECT_NAME).End(xlUp).Row + 1
        If rowNum < 2 Then rowNum = 2
    Else
        rowNum = CLng(rowToken)
    End If

    ws.Cells(rowNum, HM_COL_PROJECT_NAME).Value = Trim$(frm.Controls("txtProjectName").Text)
    ws.Cells(rowNum, HM_COL_START_DATE).Value = HM_CleanDateValue(frm.Controls("txtStartDate").Text)
    ws.Cells(rowNum, HM_COL_HOMOLOGATION_TYPE).Value = Trim$(frm.Controls("cboType").Text)
    ws.Cells(rowNum, HM_COL_HOMOLOGATION_SPEC).Value = Trim$(frm.Controls("cboSpec").Text)
    ws.Cells(rowNum, HM_COL_APPLICATION_NO).Value = Trim$(frm.Controls("txtApplicationNo").Text)
    ws.Cells(rowNum, HM_COL_PO_NO).Value = Trim$(frm.Controls("txtPONo").Text)
    ws.Cells(rowNum, HM_COL_INVOICE_NO).Value = Trim$(frm.Controls("txtInvoiceNo").Text)
    ws.Cells(rowNum, HM_COL_CERTIFICATE_NO).Value = Trim$(frm.Controls("txtCertificateNo").Text)
    ws.Cells(rowNum, HM_COL_CLOSE_DATE).Value = HM_CleanDateValue(frm.Controls("txtCloseDate").Text)
    ws.Cells(rowNum, HM_COL_COMMENT).Value = frm.Controls("txtComment").Text

    closeDate = Trim$(frm.Controls("txtCloseDate").Text)
    If closeDate = "" Then
        statusValue = "Open"
    Else
        statusValue = "Closed"
    End If
    ws.Cells(rowNum, HM_COL_STATUS).Value = statusValue
    ws.Cells(rowNum, HM_COL_LAST_UPDATED).Value = Now

    frm.Controls("txtRecordRow").Text = CStr(rowNum)
    frm.Controls("txtStatus").Text = statusValue

    HM_AppendListValue HM_LIST_SHEET, 1, Trim$(frm.Controls("cboType").Text)
    HM_AppendListValue HM_LIST_SHEET, 2, Trim$(frm.Controls("cboSpec").Text)

    HM_SelectListRowBySheetRow frm, rowNum
    MsgBox "Record saved.", vbInformation
End Sub

Private Function HM_CleanDateValue(ByVal inputValue As String) As Variant
    Dim trimmedValue As String
    trimmedValue = Trim$(inputValue)

    If trimmedValue = "" Then
        HM_CleanDateValue = ""
        Exit Function
    End If

    If IsDate(trimmedValue) Then
        HM_CleanDateValue = CDate(trimmedValue)
    Else
        HM_CleanDateValue = trimmedValue
    End If
End Function

Private Sub HM_LoadSelectedRecord(ByVal frm As Object)
    Dim ws As Worksheet
    Dim lst As Object
    Dim rowNum As Long

    Set lst = frm.Controls("lstRecords")
    If lst.ListIndex < 0 Then Exit Sub

    rowNum = CLng(lst.List(lst.ListIndex, 0))
    Set ws = HM_EnsureDataSheet()

    frm.Controls("txtRecordRow").Text = CStr(rowNum)
    frm.Controls("txtProjectName").Text = CStr(ws.Cells(rowNum, HM_COL_PROJECT_NAME).Value)
    frm.Controls("txtStartDate").Text = HM_DisplayDate(ws.Cells(rowNum, HM_COL_START_DATE).Value)
    frm.Controls("cboType").Text = CStr(ws.Cells(rowNum, HM_COL_HOMOLOGATION_TYPE).Value)
    frm.Controls("cboSpec").Text = CStr(ws.Cells(rowNum, HM_COL_HOMOLOGATION_SPEC).Value)
    frm.Controls("txtApplicationNo").Text = CStr(ws.Cells(rowNum, HM_COL_APPLICATION_NO).Value)
    frm.Controls("txtPONo").Text = CStr(ws.Cells(rowNum, HM_COL_PO_NO).Value)
    frm.Controls("txtInvoiceNo").Text = CStr(ws.Cells(rowNum, HM_COL_INVOICE_NO).Value)
    frm.Controls("txtCertificateNo").Text = CStr(ws.Cells(rowNum, HM_COL_CERTIFICATE_NO).Value)
    frm.Controls("txtCloseDate").Text = HM_DisplayDate(ws.Cells(rowNum, HM_COL_CLOSE_DATE).Value)
    frm.Controls("txtComment").Text = CStr(ws.Cells(rowNum, HM_COL_COMMENT).Value)
    frm.Controls("txtStatus").Text = CStr(ws.Cells(rowNum, HM_COL_STATUS).Value)
End Sub

Private Function HM_DisplayDate(ByVal cellValue As Variant) As String
    If IsDate(cellValue) Then
        HM_DisplayDate = Format$(CDate(cellValue), "yyyy-mm-dd")
    Else
        HM_DisplayDate = Trim$(CStr(cellValue))
    End If
End Function

Private Sub HM_PopulateList(ByVal frm As Object)
    Dim ws As Worksheet
    Dim lst As Object
    Dim lastRow As Long
    Dim r As Long
    Dim statusText As String
    Dim projectText As String

    Set ws = HM_EnsureDataSheet()
    Set lst = frm.Controls("lstRecords")
    lst.Clear

    lastRow = ws.Cells(ws.Rows.Count, HM_COL_PROJECT_NAME).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        projectText = Trim$(CStr(ws.Cells(r, HM_COL_PROJECT_NAME).Value))
        If projectText <> "" Then
            statusText = Trim$(CStr(ws.Cells(r, HM_COL_STATUS).Value))
            lst.AddItem CStr(r)
            lst.List(lst.ListCount - 1, 1) = projectText
            lst.List(lst.ListCount - 1, 2) = statusText
        End If
    Next r
End Sub

Private Sub HM_SelectListRowBySheetRow(ByVal frm As Object, ByVal rowNum As Long)
    Dim lst As Object
    Dim i As Long

    Set lst = frm.Controls("lstRecords")
    For i = 0 To lst.ListCount - 1
        If CLng(lst.List(i, 0)) = rowNum Then
            lst.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub HM_PopulateDropdowns(ByVal frm As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As String

    Set ws = HM_EnsureListSheet()

    frm.Controls("cboType").Clear
    frm.Controls("cboSpec").Clear

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        For r = 2 To lastRow
            v = Trim$(CStr(ws.Cells(r, 1).Value))
            If v <> "" Then frm.Controls("cboType").AddItem v
        Next r
    End If

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow >= 2 Then
        For r = 2 To lastRow
            v = Trim$(CStr(ws.Cells(r, 2).Value))
            If v <> "" Then frm.Controls("cboSpec").AddItem v
        Next r
    End If
End Sub

Private Function HM_EnsureDataSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HM_DATA_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HM_DATA_SHEET
    End If

    If Trim$(CStr(ws.Cells(1, HM_COL_PROJECT_NAME).Value)) = "" Then
        ws.Cells(1, HM_COL_PROJECT_NAME).Value = "Project Name"
        ws.Cells(1, HM_COL_START_DATE).Value = "Project Start Date"
        ws.Cells(1, HM_COL_HOMOLOGATION_TYPE).Value = "Homologation Type"
        ws.Cells(1, HM_COL_HOMOLOGATION_SPEC).Value = "Homologation Specification"
        ws.Cells(1, HM_COL_APPLICATION_NO).Value = "Application Number"
        ws.Cells(1, HM_COL_PO_NO).Value = "PO #"
        ws.Cells(1, HM_COL_INVOICE_NO).Value = "Invoice #"
        ws.Cells(1, HM_COL_CERTIFICATE_NO).Value = "Certificate #"
        ws.Cells(1, HM_COL_CLOSE_DATE).Value = "Close Date"
        ws.Cells(1, HM_COL_COMMENT).Value = "Comment"
        ws.Cells(1, HM_COL_STATUS).Value = "Status"
        ws.Cells(1, HM_COL_LAST_UPDATED).Value = "Last Updated"
        ws.Rows(1).Font.Bold = True
    End If

    Set HM_EnsureDataSheet = ws
End Function

Private Function HM_EnsureListSheet() As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim seedTypes As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HM_LIST_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HM_LIST_SHEET
    End If

    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Cells(1, 1).Value = "Homologation Types"
        ws.Cells(1, 2).Value = "Homologation Specifications"
        ws.Rows(1).Font.Bold = True
    End If

    seedTypes = Array("ECE", "VSCC", "ICAT", "CCC")
    For i = LBound(seedTypes) To UBound(seedTypes)
        HM_AppendListValue HM_LIST_SHEET, 1, CStr(seedTypes(i))
    Next i

    Set HM_EnsureListSheet = ws
End Function

Private Sub HM_AppendListValue(ByVal sheetName As String, ByVal colNum As Long, ByVal valueText As String)
    Dim ws As Worksheet
    Dim v As String
    Dim lastRow As Long
    Dim r As Long

    v = Trim$(valueText)
    If v = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, colNum).Value)), v, vbTextCompare) = 0 Then Exit Sub
    Next r

    ws.Cells(lastRow + 1, colNum).Value = v
End Sub

Private Sub HM_CreateStatusReport(ByVal statusFilter As String)
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim srcRow As Long
    Dim dstRow As Long
    Dim sheetName As String
    Dim normalized As String

    normalized = Trim$(statusFilter)
    If normalized = "" Then Exit Sub

    Set wsData = HM_EnsureDataSheet()
    sheetName = "Report_" & normalized & "_" & Format$(Now, "yyyymmdd_hhnnss")

    Set wsReport = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsReport.Name = Left$(sheetName, 31)

    wsData.Rows(1).Copy wsReport.Rows(1)
    wsReport.Rows(1).Font.Bold = True

    lastRow = wsData.Cells(wsData.Rows.Count, HM_COL_PROJECT_NAME).End(xlUp).Row
    dstRow = 2

    For srcRow = 2 To lastRow
        If StrComp(Trim$(CStr(wsData.Cells(srcRow, HM_COL_STATUS).Value)), normalized, vbTextCompare) = 0 Then
            wsData.Rows(srcRow).Copy wsReport.Rows(dstRow)
            dstRow = dstRow + 1
        End If
    Next srcRow

    wsReport.Columns("A:L").AutoFit

    MsgBox normalized & " project report created on sheet '" & wsReport.Name & "'.", vbInformation
End Sub
