VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocumentObligations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public Args As AccountArgs
Private m_Dialog As Form_StatusOfFundsDialog
Private FundNameFilter As String
Private RpioNameFilter As String
Private DocumentNumberFilter As String
Private FiscalYearFilter As String
Private BocNameFilter As String
Private m_FormFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set m_Dialog = New Form_StatusOfFundsDialog
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="StatusOfFundsDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("StatusOfFundsDialog")
    Set Args = m_Dialog.Args
    RpioNameFilter = "[RpioNameFilter] = '" & Args.RpioName & "'"
    Me.RecordSource = "SELECT * FROM Obligations WHERE NOT" _
        & " IsNull(Obligations.DocumentControlNumber)" _
        & " AND Obligations.BocCode NOT IN ('10', '17')" _
        & " AND Obligations.RcCode = '" & Args.RcCode & "';"
    SetRpioIcon (Args.RpioCode)
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Event / Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 0.5
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
    pAnd = " AND "
    Me.Filter = RpioNameFilter & pAnd & FiscalYearFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
    Me.DocumentNumberComboBox.RowSource = "SELECT DISTINCT Obligations.DocumentControlNumbers FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetButtonVisibility
'   Parameters:      Void
'   Retval:          Void
'   Purpose:         Toggles button visibility on and off
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = Not Me.AddButton.Visible
    Me.FirstButton.Visible = Not Me.FirstButton.Visible
    Me.PreviousButton.Visible = Not Me.PreviousButton.Visible
    Me.NextButton.Visible = Not Me.NextButton.Visible
    Me.LastButton.Visible = Not Me.LastButton.Visible
    Me.EditButton.Visible = Not Me.EditButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
    Me.DeleteButton.Visible = Not Me.DeleteButton.Visible
    Me.CalculatorButton.Visible = Not Me.CalculatorButton.Visible
    Me.ExcelButton.Visible = Not Me.ExcelButton.Visible
    Me.UndoButton.Visible = Not Me.UndoButton.Visible
    Me.SaveButton.Visible = Not Me.SaveButton.Visible
    Me.DataButton.Visible = Not Me.DataButton.Visible
    Me.CompassDocumentQueryButton.Visible = Not Me.CompassDocumentQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Hides buttons
'   Parameters:  Void
'   Purpose:     Toggles the toolbar button on/off
'---------------------------------------------------------------------------------
Private Sub HideButtons()
    On Error GoTo ErrorHandler:
    Me.AddButton.Visible = False
    Me.DataButton.Visible = False
    Me.FirstButton.Visible = False
    Me.PreviousButton.Visible = False
    Me.NextButton.Visible = False
    Me.LastButton.Visible = False
    Me.EditButton.Visible = False
    Me.RefreshButton.Visible = False
    Me.DeleteButton.Visible = False
    Me.CalculatorButton.Visible = False
    Me.ExcelButton.Visible = False
    Me.UndoButton.Visible = False
    Me.SaveButton.Visible = False
    Me.CompassDocumentQueryButton.Visible = False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            SetComboBoxColors
'   Parameters:      Void
'   Purpose:         Toggles button visibility ON and OFF
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.DocumentNumberComboBox.ForeColor = RGB(255, 255, 255)
    Me.DocumentNumberComboBox.BackColor = RGB(33, 33, 33)
    Me.DocumentNumberComboBox.BorderColor = RGB(68, 114, 196)
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.DocumentNumberComboBox.Value = vbNullString
    Me.BocNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearOptionGroupValues()
    On Error GoTo ErrorHandler:
    FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    DocumentNumberFilter = vbNullString
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    BocNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
    Me.Filter = RpioNameFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Obligations" _
        & " WHERE Obligations.BocCode IN ('36', '37', '38', '41')" _
        & pAnd & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName" _
        & " FROM Obligations" _
        & " WHERE Obligations.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.DocumentNumberComboBox.RowSource = "SELECT DISTINCT Obligations.DocumentControlNumbers" _
        & " FROM Obligations" _
        & " WHERE Obligations.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations" _
        & " WHERE Obligations.BocCode IN ('36', '37', '38', '41')" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    Args.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = RpioNameFilter & pAnd & GetFiscalYearFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
    Me.DocumentNumberComboBox.RowSource = "SELECT DISTINCT Obligations.DocumentControlNumbers" _
        & " FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName FROM Obligations" _
        & " WHERE NOT IsNull(Obligations.DocumentControlNumbers)" _
        & pAnd & "Obligations.BocCode NOT IN ('10', '17')" _
        & pAnd & Me.Filter & ";"
    Me.RecordSource = "SELECT * FROM Obligations WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" & pAnd & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFiscalYearFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & DocumentNumberFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DocumentNumberFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    BocNameFilter = vbNullString
    DocumentNumberFilter = vbNullString
    Args.FundCode = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = RpioNameFilter & pAnd & GetFundNameFilter
    Me.DocumentNumberComboBox.RowSource = "SELECT DISTINCT Obligations.DocumentControlNumbers" _
        & " FROM Obligations WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.RecordSource = "SELECT * FROM Obligations" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetFundNameFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FundNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetFundNameFilter = DocumentNumberFilter & pAnd _
                & FundNameFilter & pAnd _
                & FiscalYearFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        FiscalYearFilter & "" = "" Then
            GetFundNameFilter = DocumentNumberFilter & pAnd _
                & FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        FiscalYearFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function

'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DocumentNumberComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DocumentNumberFilter = vbNullString
    BocNameFilter = vbNullString
    DocumentNumberFilter = "[DocumentControlNumbers] = '" & Me.DocumentNumberComboBox.Value & "'"
    Me.Filter = RpioNameFilter & pAnd & GetDocumentNumberFilter
    Me.RecordSource = "SELECT * FROM Obligations" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations WHERE" _
        & " NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetDocumentNumberFilter
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDocumentNumberFilter() As String
    On Error GoTo ErrorHandler:
    If Not DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & FundNameFilter
    End If
    If Not DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetDocumentNumberFilter = DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
         FiscalYearFilter & "" = "" Then
            GetDocumentNumberFilter = DocumentNumberFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:            Event / Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    BocNameFilter = vbNullString
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = RpioNameFilter & pAnd & GetBocNameFilter
    Me.RecordSource = "SELECT * FROM Obligations" _
        & " WHERE NOT IsNull(DocumentControlNumbers)" _
        & pAnd & "[BocCode] NOT IN ('10', '17')" _
        & pAnd & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Function
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If BocNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetBocNameFilter = DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & DocumentNumberFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        DocumentNumberFilter & "" = "" And _
        FundNameFilter & "" = "" And _
         FiscalYearFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    RpioNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassDocumentQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery", , , , , acDialog
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub OriginalActionDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetColumnNames
'   Parameters:  Void
'   Retval:      Collection
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetColumnNames() As String()
    On Error GoTo ErrorHandler:
    Dim mFields As Collection
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim mArray() As String
    j = mData.Fields.count - 1
    ReDim mArray(j)
    For i = LBound(mArray()) To UBound(mArray())
        If Not mData.Fields(i).Name & "" = "" Then
            mArray(i) = mData.Fields(i).Name
        End If
    Next i
    GetColumnNames = mArray()
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetReportData
'   Parameters:  Void
'   Retval:      DAO Recordset
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetReportData() As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim mData As DAO.Recordset
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    Set GetReportData = mData
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExcelButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExcelButton_Click()
    On Error GoTo ErrorHandler:
    Dim mFields() As String
    mFields() = GetColumnNames
    Dim mBudgetPath As BudgetPath
    Dim mExcel As Excel.Application
    Dim mAllocations As Excel.Workbook
    Dim mWorksheet As Excel.Worksheet
    Dim mList As Excel.ListObject
    Dim mRange As Excel.Range
    Dim mCell As Object
    Dim mHeader As Excel.Range
    Dim mStart As Excel.Range
    Dim mEnd As Excel.Range
    Dim field As DAO.field
    Dim mData As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set mBudgetPath = New BudgetPath
    Set mExcel = CreateObject("Excel.Application")
    Set mAllocations = mExcel.Workbooks.Open(mBudgetPath.ReportTemplate)
    mAllocations.Worksheets(2).Visible = False
    Set mWorksheet = mAllocations.Worksheets(1)
    mWorksheet.Name = "Obligations By DCN"
    mWorksheet.Cells.HorizontalAlignment = xlHAlignLeft
    mWorksheet.Cells.Font.Name = "Source Code Pro"
    mWorksheet.Cells.Font.Size = 8
    Set mStart = mWorksheet.Cells(1, 1)
    Set mEnd = mWorksheet.Cells(1, UBound(mFields) - 1)
    Set mHeader = mWorksheet.Range(mStart, mEnd)
    mHeader.HorizontalAlignment = xlHAlignLeft
    mHeader.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    mHeader.Font.Name = "Source Code Pro"
    mHeader.Font.Color = vbBlack
    mHeader.Font.Bold = True
    mHeader.Font.Size = 8
    For i = LBound(mFields) To UBound(mFields)
        mHeader.Cells(i + 1).Value = mFields(i)
    Next i
    Set mData = Me.Recordset
    mData.Filter = Me.Filter
    mWorksheet.Cells(2, 1).CopyFromRecordset mData
    mHeader.Font.Color = vbBlack
    mExcel.WindowState = xlMaximized
    mExcel.Visible = True
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub UndoButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdUndo
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SaveButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SaveButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSave
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DeleteButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DeleteButton_Click()
    On Error GoTo ErrorHandler:
    If Me.Dirty Then
        DoCmd.RunCommand acCmdDelete
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub








'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetNpmIcon
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetRpioIcon(pRpioName As String)
    On Error GoTo ErrorHandler:
    Dim p_RpioCode As String
    If Not pRpioName & "" = "" Then
        p_RpioCode = DLookup("RpioCode", "StatusOfFunds", _
            "[RpioName] = '" & pRpioName & "'")
        Me.Caption = pRpioName
        Select Case CStr(p_RpioCode)
            Case "01"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R1.png"
            Case "02"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R2.png"
            Case "03"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R3.png"
            Case "04"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R4.png"
            Case "05"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R5.png"
            Case "06"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R6.png"
            Case "07"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R7.png"
            Case "08"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R8.png"
            Case "09"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R9.png"
            Case "10"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\R10.png"
            Case "11"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OA.png"
            Case "13"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OITA.png"
            Case "16"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OMS.png"
            Case "17"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OCFO.png"
            Case "20"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OCSPP.png"
            Case "26"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\ORD.png"
            Case "27"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OAR.png"
            Case "30"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OW.png"
            Case "35"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OIG.png"
            Case "39"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OGC.png"
            Case "75"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OLEM.png"
            Case "77"
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\RPIO\OECA.png"
            Case Else
                Me.RpioIcon.Picture = CurrentProject.Path _
                    & "\etc\png\AppIcons\interface\ui\capital.png"
        End Select
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProcessError
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProcessError()
    If Err.Number <> 0 Then
        m_Error = "Source:      " & Err.Source _
            & vbCrLf & "Number:     " & Err.Number _
            & vbCrLf & "Issue:      " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
End Sub





