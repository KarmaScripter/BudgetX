VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ProgrammaticObligations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Public m_Dialog As Form_StatusOfFundsDialog
Private pFilter As String
Private FiscalYearFilter As String
Private FundNameFilter As String
Private BocNameFilter As String
Private RpioNameFilter As String
Private RpioCodeFilter As String
Private pAnd As String
Private m_SQL As String
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
    If Args Is Nothing Then
        Set Args = New AccountArgs
    End If
    DoCmd.OpenForm FormName:="StatusOfFundsDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("StatusOfFundsDialog")
    Set Args = m_Dialog.Args
    Args.RpioCode = m_Dialog.Args.RpioCode
    Args.RpioName = m_Dialog.Args.RpioName
    Args.BFY = m_Dialog.Args.BFY
    Args.FundName = m_Dialog.Args.FundName
    Args.FundCode = m_Dialog.Args.FundCode
    FiscalYearFilter = "[BFY] = '" & Args.BFY & "'"
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Not Args.FundName & "" = "" Then
            m_SQL = "SELECT * FROM Obligations " _
                & " WHERE NOT IsNull(Obligations.DocumentControlNumber)" _
                & " AND Obligations.BFY = '" & Args.BFY & "'" _
                & " AND Obligations.RpioName = '" & Args.RpioName & "'" _
                & " AND Obligations.FundName = '" & Args.FundName & "';"
    End If
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Args.FundName & "" = "" Then
            m_SQL = "SELECT * FROM Obligations " _
                & " WHERE NOT IsNull(Obligations.DocumentControlNumber)" _
                & " AND Obligations.BFY = '" & Args.BFY & "'" _
                & " AND Obligations.RpioName = '" & Args.RpioName & "';"
    End If
    If Not Args.BFY & "" = "" And _
        Args.RpioName & "" = "" And _
        Not Args.FundName & "" = "" Then
            m_SQL = "SELECT * FROM Obligations " _
                & " WHERE NOT IsNull(Obligations.DocumentControlNumber)" _
                & " AND Obligations.BFY = '" & Args.BFY & "'" _
                & " AND Obligations.FundName = '" & Args.FundName & "';"
    End If
    Me.RecordSource = m_SQL
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    Me.Section(acHeader).AutoHeight = False
    Me.Section(acHeader).Height = 2
    Me.Section(acDetail).AutoHeight = False
    Me.Section(acDetail).Height = 4
    Me.Section(acFooter).AutoHeight = False
    Me.Section(acFooter).Height = 0.5
    SetComboBoxColors
    SetButtonVisibility
    pAnd = " AND "
    RpioNameFilter = "[RpioName] = '" & Args.RpioName & "'"
    RpioCodeFilter = "[RpioCode] = '" & Args.RpioCode & "'"
    Me.Filter = FiscalYearFilter & pAnd & RpioNameFilter
    Me.RpioNameComboBox.RowSource = "SELECT DISTINCT Obligations.RpioName" _
        & " FROM Obligations" _
        & " WHERE" & FiscalYearFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName" _
        & " FROM Obligations" _
        & " WHERE" & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations" _
        & " WHERE" & Me.Filter
    SetRpioIcon (Args.RpioName)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Timer()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Timer()
    On Error GoTo ErrorHandler:
    HideButtons
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
    Me.CompassSpendingQueryButton.Visible = False
    Me.SaveButton.Visible = False
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxColors
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.RpioNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.RpioNameComboBox.BackColor = RGB(33, 33, 33)
    Me.RpioNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.BocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.BocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.BocNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetButtonVisibility
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
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
    Me.CompassSpendingQueryButton.Visible = Not Me.CompassSpendingQueryButton.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Close
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Args.BFY = vbNullString
    FiscalYearFilter = vbNullString
    Me.Filter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            Args.BFY = "2020"
            FiscalYearFilter = "[BFY] = '2020'"
        Case 2
            Args.BFY = "2021"
            FiscalYearFilter = "[BFY] = '2021'"
    End Select
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Obligations WHERE " & Me.Filter
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
     If Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & RpioNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RpioNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RpioNameComboBox_Change()
    On Error GoTo ErrorHandler:
    RpioNameFilter = vbNullString
    Me.Filter = vbNullString
    Args.RpioName = Me.RpioNameComboBox.Value
    RpioNameFilter = "[RpioName] = '" & Me.RpioNameComboBox.Value & "'"
    Me.Filter = GetRpioNameFilter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName" _
        & " FROM Obligations" _
        & " WHERE " & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM Obligations WHERE " & Me.Filter
    Me.Requery
    SetRpioIcon (Args.RpioName)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetRpioNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetRpioNameFilter()
    On Error GoTo ErrorHandler:
    If Not RpioNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = FiscalYearFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = BocNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not RpioNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString Then
            GetRpioNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    BocNameFilter = vbNullString
    Args.BocCode = Me.BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * FROM Obligations WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetBocNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not FundNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        FundNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetBocNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & RpioNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    Args.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = RpioNameFilter _
        & pAnd & FiscalYearFilter _
        & pAnd & FundNameFilter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT BocName" _
        & " FROM Obligations" _
        & " WHERE " & Me.Filter
    Me.RecordSource = "SELECT * FROM Obligations WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetFundNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetFundNameFilter()
    On Error GoTo ErrorHandler:
    If Not FundNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            Me.Filter = FundNameFilter _
                & pAnd & RpioNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        FiscalYearFilter & "" = vbNullString And _
        Not BocNameFilter & "" = vbNullString And _
        RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = vbNullString And _
        Not FiscalYearFilter & "" = vbNullString And _
        BocNameFilter & "" = vbNullString And _
        Not RpioNameFilter & "" = vbNullString Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & RpioNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
'   Retval:      Void
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
'   Type:        Sub-Procedure
'   Name:        CompassSpendingQueryButton_Click()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'----------------------------------------------------------------------------------
Private Sub CompassSpendingQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        MenuButton_Click
'   Parameters:  Void
'   Retval:      Void
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
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
    Me.FiscalYearOptionGroup.Value = 2
    FiscalYearFilter = "[BFY] = '2021'"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    ClearComboBoxValues
    ClearFilterValues
    ResetOptionGroupValues
    Me.Filter = GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Obligations" _
        & " WHERE Not IsNull(Obligations.DocumentControlNumber)" _
        & pAnd & Me.Filter
    Me.RpioNameComboBox.RowSource = "SELECT DISTINCT Obligations.RpioName" _
        & " FROM Obligations WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Obligations.FundName" _
        & " FROM Obligations WHERE " & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Obligations.BocName" _
        & " FROM Obligations WHERE " & Me.Filter
    Me.Requery
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
    Dim m_Fields As Collection
    Dim field As DAO.field
    Dim m_Data As DAO.Recordset
    Set m_Data = Me.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim m_Array() As String
    j = m_Data.Fields.count - 1
    ReDim m_Array(j)
    For i = LBound(m_Array()) To UBound(m_Array())
        If Not m_Data.Fields(i).Name & "" = "" Then
            m_Array(i) = m_Data.Fields(i).Name
        End If
    Next i
    GetColumnNames = m_Array()
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
    Dim m_Data As DAO.Recordset
    Set m_Data = Me.Recordset
    m_Data.Filter = Me.Filter
    Set GetReportData = m_Data
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
    Dim p_Fields() As String
    p_Fields() = GetColumnNames
    Dim p_BudgetPath As BudgetPath
    Dim p_Excel As Excel.Application
    Dim p_Allocations As Excel.Workbook
    Dim p_Worksheet As Excel.Worksheet
    Dim p_List As Excel.ListObject
    Dim p_Range As Excel.Range
    Dim p_Cell As Object
    Dim p_Header As Excel.Range
    Dim p_Start As Excel.Range
    Dim p_End As Excel.Range
    Dim field As DAO.field
    Dim p_Data As DAO.Recordset
    Dim i As Integer
    Dim j As Integer
    Set p_BudgetPath = New BudgetPath
    Set p_Excel = CreateObject("Excel.Application")
    Set p_Allocations = p_Excel.Workbooks.Open(p_BudgetPath.ReportTemplate)
    p_Allocations.Worksheets(2).Visible = False
    Set p_Worksheet = p_Allocations.Worksheets(1)
    p_Worksheet.Name = "Programmatic Obligations"
    p_Worksheet.Cells.HorizontalAlignment = xlHAlignLeft
    p_Worksheet.Cells.Font.Name = "Source Code Pro"
    p_Worksheet.Cells.Font.Size = 8
    Set p_Start = p_Worksheet.Cells(1, 1)
    Set p_End = p_Worksheet.Cells(1, UBound(p_Fields) - 1)
    Set p_Header = p_Worksheet.Range(p_Start, p_End)
    p_Header.HorizontalAlignment = xlHAlignLeft
    p_Header.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    p_Header.Font.Name = "Source Code Pro"
    p_Header.Font.Color = vbBlack
    p_Header.Font.Bold = True
    p_Header.Font.Size = 8
    For i = LBound(p_Fields) To UBound(p_Fields)
        p_Header.Cells(i + 1).Value = p_Fields(i)
    Next i
    Set p_Data = Me.Recordset
    p_Data.Filter = Me.Filter
    p_Worksheet.Cells(2, 1).CopyFromRecordset p_Data
    p_Header.Font.Color = vbBlack
    p_Excel.WindowState = xlMaximized
    p_Excel.Visible = True
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
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.BocNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
    Me.RpioNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Retval:      Void
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    BocNameFilter = vbNullString
    RpioNameFilter = vbNullString
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





