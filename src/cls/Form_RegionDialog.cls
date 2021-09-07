VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_RegionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Public Args As AccountArgs
Private FiscalYearFilter As String
Private DivisionNameFilter As String
Private FundNameFilter As String
Private m_FundCode As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    SetComboBoxColors
    pAnd = " AND "
    Set Args = New AccountArgs
    FiscalYearFilter = "[BFY] = '2021'"
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT Divisions.Title FROM Divisions " _
        & "WHERE Divisions.Code NOT IN ('06', '06R')"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT Allocations.BocName FROM Allocations;"
    Me.FiscalYearOptionGroup.DefaultValue = 2
    Me.FiscalYearOptionGroup.Value = 2
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    FiscalYearFilter = vbNullString
    Select Case Me.FiscalYearOptionGroup.Value
        Case 1
            FiscalYearFilter = "[BFY] = '2020'"
            Args.BFY = "2020"
            Me.Filter = GetFiscalYearFilter
        Case 2
            FiscalYearFilter = "[BFY] = '2021'"
            Args.BFY = "2021"
            Me.Filter = GetFiscalYearFilter
    End Select
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
    If Not FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        DivisionNameFilter = "" And _
        Not FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter = "" And _
        Not DivisionNameFilter = "" And _
        FundNameFilter = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & DivisionNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DivisionNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter = vbNullString
    DivisionNameFilter = vbNullString
    FundNameFilter = vbNullString
    Args.DivisionName = Me.DivisionNameComboBox.Value
    Args.RcCode = DLookup("Code", "Divisions", "Divisions.Title = '" _
        & Me.DivisionNameComboBox.Value & "'")
    DivisionNameFilter = "[DivisionName] = '" & Me.DivisionNameComboBox.Value & "'"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Allocations.FundName" _
        & " FROM Allocations" _
        & " WHERE " & DivisionNameFilter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetDivisionNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetDivisionNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not DivisionNameFilter = "" And _
        Not FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DivisionNameFilter = "" And _
        FundNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not DivisionNameFilter = "" And _
        Not FundNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetDivisionNameFilter = DivisionNameFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    m_FundCode = DLookup("FundCode", "Allocations", _
        "Allocations.FundName = '" & Me.FundNameComboBox.Value & "'")
    Me.Filter = vbNullString
    FundNameFilter = vbNullString
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Args.FundName = Me.FundNameComboBox.Value
    Args.FundCode = m_FundCode
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
Private Function GetFundNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FundNameFilter = "" And _
        Not DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & DivisionNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = "" And _
        DivisionNameFilter = "" And _
        Not FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter = "" And _
        Not DivisionNameFilter = "" And _
        FiscalYearFilter = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & DivisionNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExecuteButton_OnClick
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler:
    Me.Visible = Not Me.Visible
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ClearOptionGroupValues
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!DivisionNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!DivisionNameComboBox.BackColor = RGB(33, 33, 33)
    Me!DivisionNameComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me!FundNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearOptionGroupValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearOptionGroupValues()
    If Not IsNull(Me.FiscalYearOptionGroup) And _
        Not Me.FiscalYearOptionGroup.Value & "" = vbNullString Then
            Me.FiscalYearOptionGroup.Value = vbNullString
            FiscalYearFilter = vbNullString
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    Me.DivisionNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilterValues
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    FiscalYearFilter = vbNullString
    DivisionNameFilter = vbNullString
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    Set Args = Nothing
    DoCmd.Close acForm, Me.Name
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





