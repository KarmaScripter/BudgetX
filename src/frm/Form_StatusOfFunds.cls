VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StatusOfFunds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Args As AccountArgs
Private FiscalYearFilter As String
Private FundNameFilter As String
Private ProgramProjectNameFilter As String
Private BocNameFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Open
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Me.RecordSource = "SELECT *" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND StatusOfFunds.BocCode NOT IN ('10', '17', '21', '28');"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    pAnd = " AND "
    SetButtonVisibility
    SetComboBoxColors
    ResetOptionGroupValues
    ClearComboBoxes
    Set Args = New AccountArgs
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.FundName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName" _
        & " FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetButtonVisibility()
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
    Me.CompassStatusOfFundsButton.Visible = Not Me.CompassStatusOfFundsButton.Visible
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        SetComboBoxColors
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.ProgramProjectNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.ProgramProjectNameComboBox.BackColor = RGB(33, 33, 33)
    Me.ProgramProjectNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.BocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.BocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.BocNameComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
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
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearFilters
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearFilters()
    On Error GoTo ErrorHandler:
    BocNameFilter = vbNullString
    FiscalYearFilter = vbNullString
    FundNameFilter = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ClearComboBoxes
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxes()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value = vbNullString
    Me.ProgramProjectNameComboBox.Value = vbNullString
    Me.BocNameComboBox.Value = vbNullString
    ClearFilters
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ResetOptionGroupValues
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ResetOptionGroupValues()
    On Error GoTo ErrorHandler:
    Me.FiscalYearOptionGroup.Value = 2
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
Private Function GetFiscalYearFilter()
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function

'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FundNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FundNameFilter = vbNullString
    Me.Filter = vbNullString
    Args.FundName = Me.FundNameComboBox.Value
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = GetFundNameFilter
    Me.ProgramProjectNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.ProgramProjectName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
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
    If Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = FiscalYearFilter & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = FiscalYearFilter & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            Me.Filter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetFundNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProgramProjectNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.ProgramProjectName = vbNullString
    ProgramProjectNameFilter = vbNullString
    Me.Filter = vbNullString
    Args.ProgramProjectName = Me.ProgramProjectNameComboBox.Value
    ProgramProjectNameFilter = "[ProgramProjectName] = '" & Me.ProgramProjectNameComboBox.Value & "'"
    Me.Filter = GetProgramProjectNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.Requery
    Me.BocNameComboBox.RowSource = "SELECT DISTINCT StatusOfFunds.BocName FROM StatusOfFunds WHERE" _
        & " StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetProgramProjectNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetProgramProjectNameFilter()
    On Error GoTo ErrorHandler:
    If Not ProgramProjectNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not ProgramProjectNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        BocNameFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetProgramProjectNameFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        BocNameComboBox_Change
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Args.BocName = vbNullString
    BocNameFilter = vbNullString
    Me.Filter = vbNullString
    Args.BocName = Me.BocNameComboBox.Value
    BocNameFilter = "[BocName] = '" & Me.BocNameComboBox.Value & "'"
    Me.Filter = GetBocNameFilter
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7' AND " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Function
'   Name:        GetProgramProjectNameFilter
'   Parameters:  Void
'   Retval:      String
'   Purpose:
'---------------------------------------------------------------------------------
Private Function GetBocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not BocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = FiscalYearFilter _
                & pAnd & BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & ProgramProjectNameFilter
    End If
    If Not BocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        ProgramProjectNameFilter & "" = "" Then
            GetBocNameFilter = BocNameFilter _
                & pAnd & FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxes
    ResetOptionGroupValues
    ClearFilters
    Me.RecordSource = "SELECT * FROM StatusOfFunds" _
        & " WHERE StatusOfFunds.BudgetLevel = '7'" _
        & " AND BocCode NOT IN ('10', '17', '21', '28');"
    Me.Requery
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
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
'   Type:        Event Sub-Procedure
'   Name:        CalculatorButton_Click
'   Parameters:  Void
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
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'----------------------------------------------------------------------------------
Private Sub CompassStatusOfFundsButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "CompassBudgetStatusQuery"
ErrorHandler:
    ProcessError
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetRpioIcon
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




