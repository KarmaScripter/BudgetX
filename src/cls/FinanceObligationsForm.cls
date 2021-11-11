Option Compare Database
Option Explicit



Public Args As AccountArgs
Private m_Dialog As Form_DivisionSelector
Private FundNameFilter As String
Private AccountCodeFilter As String
Private BocNameFilter As String
Private FocNameFilter As String
Private FiscalYearFilter As String
Private RcCodeFilter As String
Private pAnd As String
Private m_Error As String
Private m_Notification As String




'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler::
    Set m_Dialog = New Form_DivisionSelector
    Set Args = New AccountArgs
    DoCmd.OpenForm FormName:="DivisionSelector", WindowMode:=acDialog
    Set m_Dialog = Forms("DivisionSelector")
    Set Args = m_Dialog.Args
    RcCodeFilter = "[RcCode] = '" & Args.RcCode & "'"
    Me.RecordSource = "SELECT * FROM Purchases" _
        & " WHERE Purchases.BocCode NOT IN ('10', '17')" _
        & " AND Purchases.RcCode = '" & Args.RcCode & "';"
    SetDivisionIcon
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:=""
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Load
'   Parameters  Void
'   Retval      Void
'   Purpose
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
    FiscalYearFilter = "[BFY] = '2022'"
    ClearComboBoxValues
    SetComboBoxColors
    pAnd = " AND "
    ClearComboBoxValues
    SetComboBoxColors
    SetButtonVisibility
    Me.Filter = RcCodeFilter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT Purchases.AccountCode" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT Purchases.FocName" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Purchases.FundName" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
    Me.FilterOn = True
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:=""
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        Form_Close
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Form_Close()
    On Error GoTo ErrorHandler:
    Me.Filter & "" = ""
    RcCodeFilter & "" = ""
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="Form_Close()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Sub-Procedure
'   Name        SetButtonVisibility
'   Parameters  Void
'   Retval      Void
'   Purpose
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="SetButtonVisibility()"
    Exit Sub
End Sub






'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        SetDivisionIcon
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetDivisionIcon()
    On Error GoTo ErrorHandler::
    If Not Args.RcCode & "" = "" Then
        Me.Caption = "Obligations By Finance Object Class"
        Select Case Args.RcCode
            Case "06A"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ORA.png"
            Case "06B"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\LCARD.png"
            Case "06C"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\MSD.png"
            Case "06D"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ORC.png"
            Case "06F"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\EJ.png"
            Case "06G"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WCF.png"
            Case "06H"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\LSASD.png"
            Case "06J"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ARD.png"
            Case "06K"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WD.png"
            Case "06L"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\SEMD.png"
            Case "06M"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\ECAD.png"
            Case "06N"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\WSA.png"
            Case "06R"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\MSR.png"
            Case "06X"
                Me.DivisionIcon.Picture = CurrentProject.Path & "\etc\png\DivisionLogo\XA.png"
        End Select
    Else
        Me.DivisionIcon.Picture = _
            CurrentProject.Path & "\etc\png\AppIcons\interface\ui\Reports.png"
    End If
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="SetDivisionIcon()"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="Form_Timer()"
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
    Me.CompassSpendingQueryButton.Visible = False
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="HideButtons()"
    Exit Sub
End Sub




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        SetComboBoxColors
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FundNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FundNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.FocNameComboBox.ForeColor = RGB(255, 255, 255)
    Me.FocNameComboBox.BackColor = RGB(33, 33, 33)
    Me.FocNameComboBox.BorderColor = RGB(68, 114, 196)
    Me.AccountCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me.AccountCodeComboBox.BackColor = RGB(33, 33, 33)
    Me.AccountCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="SetComboBoxColors()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearComboBoxValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    Me.FundNameComboBox.Value & "" = ""
    Me.FocNameComboBox.Value & "" = ""
    Me.AccountCodeComboBox.Value & "" = ""
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="ClearComboBoxValues()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        ClearFilterValues
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ClearFilterValues()
    On Error GoTo ErrorHandler:
    FundNameFilter & "" = ""
    FocNameFilter & "" = ""
    AccountCodeFilter & "" = ""
    FiscalYearFilter & "" = ""
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="ClearFilterValues()"
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
     FiscalYearFilter = "[BFY] = '2022'"
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="ResetOptionGroupValues()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FocNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FocNameComboBox_Change()
    On Error GoTo ErrorHandler:
    FocNameFilter & "" = ""
    Me.Filter & "" = ""
    Me.AccountCodeComboBox.Value & "" = ""
    FocNameFilter = "[FocName] = '" & Me.FocNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetFocNameFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT Purchases.AccountCode" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="FocNameComboBox_Change()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFocNameFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFocNameFilter() As String
    On Error GoTo ErrorHandler:
    If Not FocNameFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter
    End If
    If Not FocNameFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
            & pAnd & AccountCodeFilter
    End If
    If Not FocNameFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FocNameFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FocNameFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not FocNameFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not FocNameFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FocNameFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetFocNameFilter = FocNameFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetFocNameFilter()"
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        AccountCodeComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter & "" = ""
    AccountCodeFilter & "" = ""
    AccountCodeFilter = "[AccountCode] = '" & Me.AccountCodeComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetAccountCodeFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="AccountCodeComboBox_Change()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetAccountCodeFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetAccountCodeFilter() As String
    On Error GoTo ErrorHandler:
    If Not AccountCodeFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
            & pAnd & FocNameFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FocNameFilter _
                & pAnd & FundNameFilter
    End If
    If Not AccountCodeFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not FundNameFilter & "" = "" Then
            GetAccountCodeFilter = AccountCodeFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & FundNameFilter
    End If
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetAccountCodeFilter()"
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FundNameComboBox_Change
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    Me.Filter & "" = ""
    FundNameFilter & "" = ""
    Me.FocNameComboBox.Value & "" = ""
    Me.AccountCodeComboBox.Value & "" = ""
    FundNameFilter = "[FundName] = '" & Me.FundNameComboBox.Value & "'"
    Me.Filter = RcCodeFilter & pAnd & GetFundNameFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE" & Me.Filter
    Me.Requery
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT Purchases.FocName" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT Purchases.AccountCode" _
        & " FROM Purchases" _
        & " WHERE " & Me.Filter
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="FundNameComboBox_Change()"
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
    If Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
            & pAnd & FocNameFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" And _
        Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FiscalYearFilter
    End If
    If Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" And _
        FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" Then
            GetFundNameFilter = FundNameFilter _
                & pAnd & FocNameFilter _
                & pAnd & AccountCodeFilter
    End If
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetFundNameFilter()"
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        FiscalYearOptionGroup_AfterUpdate
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FiscalYearOptionGroup_AfterUpdate()
    On Error GoTo ErrorHandler:
    Me.Filter & "" = ""
    FiscalYearFilter & "" = ""
    Select Case FiscalYearOptionGroup.Value
        Case 1
            FiscalYearFilter = "[BFY] = '2021'"
        Case 2
            FiscalYearFilter = "[BFY] = '2022'"
    End Select
    Me.Filter = RcCodeFilter & pAnd & GetFiscalYearFilter
    Me.RecordSource = "SELECT * FROM Purchases WHERE " & Me.Filter
    Me.Requery
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="FiscalYearOptionGroup_AfterUpdate()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetFiscalYearFilter
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Function GetFiscalYearFilter() As String
    On Error GoTo ErrorHandler:
    If Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
            & pAnd & FundNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FundNameFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        Not FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FundNameFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        AccountCodeFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & FocNameFilter
    End If
    If Not FiscalYearFilter & "" = "" And _
        Not AccountCodeFilter & "" = "" And _
        FundNameFilter & "" = "" And _
        Not FocNameFilter & "" = "" Then
            GetFiscalYearFilter = FiscalYearFilter _
                & pAnd & AccountCodeFilter _
                & pAnd & FocNameFilter
    End If
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetFiscalYearFilter()"
    Exit Function
End Function



'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        CalculatorButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub CalculatorButton_Click()
    On Error GoTo ErrorHandler:
    Calculator.Calculate
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="CalculatorButton_Click()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub CompassSpendingQueryButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetDetailedSpendingQuery"
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="CompassSpendingQueryButton_Click()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        MenuButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub MenuButton_Click()
    On Error GoTo ErrorHandler:
    SetButtonVisibility
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:=""
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        OriginalActionDateTextBox_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub OriginalActionDateTextBox_Click()
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm "BudgetCalendar", , , , , , Me.OriginalActionDateTextBox
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="OriginalActionDateTextBox_Click()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Event Sub-Procedure
'   Name        RefreshButton_Click
'   Parameters  Void
'   Retval      Void
'   Purpose
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    ClearComboBoxValues
    ClearFilterValues
    ResetOptionGroupValues
    Me.AccountCodeComboBox.RowSource = "SELECT DISTINCT Purchases.AccountCode FROM Purchases"
    Me.FocNameComboBox.RowSource = "SELECT DISTINCT Purchases.FocName FROM Purchases"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT Purchases.FundName FROM Purchases"
    Me.RecordSource = "SELECT * FROM Purchases;"
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="RefreshButton_Click()"
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
    j = m_Data.Fields.Count - 1
    ReDim m_Array(j)
    For i = LBound(m_Array()) To UBound(m_Array())
        If Not m_Data.Fields(i).Name & "" = "" Then
            m_Array(i) = m_Data.Fields(i).Name
        End If
    Next i
    GetColumnNames = m_Array()
ErrorHandler:
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetColumnNames()"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="GetReportData() "
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
    p_Worksheet.Name = "Obligations By FOC"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="ExcelButton_Click()"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="UndoButton_Click()"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="SaveButton_Click()"
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
    ProcessError Name:="FinanceObligationsForm", _ 
        Member:="DeleteButton_Click()"
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Process
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ProcessError(Optional Name As String, Optional Member As String)
    If Err.Number <> 0 And _
        Not IsMissing(Name) And _
        Not IsMissing(Member) Then
            m_Error = "Source:      " & Err.Source _
                & vbCrLf & "Number:     " & Err.Number _
                & vbCrLf & "Issue:      " & Err.Description _
                & vbCrLf & "Class:      " & Name _
                & vbCrLf & "Member:     " & Member
    End If
    If Err.Number <> 0 And _
        IsMissing(Name) And _
        IsMissing(Member) Then
            m_Error = "Source:      " & Err.Source _
                & vbCrLf & "Number:     " & Err.Number _
                & vbCrLf & "Issue:      " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
    Err.Clear
End Sub





