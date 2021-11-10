Option Compare Database
Option Explicit


Public Args As PayrollArgs
Private mStartDate As String
Private mEndDate As String
Private mDivisionName As String
Private mFundName As String
Private pAnd As String
Private mError As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    Set Args = New PayrollArgs
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="Form_Open"
    Exit Sub
End Sub


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
    Me.DivisionNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.DivisionName" _
        & " FROM PayrollObligations;"
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations;"
    Me.StartDateTextBox.Value = #10/1/2020#
    Me.EndDateTextBox.Value = Date
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="Form_Load()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        DivisionNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub DivisionNameComboBox_Change()
    On Error GoTo ErrorHandler:
    mDivisionName = vbNullString
    Args.DivisionName = Me.DivisionNameComboBox.Value
    Args.RcCode = DLookup("RcCode", "PayrollObligations", _
        "PayrollObligations.DivisionName = '" & Me.DivisionNameComboBox.Value & "'")
    Me.FundNameComboBox.RowSource = "SELECT DISTINCT PayrollObligations.FundName" _
        & " FROM PayrollObligations WHERE" _
        & " PayrollObligations.DivisionName = '" & Me.DivisionNameComboBox.Value & "'"
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="DivisionNameComboBox_Change()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProgramProjectNameComboBox_Change
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundNameComboBox_Change()
    On Error GoTo ErrorHandler:
    mFundName = vbNullString
    mFundName = Me.FundNameComboBox.Value
    Args.FundName = mFundName
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="FundNameComboBox_Change()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    mStartDate = vbNullString
    mStartDate = Me.StartDateTextBox.text
    Args.StartDate = mStartDate
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="StartDateTextBox_AfterUpdate()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        FiscalYearOptionGroup_AfterUpdate
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Change()
    On Error GoTo ErrorHandler:
    mEndDate = vbNullString
    mEndDate = Me.EndDateTextBox.text
    Args.EndDate = mEndDate
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="EndDateTextBox_Change()"
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AcceptButton_Click()
    On Error GoTo ErrorHandler
    Me.Visible = Not Me.Visible
ErrorHandler:   
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="AcceptButton_Click()"
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
    ClearTextBoxValues
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="RefreshButton_Click()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
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
    Me!StartDateTextBox.ForeColor = RGB(255, 255, 255)
    Me!StartDateTextBox.BackColor = RGB(33, 33, 33)
    Me!StartDateTextBox.BorderColor = RGB(68, 114, 196)
    Me!EndDateTextBox.ForeColor = RGB(255, 255, 255)
    Me!EndDateTextBox.BackColor = RGB(33, 33, 33)
    Me!EndDateTextBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="SetComboBoxColors()"
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
    On Error GoTo ErrorHandler:
    Me.DivisionNameComboBox.Value = vbNullString
    Me.FundNameComboBox.Value = vbNullString
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="ClearComboBoxValues()"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        RefreshButton_Click
'   Parameters:  Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearTextBoxValues()
    On Error GoTo ErrorHandler:
    Me.StartDateTextBox.Value = vbNullString
    Me.EndDateTextBox.Value = vbNullString
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="ClearTextBoxValues()"
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
    On Error GoTo ErrorHandler:
    mEndDate = vbNullString
    mStartDate = vbNullString
    mDivisionName = vbNullString
    mFundName = vbNullString
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="ClearFilterValues()"
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
    On Error GoTo ErrorHandler:
    Set Args = Nothing
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="ExitButton_Click()"
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub StartDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    mStartDate = vbNullString
    Dim calendar As New Form_BudgetCalendar
    DoCmd.OpenForm FormName:="BudgetCalendar", WindowMode:=acDialog, OpenArgs:=Me.StartDateTextBox
    Set calendar = Forms("BudgetCalendar")
    Me.StartDateTextBox.Value = calendar.txtCalendarDate
    mStartDate = Me.StartDateTextBox.Value
    Args.StartDate = mStartDate
    DoCmd.Close ObjectType:=acForm, ObjectName:="BudgetCalendar", Save:=acSaveNo
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="StartDateTextBox_Enter()"
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        ExitButton_Click
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub EndDateTextBox_Enter()
    On Error GoTo ErrorHandler:
    Dim calendar As New Form_BudgetCalendar
    DoCmd.OpenForm FormName:="BudgetCalendar", WindowMode:=acDialog, OpenArgs:=Me.EndDateTextBox
    Set calendar = Forms("BudgetCalendar")
    Me.EndDateTextBox.Value = calendar.txtCalendarDate
    mEndDate = Me.EndDateTextBox.Value
    Args.EndDate = mEndDate
    DoCmd.Close ObjectType:=acForm, ObjectName:="BudgetCalendar", Save:=acSaveNo
ErrorHandler:
    ProcessError Name:="PayPeriodDialog", _ 
        Member:="EndDateTextBox_Enter()"
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
