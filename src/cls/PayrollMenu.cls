Option Compare Database
Option Explicit


Private m_Error As String





'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Form_Load()
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.Caption = "Payroll Menu"
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="Form_Load()"
    Exit Sub
End Sub







'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollAuthorityButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollAuthorityForm", _
        WindowMode:=acDialog, _
        OpenArgs:="[BocCode] ='10' AND [BudgetLevel] = '7'"
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="PayrollAuthorityButton_Click()"
    Exit Sub
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FteButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="FullTimeAllocations", WindowMode:=acDialog
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="FteButton_Click()"
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollAccrualButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollAccrual", WindowMode:=acDialog
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="PayrollAccrualButton_Click()"
    Exit Sub
End Sub







'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollUtilizationButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollUtilization", WindowMode:=acDialog
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="PayrollUtilizationButton_Click()"
    Exit Sub
End Sub






'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PayrollObligationButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PayrollObligations", WindowMode:=acDialog
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="PayrollObligationButton_Click()"
    Exit Sub
End Sub







'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub TimeAttendanceButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, "Payroll", acSaveNo
    DoCmd.OpenForm FormName:="PeoplePlusHours", WindowMode:=acDialog
ErrorHandler:
    ProcessError Name:="PayrollMenu", _
        Member:="TimeAttendanceButton_Click()"
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




