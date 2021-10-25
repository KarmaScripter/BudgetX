VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ObligationsByFundingDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public RcCodeArgs As String
Public FundCodeArgs As String
Public FiscalYearArgs As String
Private m_SQL As String
Public m_Dialog As Form_StatusOfFundsDialog
Public pAnd As String
Public Args As AccountArgs
Private m_Error As String











'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
ErrorHandler:
    ProcessError Name:="ObligationByFundingDocument", _
        Member:="Report_Load"
    Exit Sub
End Sub







'---------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    Set m_Dialog = New Form_StatusOfFundsDialog
    Set Args = New AccountArgs
    pAnd = " AND "
    DoCmd.OpenForm FormName:="StatusOfFundsDialog", WindowMode:=acDialog
    Set m_Dialog = Forms("StatusOfFundsDialog")
    Set Args = m_Dialog.Args
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Not Args.FundName & "" = "" Then
            m_SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.RpioName = '" & Args.RpioName & "'" _
            & " AND Obligations.FundName = '" & Args.FundName & "';"
            Me.RecordSource = m_SQL
    End If
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Args.FundName & "" = "" Then
            m_SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.RpioName = '" & Args.RpioName & "';"
            Me.RecordSource = m_SQL
    End If
    If Not Args.BFY & "" = "" And _
        Args.RpioName & "" = "" And _
        Args.FundName & "" = "" Then
            m_SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "';"
            Me.RecordSource = m_SQL
    End If
    If Not Args.BFY & "" = "" And _
        Args.RpioName & "" = "" And _
        Not Args.FundName & "" = "" Then
            m_SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.FundName = '" & Args.FundName & "';"
            Me.RecordSource = m_SQL
    End If
    Me.Requery
    DoCmd.Close ObjectType:=acForm, _
        ObjectName:="StatusOfFundsDialog", Save:=acSaveNo
ErrorHandler:
    ProcessError Name:="ObligationByFundingDocument", _
        Member:="Report_Load"
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Function
'   Name:        VerifyArgs
'   Parameters:  AccountArgs
'   RetVal:      Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Function VeryifyArgs(pArg As AccountArgs) As Boolean
    On Error GoTo ErrorHandler
    Dim p_Text As String
    If pArg Is Nothing Then
        p_Text = "AccountArgs Object is NULL or Nothing"
        MessageFactory.Notify (p_Text)
        VeryifyArgs = False
        Exit Function
    End If
    If pArg.BFY & "" = "" Then
        p_Text = "AccountArgs BFY Property is not NULL or Empty"
        MessageFactory.Notify (p_Text)
        VeryifyArgs = False
        Exit Function
    End If
    If pArg.RpioName & "" = "" Then
        p_Text = "AccountArgs Fiscal Year is not NULL or Empty"
        MessageFactory.Notify (p_Text)
        VeryifyArgs = False
        Exit Function
    End If
    If pArg.FundName & "" = "" Then
        p_Text = "AccountArgs FundName Property is NULL or Empty"
        MessageFactory.Notify (p_Text)
        VeryifyArgs = False
        Exit Function
    End If
ErrorHandler:
    ProcessError Name:="ObligationByFundingDocument", _
        Member:="VeryifyArgs"
End Function






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
