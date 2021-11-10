VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_OutlayActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Public Args As AccountArgs
Public m_Dialog As Form_StatusOfFundsDialog
Public m_FundName As String
Public m_RpioName As String
Public m_BFY As String
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Load()
    On Error GoTo ErrorHandler:
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Event Sub-Procedure
'   Name:        Report_Open
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Report_Open(Cancel As Integer)
    On Error GoTo ErrorHandler:
    DoCmd.OpenForm FormName:="StatusOfFundsDialog", WindowMode:=acDialog
    Set m_Dialog = New Form_StatusOfFundsDialog
    Set m_Dialog = Forms("StatusOfFundsDialog")
    Set Args = New AccountArgs
    Set Args = m_Dialog.Args
    Dim SQL As String
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Not Args.FundName & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.RpioName = '" & Args.RpioName & "'" _
            & " AND Obligations.FundName = '" & Args.FundName & "';"
            Me.RecordSource = SQL
            Me.Requery
    End If
    If Not Args.BFY & "" = "" And _
        Not Args.RpioName & "" = "" And _
        Args.FundName & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.RpioName = '" & Args.RpioName & "';"
            Me.RecordSource = SQL
            Me.Requery
    End If
    If Not Args.BFY & "" = "" And _
        Args.RpioName & "" = "" And _
        Args.FundName & "" = "" Then
            SQL = "SELECT *" _
            & " FROM Obligations" _
            & " WHERE Obligations.BFY = '" & Args.BFY & "'" _
            & " AND Obligations.FundName = '" & Args.FundName & "';"
            Me.RecordSource = SQL
            Me.Requery
    End If
    DoCmd.Close ObjectType:=acForm, ObjectName:=m_Dialog.Name, Save:=acSaveNo
ErrorHandler:
    ProcessError
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


