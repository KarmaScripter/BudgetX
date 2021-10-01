VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ProgramProjectDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Args As AccountArgs
Private pAnd As String
Private m_DataSet As DAO.Recordset
Private m_ProgramProjectCode As String
Private m_Error As String









'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Me.RecordSource = "SELECT * FROM ProgramDescriptions;"
    Me.NextButton.SetFocus
ErrorHandler:
    ProcessError
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub NextButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.FindNext
ErrorHandler:
    ProcessError
    Exit Sub
End Sub










'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub PreviousButton_Click()
    On Error GoTo ErrorHandler:
ErrorHandler:
    ProcessError
    Exit Sub
End Sub








'---------------------------------------------------------------------------------
'   Type:        Property
'   Name:
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, _
        Save:=acSaveNo
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

