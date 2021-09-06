VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MessageDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS   ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Private mMessage As String
Private mTitle As String
Private mNumber As Integer
Private mDescription As String
Private mErrorArgs As Variant
Private m_Error As String

'---------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------- METHODS  ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------





'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
        Me.MessageBox.Value = ""
        Me.MessageBox.Value = CStr(Me.OpenArgs)
    End If
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let MessageArgs(msg As String)
    If Not msg & "" = "" Then
        mMessage = msg
    End If
End Property







'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get MessageArgs() As String
    On Error GoTo ErrorHandler:
    If Not mMessage & "" = "" Then
        MessageArgs = mMessage
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property







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





