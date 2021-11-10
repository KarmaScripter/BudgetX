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


Private m_Message As String
Private m_Title As String
Private m_Number As Integer
Private m_Description As String
Private m_ErrorArgs As Variant
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type:        Constructor
'   Name:        Class_Initialize
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub








'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    If Not IsNull(Me.OpenArgs) Then
        Me.MessageBox.Value = ""
        Me.MessageBox.Value = CStr(Me.OpenArgs)
    End If
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="Form_Load"
    Exit Sub
End Sub





'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Text(pMessage As String)
    On Error GoTo ErrorHandler:
    If Not pMessage & "" = "" Then
        m_Message = pMessage
    End If
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="Text"
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type:        Event Procedure
'   Name:        Form_Load
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Text() As String
    On Error GoTo ErrorHandler:
    If Not m_Message & "" = "" Then
        Text = m_Message
    End If
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="Text"
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Sub-Procedure
'   Name:        ProcessError
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ExitButton_Click()
    On Error GoTo ErrorHandler:
    DoCmd.Close acForm, Me.Name, acSaveNo
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="ExitButton_Click"
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ErrorArgs(pError As Variant)
    On Error GoTo ErrorHandler:
    If Not IsNull(pError) Then
        m_ErrorArgs = pError
    End If
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="ExitButton_Click"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type:        Property Procedure
'   Name:
'   Parameters:  Void
'   Retval:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ErrorArgs() As Variant
    On Error GoTo ErrorHandler:
    If Not IsNull(m_ErrorArgs) Then
        ErrorArgs = m_ErrorArgs
        Text = CStr(ErrorArgs)
    End If
ErrorHandler:
    ProcessError Name:="MessageDialog", _
        Member:="ExitButton_Click"
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:        Function
'   Name:        VerifyArgs
'   Parameters:  AccountArgs
'   RetVal:      Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Function VerifyArgs(pArg As Variant) As Boolean
    On Error GoTo ErrorHandler
    Dim p_Text As String
    If pArg Is Nothing Then
        p_Text = "AccountArgs Object is NULL or Nothing"
        MessageFactory.Notify (p_Text)
        VerifyArgs = False
        Exit Function
    End If
    If pArg.BFY & "" = "" Then
        p_Text = "AccountArgs BFY Property is not NULL or Empty"
        MessageFactory.Notify (p_Text)
        VerifyArgs = False
        Exit Function
    End If
    If pArg.RpioName & "" = "" Then
        p_Text = "AccountArgs Fiscal Year is not NULL or Empty"
        MessageFactory.Notify (p_Text)
        VerifyArgs = False
        Exit Function
    End If
    If pArg.FundName & "" = "" Then
        p_Text = "AccountArgs FundName Property is NULL or Empty"
        MessageFactory.Notify (p_Text)
        VerifyArgs = False
        Exit Function
    End If
    VerifyArgs = True
ErrorHandler:
    ProcessError Name:="ObligationByFundingDocument", _
        Member:="VerifyArgs"
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

