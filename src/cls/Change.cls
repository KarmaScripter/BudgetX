VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private m_TableName As String
Private m_FieldName As String
Private m_Action As String
Private m_TimeStamp As String
Private m_OldValue As String
Private m_NewValue As String
Private m_ChangeId As Integer
Private m_Error As String


Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub




'----------------------------------------------------------------------------------
'   Type        Property
'   Name        ID(Integer)
'   Parameters  Integer
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let ID(cid As Integer)
    On Error GoTo ErrorHandler:
    If cid > 0 Then
        m_ChangeId = cid
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="ID"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        ID()
'   Parameters  Void
'   RetVal      Integer
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If m_ChangeId > 0 Then
        ID = m_ChangeId
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="ID"
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type        Property
'   Name        TableName(String)
'   Parameters  Void
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let TableName(table As String)
    On Error GoTo ErrorHandler:
    If Not table & "" = "" Then
        m_TableName = table
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="TableName"
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property
'   Name        FieldName()
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get TableName() As String
    On Error GoTo ErrorHandler:
    If Not m_TableName & "" = "" Then
        TableName = m_TableName
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="TableName"
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type        Property
'   Name        FieldName(String)
'   Parameters  String
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let FieldName(field As String)
    On Error GoTo ErrorHandler:
    If Not field & "" = "" Then
        m_FieldName = field
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="FieldName"
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property
'   Name        FieldName()
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get FieldName() As String
    On Error GoTo ErrorHandler:
    If Not m_FieldName & "" = "" Then
        FieldName = m_FieldName
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="FieldName"
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property
'   Name        Action(String)
'   Parameters  Void
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let Action(act As String)
    On Error GoTo ErrorHandler:
    If Not act & "" = "" Then
        m_Action = act
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="Action"
    Exit Property
End Property





'----------------------------------------------------------------------------------
'   Type        Property
'   Name        Action()
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Action() As String
    On Error GoTo ErrorHandler:
    If Not m_Action & "" = "" Then
        Action = m_Action
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="Action"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        TimeStamp(String)
'   Parameters  String
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let TimeStamp(tm As String)
    On Error GoTo ErrorHandler:
    If Not tm & "" = "" Then
        m_TimeStamp = tm
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="TimeStamp"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        TimeStamp
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get TimeStamp() As String
    On Error GoTo ErrorHandler:
    If Not m_TimeStamp & "" = "" Then
        TimeStamp = m_TimeStamp
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="TimeStamp"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        OldValue(String)
'   Parameters  Void
'   RetVal      Void
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let OldValue(ov As String)
    On Error GoTo ErrorHandler:
    If Not ov & "" = "" Then
        m_OldValue = ov
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="OldValue"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        OldValue()
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get OldValue() As String
    If Not m_OldValue & "" = "" Then
        OldValue = m_OldValue
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="OldValue"
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type        Property
'   Name
'   Parameters  Void
'   RetVal
'   Purpose
'---------------------------------------------------------------------------------
Public Property Let NewValue(nv As String)
    On Error GoTo ErrorHandler:
    If Not nv & "" = "" Then
        m_NewValue = nv
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="NewValue"
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property
'   Name        Void
'   Parameters  Void
'   RetVal      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get NewValue() As String
    On Error GoTo ErrorHandler:
    If Not m_NewValue & "" = "" Then
        NewValue = m_NewValue
    End If
ErrorHandler:
    ProcessError Name:="Change"
        Member:="NewValue"
    Exit Property
End Property












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

