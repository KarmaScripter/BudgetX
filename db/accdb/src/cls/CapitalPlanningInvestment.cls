VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapitalPlanningInvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit




Private m_Code As String
Private m_Name As String
Private m_Value As Double
Private m_ID As Integer
Private m_Error As String


'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ID(bocid As Integer)
    On Error GoTo ErrorHandler:
    If bocid > 0 Then
        m_ID = bocid
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ID() As Integer
    On Error GoTo ErrorHandler:
    If m_ID > 0 Then
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      ID" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Code(cd As String)
    On Error GoTo ErrorHandler:
    If Not cd & "" = "" Then
        m_Code = cd
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Code() As String
    On Error GoTo ErrorHandler:
    If Not m_Code & "" = "" Then
        Code = m_Code
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      Code" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Name(nm As String)
    On Error GoTo ErrorHandler:
    If Not nm & "" = "" Then
        m_Name = nm
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_Name & "" = "" Then
        Name = m_Name
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:     Name" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Value(vl As Double)
    On Error GoTo ErrorHandler:
    If Not vl & "" = "" Then
        m_Value = vl
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      Value" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Value() As Double
    On Error GoTo ErrorHandler:
    If Not m_Value & "" = "" Then
        Value = m_Value
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:      CapitalPlanningInvestment" _
            & vbCrLf & "Member:      Value" _
            & vbCrLf & "Descript: " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property