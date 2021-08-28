VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_DataAccess As DataAccess
Private m_Base As DAO.field
Private m_DataRow As DataRow
Private m_Value  As Variant
Private m_Name As String
Private m_Caption As String
Private m_Type As String
Private m_Ordinal As Integer
Private m_Error As String
Private m_Notification As String





Private Sub Class_Initialize()
    Set m_DataAccess = New DataAccess
    Set m_Base = New DAO.field
End Sub

Private Sub Class_Terminate()
    If Not m_DataAccess Is Nothing Then
        Set m_DataAccess = Nothing
    End If
    If Not m_Base Is Nothing Then
        Set m_Base = Nothing
    End If
End Sub



'----------------------------------------------------------------------------------
'   Type:        Property Let
'   Name:        Base
'   Parameters:  Void
'   RetVal:      Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Base(pBase As DAO.field)
    On Error GoTo ErrorHandler:
    If Not pBase Is Nothing Then
        Set m_Base = pBase
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Base" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Base
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As DAO.field
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Set Base = m_Base
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Base" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        DAO.Field
'   Parameters:  Void
'   RetVal:
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Name() As String
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Name" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Ordinal
'   Parameters:  Void
'   RetVal:      Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Ordinal() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        Ordinal = m_Base.OrdinalPosition
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Ordinal" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Ordinal
'   Parameters:  Void
'   RetVal:      Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Value() As Variant
    On Error GoTo ErrorHandler:
    If Not m_Base.Value Is Nothing Then
        Value = m_Base.Value
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Value" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        DataType
'   Parameters:  Void
'   RetVal:      Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get DataType() As String
    On Error GoTo ErrorHandler:
    If Not m_Base.Type & "" = "" Then
        DataType = m_Base.Type
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     DataType" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Ordinal
'   Parameters:  Void
'   RetVal:      Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Caption(pCaption As String)
    On Error GoTo ErrorHandler:
    If Not pCaption & "" = "" Then
        m_Caption = pCaption
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Caption" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property







'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        Ordinal
'   Parameters:  Void
'   RetVal:      Integer
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Caption() As String
    On Error GoTo ErrorHandler:
    If Not m_Caption & "" = "" Then
        Caption = m_Caption
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Ordinal" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        IsConnected
'   Parameters:  Void
'   RetVal:      Boolean
'   Purpose:     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsConnected() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base Is Nothing Then
        IsConnected = True
    End If
    If m_Base Is Nothing Then
        IsConnected = False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       DataColumn" _
            & vbCrLf & "Member:     Base" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




