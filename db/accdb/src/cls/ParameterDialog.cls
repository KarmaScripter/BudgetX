VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParameterDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private SelectedYear As String
Private SelectedLevel As String
Private SelectedFund As String
Private SelectedAccount As String
Private SelectedObjectClass As String
Private SelectedRcCode As String
Private m_Error As String


Public Sub Class_Initialize()
End Sub

Public Sub Class_Terminate()
End Sub






'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get FiscalYear()
    If Not IsNull(SelectedYear) And SelectedYear <> "" Then
        FiscalYear = SelectedYear
    End If
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Level()
    If Not IsNull(SelectedLevel) And SelectedLevel <> "" Then
        Level = SelectedLevel
    End If
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get Fund()
    If Not IsNull(SelectedFund) And SelectedFund <> "" Then
        Fund = SelectedFund
    End If
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get PRC()
    If Not IsNull(SelectedAccount) And SelectedAccount <> "" Then
        PRC = SelectedAccount
    End If
End Property



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get BOC()
    If Not IsNull(SelectedObjectClass) And SelectedObjectClass <> "" Then
        BOC = SelectedObjectClass
    End If
End Property

'----------------------------------------------------------------------------------
'   Type
'   Name
'   Parameters
'   Purpose
'---------------------------------------------------------------------------------
Private Sub CloseButton_Click()
    If Not IsNull(SelectedYear) Then
        SelectedYear = vbNull
    End If
    
    
    If Not IsNull(SelectedLevel) Then
        SelectedLevel = vbNull
    End If
    
    
    If Not IsNull(SelectedFund) Then
        SelectedFund = vbNull
    End If
        
    If Not IsNull(SelectedAccount) Then
        SelectedAccount = vbNull
    End If
    
    If Not IsNull(SelectedObjectClass) Then
        SelectedObjectClass = vbNull
    End If
    
    If Not IsNull(SelectedRcCode) Then
        SelectedRcCode = vbNull
    End If
    
End Sub


'----------------------------------------------------------------------------------
'   Type
'   Name
'   Parameters
'   Purpose
'---------------------------------------------------------------------------------
Private Sub ExecuteButton_Click()
    '---- Private variable declaration
End Sub


'----------------------------------------------------------------------------------
'   Type
'   Name
'   Parameters
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FiscalYearComboBox_Change()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = Me!BocCodeComboBox
        SelectedYear = cbo.SelText
End Sub


'----------------------------------------------------------------------------------
'   Type
'   Name
'   Parameters
'   Purpose
'---------------------------------------------------------------------------------
Private Sub BudgetLevelComboBox_Change()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = Me!BocCodeComboBox
        SelectedLevel = cbo.SelText
End Sub


'----------------------------------------------------------------------------------
'   Type
'   Name
'   Parameters
'   Purpose
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_Change()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = Me!BocCodeComboBox
        SelectedFund = cbo.SelText
End Sub




'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Sub AccountCodeComboBox_Change()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = Me!AccountCodeComboBox
    
    If Not IsNull(SelectedAccount) And SelectedAccount <> "" Then
         SelectedAccount = cbo.SelText
    End If
End Sub



'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name
'   Parameters  Void
'   Retval      String
'   Purpose
'---------------------------------------------------------------------------------
Private Sub BocCodeComboBox_Change()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = Me!BocCodeComboBox
    
    If Not IsNull(SelectedObjectClass) And SelectedObjectClass <> "" Then
        SelectedObjectClass = cbo.SelText
    End If
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


