VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_ReferenceModels As String
Private m_ViewModels As String
Private m_DataModels As String
Private m_ExcelReport As String
Private m_ExcelBudget As String
Private m_BudgetFolder As String
Private m_DivisionModels As String
Private m_Path As String
Private m_Error As String



Private Sub Class_Initialize()
    m_ViewModels = "Provider=Microsoft.Access.OLEDB.10.0;" _
        & "Persist Security Info=False;" _
        & "Data Source=" _
        & CurrentProject.path _
        & "\Views.accdb;" _
        & "User ID=Admin;" _
        & "Data Provider=Microsoft.ACE.OLEDB.12.0"
        
    m_DataModels = "Provider=Microsoft.Access.OLEDB.10.0;" _
        & "Persist Security Info=False;" _
        & "Data Source=" _
        & CurrentProject.path _
        & "\Data.accdb;" _
        & "User ID=Admin;" _
        & "Data Provider=Microsoft.ACE.OLEDB.12.0"
     
    m_ReferenceModels = "Provider=Microsoft.Access.OLEDB.10.0;" _
        & "Persist Security Info=False;" _
        & "Data Source=" _
        & CurrentProject.path _
        & "\References.accdb;" _
        & "User ID=Admin;" _
        & "Data Provider=Microsoft.ACE.OLEDB.12.0"
        
    m_DivisionModels = "Provider=Microsoft.Access.OLEDB.10.0;" _
        & "Persist Security Info=False;" _
        & "Data Source=" _
        & CurrentProject.path _
        & "\Divisions.accdb;" _
        & "User ID=Admin;" _
        & "Data Provider=Microsoft.ACE.OLEDB.12.0"
      
    m_ExcelReport = CurrentProject.path & "\etc\templates\reports\ReportBase.xlsx"
        
    m_ExcelBudget = CurrentProject.path & "\etc\templates\reports\budgets\BudgetBase.xlsx"
End Sub


'---------------------------------------------------------------------------------
'   Type:            Property-Get
'   Name:            References
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get References() As String
    On Error GoTo ErrorHandler:
    If Not m_ReferenceModels & "" = "" Then
        References = m_ReferenceModels
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      References" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Property-Get
'   Name:            Data
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Data() As String
    On Error GoTo ErrorHandler:
    If Not m_DataModels & "" = "" Then
        Data = m_DataModels
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      Data" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type:            Property-Get
'   Name:            Divisions
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Divisions() As String
    On Error GoTo ErrorHandler:
    If Not m_DivisionModels & "" = "" Then
        Divisions = m_DivisionModels
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      Divisions" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Property-Get
'   Name:            ReportTemplate
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get ReportTemplate() As String
    On Error GoTo ErrorHandler:
    If Not m_ExcelReport & "" = "" Then
        ReportTemplate = m_ExcelReport
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      ReportTemplate" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property





'---------------------------------------------------------------------------------
'   Type:            Property-Get
'   Name:            ExcelBudget
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get BudgetTemplate() As String
    On Error GoTo ErrorHandler:
    If Not m_ExcelBudget & "" = "" Then
        BudgetTemplate = m_ExcelBudget
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      BudgetTemplate" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property




'---------------------------------------------------------------------------------
'   Type:            Property-Let
'   Name:            External
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let ExternalFile(fp As String)
    On Error GoTo ErrorHandler:
    If Not fp & "" = "" And Not Dir(fp) = "" Then
        m_Path = fp
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      ExternalFile" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Property
End Property



'---------------------------------------------------------------------------------
'   Type            Property-Get
'   Name            ExternalFile
'   Parameters      Void
'   Retval          String
'   Purpose
'---------------------------------------------------------------------------------
Public Property Get ExternalFile() As String
    On Error GoTo ErrorHandler:
    If Not m_Path & "" = "" Then
        ExternalFile = m_Path
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetPath" _
            & vbCrLf & "Member:      ExternalFile" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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



