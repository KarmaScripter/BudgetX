VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private m_BudgetExecution As DAO.Workspace
Private m_Data As DAO.Database
Private m_DataSet As DAO.Recordset
Private m_Error As String
Private m_Notification As String







'----------------------------------------------------------------------------------
'   Type        Contructor/Finalizer
'   Name        Class_Initialize()/Class_Terminate
'   Parameters  Void
'   Retval      AccessData
'   Purpose
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_Data = CurrentDb
    Set m_BudgetExecution = DBEngine.Workspaces(0)
End Sub

Private Sub Class_Terminate()
    If Not IsNull(m_Data) Then
        Set m_Data = Nothing
    End If
    If Not IsNull(m_BudgetExecution) Then
        Set m_BudgetExecution = Nothing
    End If
End Sub





'----------------------------------------------------------------------------------
'   Type:        Property Get
'   Name:        IsConnected
'   Parameters:  Void
'   RetVal:      Boolean
'   Purpose:     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsConnected() As Boolean
    On Error GoTo ErrorHandler:
    If Not IsObject(m_Data) Then
        IsConnected = True
    End If
    If IsObject(m_Data) Then
        IsConnected = False
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Property Get
'   Name        IsOpen
'   Parameters  Void
'   RetVal      Boolean
'   Purpose     Function returns boolean value indicating if Workspace is open
'---------------------------------------------------------------------------------
Public Property Get IsOpen() As Boolean
    On Error GoTo ErrorHandler:
    If Not IsObject(m_BudgetExecution) Then
        IsOpen = True
    End If
    If IsObject(m_BudgetExecution) Then
        IsOpen = False
    End If
ErrorHandler:
    ProcessError
    Exit Property
End Property


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetDataSet
'   Parameters  String - table;
'   RetVal      DAO.Recordset
'   Purpose     Function returns the records from the Data Table using the table's
'               name or a sql select query
'---------------------------------------------------------------------------------
Public Function GetDataSet(table As String) As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim rs As DAO.Recordset
    If Not IsNull(m_Data) And _
        Not table & "" = "" Then
           Set rs = m_Data.OpenRecordset(table, dbOpenDynaset)
    End If
    Set GetDataSet = rs
ErrorHandler:
    ProcessError
    Exit Function
End Function


'----------------------------------------------------------------------------------
'   Type        Function
'   Name        GetRecord
'   Parameters  String - table;
'   RetVal      DAO.Recordset
'   Purpose     Function returns record using ID parameter and Primary Key
'               defined in a sql select query
'---------------------------------------------------------------------------------
Public Function GetDataRecord(table As String, primarykey As String, ID As Integer) As DAO.Recordset
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Set DB = m_Data
    Dim rs As DAO.Recordset
    Dim SQL As String
    If Not IsNull(m_Data) And _
        Not table & "" = "" And _
        Not primarykey & "" = "" And _
        ID > 0 Then
            SQL = "SELECT * FROM table WHERE key = 'id'"
            Set rs = m_Data.OpenRecordset(SQL, dbOpenDynaset)
            Set GetDataRecord = rs
    End If
ErrorHandler:
    ProcessError
    Exit Function
End Function




'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        CloseDatabase
'   Parameters  Void
'   RetVal      Void
'   Purpose     Method used to close the database
'---------------------------------------------------------------------------------
Public Sub CloseDatabase()
    On Error GoTo ErrorHandler:
    If Not m_Data Is Nothing Then
        Set m_Data = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub



'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        CloseWorkspace
'   Parameters  Void
'   RetVal      Void
'   Purpose     Method used to close the database
'---------------------------------------------------------------------------------
Public Sub CloseWorkspace()
    On Error GoTo ErrorHandler:
    If Not IsNull(m_BudgetExecution) Then
        Set m_BudgetExecution = Nothing
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        UpdateDate
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to update data in current recordset
'---------------------------------------------------------------------------------
Public Sub UpdateData(table As String, field As String, Value As Variant)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Set DB = m_Data
    Dim rs As DAO.Recordset
    If Not table & "" = "" Then
        Set rs = GetDataSet(table)
    End If
    If Not field & "" = "" _
        And Not IsNull(Value) Then
            rs.Edit
            rs.Fields(field) = Value
            rs.Update
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        AddRecord
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to add data to recordset
'---------------------------------------------------------------------------------
Public Sub AddRecord(table As String, field As String, Value As Variant)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = m_Data
    Set rs = GetDataSet(table)
    With rs
      .AddNew
      !Fields(field) = Value
      .Update
    End With
ErrorHandler:
    ProcessError
    Exit Sub
End Sub

'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        UpdateRecord
'   Parameters  table - String, field - String, value - Variant
'   RetVal      Void
'   Purpose     Method used to add data to recordset
'---------------------------------------------------------------------------------
Public Sub UpdateRecord(table As String, field As String, NewValue As Variant, ID As Integer)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = m_Data
    Set rs = GetDataRecord(table, field, ID)
    If Not IsNull(rs) Then
        With rs
          .AddNew
          !Fields(field) = NewValue
          .Update
        End With
    End If
ErrorHandler:
    ProcessError
    Exit Sub
End Sub


'----------------------------------------------------------------------------------
'   Type        Subprocedure
'   Name        DeleteRecord
'   Parameters  table - String, pk - String, id - integer
'   RetVal      Void
'   Purpose     Method used to delete records from the data set
'---------------------------------------------------------------------------------
Public Sub DeleteRecord(table As String, pk As String, ID As Integer)
    On Error GoTo ErrorHandler:
    Dim DB As DAO.Database
    Dim rs As DAO.Recordset
    Set DB = m_Data
    Dim SQL As String
    SQL = "DELETE FROM table WHERE pk = id"
    Set rs = DB.OpenRecordset(SQL)
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



