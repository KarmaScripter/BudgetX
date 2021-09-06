Attribute VB_Name = "SQLite"
Option Compare Database


Public pid As Variant
Private m_SQLitePath As String
Private m_SQLiteArg As String
Private m_ShellArgPath As String
Private m_Error As String




'----------------------------------------------------------------------------------
'   Type        SubProcedure
'   Name        Calculate
'   Parameters  Void
'   Purpose     Launches the Windows 10 calculator 'calc.exe'
'----------------------------------------------------------------------------------
Public Sub Run()
    On Error GoTo ErrorHandler:
    m_SQLitePath = Replace(CurrentProject.path, "accdb", "sqlite\gui\SQLiteDatabaseBrowserPortable.exe")
    m_SQLiteArg = " " & Replace(CurrentProject.path, "accdb", "sqlite\gui\Data.db")
    m_ShellArgPath = m_SQLitePath & m_SQLiteArg
    vPID = Shell(mShellArgPath, 3)
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       SQLite" _
            & vbCrLf & "Member:     Run()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
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



