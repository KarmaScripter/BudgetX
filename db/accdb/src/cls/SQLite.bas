Attribute VB_Name = "SQLite"
Option Compare Database


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS   ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Public pid As Variant
Private mSQLitePath As String
Private mSQLiteArg As String
Private mShellArgPath As String
Private mError As String


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   METHODS  ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------



'----------------------------------------------------------------------------------
'   Type        SubProcedure
'   Name        Calculate
'   Parameters  Void
'   Purpose     Launches the Windows 10 calculator 'calc.exe'
'----------------------------------------------------------------------------------
Public Sub Run()
    On Error GoTo ErrorHandler:
    mSQLitePath = Replace(CurrentProject.path, "accdb", "sqlite\gui\SQLiteDatabaseBrowserPortable.exe")
    mSQLiteArg = " " & Replace(CurrentProject.path, "accdb", "sqlite\gui\Data.db")
    mShellArgPath = mSQLitePath & mSQLiteArg
    vPID = Shell(mShellArgPath, 3)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:       SQLite" _
            & vbCrLf & "Member:     Run()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub


