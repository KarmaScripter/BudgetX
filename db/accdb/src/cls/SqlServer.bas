Attribute VB_Name = "SqlServer"
Option Compare Database


'---------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------   FIELDS   ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------

Public pid As Variant
Private mCompactPath As String
Private mCompactArg As String
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
Public Sub RunCompact()
    On Error GoTo ErrorHandler:
    mCompactPath = Replace(CurrentProject.path, "accdb", "sqlce\gui\CompactView.exe")
    mCompactArg = " " & Replace(CurrentProject.path, "accdb", "sqlce\gui\Data.sdf")
    mShellArgPath = mCompactPath & mCompactArg
    vPID = Shell(mShellArgPath, 3)
ErrorHandler:
    If Err.Number > 0 Then
        mError = "Source:   SqlServer" _
            & vbCrLf & "Member:     RunCompact()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (mError)
    Exit Sub
End Sub
