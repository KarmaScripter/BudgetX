VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit





Private m_Base As String
Private m_Path As String
Private m_Name As String
Private m_Extension As String
Private m_ParentFolder As String
Private m_FileSys As Scripting.FileSystemObject
Private m_File As Scripting.File
Private m_CurrentDirectory As String
Private m_Drive As String
Private m_Type As String
Private m_DateCreated As Date
Private m_LastAccessed As Date
Private m_LastModified As Date
Private m_Stream As Scripting.TextStream
Private m_Error As String
Private m_Notification As String





'---------------------------------------------------------------------------------
'   Type:            Constructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub



'---------------------------------------------------------------------------------
'   Type:            Destructor
'   Name:
'   Parameters:      Void
'   Retval:          BudgetFile
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Class_Terminate()
End Sub



'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Let Base(fp As String)
    If Not fp & "" = "" Then
        m_Base = fp
    End If
End Property



'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Property Get Base() As String
    If Not m_Base & "" = "" Then
        Base = m_Base
    End If
End Property




'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            FileExists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function Exists(pPath As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not pPath & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        If m_FileSys.FileExists(pPath) = True Then
            Exists = True
        Else
            Exists = False
        End If
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Exists(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetStreamReader
'   Parameters:      String
'   Retval:          TextStream
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetStreamReader() As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_File = m_FileSys.GetFile(m_Base)
        Set GetStreamReader = m_File.OpenAsTextStream(IOMode:=ForReading, _
            Format:=TristateUseDefault)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetStreamReader" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetStreamReader
'   Parameters:      String
'   Retval:          TextStream
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetStreamWriter() As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set m_File = m_FileSys.GetFile(m_Base)
            Set GetStreamReader = m_File.OpenAsTextStream(IOMode:=ForWriting, _
                Format:=TristateUseDefault)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetStreamReader" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            Create()
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function CreateText(pFileName As String) As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not pFileName & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set CreateText = m_FileSys.CreateTextFile(filename:=pFileName, _
                Overwrite:=True)
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Create()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetDrive() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_Drive = m_FileSys.GetFile(m_Base).Drive
        GetDrive = m_Drive
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetDrive()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetParentFolder() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_ParentFolder = m_FileSys.GetFile(m_Base).ParentFolder
        GetParentFolder = m_ParentFolder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetPath()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            GetFullPath()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetFullPath() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_Path = m_FileSys.GetAbsolutePathName(m_Base)
        GetFullPath = m_Path
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetFullPath()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetDateCreated() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_DateCreated = m_FileSys.GetFile(m_Base).DateCreated
        GetDateCreated = m_DateCreated
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetDateCreated()" _
            & vbCrLf & "Descript:   " & Err.Description
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetLastAccessed() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_LastAccessed = m_FileSys.GetFile(m_Base).DateLastAccessed
        GetLastAccessed = m_LastAccessed
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetLastAccessed()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetLastModified() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_LastModified = m_FileSys.GetFile(m_Base).DateLastModified
        GetLastModified = m_LastModified
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetLastModified()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetName() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_Name = m_FileSys.GetFile(m_Base).ShortName
        GetName = m_Name
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetFileName()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetType() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_Type = m_FileSys.GetFile(m_Base).Type
        GetType = m_Type
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetType()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetSize() As Variant
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        GetSize = m_FileSys.GetFile(m_Base).Size
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetFileName()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



 
'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            Extension
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetExtension() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        m_Extension = m_FileSys.GetExtensionName(m_Base)
        GetExtension = m_Extension
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetExtension()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetLength() As Long
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        GetLength = FileLen(m_Base)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     GetLength()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function ReadAll() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_File = m_FileSys.GetFile(m_Base)
        Dim m_Out As String
        Dim m_Stream As Scripting.TextStream
        Set m_Stream = m_File.OpenAsTextStream
        m_Out = m_Stream.ReadAll
        Set m_FileSys = Nothing
        Set m_File = Nothing
        Set m_Stream = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     ReadAll()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function Readline() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_File = m_FileSys.GetFile(m_Base)
        Dim m_Out As String
        Dim m_Stream As Scripting.TextStream
        Set m_Stream = m_File.OpenAsTextStream
        m_Out = m_Stream.Readline
        Set m_FileSys = Nothing
        Set m_File = Nothing
        Set m_Stream = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Readline" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Move
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Move(des As String)
    If Not m_Path & "" = "" And _
        Not des & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_File = m_FileSys.GetFile(m_Base)
            m_File.Move (des)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Move(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileCopy
'   Parameters:      String, String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Copy(dest As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And Not dest & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_File = m_FileSys.GetFile(m_Base)
        m_File.Copy Destination:=dest, OverWriteFiles:=True
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Copy(String, String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Delete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Delete()
    On Error GoTo ErrorHandler:
    Dim m_FileSys As FileSystemObject
    Set m_FileSys = New FileSystemObject
    Set m_FileSys = CreateObject("Scripting.FileSystemObject")
    If Not m_Base & "" = "" Then
        m_FileSys.DeleteFile m_Path
        Kill (m_Path)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Delete()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileRename
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Rename(pName As String)
    On Error GoTo ErrorHandler:
    If Not m_Path & "" = "" And _
        Not pName & "" = "" Then
            Name m_Path As pName
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     Rename(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            RecursiveDelete
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub RecursiveDelete(pFilePath As String, pFileName As String)
    On Error GoTo ErrorHandler:
    Dim p_CurrentDirectory As String
    Dim p_CurrentPath As String
    Dim p_File As Variant
    Dim p_Folders As Collection
    Set p_Folders = New Collection
    If Len(Dir(pFilePath & pFileName)) > 0 Then
       SetAttr pFilePath & pFileName, vbNormal
       Kill pFilePath & pFileName
    End If
    p_CurrentDirectory = Dir(pFilePath, vbDirectory)
    Do Until p_CurrentDirectory = vbNullString
        If Left(p_CurrentDirectory, 1) <> "." _
            And (GetAttr(pFilePath & p_CurrentDirectory) _
            And vbDirectory) = vbDirectory Then
                p_Folders.add pFilePath & p_CurrentDirectory & "\"
        End If
        p_CurrentDirectory = Dir()
    Loop
    For Each p_File In p_Folders
        RecursiveDelete CStr(p_File), pFileName
    Next p_File
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     RecursiveDelete(String, String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Browse()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select a file"
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    p_Dialog.Filters.add "All files", "*.*"
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     BrowseFile()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub
 
 
 


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function UnZip(pSource As String, pDestination As String) As String
    On Error GoTo ErrorHandler:
    Dim p_FailedReason As String
    Dim p_ZipToolPath As String
    Dim p_ShellCmd As String
    Dim p_Success As Boolean
    p_ZipToolPath = CurrentProject.path & "\etc\7za"
    p_ShellCmd = p_ZipToolPath & " x " & pSource & " -o" & pDestination & " -ry"
    p_Success = Shell(p_ShellCmd, vbHide)
    If p_Success = True Then
        Kill pSource
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFile" _
            & vbCrLf & "Member:     BrowseFile()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    p_FailedReason = Err.Description
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function
'   Name:            GetCurrentDirectory()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function GetCurrentDirectory() As String
    On Error GoTo ErrorHandler:
    m_CurrentDirectory = CurDir()
    GetCurrentDirectory = m_CurrentDirectory
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetCurrentDirectory()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub OpenSaveDialog()
    On Error GoTo ErrorHandler:
    Dim p_Dialog As FileDialog
    Dim result As Integer
    Set p_Dialog = Application.FileDialog(msoFileDialogSaveAs)
    p_Dialog.AllowMultiSelect = False
    p_Dialog.Title = "Select location to save file."
    p_Dialog.InitialFileName = "C:\"
    p_Dialog.Filters.Clear
    If p_Dialog.Show = -1 Then
       Debug.Print p_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetFile" _
            & vbCrLf & "Member:     OpenSaveDialog()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub




'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            GetAttribute
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetAttribute() As VbFileAttribute
    On Error GoTo ErrorHandler:
    If Not Dir(m_Base) & "" = "" Then
        Dim p_Attribute As VbFileAttribute
        p_Attribute = GetAttr(m_Base)
        GetAttribute = p_Attribute
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetFile" _
            & vbCrLf & "Member:     GetAttribute()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function
