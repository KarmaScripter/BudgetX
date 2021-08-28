VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BudgetFolder"
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
Private m_Folder As Scripting.folder
Private m_File As Scripting.File
Private m_Files As Scripting.Files
Private m_SubFolders As Scripting.folders
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
'   Name:            Exists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Function Exists() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        If m_FileSys.FolderExists(m_Base) = True Then
            Exists = True
        Else
            Exists = False
        End If
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     Exists()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            Exists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Function IsRooted() As Boolean
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        If m_FileSys.GetFolder(m_Base).IsRootFolder = True Then
            IsRooted = True
        Else
            IsRooted = False
        End If
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     IsRooted()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function


'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:            FolderExists
'   Parameters:      String
'   Retval:          Boolean
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function VerifyExists(pPath As String) As Boolean
    On Error GoTo ErrorHandler:
    If Not pPath & "" = "" And _
        Not m_Base & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            If m_FileSys.FolderExists(pPath) = True Then
                VerifyExists = True
            Else
                VerifyExists = False
            End If
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     VerifyExists(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function Create() As Scripting.folder
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_Folder = m_FileSys.CreateFolder(m_Base)
        Set Create = m_Folder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     Create()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type:            Property Let
'   Name:
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Function CreateDirectory(pLocation As String) As Scripting.folder
    On Error GoTo ErrorHandler:
    Dim folder As Scripting.folder
    If Not pLocation & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set folder = m_FileSys.CreateFolder(pLocation)
        Set CreateDirectory = folder
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     CreateDirectory(String)" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
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
    Dim m_Current As String
    m_Current = CurDir()
    GetCurrentDirectory = m_Current
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
'   Type:            Function
'   Name:            GetCurrentDirectory()
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Static Sub ChangeDirectory(pLocation As String)
    On Error GoTo ErrorHandler:
    If Not pLocation & "" = "" Then
        ChDir (m_Base)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     ChangeDirectory()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Function Procedure
'   Name:
'   Parameters:      Void
'   Retval:          String
'   Purpose:
'---------------------------------------------------------------------------------
Public Function GetPath() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Path = m_FileSys.GetFolder(m_Base).ShortPath
        GetPath = m_Path
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetPath()" _
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
Public Function GetDrive() As String
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Drive = m_FileSys.GetFolder(m_Base).Drive
        GetDrive = m_Drive
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
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
    If Not m_Base & "" = "" And _
        m_ParentFolder & "" = "" Then
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            m_ParentFolder = m_FileSys.GetParentFolderName(m_Base)
            GetParentFolder = m_ParentFolder
            Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetParentFolder()" _
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
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_DateCreated = m_FileSys.GetFolder(m_Base).DateCreated
        GetDateCreated = m_DateCreated
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetDateCreated()" _
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
Public Function GetLastAccessed() As Date
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_LastAccessed = m_FileSys.GetFolder(m_Base).DateLastAccessed
        GetLastAccessed = m_LastAccessed
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
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
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_LastModified = m_FileSys.GetFolder(m_Base).DateLastModified
        GetLastModified = m_LastModified
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
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
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Name = m_FileSys.GetFolder(m_Base).ShortName
        GetName = m_Name
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetName()" _
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
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        m_Type = m_FileSys.GetFile(m_Base).Type
        GetType = m_Type
        Set m_FileSys = Nothing
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetType()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Function GetDirectorySize() As Variant
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set GetDirectorySize = m_FileSys.GetFolder(m_Base).Size
    End If
    Set m_FileSys = Nothing
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetSize()" _
            & vbCrLf & "Descript:   " & Err.Description
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
Public Function GetSubFolders() As Scripting.folders
    On Error GoTo ErrorHandler:
    Dim count As Integer
    Dim folders As Scripting.folders
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        Set GetSubFolders = m_SubFolders
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetSubFolders()" _
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
Public Function CountSubFolders() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        CountSubFolders = m_SubFolders.count
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetSubFolders()" _
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
Public Sub AddSubFolder(pFolderName As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set m_SubFolders = m_FileSys.GetFolder(m_Base).SubFolders
        Dim p_Folder As New Scripting.FileSystemObject
        Set p_Folder = m_SubFolders.add(pFolderName)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetSubFolders()" _
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
Public Function GetFiles() As Scripting.Files
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        Set GetFiles = m_FileSys.GetFolder(m_Base).Files
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetFiles()" _
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
Public Function CountFiles() As Integer
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" Then
        Set m_FileSys = New FileSystemObject
        Set m_FileSys = CreateObject("Scripting.FileSystemObject")
        CountFiles = m_FileSys.GetFolder(m_Base).Files.count
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     GetFiles()" _
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
    On Error GoTo ErrorHandler:
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
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
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
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
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     Delete()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            CreateTextFile
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Function CreateTextFile(pFileName As String) As Scripting.TextStream
    On Error GoTo ErrorHandler:
    If Not pFileName & "" = "" And _
        Not m_Base & "" = "" Then
            Dim m_TextFile As Scripting.TextStream
            Set m_FileSys = New FileSystemObject
            Set m_FileSys = CreateObject("Scripting.FileSystemObject")
            Set m_TextFile = m_FileSys.CreateTextFile(filename:=pFileName, Overwrite:=True)
            CreateTextFile = m_TextFile
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     CreateTextFile(String)" _
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
Public Sub Browse()
    On Error GoTo ErrorHandler:
    Set m_FileSys = New FileSystemObject
    Set m_FileSys = CreateObject("Scripting.FileSystemObject")
    Dim m_Dialog As FileDialog
    Dim result As Integer
    Set m_Dialog = Application.FileDialog(msoFileDialogFilePicker)
    m_Dialog.AllowMultiSelect = False
    m_Dialog.Title = "Select a file"
    m_Dialog.InitialFileName = "C:\"
    m_Dialog.Filters.Clear
    m_Dialog.Filters.add "Text files", "*.txt"
    m_Dialog.Filters.add "All files", "*.*"
    If m_Dialog.Show = -1 Then
       m_FileSys.OpenTextFile (m_Dialog.Item)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     Browse()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub

 


'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            FileRename
'   Parameters:      String
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Rename(pName As String)
    On Error GoTo ErrorHandler:
    If Not m_Base & "" = "" And _
        Not pName & "" = "" Then
         If Dir(m_Base) <> "" Then
            Name m_Base As pName
        End If
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     Rename(String)" _
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
Public Sub OpenFolderDialog()
    On Error GoTo ErrorHandler:
    Dim m_Dialog As FileDialog
    Dim result As Integer
    Set m_Dialog = Application.FileDialog(msoFileDialogFolderPicker)
    m_Dialog.AllowMultiSelect = False
    m_Dialog.Title = "Select an Folder"
    m_Dialog.InitialFileName = "C:\"
    m_Dialog.Filters.Clear
    If m_Dialog.Show = -1 Then
       Debug.Print m_Dialog.SelectedItems(1)
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:       BudgetFolder" _
            & vbCrLf & "Member:     OpenFolderDialog()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub





'---------------------------------------------------------------------------------
'   Type:            Sub-Procedure
'   Name:            Zip
'   Parameters:      Void
'   Retval:          Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub Zip(pFolderPath As Variant, p_ZippedPath As Variant)
    On Error GoTo ErrorHandler:
    Dim p_Shell As Object
    Open p_ZippedPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    Set p_Shell = CreateObject("Shell.Application")
    p_Shell.Namespace(p_ZippedPath).CopyHere p_Shell.Namespace(pFolderPath).Items
    Do Until p_Shell.Namespace(p_ZippedPath).Items.count = p_Shell.Namespace(pFolderPath).Items.count
        p_Shell.Wait (Now + TimeValue("0:00:01"))
    Loop
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   BudgetFolder" _
            & vbCrLf & "Member:     BrowseFile()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    Exit Sub
End Sub


