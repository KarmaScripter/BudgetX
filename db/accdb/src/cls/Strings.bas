Attribute VB_Name = "Strings"
Option Compare Database
Option Explicit


Private m_Error As String





'---------------------------------------------------------------------------------
'   Type            Function
'   Name            ToProperCase
'   Parameters      Void
'   Return          Void
'   Purpose
'                   Capitalize the first character and add a space before
'                   each capitalized letter (except the first character).
'---------------------------------------------------------------------------------
Public Static Function ToProperCase(ByVal pString As String) As String
    On Error GoTo ErrorHandler:
    Dim result As String
    Dim i As Integer
    Dim ch As String
    If Len(pString) < 2 Then
        ToProperCase = UCase$(pString)
        Exit Function
    End If
    result = UCase$(mID$(pString, 1, 1))
    For i = 2 To Len(pString)
        ch = mID$(pString, i, 1)
        If (UCase$(ch) = ch) Then result = result & " "
        result = result & ch
    Next i
    ToProperCase = result
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type            Function
'   Name            ToCamelCase
'   Parameters      Void
'   Return          Void
'   Purpose
'                   Capitalize the first character and add a space before
'                   each capitalized letter (except the first character).
'---------------------------------------------------------------------------------
Public Function ToCamelCase(ByVal pString As String) As String
    On Error GoTo ErrorHandler:
    Dim result As String
    result = ToPascalCase(pString)
    If Len(result) > 0 Then
        Mid$(result, 1, 1) = LCase$(mID$(result, 1, 1))
    End If
    ToCamelCase = result
ErrorHandler:
    ProcessError
    Exit Function
End Function



'---------------------------------------------------------------------------------
'   Type            Function
'   Name            ToPascalCase
'   Parameters      Void
'   Return          Void
'   Purpose
'                   Capitalize the first character and add a space before
'                   each capitalized letter (except the first character).
'---------------------------------------------------------------------------------
Public Function ToPascalCase(ByVal pString As String) As String
    On Error GoTo ErrorHandler:
    Dim words() As String
    Dim i As Integer
    If Len(pString) < 2 Then
        ToPascalCase = UCase$(pString)
        Exit Function
    End If
    words = Split(pString)
    For i = LBound(words) To UBound(words)
        If (Len(words(i)) > 0) Then
            Mid$(words(i), 1, 1) = UCase$(mID$(words(i), 1, _
                1))
        End If
    Next i
    ToPascalCase = Join(words, "")
ErrorHandler:
    ProcessError
    Exit Function
End Function





'---------------------------------------------------------------------------------
'   Type            Function
'   Name            SearchArray
'   Parameters      Void
'   Return          Void
'   Purpose
'                   Capitalize the first character and add a space before
'                   each capitalized letter (except the first character).
'---------------------------------------------------------------------------------
Public Function SearchArray(pArray As Variant, pString As String) As Integer
    Dim FindStrInArray As Integer
    FindStrInArray = -1
    Dim i As Integer
    For i = LBound(pArray) To UBound(pArray)
        If pString = pArray(i) Then
            FindStrInArray = i
            Exit For
        End If
    Next i
ErrorHandler:
    ProcessError
    Exit Function
End Function




'---------------------------------------------------------------------------------
'   Type            Function
'   Name            SplitToArray
'   Parameters      Void
'   Return          Void
'   Purpose
'                   Capitalize the first character and add a space before
'                   each capitalized letter (except the first character).
'---------------------------------------------------------------------------------
Public Function SplitIntoArray(pString As String, m_Separator As String) As Variant
    Dim Arr As Variant
    If Len(pString) > 0 Then
        Arr = Split(pString, m_Separator)
        Dim i As Integer
        For i = LBound(Arr) To UBound(Arr)
            Arr(i) = Trim(Arr(i))
        Next i
    Else
        Arr = Array()
    End If
    SplitIntoArray = Arr
ErrorHandler:
    ProcessError
    Exit Function
End Function





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



