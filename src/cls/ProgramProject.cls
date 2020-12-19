VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgramProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Code As String

Public Name As String

Public StatutoryAuthority As String

Private Sub Class_Initialize()
    Code = Null
    Name = Null
    StatutoryAuthority = Null
End Sub
