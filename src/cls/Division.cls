VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Division"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

'Public Properties of the Division Class

Public Name As String
Public Code As String
Public Caption As String
Public Title As String


'Initalizer setting all inital properties to null for later assingment
Private Sub Class_Initialize()
    Name = Null
    Code = Null
    Caption = Null
    Title = Null
End Sub
