VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Appropriation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

'Public Properties of the Appropriation Class

Public Code As String

Public Name As String

Public Title As String

Public TreasurySymbol As String


Private Sub Class_Initialize()
    Code = Null
    Name = Null
    Title = Null
    TreasurySymbol = Null
End Sub
