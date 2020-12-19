VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

'Public Properties of the PRC Class

Public RpioCode As String
Public BFY As String
Public BudgetLevel As String
Public FundCode As String
Public AhCode As String
Public OrgCode As String
Public AccountCode As String
Public ActivityCode As String
Public BocCode As String
Public RcCode As String
Public Amount As Double




'Default constructor setting all values to null

Private Sub Class_Initialize()
    RpioCode = Null
    BFY = Null
    BudgetLevel = Null
    FundCode = Null
    AhCode = Null
    OrgCode = Null
    AccountCode = Null
    ActivityCode = Null
    BocCode = Null
    RcCode = Null
    Amount = 0#
End Sub
