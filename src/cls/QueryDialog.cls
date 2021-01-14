VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QueryDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---- Private Backing Field Declarations

Private strBFY As String
Private strCompassLevel As String
Private strFundCode As String
Private strAccountCode As String
Private strBocCode As String


Private Sub OnCloseButtonClick()
    Dim Action As FormAction
    Set Action = New FormAction
    Action.OnCloseButtonClick
End Sub

'------  Event handler for the Query Dialog Execute Button ----

Private Sub OnExecuteButtonClick()
    '---- Private variable declaration

    
End Sub

'---- Public Properties

Public Property Get FiscalYear()
    If Not IsNull(strBFY) And strBFY <> "" Then
        FiscalYear = strBFY
    End If
End Property

Public Property Get Level()
    If Not IsNull(strCompassLevel) And strCompassLevel <> "" Then
        Level = strCompassLevel
    End If
End Property

Public Property Get Fund()
    If Not IsNull(strFundCode) And strFundCode <> "" Then
        Fund = strFundCode
    End If
End Property


Public Property Get PRC()
    If Not IsNull(strAccountCode) And strAccountCode <> "" Then
        PRC = strAccountCode
    End If
End Property


Public Property Get BOC()
    If Not IsNull(strBocCode) And strBocCode <> "" Then
        BOC = strBocCode
    End If
End Property

'---- Event Handlers for the Query Dialog ComboBoxes

Private Sub OnBfyComboBoxChange()
    strBFY = BfyComboBox.SelText
End Sub


Private Sub OnBudgetLevelComboBoxChange()
    strCompassLevel = BudgetLevelComboBox.SelText
End Sub



Private Sub OnFundCodeComboBoxChange()
    strFundCode = FundCodeComboBox.SelText
End Sub



Private Sub OnAccountCodeComboBoxChange()
    strAccountCode = AccountCodeComboBox.SelText
End Sub


Private Sub OnBocCodeComboBoxChange()
    strBocCode = BocCodeComboBox.SelText
End Sub

