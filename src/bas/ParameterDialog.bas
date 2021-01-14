Option Compare Database
Option Explicit

'-------------------------------------------------------------------
'---------      Private Backing Fields                --------------
'-------------------------------------------------------------------

Private SelectedYear As String
Private SelectedLevel As String
Private SelectedFund As String
Private SelectedAccount As String
Private SelectedObjectClass As String

'-------------------------------------------------------------------
'---------               Public Properties            --------------
'-------------------------------------------------------------------

Public Property Get FiscalYear()
    If Not IsNull(SelectedYear) And SelectedYear <> "" Then
        FiscalYear = SelectedYear
    End If
End Property

Public Property Get Level()
    If Not IsNull(SelectedLevel) And SelectedLevel <> "" Then
        Level = SelectedLevel
    End If
End Property

Public Property Get Fund()
    If Not IsNull(SelectedFund) And SelectedFund <> "" Then
        Fund = SelectedFund
    End If
End Property


Public Property Get PRC()
    If Not IsNull(SelectedAccount) And SelectedAccount <> "" Then
        PRC = SelectedAccount
    End If
End Property


Public Property Get BOC()
    If Not IsNull(SelectedObjectClass) And SelectedObjectClass <> "" Then
        BOC = SelectedObjectClass
    End If
End Property

'-------------------------------------------------------------------
'---------   Event Handlers for the Query Dialog ComboBoxes         
'-------------------------------------------------------------------

Private Sub OnCloseButtonClick()
    Dim Action As FormAction
    Set Action = New FormAction
    Action.OnCloseButtonClick
End Sub


Private Sub OnExecuteButtonClick()
    '---- Private variable declaration    
End Sub


Private Sub OnFiscalYearComboBoxChange()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = BfyComboBox
    
    If Not IsNull(SelectedYear) And SelectedYear <> "" Then
        SelectedYear = cbo.SelText
    End If
End Sub


Private Sub OnBudgetLevelComboBoxChange()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = BudgetLevelComboBox
    
    If Not IsNull(SelectedLevel) Or SelectedLevel = "" Then
        SelectedLevel = cbo.SelText
    End If
End Sub



Private Sub OnFundCodeComboBoxChange()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = FundCodeComboBox
    
    If Not IsNull(SelectedFund) And SelectedFund <> "" Then
        SelectedFund = cbo.SelText
    End If
End Sub



Private Sub OnAccountCodeComboBoxChange()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = AccountCodeComboBox
    
    If Not IsNull(SelectedAccount) And SelectedAccount <> "" Then
         SelectedAccount = cbo.SelText
    End If
End Sub


Private Sub OnBocCodeComboBoxChange()
    Dim cbo As ComboBox
    Set cbo = New ComboBox
    cbo = BocCodeComboBox
    
    If Not IsNull(SelectedObjectClass) And SelectedObjectClass <> "" Then
        SelectedObjectClass = cbo.SelText
    End If
End Sub