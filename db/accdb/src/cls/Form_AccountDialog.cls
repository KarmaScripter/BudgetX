VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AccountDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private PRC As AccountArgs
Private pAnd As String
Private m_BudgetLevel As String
Private m_FiscalYear As String
Private m_FundCode As String
Private m_AhCode As String
Private m_OrgCode As String
Private m_BocCode As String
Private m_AccountCode As String
Private m_ActivityCode As String
Private m_RcCode As String


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo ErrorHandler:
    Set PRC = New AccountArgs
    ClearComboBoxValues
    SetButtonVisibility
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub SetComboBoxColors()
    On Error GoTo ErrorHandler:
    Me!BfyComboBox.ForeColor = RGB(255, 255, 255)
    Me!BfyComboBox.BackColor = RGB(33, 33, 33)
    Me!BfyComboBox.BorderColor = RGB(68, 114, 196)
    Me!FundCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!FundCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!FundCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!AhCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!AhCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!AhCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!OrgCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!OrgCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!OrgCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!AccountCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!AccountCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!AccountCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!ActivityCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!ActivityCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!ActivityCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!RcCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!RcCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!RcCodeComboBox.BorderColor = RGB(68, 114, 196)
    Me!BocCodeComboBox.ForeColor = RGB(255, 255, 255)
    Me!BocCodeComboBox.BackColor = RGB(33, 33, 33)
    Me!BocCodeComboBox.BorderColor = RGB(68, 114, 196)
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ClearComboBoxValues()
    On Error GoTo ErrorHandler:
    SetComboBoxColors
    Me.BfyComboBox.Value = ""
    Me.AccountCodeComboBox.Value = ""
    Me.BocCodeComboBox.Value = ""
    Me.RcCodeComboBox.Value = ""
    Me.AhCodeComboBox.Value = ""
    Me.ActivityCodeComboBox.Value = ""
    Me.FundCodeComboBox.Value = ""
    Me.OrgCodeComboBox.Value = ""
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub SetButtonVisibility()
    On Error GoTo ErrorHandler:
    Me.AcceptButton.Visible = Not Me.AcceptButton.Visible
    Me.RefreshButton.Visible = Not Me.RefreshButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Public Sub SetGoButtonVisibility()
    On Error GoTo ErrorHandler:
    If IsNull(PRC.BFY) And _
        IsNull(PRC.FundCode) And _
        Me!GoButton.Visible = False Then
            Me!GoButton.Visible = True
        Else: Me!GoButton.Visible = False
    End If
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub ActivityCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.ActivityCode = Me!ActivityCodeComboBox.SelText
    ActivityCode = PRC.ActivityCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub AhCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.AhCode = Me!AhCodeComboBox.SelText
    AH = PRC.AhCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub OrgCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.OrgCode = Me!OrgCodeComboBox.SelText
    ORG = PRC.OrgCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RcCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.RcCode = Me!RcCodeComboBox.SelText
    rc = PRC.RcCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub



'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BfyComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.BFY = Me!BfyComboBox.SelText
    FiscalYear = PRC.BFY
    Me!AcceptButton.Visible = Not Me!AcceptButton.Visible
    Me!RefreshButton.Visible = Not Me!RefreshButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub BocCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.BocCode = Me!BocCodeComboBox.SelText
    BOC = PRC.BocCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub FundCodeComboBox_AfterUpdate()
    On Error GoTo ErrorHandler:
    PRC.FundCode = Me!FundCodeComboBox.SelText
    FundCode = PRC.FundCode
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


'---------------------------------------------------------------------------------
'   Type:           Procedure
'   Name:
'   Parameters:     Void
'   Retval:         Void
'   Purpose:
'---------------------------------------------------------------------------------
Private Sub RefreshButton_Click()
    On Error GoTo ErrorHandler:
    Me!ExitButton.SetFocus
    ClearComboBoxValues
    Me!GoButton.Visible = Not Me!GoButton.Visible
    Me!RefreshButton.Visible = Not Me!RefreshButton.Visible
ErrorHandler:
    If Err.Number > 0 Then
        m_Error = "Source:   AccountDialog" _
            & vbCrLf & "Member:     SetComboBoxColors()" _
            & vbCrLf & "Descript:   " & Err.Description
        Err.Clear
    End If
    MessageFactory.ShowError (m_Error)
    Exit Sub
End Sub


Public Property Let Level(bl As String)
    If (bl & "") = "" Then
        m_BudgetLevel = bl
    End If
End Property


Public Property Get Level() As String
    If Not (m_BudgetLevel & "") = "" Then
        BudgetLevel = m_BudgetLevel
    End If
End Property

Public Property Let FiscalYear(Code As String)
    If (Code <> "") Then
        m_FiscalYear = Code
    End If
End Property


Public Property Get FiscalYear() As String
    If (m_FiscalYear <> "" And Not IsNull(mFiscalYear)) Then
        FiscalYear = m_FiscalYear
    End If
End Property

Public Property Let Fund(Code As String)
    If (Code <> "") Then
        m_FundCode = Code
    End If
End Property


Public Property Get Fund() As String
    If (m_FundCode <> "" And Not IsNull(mFundCode)) Then
        Fund = m_FundCode
    End If
End Property

Public Property Let BOC(Code As String)
    If (Code <> "") Then
        m_BocCode = Code
    End If
End Property


Public Property Get BOC() As String
    If (m_BocCode <> "" And Not IsNull(mBocCode)) Then
        BOC = m_BocCode
    End If
End Property


Public Property Let rc(Code As String)
    If (Code <> "") Then
        m_RcCode = Code
    End If
End Property


Public Property Get rc() As String
    If (m_RcCode <> "" And Not IsNull(mRcCode)) Then
        rc = m_RcCode
    End If
End Property


Public Property Get Code() As String
    If (m_AccountCode <> "" And Not IsNull(mAccountCode)) Then
        Code = m_AccountCode
    End If
End Property

Public Property Let Code(Code As String)
    If (Code <> "") Then
        m_AccountCode = Code
    End If
End Property


Public Property Let AH(ac As String)
    If Not (ac & "" = "") Then
        m_AhCode = ac
    End If
End Property


Public Property Get AH() As String
    If Not m_AhCode & "" = "" Then
        AH = m_AhCode
    End If
End Property


Public Property Get Activity() As String
    If Not m_ActivityCode & "" = "" Then
        ActivityCode = m_ActivityCode
    End If
End Property

Public Property Let Activity(ac As String)
    If Not ac & "" = "" Then
        m_Activity = ac
    End If
End Property


Public Property Let ORG(Code As String)
    If (Code <> "") Then
        m_OrgCode = Code
    End If
End Property


Public Property Get ORG() As String
    If (m_OrgCode <> "" And Not IsNull(mOrgCode)) Then
        ORG = m_OrgCode
    End If
End Property

