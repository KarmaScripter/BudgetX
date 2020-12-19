VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutlayType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Public Enum OutlayType

    '/// <summary>
    '/// Defines the None
    '/// </summary>
    NS = 0

    '/// <summary>
    '/// Defines the Commitment
    '/// </summary>
    Commitment

    '/// <summary>
    '/// Defines the OpenCommitment
    '/// </summary>
    OpenCommitment

    '/// <summary>
    '/// Defines the Obligation
    '/// </summary>
    Obligation

    '/// <summary>
    '/// Defines the ULO
    '/// </summary>
    ULO

    '/// <summary>
    '/// The deobligation
    '/// </summary>
    Deobligation

    '/// <summary>
    '/// Defines the Expenditure
    '/// </summary>
    Expenditure

    '/// <summary>
    '/// Defines the All
    '/// </summary>
    All
    
End Enum

