VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database


Public Enum STAT

    '/// <summary>
    '/// The ns
    '/// </summary>
    NS = 0

    '/// <summary>
    '/// Defines the Total
    '/// </summary>
    Total = 1

    '/// <summary>
    '/// Defines the Count
    '/// </summary>
    Count = 2

    '/// <summary>
    '/// Defines the Average
    '/// </summary>
    Average = 3

    '/// <summary>
    '/// Defines the Percentage
    '/// </summary>
    Percentage = 4

    '/// <summary>
    '/// The standard deviation
    '/// </summary>
    StandardDeviation = 5

    '/// <summary>
    '/// The variance
    '/// </summary>
    Variance = 6
End Enum

