VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Public Enum SQL

    '/// <summary>
    '/// Defines the None
    '/// </summary>
    NS = 0

    '/// <summary>
    '/// Defines the SELECT
    '/// </summary>
    SELCT

    '/// <summary>
    '/// Defines the INSERT
    '/// </summary>
    Insert

    '/// <summary>
    '/// Defines the UPDATE
    '/// </summary>
    Update

    '/// <summary>
    '/// Defines the DELETE
    '/// </summary>
    Delete

    '/// <summary>
    '/// Defines the CREATE
    '/// </summary>
    Create

    '/// <summary>
    '/// Defines the DROP
    '/// </summary>
    Drop

    '/// <summary>
    '/// Defines the ALTER
    '/// </summary>
    ALTER

    '/// <summary>
    '/// The attach
    '/// </summary>
    ATTACH

    '/// <summary>
    '/// The detach
    '/// </summary>
    DETACH
End Enum

