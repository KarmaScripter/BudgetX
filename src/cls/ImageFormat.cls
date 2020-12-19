VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Public Enum ImageFormat

    '/// <summary>
    '/// Defines the None
    '/// </summary>
    NS = 0

    '/// <summary>
    '/// The PNG
    '/// </summary>
    PNG

    '/// <summary>
    '/// The JPEG
    '/// </summary>
    JPEG

    '/// <summary>
    '/// The BMP
    '/// </summary>
    BMP

    '/// <summary>
    '/// The icon
    '/// </summary>
    ICO

    '/// <summary>
    '/// The GIF
    '/// </summary>
    GIF

    '/// <summary>
    '/// The RESX
    '/// </summary>
    RESX
    
End Enum

