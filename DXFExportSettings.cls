Option Explicit

' Class Module (Name): DXFExportSettings.
' Nilai enum berasal dari FilterDXF.tlb CorelDRAW 2024, bukan cdrUnit/ListIndex.
Public DisplayName As String
Public AutoCADVersion As Long
Public AutoCADUnits As Long
Public BitmapType As Long
Public CurvesAs As String
Public Tolerance As Long
Public TextAsCurves As Boolean
Public FillUnmapped As Boolean
Public FillRed As Long
Public FillGreen As Long
Public FillBlue As Long

Private Sub Class_Initialize()
    DisplayName = "Default"
    AutoCADVersion = 12 ' dxfVersion2007 / R2007
    AutoCADUnits = 4 ' dxfCentimeters
    BitmapType = 0 ' dxfBitmapJPEG
    CurvesAs = "splines"
    Tolerance = 100
    TextAsCurves = True
    FillUnmapped = True
    FillRed = 0
    FillGreen = 0
    FillBlue = 0
End Sub
