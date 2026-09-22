Option Explicit

' Class Module (Name): PNGPreset. Data hasil parser, belum terhubung ke form.
Public DisplayName As String
Public SourcePath As String
Public SourceXML As String
Public ColorMode As String
Public Transparency As Boolean
Public Antialiased As Boolean
Public EmbedColorProfile As Boolean
Public Interlaced As Boolean
Public CropToPage As Boolean
Public Resolution As Double
Public MatteRed As Long
Public MatteGreen As Long
Public MatteBlue As Long

Private Sub Class_Initialize()
    ' Default yang sama dengan skeleton PNGSettings untuk atribut yang absen.
    ColorMode = BITMAP_DEFAULT_COLOR_MODE
    Transparency = True
    Antialiased = BITMAP_DEFAULT_ANTIALIASED
    EmbedColorProfile = BITMAP_DEFAULT_EMBED_PROFILE
    Interlaced = False
    CropToPage = BITMAP_DEFAULT_CROP_TO_PAGE
    Resolution = BITMAP_DEFAULT_RESOLUTION
    MatteRed = BITMAP_DEFAULT_MATTE_COMPONENT
    MatteGreen = BITMAP_DEFAULT_MATTE_COMPONENT
    MatteBlue = BITMAP_DEFAULT_MATTE_COMPONENT
End Sub
