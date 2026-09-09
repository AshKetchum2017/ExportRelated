Option Explicit

' Class Module (Name): JPEGPreset.
Public DisplayName As String
Public SourcePath As String
Public SourceXML As String
Public ColorMode As String
Public Antialiased As Boolean
Public EmbedColorProfile As Boolean
Public CropToPage As Boolean
Public Resolution As Double
Public SubFormat As String
Public Quality As Long
Public Blur As Long
Public UseColorProof As Boolean
Public Optimize As Boolean
Public Progressive As Boolean
Public MatteRed As Long
Public MatteGreen As Long
Public MatteBlue As Long

Private Sub Class_Initialize()
    ColorMode = BITMAP_DEFAULT_COLOR_MODE
    Antialiased = BITMAP_DEFAULT_ANTIALIASED
    EmbedColorProfile = BITMAP_DEFAULT_EMBED_PROFILE
    CropToPage = BITMAP_DEFAULT_CROP_TO_PAGE
    Resolution = BITMAP_DEFAULT_RESOLUTION
    SubFormat = "standard"
    Quality = 80
    Blur = 0
    UseColorProof = False
    Optimize = False
    Progressive = False
    MatteRed = BITMAP_DEFAULT_MATTE_COMPONENT
    MatteGreen = BITMAP_DEFAULT_MATTE_COMPONENT
    MatteBlue = BITMAP_DEFAULT_MATTE_COMPONENT
End Sub
