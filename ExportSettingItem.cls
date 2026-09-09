Option Explicit

' Satu item queue. Summary bukan sumber data; opsi format berupa snapshot per item.
Private pSettings As ExportSettings
Public SourceDocument As Document
Public PNGSettingsXML As String
Public JPEGSettingsXML As String
Public DXFSettingsXML As String
Public PDFPreset As String
Public PDFPresetDisplay As String
Public PDFCompatibility As String
Public Layer1 As Boolean
Public Layer2 As Boolean
Public Layer3 As Boolean
Private pOutputCount As Long

Private Sub Class_Initialize()
    Set pSettings = New ExportSettings
    Layer1 = True
End Sub

Public Property Get Settings() As ExportSettings
    Set Settings = pSettings
End Property

Public Property Get OutputCount() As Long
    OutputCount = pOutputCount
End Property

Public Function Clone() As ExportSettingItem
    Dim result As ExportSettingItem
    Set result = New ExportSettingItem
    result.CopyFrom Me
    Set Clone = result
End Function

Public Sub CopyFrom(ByVal source As ExportSettingItem)
    pSettings.Directory = source.Settings.Directory
    pSettings.FormatText = source.Settings.FormatText
    pSettings.PageText = source.Settings.PageText
    pSettings.TemplateText = source.Settings.TemplateText
    Set SourceDocument = source.SourceDocument
    PNGSettingsXML = source.PNGSettingsXML
    JPEGSettingsXML = source.JPEGSettingsXML
    DXFSettingsXML = source.DXFSettingsXML
    PDFPreset = source.PDFPreset
    PDFPresetDisplay = source.PDFPresetDisplay
    PDFCompatibility = source.PDFCompatibility
    Layer1 = source.Layer1
    Layer2 = source.Layer2
    Layer3 = source.Layer3
    pOutputCount = source.OutputCount
End Sub

' Registry hanya dibaca sebagai default ketika format belum memiliki snapshot.
Public Sub EnsureFormatSettings()
    Dim pngStore As PNGSettingsStore
    Dim png As PNGPreset
    Dim jpegStore As JPEGSettingsStore
    Dim jpeg As JPEGPreset
    Dim dxfStore As DXFSettingsStore
    Dim dxf As DXFExportSettings

    Select Case pSettings.FormatText
        Case ".png"
            If Len(PNGSettingsXML) > 0 Then Exit Sub
            Set pngStore = New PNGSettingsStore
            Set png = pngStore.LoadSettings()
            If png Is Nothing Then Set png = New PNGPreset
            PNGSettingsXML = pngStore.SerializeSettings(png)
        Case ".jpg"
            If Len(JPEGSettingsXML) > 0 Then Exit Sub
            Set jpegStore = New JPEGSettingsStore
            Set jpeg = jpegStore.LoadSettings()
            If jpeg Is Nothing Then Set jpeg = New JPEGPreset
            JPEGSettingsXML = jpegStore.SerializeSettings(jpeg)
        Case ".pdf"
            If Len(PDFPreset) = 0 Then
                PDFPreset = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFPreset", "Used by Rin's Macro")
                If Len(Trim$(PDFPreset)) = 0 Then PDFPreset = "Used by Rin's Macro"
                PDFPresetDisplay = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFPresetDisplay", PDFPreset)
                PDFCompatibility = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFCompatibility", "Acrobat 9.0")
            End If
        Case ".dxf"
            If Len(DXFSettingsXML) > 0 Then Exit Sub
            Set dxfStore = New DXFSettingsStore
            Set dxf = dxfStore.LoadSettings()
            If dxf Is Nothing Then Set dxf = New DXFExportSettings
            DXFSettingsXML = dxfStore.SerializeSettings(dxf)
        Case Else
            Err.Raise 5, "ExportSettingItem", "Pilih format PDF, DXF, PNG, atau JPG."
    End Select
End Sub

Public Property Get PresetDisplay() As String
    Dim xml As Object
    Select Case pSettings.FormatText
        Case ".pdf"
            PresetDisplay = PDFPresetDisplay
            If Len(PresetDisplay) = 0 Then PresetDisplay = PDFPreset
        Case ".png", ".jpg"
            If pSettings.FormatText = ".png" Then
                Set xml = ParseExportPresetXML(PNGSettingsXML, "ExportSettingItem")
            Else
                Set xml = ParseExportPresetXML(JPEGSettingsXML, "ExportSettingItem")
            End If
            PresetDisplay = ExportPresetText(xml.documentElement, "name", "Default")
        Case ".dxf"
            Set xml = ParseExportPresetXML(DXFSettingsXML, "ExportSettingItem")
            PresetDisplay = ExportPresetText(xml.documentElement, "name", "Default")
    End Select
    If Len(PresetDisplay) = 0 Then PresetDisplay = "Default"
End Property

Public Property Get Summary() As String
    Summary = UCase$(Mid$(pSettings.FormatText, 2)) & " | " & PresetDisplay & " | " & CStr(pOutputCount)
End Property

' Export queue wajib memakai snapshot; jangan diam-diam mengambil default registry.
Public Function ReadPNGSnapshot() As PNGPreset
    Dim reader As PNGPresetReader
    Dim store As PNGSettingsStore
    Dim snapshot As PNGPreset
    If Len(PNGSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot PNG kosong. Buka Modify lalu Save."
    Set reader = New PNGPresetReader
    Set store = New PNGSettingsStore
    Set snapshot = reader.ParseXML(PNGSettingsXML)
    If snapshot Is Nothing Then Err.Raise 5, "ExportSettingItem", "Snapshot item bukan PNG."
    store.ValidateSettings snapshot
    Set ReadPNGSnapshot = snapshot
End Function

Public Function ReadJPEGSnapshot() As JPEGPreset
    Dim reader As JPEGPresetReader
    Dim store As JPEGSettingsStore
    Dim snapshot As JPEGPreset
    If Len(JPEGSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot JPEG kosong. Buka Modify lalu Save."
    Set reader = New JPEGPresetReader
    Set store = New JPEGSettingsStore
    Set snapshot = reader.ParseXML(JPEGSettingsXML)
    If snapshot Is Nothing Then Err.Raise 5, "ExportSettingItem", "Snapshot item bukan JPEG."
    store.ValidateSettings snapshot
    Set ReadJPEGSnapshot = snapshot
End Function

' Snapshot DXF wajib tersedia dan valid sebelum export queue.
Public Function ReadDXFSnapshot() As DXFExportSettings
    Dim store As DXFSettingsStore
    If Len(DXFSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot DXF kosong. Buka Modify lalu Save."
    Set store = New DXFSettingsStore
    Set ReadDXFSnapshot = store.ParseXML(DXFSettingsXML)
End Function

' Hitung file memakai parser existing, bukan menghitung page atau memecah summary.
Public Sub ValidateForQueue()
    Dim pageParser As ExportPageParser
    Dim nameParser As ExportTemplateParser
    Dim groups As Collection
    Dim group As Variant
    Dim groupIndex As Long
    Dim pageIndex As Long
    Dim fso As Object

    If SourceDocument Is Nothing Then Err.Raise 5, "ExportSettingItem", "Tidak ada dokumen sumber."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pSettings.Directory) Then Err.Raise 5, "ExportSettingItem", "Pilih folder export yang valid."
    EnsureFormatSettings
    ' Input kosong menunjuk page saat item disimpan, bukan page aktif saat export nanti.
    If Len(pSettings.PageText) = 0 Then pSettings.PageText = CStr(SourceDocument.ActivePage.Index)
    Set pageParser = New ExportPageParser
    If Not pageParser.ResolvePageGroups(SourceDocument, pSettings.PageText, groups) Then
        Err.Raise 5, "ExportSettingItem", "Format page tidak valid atau kosong."
    End If
    For groupIndex = 1 To groups.Count
        group = groups.Item(groupIndex)
        If pSettings.FormatText <> ".pdf" And pageParser.PageGroupCount(group) > 1 Then
            Err.Raise 5, "ExportSettingItem", "Grup multi-page hanya didukung untuk PDF."
        End If
        For pageIndex = LBound(group) To UBound(group)
            If group(pageIndex) < 1 Or group(pageIndex) > SourceDocument.Pages.Count Then
                Err.Raise 5, "ExportSettingItem", "Page " & CStr(group(pageIndex)) & " tidak tersedia pada dokumen sumber."
            End If
        Next pageIndex
    Next groupIndex
    Set nameParser = New ExportTemplateParser
    If Len(pSettings.TemplateText) > 0 Then
        If Not nameParser.ValidateNameTemplate(pSettings.TemplateText, groups.Count, pageParser.SyncGroupCount) Then
            Err.Raise 5, "ExportSettingItem", "Jumlah value nama tidak sesuai dengan grup page."
        End If
    End If
    pOutputCount = groups.Count
End Sub
