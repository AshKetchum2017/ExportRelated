Option Explicit

' Code-behind UserForm (Name) = ExportRelatedSettings.

Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "ExportRelatedMacro"
Private Const REG_LAST_DIRECTORY_KEY As String = "LastDirectory"
Private Const REG_LAST_EXPORT_FORMAT_KEY As String = "LastExportFormat"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_USENEWUI As Long = &H50
Private Const FILE_DIALOG_SAVE_AS As Long = 1
Private Const SELECT_FOLDER_DUMMY_FILE As String = "Select this folder"

Private isLoadingSettings As Boolean
Private pQueueDraft As ExportSettingItem
Private pQueueResult As ExportSettingItem

' Tanpa pemanggilan ini, cmdSave tetap menjalankan single export existing.
Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo BeginFailed
    operation = "Membuat draft item"
    Set pQueueResult = Nothing
    If item Is Nothing Then
        Set pQueueDraft = New ExportSettingItem
        operation = "Membaca ActiveDocument"
        Set pQueueDraft.SourceDocument = ActiveDocument
        operation = "ExportSettings.LoadFromForm Me"
        pQueueDraft.Settings.LoadFromForm Me
    Else
        operation = "Clone item untuk Modify"
        Set pQueueDraft = item.Clone()
    End If
    operation = "Memeriksa SourceDocument draft"
    If pQueueDraft.SourceDocument Is Nothing Then Err.Raise 5, "ExportRelatedSettings", "Tidak ada dokumen sumber."

    isLoadingSettings = True
    operation = "Mengisi txbDirectory.Text"
    txbDirectory.Text = pQueueDraft.Settings.Directory
    operation = "Mengisi cmbExFormat.Value"
    cmbExFormat.Value = pQueueDraft.Settings.FormatText
    operation = "Mengisi txbPage.Text"
    txbPage.Text = pQueueDraft.Settings.PageText
    operation = "Mengisi txbName.Text"
    txbName.Text = pQueueDraft.Settings.TemplateText
    isLoadingSettings = False
    operation = "EnsureFormatSettings pada draft"
    pQueueDraft.EnsureFormatSettings
    Exit Sub

BeginFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    isLoadingSettings = False
    Err.Raise errorNumber, "ExportRelatedSettings.BeginQueueEdit", _
        operation & vbCrLf & "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Public Property Get QueueResult() As ExportSettingItem
    Set QueueResult = pQueueResult
End Property

Private Sub SaveQueueItem()
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo SaveFailed
    pQueueDraft.Settings.LoadFromForm Me
    pQueueDraft.Settings.FormatText = NormalizeExportFormatText(cmbExFormat.Value)
    pQueueDraft.ValidateForQueue
    SaveCurrentDirectorySetting
    Set pQueueResult = pQueueDraft.Clone()
    Me.Hide
    Exit Sub
SaveFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox "Gagal menyimpan item export (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "Export Queue"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If pQueueDraft Is Nothing Then Exit Sub
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Set pQueueResult = Nothing
        Me.Hide
    End If
End Sub

Private Sub cmdFormatSettings_Click()

    Select Case ResolveExportFormat(Trim$(cmbExFormat.value))
        Case cdrPDF
            ShowFormatSettings "PDFSettings"
        Case cdrDXF
            MsgBox "Pengaturan DXF belum tersedia.", vbInformation, "Export Related"
        Case cdrPNG
            ShowFormatSettings "PNGSettings"
        Case cdrJPEG
            ShowFormatSettings "JPEGSettings"
        Case Else
            MsgBox "Pengaturan format ini belum tersedia.", vbInformation, "Export Related"
    End Select

End Sub

Private Sub ShowFormatSettings(ByVal formName As String)

    Dim presetForm As Object
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo SettingsFailed
    operation = "UserForms.Add(" & formName & ") / UserForm_Initialize"
    Set presetForm = UserForms.Add(formName)

    If presetForm Is Nothing Then
        Err.Raise 91, "ExportRelatedSettings.ShowFormatSettings", "UserForms.Add tidak mengembalikan instance " & formName & "."
    End If

    If Not pQueueDraft Is Nothing Then
        operation = "Menyiapkan snapshot pengaturan " & formName
        pQueueDraft.Settings.FormatText = NormalizeExportFormatText(cmbExFormat.Value)
        pQueueDraft.EnsureFormatSettings
        operation = formName & ".BeginQueueEdit"
        presetForm.BeginQueueEdit pQueueDraft
    End If
    operation = formName & ".Show vbModal"
    presetForm.Show vbModal
    operation = "Unload " & formName
    Unload presetForm
    Exit Sub

SettingsFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If Not presetForm Is Nothing Then Unload presetForm
    On Error GoTo 0
    MsgBox "Gagal membuka pengaturan " & formName & vbCrLf & _
        "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbExclamation, "Export Related"

End Sub

Private Sub txbDirectory_Change()

End Sub

Private Sub txbDirectory_AfterUpdate()
    SaveCurrentDirectorySetting
End Sub

Private Sub cmbExFormat_Change()

    If isLoadingSettings Then Exit Sub
    SaveCurrentExportFormatSetting

End Sub

Private Sub UserForm_Initialize()
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo InitializeFailed
    isLoadingSettings = True
    operation = "PopulateExportFormats / cmbExFormat"
    PopulateExportFormats
    operation = "LoadSavedSettings"
    LoadSavedSettings
    isLoadingSettings = False
    Exit Sub

InitializeFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    isLoadingSettings = False
    ' Tampilkan error asli sebelum COM meneruskannya melalui UserForms.Add.
    MsgBox "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbExclamation, "ExportRelatedSettings - Initialize"
    Err.Raise errorNumber, "ExportRelatedSettings.UserForm_Initialize", _
        operation & vbCrLf & "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Private Sub LoadSavedSettings()

    Dim savedDirectory As String
    Dim savedExportFormat As String
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo LoadFailed
    operation = "GetSetting LastDirectory"
    savedDirectory = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_LAST_DIRECTORY_KEY, "")
    If Len(savedDirectory) > 0 Then
        operation = "Memeriksa directory tersimpan dengan Dir$"
        If Len(Dir$(savedDirectory, vbDirectory)) > 0 Then
            operation = "Mengisi txbDirectory.Text dari LastDirectory"
            txbDirectory.Text = savedDirectory
        End If
    End If

    operation = "GetSetting LastExportFormat"
    savedExportFormat = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_LAST_EXPORT_FORMAT_KEY, "")
    operation = "IsSupportedExportFormatText"
    If IsSupportedExportFormatText(savedExportFormat) Then
        operation = "Mengisi cmbExFormat.Value dari LastExportFormat"
        cmbExFormat.value = NormalizeExportFormatText(savedExportFormat)
    End If
    Exit Sub

LoadFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "ExportRelatedSettings.LoadSavedSettings", _
        operation & vbCrLf & "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Private Sub SaveCurrentDirectorySetting()

    Dim currentDirectory As String
    Dim fso As Object

    ' LastDirectory adalah default untuk Add berikutnya, bukan directory semua item.
    If isLoadingSettings Then Exit Sub
    currentDirectory = Trim$(txbDirectory.Text)
    If Len(currentDirectory) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(currentDirectory) Then
            SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_LAST_DIRECTORY_KEY, currentDirectory
        End If
    End If

End Sub

Private Sub SaveCurrentExportFormatSetting()

    Dim currentExportFormat As String

    If Not pQueueDraft Is Nothing Then Exit Sub
    currentExportFormat = NormalizeExportFormatText(cmbExFormat.value)
    If IsSupportedExportFormatText(currentExportFormat) Then
        SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_LAST_EXPORT_FORMAT_KEY, currentExportFormat
    End If

End Sub

Private Sub PopulateExportFormats()

    cmbExFormat.Clear
    cmbExFormat.AddItem ".pdf"
    cmbExFormat.AddItem ".dxf"
    cmbExFormat.AddItem ".png"
    cmbExFormat.AddItem ".jpg"
    cmbExFormat.ListIndex = 0

End Sub

Private Sub cmdBrowse_Click()

    Dim selectedPath As String

    selectedPath = SelectExportFolder(Trim$(txbDirectory.Text))

    If Len(selectedPath) > 0 And Len(Dir$(selectedPath, vbDirectory)) > 0 Then
        txbDirectory.value = selectedPath
        SaveCurrentDirectorySetting
    End If

End Sub

Private Function SelectExportFolder(ByVal initialPath As String) As String
    Dim selectedPath As String
    Dim fileDialogHandled As Boolean

    If Len(initialPath) > 0 Then
        If Len(Dir$(initialPath, vbDirectory)) = 0 Then
            initialPath = vbNullString
        End If
    End If

    fileDialogHandled = TrySelectFolderWithFileDialog(initialPath, selectedPath)
    If fileDialogHandled Then
        SelectExportFolder = selectedPath
        Exit Function
    End If

    SelectExportFolder = SelectFolderWithBrowseForFolder(initialPath)
End Function

Private Function TrySelectFolderWithFileDialog(ByVal initialPath As String, ByRef selectedFolder As String) As Boolean
    Dim cst As Object
    Dim selectedItem As String

    On Error Resume Next
    Set cst = Application.CorelScriptTools
    If Err.Number <> 0 Or cst Is Nothing Then
        Err.Clear
        On Error GoTo 0
        TrySelectFolderWithFileDialog = False
        Exit Function
    End If

    selectedItem = cst.GetFileBox( _
        "All Files (*.*)|*.*", _
        "Select Export Folder", _
        FILE_DIALOG_SAVE_AS, _
        SELECT_FOLDER_DUMMY_FILE, _
        vbNullString, _
        initialPath, _
        "Select Folder")

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        TrySelectFolderWithFileDialog = False
        Exit Function
    End If
    On Error GoTo 0

    selectedFolder = ResolveFolderFromDialogPath(selectedItem)
    TrySelectFolderWithFileDialog = True
End Function

Private Function SelectFolderWithBrowseForFolder(ByVal initialPath As String) As String
    Dim shellApp As Object
    Dim folder As Object
    Dim rootFolder As Variant
    Dim selectedPath As String

    rootFolder = initialPath
    If Len(initialPath) = 0 Then
        rootFolder = 0
    End If

    Set shellApp = CreateObject("Shell.Application")
    Set folder = shellApp.BrowseForFolder(0, "Select Export Folder", BIF_RETURNONLYFSDIRS Or BIF_USENEWUI, rootFolder)

    If Not folder Is Nothing Then
        selectedPath = folder.Self.Path
        If Len(selectedPath) > 0 And Len(Dir$(selectedPath, vbDirectory)) > 0 Then
            SelectFolderWithBrowseForFolder = selectedPath
        End If
    End If

    Set folder = Nothing
    Set shellApp = Nothing

End Function

Private Function ResolveFolderFromDialogPath(ByVal selectedItem As String) As String
    Dim slashPosition As Long
    Dim selectedPath As String

    selectedPath = Trim$(selectedItem)
    If Len(selectedPath) = 0 Then
        Exit Function
    End If

    If Len(Dir$(selectedPath, vbDirectory)) > 0 Then
        ResolveFolderFromDialogPath = selectedPath
        Exit Function
    End If

    slashPosition = InStrRev(selectedPath, "\")
    If slashPosition <= 1 Then
        Exit Function
    End If

    selectedPath = Left$(selectedPath, slashPosition - 1)
    If Len(Dir$(selectedPath, vbDirectory)) > 0 Then
        ResolveFolderFromDialogPath = selectedPath
    End If
End Function

Private Sub cmdCancel_Click()

    If Not pQueueDraft Is Nothing Then
        Set pQueueResult = Nothing
        Me.Hide
        Exit Sub
    End If
    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    Dim doc As Document
    Dim exportSettings As exportSettings
    Dim exportParser As ExportTemplateParser
    Dim pageParser As ExportPageParser
    Dim exportRunner As exportRunner
    Dim originalPage As Page
    Dim exportDir As String
    Dim exportFormat As cdrFilter
    Dim pageGroups As Collection
    Dim pageGroup As Variant
    Dim exportIndex As Long
    Dim outputCount As Long
    Dim pageNumber As Long
    Dim firstPageNumber As Long
    Dim groupPageIndex As Long
    Dim pageRangeText As String
    Dim outputName As String
    Dim outputPath As String
    Dim templateText As String
    Dim baseName As String
    Dim generatedNames() As String
    Dim exportError As Long
    Dim exportDescription As String

    If Not pQueueDraft Is Nothing Then
        SaveQueueItem
        Exit Sub
    End If

    On Error GoTo ExportFailed

    If ActiveDocument Is Nothing Then
        MsgBox "Tidak ada dokumen aktif.", vbExclamation, "Export Related"
        Exit Sub
    End If

    Set doc = ActiveDocument
    Set originalPage = doc.ActivePage

    Set exportSettings = New exportSettings
    exportSettings.LoadFromForm Me

    exportDir = exportSettings.Directory
    If Len(exportDir) = 0 Then
        MsgBox "Pilih folder tujuan export terlebih dahulu.", vbExclamation, "Export Related"
        Exit Sub
    End If

    If Not exportSettings.IsValidDirectory Then
        MsgBox "Folder tujuan export tidak valid atau tidak ditemukan.", vbExclamation, "Export Related"
        Exit Sub
    End If

    If Right$(exportDir, 1) <> "\" Then
        exportDir = exportDir & "\"
    End If

    Set exportRunner = New exportRunner
    Set exportParser = New ExportTemplateParser
    Set pageParser = New ExportPageParser

    templateText = exportSettings.templateText
    If Not pageParser.ResolvePageGroups(doc, exportSettings.pageText, pageGroups) Then
        MsgBox "Format page di txbPage tidak valid atau kosong.", vbExclamation, "Export Related"
        Exit Sub
    End If

    outputCount = pageGroups.Count
    If Len(templateText) > 0 Then
        If Not exportParser.ValidateNameTemplate(templateText, outputCount, pageParser.SyncGroupCount) Then
            MsgBox "Jumlah value di txbName tidak sesuai dengan jumlah page di txbPage.", vbExclamation, "Export Related"
            Exit Sub
        End If
    End If

    baseName = exportParser.RemoveDocumentExtension(doc.Name)
    If Len(baseName) = 0 Then
        baseName = "Export"
    End If

    exportFormat = ResolveExportFormat(Trim$(cmbExFormat.value))
    ReDim generatedNames(0 To outputCount - 1)

    For exportIndex = 1 To outputCount
        pageGroup = pageGroups(exportIndex)
        If exportFormat <> cdrPDF And pageParser.PageGroupCount(pageGroup) > 1 Then
            MsgBox "Grup multi-page di txbPage hanya didukung untuk export PDF.", vbExclamation, "Export Related"
            Exit Sub
        End If

        For groupPageIndex = LBound(pageGroup) To UBound(pageGroup)
            pageNumber = pageGroup(groupPageIndex)
            If pageNumber < 1 Or pageNumber > doc.pages.Count Then
                MsgBox "Page " & pageNumber & " tidak valid untuk dokumen ini.", vbExclamation, "Export Related"
                Exit Sub
            End If
        Next groupPageIndex

        firstPageNumber = pageParser.FirstPageInGroup(pageGroup)
        generatedNames(exportIndex - 1) = exportParser.BuildExportFileName(baseName, templateText, exportIndex - 1, firstPageNumber, pageParser.SyncGroupIndex(exportIndex) - 1)
    Next exportIndex
    For exportIndex = 1 To outputCount
        pageGroup = pageGroups(exportIndex)
        firstPageNumber = pageParser.FirstPageInGroup(pageGroup)
        pageRangeText = pageParser.PageGroupToRangeText(pageGroup)
        outputName = exportParser.BuildUniqueExportName(generatedNames, outputCount, exportIndex)

        outputPath = exportDir & outputName & GetExportExtension(exportFormat)

        If exportFormat <> cdrPNG And exportFormat <> cdrJPEG And Len(Dir$(outputPath)) > 0 Then
            Kill outputPath
        End If

        doc.pages(firstPageNumber).Activate

        If Not exportRunner.ExportDocument(doc, outputPath, exportFormat, pageRangeText) Then
            exportError = exportRunner.LastErrorNumber
            exportDescription = exportRunner.LastErrorDescription
            If exportFormat = cdrPNG Or exportFormat = cdrJPEG Or Len(Dir$(outputPath)) = 0 Then
                On Error Resume Next
                If Not originalPage Is Nothing Then originalPage.Activate
                On Error GoTo ExportFailed

                MsgBox "Error export " & exportError & vbCrLf & _
                    "Description: [" & exportDescription & "]", _
                    vbCritical, "Export Related"
                Exit Sub
            End If
        End If
    Next exportIndex

    On Error Resume Next
    If Not originalPage Is Nothing Then originalPage.Activate
    On Error GoTo ExportFailed

    MsgBox "Ekspor berhasil dilakukan untuk " & outputCount & " file.", vbInformation, "Export Related"
    Unload Me
    Exit Sub

ExportFailed:
    exportError = Err.Number
    exportDescription = Err.Description

    On Error Resume Next
    If Not originalPage Is Nothing Then originalPage.Activate
    On Error GoTo 0

    MsgBox "Error " & exportError & vbCrLf & _
        "Description: [" & exportDescription & "]", _
        vbCritical, "Export Related"
End Sub


Private Sub cmdHintName_Click()

    NameHintMenu.Show vbModeless
    
End Sub

Private Sub cmdHintPage_Click()

    PageHintMenu.Show vbModeless

End Sub

Private Function ResolveExportFormat(ByVal formatText As String) As cdrFilter

    Dim normalizedText As String

    normalizedText = UCase$(NormalizeExportFormatText(formatText))

    Select Case normalizedText
        Case ".PDF"
            ResolveExportFormat = cdrPDF
        Case ".DXF"
            ResolveExportFormat = cdrDXF
        Case ".PNG"
            ResolveExportFormat = cdrPNG
        Case ".JPG"
            ResolveExportFormat = cdrJPEG
        Case Else
            ResolveExportFormat = cdrPDF
    End Select

End Function

Private Function NormalizeExportFormatText(ByVal formatText As String) As String

    Dim normalizedText As String

    normalizedText = UCase$(Trim$(formatText))
    If Left$(normalizedText, 1) = "." Then
        normalizedText = Mid$(normalizedText, 2)
    End If

    Select Case normalizedText
        Case "PDF", "CDRPDF", "PDF FILE"
            NormalizeExportFormatText = ".pdf"
        Case "DXF", "CDRDXF", "DXF FILE"
            NormalizeExportFormatText = ".dxf"
        Case "PNG", "CDRPNG", "PNG FILE"
            NormalizeExportFormatText = ".png"
        Case "JPG", "CDRJPEG", "JPEG FILE"
            NormalizeExportFormatText = ".jpg"
    End Select

End Function

Private Function IsSupportedExportFormatText(ByVal formatText As String) As Boolean

    IsSupportedExportFormatText = (Len(NormalizeExportFormatText(formatText)) > 0)

End Function

Private Function GetExportExtension(ByVal filterValue As cdrFilter) As String

    Select Case filterValue
        Case cdrPDF
            GetExportExtension = ".pdf"
        Case cdrDXF
            GetExportExtension = ".dxf"
        Case cdrPNG
            GetExportExtension = ".png"
        Case cdrJPEG
            GetExportExtension = ".jpg"
        Case Else
            GetExportExtension = ".pdf"
    End Select

End Function

