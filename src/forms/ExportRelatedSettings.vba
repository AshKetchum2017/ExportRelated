Option Explicit

' Code-behind UserForm (Name) = ExportRelatedSettings.

Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "ExportRelatedMacro"
Private Const REG_LAST_DIRECTORY_KEY As String = "LastDirectory"
Private Const REG_PARENT_DIRECTORY_KEY As String = "UseParentDirectory"
Private Const REG_LAST_EXPORT_FORMAT_KEY As String = "LastExportFormat"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_USENEWUI As Long = &H50
Private Const FILE_DIALOG_SAVE_AS As Long = 1
Private Const SELECT_FOLDER_DUMMY_FILE As String = "Select this folder"

Private isLoadingSettings As Boolean
Private pQueueDraft As ExportSettingItem
Private pQueueResult As ExportSettingItem
Private pManualDirectory As String
Private pMRBehaviorObserver As Object
Private pMRAction As Boolean
Private pHintReady As Boolean
Private pUpdatingHint As Boolean
Private pHintActive(1 To 3) As Boolean
Private pInputColor(1 To 3) As Long
Private pFocusedInput As String

' Tanpa pemanggilan ini, cmdSave tetap menjalankan single export existing.
Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim sourceDocument As Document

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

    operation = "Resolve SourceDocument draft"
    Set sourceDocument = pQueueDraft.ResolveSourceDocument()

    If sourceDocument Is Nothing Then
        If Len(pQueueDraft.SourceDocumentPath) > 0 Then
            Err.Raise 5, "ExportRelatedSettings", _
                "Dokumen sumber item tidak sedang terbuka." & vbCrLf & _
                "Source: " & pQueueDraft.SourceDocumentPath
        Else
            Err.Raise 5, "ExportRelatedSettings", _
                "Dokumen sumber item tidak tersedia dan tidak memiliki path yang dapat dipulihkan."
        End If
    End If

    isLoadingSettings = True
    operation = "Mengisi txbDirectory.Text"
    ERSetInput txbDirectory, pQueueDraft.Settings.Directory
    If Not pQueueDraft.Settings.UseParentDirectory Then pManualDirectory = ERInputText(txbDirectory)
    chkParentDirectory.Value = pQueueDraft.Settings.UseParentDirectory
    UpdateParentDirectoryControls
    operation = "Mengisi cmbExFormat.Value"
    cmbExFormat.Value = pQueueDraft.Settings.FormatText
    operation = "Mengisi txbPage.Text"
    ERSetInput txbPage, pQueueDraft.Settings.PageText
    operation = "Mengisi txbName.Text"
    ERSetInput txbName, pQueueDraft.Settings.TemplateText
    isLoadingSettings = False
    operation = "EnsureFormatSettings pada draft"
    pQueueDraft.EnsureFormatSettings
    operation = "Memperbarui preview nama file"
    UpdateFilePreview
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
    Dim isDayDirectory As Boolean

    On Error GoTo SaveFailed
    pQueueDraft.Settings.LoadFromForm Me
    pQueueDraft.Settings.FormatText = NormalizeExportFormatText(cmbExFormat.Value)
    pQueueDraft.Settings.Directory = ResolveSaveDirectory(pQueueDraft.Settings, pQueueDraft.SourceDocument, isDayDirectory)
    pQueueDraft.ValidateForQueue Not isDayDirectory
    If isDayDirectory Then EnsureDayDirectory pQueueDraft.Settings.Directory
    ERSetInput txbDirectory, pQueueDraft.Settings.Directory
    If isDayDirectory Then pManualDirectory = ERInputText(txbDirectory)
    SaveCurrentDirectorySetting
    SaveParentDirectorySetting
    Set pQueueResult = pQueueDraft.Clone()
    Me.Hide
    If Not pMRBehaviorObserver Is Nothing And Not pMRAction Then _
        CallByName pMRBehaviorObserver, "EditorFinished", VbMethod, True
    Exit Sub
SaveFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    If Not pMRBehaviorObserver Is Nothing Then
        If pMRAction Then Err.Raise errorNumber, "ExportRelatedSettings.cmdSave", errorDescription
        CallByName pMRBehaviorObserver, "EditorFailed", VbMethod, errorNumber, errorDescription
        Exit Sub
    End If
    MsgBox "Gagal menyimpan item export (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "Export Queue"
End Sub

Private Function ResolveSaveDirectory(ByVal settings As ExportSettings, ByVal doc As Document, _
                                      ByRef isDayDirectory As Boolean) As String
    Dim fso As Object
    Dim lastDirectory As String
    Dim parentDirectory As String
    Dim dayName As String

    isDayDirectory = False
    If settings.UseParentDirectory Or settings.Directory <> "/*dd" Then
        ResolveSaveDirectory = settings.ResolveDirectory(doc)
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    lastDirectory = Trim$(GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_LAST_DIRECTORY_KEY, ""))
    If Len(lastDirectory) = 0 Then
        Err.Raise 5, "ExportRelatedSettings", "Token /*dd memerlukan LastDirectory yang berakhir dengan folder hari, misalnya 07."
    End If
    Do While Right$(lastDirectory, 1) = "\"
        lastDirectory = Left$(lastDirectory, Len(lastDirectory) - 1)
    Loop
    dayName = fso.GetFileName(lastDirectory)
    If Len(dayName) <> 2 Or Not dayName Like "##" Then
        Err.Raise 5, "ExportRelatedSettings", "Folder terakhir LastDirectory harus berupa hari dua digit (01-31): " & lastDirectory
    End If
    If CLng(dayName) < 1 Or CLng(dayName) > 31 Then
        Err.Raise 5, "ExportRelatedSettings", "Folder hari LastDirectory harus berada pada 01-31: " & lastDirectory
    End If
    parentDirectory = fso.GetParentFolderName(lastDirectory)
    If Not fso.FolderExists(parentDirectory) Then
        Err.Raise 76, "ExportRelatedSettings", "Folder induk LastDirectory tidak ditemukan: " & parentDirectory
    End If
    ResolveSaveDirectory = fso.BuildPath(parentDirectory, Format$(Date, "dd"))
    isDayDirectory = True
End Function

Private Sub EnsureDayDirectory(ByVal directory As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(directory) Then fso.CreateFolder directory
End Sub

Private Sub chkParentDirectory_Click()
    If isLoadingSettings Then Exit Sub
    If CBool(chkParentDirectory.Value) Then pManualDirectory = ERInputText(txbDirectory)
    UpdateParentDirectoryControls
End Sub

Private Sub UpdateParentDirectoryControls()
    Dim doc As Document
    Dim settings As ExportSettings
    txbDirectory.Enabled = Not CBool(chkParentDirectory.Value)
    cmdBrowse.Enabled = Not CBool(chkParentDirectory.Value)
    If CBool(chkParentDirectory.Value) Then
        If Not pQueueDraft Is Nothing Then
            Set doc = pQueueDraft.SourceDocument
        ElseIf Application.Documents.Count > 0 Then
            Set doc = ActiveDocument
        End If
        Set settings = New ExportSettings
        ERSetInput txbDirectory, settings.DocumentDirectory(doc)
    Else
        ERSetInput txbDirectory, pManualDirectory
    End If
End Sub

Private Sub SaveParentDirectorySetting()
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_PARENT_DIRECTORY_KEY, CStr(CBool(chkParentDirectory.Value))
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If pQueueDraft Is Nothing Then Exit Sub
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Set pQueueResult = Nothing
        Me.Hide
        If Not pMRBehaviorObserver Is Nothing Then CallByName pMRBehaviorObserver, "EditorFinished", VbMethod, False
    End If
End Sub

Private Sub cmdFormatSettings_Click()

    Select Case ResolveExportFormat(Trim$(cmbExFormat.value))
        Case cdrPDF
            ShowFormatSettings "PDFSettings"
        Case cdrDXF
            ShowFormatSettings "DXFSettings"
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

' Hint hanya presentasi. Semua pembacaan settings memakai nilai asli melalui ERInputText.
Private Function ERInputIndex(ByVal box As MSForms.TextBox) As Long
    Select Case box.Name
        Case "txbDirectory": ERInputIndex = 1
        Case "txbName": ERInputIndex = 2
        Case "txbPage": ERInputIndex = 3
    End Select
End Function

Public Function ERInputText(ByVal box As MSForms.TextBox) As String
    If pHintReady Then
        If pHintActive(ERInputIndex(box)) Then Exit Function
    End If
    ERInputText = box.Text
End Function

Private Sub ERShowHint(ByVal box As MSForms.TextBox)
    Dim hint As String
    Dim index As Long
    If Not pHintReady Then Exit Sub
    If Not box.Enabled Then Exit Sub
    If pFocusedInput = box.Name Then Exit Sub
    If Len(box.Text) > 0 Then Exit Sub
    index = ERInputIndex(box)
    Select Case index
        Case 1: hint = "Enter export folder path"
        Case 2: hint = "Current document name"
        Case 3: hint = "Current page"
    End Select
    pUpdatingHint = True
    pHintActive(index) = True
    box.ForeColor = RGB(128, 128, 128)
    box.Text = hint
    pUpdatingHint = False
End Sub

' Assignment dari Browse/Modify/MB tetap input nyata meskipun teksnya sama dengan hint.
Private Sub ERSetInput(ByVal box As MSForms.TextBox, ByVal value As String)
    Dim index As Long
    index = ERInputIndex(box)
    pUpdatingHint = True
    pHintActive(index) = False
    If pHintReady Then box.ForeColor = pInputColor(index)
    box.Text = value
    pUpdatingHint = False
    ERShowHint box
    If Not isLoadingSettings And index <> 1 Then UpdateFilePreview
End Sub

Private Sub ERInputChanged(ByVal box As MSForms.TextBox)
    Dim index As Long
    If pUpdatingHint Or Not pHintReady Then Exit Sub
    index = ERInputIndex(box)
    pHintActive(index) = False
    box.ForeColor = pInputColor(index)
    ERShowHint box
End Sub

Private Sub EREnterInput(ByVal box As MSForms.TextBox)
    Dim index As Long
    If Not pHintReady Then Exit Sub
    pFocusedInput = box.Name
    index = ERInputIndex(box)
    pUpdatingHint = True
    If pHintActive(index) Then box.Text = vbNullString
    pHintActive(index) = False
    box.ForeColor = pInputColor(index)
    pUpdatingHint = False
End Sub

Private Sub ERExitInput(ByVal box As MSForms.TextBox)
    pFocusedInput = vbNullString
    ERShowHint box
End Sub

Private Sub txbDirectory_Enter()
    EREnterInput txbDirectory
End Sub

Private Sub txbDirectory_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ERExitInput txbDirectory
End Sub

Private Sub txbName_Enter()
    EREnterInput txbName
End Sub

Private Sub txbName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ERExitInput txbName
End Sub

Private Sub txbPage_Enter()
    EREnterInput txbPage
End Sub

Private Sub txbPage_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ERExitInput txbPage
End Sub

Private Sub txbDirectory_Change()
    ERInputChanged txbDirectory
End Sub

Private Sub txbDirectory_AfterUpdate()
    SaveCurrentDirectorySetting
End Sub

Private Sub txbName_Change()
    If pUpdatingHint Then Exit Sub
    ERInputChanged txbName
    If isLoadingSettings Then Exit Sub
    UpdateFilePreview
End Sub

Private Sub txbPage_Change()
    If pUpdatingHint Then Exit Sub
    ERInputChanged txbPage
    If isLoadingSettings Then Exit Sub
    UpdateFilePreview
End Sub

Private Sub cmbExFormat_Change()
    If isLoadingSettings Then Exit Sub
    SaveCurrentExportFormatSetting
    UpdateFilePreview
End Sub

Private Sub UpdateFilePreview()
    Dim doc As Document
    Dim tokenPlan As ExportTokenPlan
    Dim nameParser As ExportTemplateParser
    Dim pageParser As ExportPageParser
    Dim groups As Collection
    Dim group As Variant
    Dim generatedNames() As String
    Dim exportFormat As String
    Dim index As Long

    On Error GoTo PreviewUnavailable
    txbPreview.Value = vbNullString
    exportFormat = NormalizeExportFormatText(CStr(cmbExFormat.Value))
    If Len(exportFormat) = 0 Then Exit Sub

    If Not pQueueDraft Is Nothing Then
        Set doc = pQueueDraft.ResolveSourceDocument()
    ElseIf Application.Documents.Count > 0 Then
        Set doc = ActiveDocument
    End If
    If doc Is Nothing Then Exit Sub

    Set tokenPlan = New ExportTokenPlan
    ' Page kosong mengikuti page aktif dokumen sumber, sama seperti alur export.
    tokenPlan.Prepare doc, ERInputText(txbPage), ERInputText(txbName)
    Set groups = tokenPlan.Groups
    Set pageParser = New ExportPageParser
    If exportFormat <> ".pdf" Then
        For Each group In groups
            If pageParser.PageGroupCount(group) > 1 Then Exit Sub
        Next group
    End If

    ReDim generatedNames(0 To groups.Count - 1)
    ' Nama pertama dapat memperoleh (1) bila ada output lain bernama sama.
    For index = 1 To groups.Count
        generatedNames(index - 1) = tokenPlan.NameAt(index)
    Next index
    Set nameParser = New ExportTemplateParser
    txbPreview.Value = nameParser.BuildUniqueExportName(generatedNames, groups.Count, 1) & exportFormat
    Exit Sub

PreviewUnavailable:
    txbPreview.Value = vbNullString
End Sub

Private Sub UserForm_Initialize()
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo InitializeFailed
    isLoadingSettings = True
    operation = "Menyiapkan hint TextBox"
    pInputColor(1) = txbDirectory.ForeColor
    pInputColor(2) = txbName.ForeColor
    pInputColor(3) = txbPage.ForeColor
    pHintReady = True
    operation = "Mengunci txbPreview"
    txbPreview.Enabled = True
    txbPreview.Locked = True
    operation = "PopulateExportFormats / cmbExFormat"
    PopulateExportFormats
    operation = "LoadSavedSettings"
    LoadSavedSettings
    ERShowHint txbDirectory
    ERShowHint txbName
    ERShowHint txbPage
    isLoadingSettings = False
    operation = "Memperbarui preview nama file"
    UpdateFilePreview
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
            ERSetInput txbDirectory, savedDirectory
        End If
    End If

    operation = "GetSetting LastExportFormat"
    savedExportFormat = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_LAST_EXPORT_FORMAT_KEY, "")
    operation = "IsSupportedExportFormatText"
    If IsSupportedExportFormatText(savedExportFormat) Then
        operation = "Mengisi cmbExFormat.Value dari LastExportFormat"
        cmbExFormat.value = NormalizeExportFormatText(savedExportFormat)
    End If
    operation = "GetSetting UseParentDirectory"
    chkParentDirectory.Value = (StrComp(GetSetting(REG_APP_NAME, REG_SECTION_NAME, _
        REG_PARENT_DIRECTORY_KEY, "False"), "True", vbTextCompare) = 0)
    pManualDirectory = ERInputText(txbDirectory)
    UpdateParentDirectoryControls
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
    If CBool(chkParentDirectory.Value) Then Exit Sub
    currentDirectory = Trim$(ERInputText(txbDirectory))
    If currentDirectory = "/*dd" Then Exit Sub
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

    If CBool(chkParentDirectory.Value) Then Exit Sub
    selectedPath = SelectExportFolder(Trim$(ERInputText(txbDirectory)))

    If Len(selectedPath) > 0 And Len(Dir$(selectedPath, vbDirectory)) > 0 Then
        ERSetInput txbDirectory, selectedPath
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
        If Not pMRBehaviorObserver Is Nothing And Not pMRAction Then _
            CallByName pMRBehaviorObserver, "EditorFinished", VbMethod, False
        Exit Sub
    End If
    Unload Me
    
End Sub

Public Sub MRSetBehaviorObserver(ByVal observer As Object)
    Set pMRBehaviorObserver = observer
End Sub

Public Sub MRDetachBehavior()
    Set pMRBehaviorObserver = Nothing
End Sub

Public Sub MRBehaviorValue(ByVal target As String, ByVal value As Variant)
    Select Case LCase$(target)
        Case "txbname": ERSetInput txbName, CStr(value)
        Case "txbpage": ERSetInput txbPage, CStr(value)
        Case "cmbexformat": cmbExFormat.Value = LCase$(CStr(value))
        Case "chkparentdirectory"
            chkParentDirectory.Value = CBool(value)
            UpdateParentDirectoryControls
        Case "txbdirectory"
            If CBool(chkParentDirectory.Value) Then Err.Raise 5, , "Nonaktifkan chkParentDirectory sebelum mengisi txbDirectory."
            ERSetInput txbDirectory, CStr(value)
            pManualDirectory = CStr(value)
        Case Else: Err.Raise 5, , "Target settings tidak terdaftar: " & target
    End Select
End Sub

Public Sub MRBehaviorAction(ByVal action As String)
    Dim number As Long, description As String
    On Error GoTo Failed
    If pQueueDraft Is Nothing Then Err.Raise 5, , "Settings belum dibuka dalam konteks antrean."
    pMRAction = True
    Select Case LCase$(action)
        Case "cmdsave": cmdSave_Click
        Case "cmdcancel": cmdCancel_Click
        Case Else: Err.Raise 5, , "Action settings tidak terdaftar: " & action
    End Select
    pMRAction = False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    pMRAction = False
    Err.Raise number, "ExportRelatedSettings.MRBehaviorAction", description
End Sub

Public Function MRCreateBehaviorFormatEditor() As Object
    Dim formatForm As Object
    Dim formName As String
    Dim operation As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    If pQueueDraft Is Nothing Then
        Err.Raise 5, , "Settings belum dibuka dalam konteks antrean."
    End If

    operation = "Menentukan format settings"

    Select Case ResolveExportFormat(Trim$(cmbExFormat.Value))
        Case cdrJPEG
            formName = "JPEGSettings"
        Case Else
            Err.Raise 5, , _
                "MacroBehavior Format Settings saat ini baru mendukung JPEG."
    End Select

    operation = "UserForms.Add(" & formName & ")"
    Set formatForm = UserForms.Add(formName)

    If formatForm Is Nothing Then
        Err.Raise 91, , _
            "UserForms.Add tidak mengembalikan instance " & formName & "."
    End If

    operation = "Menyiapkan snapshot " & formName
    pQueueDraft.Settings.FormatText = _
        NormalizeExportFormatText(cmbExFormat.Value)

    pQueueDraft.EnsureFormatSettings

    operation = formName & ".BeginQueueEdit"
    formatForm.BeginQueueEdit pQueueDraft

    Set MRCreateBehaviorFormatEditor = formatForm
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    If Not formatForm Is Nothing Then Unload formatForm
    On Error GoTo 0

    Err.Raise errorNumber, _
        "ExportRelatedSettings.MRCreateBehaviorFormatEditor", _
        operation & vbCrLf & errorDescription
End Function

Private Sub cmdSave_Click()

    Dim doc As Document
    Dim exportSettings As exportSettings
    Dim exportParser As ExportTemplateParser
    Dim pageParser As ExportPageParser
    Dim tokenPlan As ExportTokenPlan
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
    Dim generatedNames() As String
    Dim exportError As Long
    Dim exportDescription As String
    Dim isDayDirectory As Boolean

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

    exportSettings.Directory = ResolveSaveDirectory(exportSettings, doc, isDayDirectory)
    exportDir = exportSettings.Directory
    If Len(exportDir) = 0 Then
        MsgBox "Pilih folder tujuan export terlebih dahulu.", vbExclamation, "Export Related"
        Exit Sub
    End If

    If Not isDayDirectory Then
        If Not exportSettings.IsValidDirectory Then
            MsgBox "Folder tujuan export tidak valid atau tidak ditemukan.", vbExclamation, "Export Related"
            Exit Sub
        End If
    End If

    If Right$(exportDir, 1) <> "\" Then
        exportDir = exportDir & "\"
    End If

    Set exportRunner = New exportRunner
    Set exportParser = New ExportTemplateParser
    Set pageParser = New ExportPageParser
    Set tokenPlan = New ExportTokenPlan

    templateText = exportSettings.templateText
    tokenPlan.Prepare doc, exportSettings.pageText, templateText
    Set pageGroups = tokenPlan.Groups

    outputCount = pageGroups.Count

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
        generatedNames(exportIndex - 1) = tokenPlan.NameAt(exportIndex)
    Next exportIndex
    If isDayDirectory Then EnsureDayDirectory exportSettings.Directory
    ERSetInput txbDirectory, exportSettings.Directory
    If isDayDirectory Then pManualDirectory = ERInputText(txbDirectory)
    SaveCurrentDirectorySetting
    SaveParentDirectorySetting
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
            If exportFormat = cdrPNG Or exportFormat = cdrJPEG Or exportFormat = cdrDXF Or Len(Dir$(outputPath)) = 0 Then
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

    If Len(tokenPlan.WarningDescription) > 0 Then
        MsgBox "Ekspor selesai." & vbCrLf & tokenPlan.WarningDescription, vbExclamation, "Export Related"
    Else
        MsgBox "Ekspor berhasil dilakukan untuk " & outputCount & " file.", vbInformation, "Export Related"
    End If
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
