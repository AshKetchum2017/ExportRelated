Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


' Code-behind UserForm (Name) = ExportRelatedMenu.
' lbxSettingLists: MSForms.ListBox, satu baris = pItems.Item(ListIndex + 1).
Private pItems As Collection
Private pList As MSForms.ListBox
Private pBinding As Boolean
Private pExporting As Boolean
Private pMRBehavior As Boolean
Private pMRExportAction As Boolean

Private Sub cmdClearLists_Click()
    If pMRBehavior Then Exit Sub
    If pExporting Then Exit Sub
    pBinding = True
    pList.Clear
    Set pItems = New Collection
    pBinding = False
    SyncSelectedLayers
End Sub

Private Sub UserForm_Initialize()
    Dim listControl As Object
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo InitializeFailed
    operation = "Membuat Collection item queue"
    Set pItems = New Collection
    operation = "Mengakses kontrol Me.lbxSettingLists"
    Set listControl = Me.lbxSettingLists
    operation = "Memeriksa tipe lbxSettingLists: " & TypeName(listControl)
    If Not TypeOf listControl Is MSForms.ListBox Then
        Err.Raise 13, "ExportRelatedMenu", "lbxSettingLists bertipe " & TypeName(listControl) & _
            "; dibutuhkan MSForms.ListBox. Ganti kontrolnya menjadi ListBox dengan (Name) tetap lbxSettingLists."
    End If
    operation = "Menghubungkan lbxSettingLists ke MSForms.ListBox"
    Set pList = listControl
    operation = "Mengosongkan lbxSettingLists"
    pList.Clear
    operation = "Mengatur ColumnCount lbxSettingLists"
    pList.ColumnCount = 1
    operation = "Mengatur MultiSelect lbxSettingLists"
    pList.MultiSelect = fmMultiSelectSingle
    operation = "Menyinkronkan checkbox layer dan tombol queue"
    SyncSelectedLayers
    Exit Sub

InitializeFailed:
    errorNumber = Err.Number
    errorSource = Err.source
    errorDescription = Err.Description
    ' Tampilkan error asli sebelum diteruskan melewati proses Load UserForm.
    MsgBox "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbCritical, "Export Queue - Initialize"
    Err.Raise errorNumber, "ExportRelatedMenu.UserForm_Initialize", _
        "Operasi: " & operation & vbCrLf & "Error " & CStr(errorNumber) & ": " & errorDescription
End Sub

Private Sub cmdAddSetting_Click()
    If pMRBehavior Then Exit Sub
    If pExporting Then Exit Sub
    EditQueueItem -1
End Sub

Private Sub cmdModify_Click()
    If pMRBehavior Then Exit Sub
    If pExporting Then Exit Sub
    If pList.ListIndex < 0 Then Exit Sub
    EditQueueItem pList.ListIndex
End Sub

Private Sub EditQueueItem(ByVal rowIndex As Long)
    Dim editor As Object
    Dim source As ExportSettingItem
    Dim result As ExportSettingItem
    Dim rowText As String
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo EditFailed
    operation = "Membaca item terpilih"
    If rowIndex >= 0 Then Set source = pItems.item(rowIndex + 1)
    operation = "UserForms.Add(ExportRelatedSettings) / Initialize"
    Set editor = UserForms.Add("ExportRelatedSettings")
    operation = "BeginQueueEdit pada " & TypeName(editor)
    editor.BeginQueueEdit source
    operation = "Show vbModal pada " & TypeName(editor)
    editor.Show vbModal
    operation = "Membaca QueueResult pada " & TypeName(editor)
    Set result = editor.QueueResult
    operation = "Unload editor"
    Unload editor
    Set editor = Nothing
    If result Is Nothing Then Exit Sub

    operation = "Membaca summary item"
    rowText = result.Summary
    operation = "Memperbarui data item dan row list"
    pBinding = True
    If rowIndex < 0 Then
        pItems.Add result
        pList.AddItem rowText
        pList.ListIndex = pItems.Count - 1
    Else
        source.CopyFrom result
        pList.List(rowIndex, 0) = rowText
        pList.ListIndex = rowIndex
    End If
    pBinding = False
    operation = "SyncSelectedLayers setelah edit"
    SyncSelectedLayers
    Exit Sub

EditFailed:
    errorNumber = Err.Number
    errorSource = Err.source
    errorDescription = Err.Description
    On Error Resume Next
    If Not editor Is Nothing Then Unload editor
    pBinding = False
    SyncSelectedLayers
    On Error GoTo 0
    MsgBox "Gagal mengedit item export." & vbCrLf & _
        "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbExclamation, "Export Queue"
End Sub

Private Sub cmdRemoveSetting_Click()
    If pMRBehavior Then Exit Sub
    Dim rowIndex As Long
    If pExporting Then Exit Sub
    rowIndex = pList.ListIndex
    If rowIndex < 0 Then Exit Sub

    pBinding = True
    pItems.Remove rowIndex + 1
    pList.RemoveItem rowIndex
    If pItems.Count > 0 Then
        If rowIndex >= pItems.Count Then rowIndex = pItems.Count - 1
        pList.ListIndex = rowIndex
    End If
    pBinding = False
    SyncSelectedLayers
End Sub

Private Sub lbxSettingLists_Change()
    If pBinding Then Exit Sub
    SyncSelectedLayers
End Sub

Private Sub SyncSelectedLayers()
    Dim item As ExportSettingItem
    Dim hasSelection As Boolean

    hasSelection = (pList.ListIndex >= 0)
    pBinding = True
    If hasSelection Then
        Set item = pItems.item(pList.ListIndex + 1)
        chkLayer1.value = item.Layer1
        chkLayer2.value = item.Layer2
        chkLayer3.value = item.Layer3
    Else
        chkLayer1.value = True
        chkLayer2.value = False
        chkLayer3.value = False
    End If
    chkLayer1.Enabled = hasSelection And Not pExporting
    chkLayer2.Enabled = hasSelection And Not pExporting
    chkLayer3.Enabled = hasSelection And Not pExporting
    cmdRemoveSetting.Enabled = hasSelection And Not pExporting
    cmdModify.Enabled = hasSelection And Not pExporting
    cmdExport.Enabled = (pItems.Count > 0) And Not pExporting
    cmdClearLists.Enabled = (pItems.Count > 0) And Not pExporting
    cmdAddSetting.Enabled = Not pExporting
    cmdClose.Enabled = Not pExporting
    pList.Enabled = Not pExporting
    pBinding = False
End Sub

Private Sub SaveSelectedLayers()
    Dim item As ExportSettingItem
    Dim rowText As String
    If pBinding Or pExporting Then Exit Sub
    If pList.ListIndex < 0 Then Exit Sub
    Set item = pItems.item(pList.ListIndex + 1)
    item.Layer1 = CBool(chkLayer1.value)
    item.Layer2 = CBool(chkLayer2.value)
    item.Layer3 = CBool(chkLayer3.value)
    rowText = item.Summary
    pBinding = True
    pList.List(pList.ListIndex, 0) = rowText
    pBinding = False
End Sub

Private Sub chkLayer1_Click()
    SaveSelectedLayers
End Sub

Private Sub chkLayer2_Click()
    SaveSelectedLayers
End Sub

Private Sub chkLayer3_Click()
    SaveSelectedLayers
End Sub

Private Sub cmdExport_Click()
    Dim runner As ExportQueueRunner
    Dim succeeded As Boolean
    Dim resultMessage As String
    Dim warningMessage As String
    Dim messageStyle As VbMsgBoxStyle
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ExportFailed
    If pMRBehavior And Not pMRExportAction Then Exit Sub
    If pExporting Then Exit Sub
    If pItems.Count = 0 Then
        If pMRBehavior Then Err.Raise 5, , "Antrean export kosong."
        Exit Sub
    End If
    SaveSelectedLayers
    pExporting = True
    SyncSelectedLayers
    Set runner = New ExportQueueRunner
    succeeded = runner.Run(pItems)
    warningMessage = runner.WarningDescription
    pExporting = False
    If succeeded Then
        SyncSelectedLayers
        resultMessage = "Export selesai: " & CStr(runner.CompletedFiles) & " file dari " & CStr(pItems.Count) & " item."
        messageStyle = vbInformation
        If Len(warningMessage) > 0 Then
            resultMessage = resultMessage & vbCrLf & vbCrLf & "Peringatan:" & vbCrLf & warningMessage
            messageStyle = vbExclamation
        End If
        If Not pMRBehavior Or Len(warningMessage) > 0 Then MsgBox resultMessage, messageStyle, "Export Queue"
    Else
        If runner.FailedItemIndex > 0 And runner.FailedItemIndex <= pItems.Count Then
            pList.ListIndex = runner.FailedItemIndex - 1
        End If
        SyncSelectedLayers
        resultMessage = "Export dihentikan. File yang sudah diekspor: " & CStr(runner.CompletedFiles) & "."
        If runner.FailedItemIndex > 0 Then resultMessage = resultMessage & vbCrLf & "Item: " & CStr(runner.FailedItemIndex)
        resultMessage = resultMessage & vbCrLf & "Error " & CStr(runner.LastErrorNumber) & ": " & runner.LastErrorDescription
        If Len(warningMessage) > 0 Then
            resultMessage = resultMessage & vbCrLf & vbCrLf & "Peringatan:" & vbCrLf & warningMessage
        End If
        If pMRBehavior Then
            errorNumber = runner.LastErrorNumber
            If errorNumber = 0 Then errorNumber = 5
            Err.Raise errorNumber, "ExportQueueRunner", resultMessage
        End If
        MsgBox resultMessage, vbExclamation, "Export Queue"
    End If
    Exit Sub

ExportFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pExporting = False
    On Error Resume Next
    SyncSelectedLayers
    On Error GoTo 0
    If pMRBehavior Then Err.Raise errorNumber, "ExportRelatedMenu.cmdExport", errorDescription
    MsgBox "Gagal menjalankan export queue (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "Export Queue"
End Sub

Private Sub cmdClose_Click()
    If pExporting Then Exit Sub
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If pExporting Then
        Cancel = 1
    ElseIf pMRBehavior Then
        If Not pMRObserver Is Nothing Then CallByName pMRObserver, "MainClosing", VbMethod
    End If
End Sub

Public Sub MRSetBehaviorMode(ByVal active As Boolean)
    pMRBehavior = active
End Sub

Public Function MRCreateBehaviorEditor() As Object
    Dim editor As ExportRelatedSettings, number As Long, description As String
    On Error GoTo Failed
    If pExporting Then Err.Raise 5, , "Export sedang berjalan."
    Set editor = New ExportRelatedSettings
    editor.BeginQueueEdit Nothing
    Set MRCreateBehaviorEditor = editor
    Exit Function
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not editor Is Nothing Then Unload editor
    On Error GoTo 0
    Err.Raise number, "ExportRelatedMenu.MRCreateBehaviorEditor", description
End Function

Public Sub MRCommitBehaviorItem(ByVal result As ExportSettingItem)
    If result Is Nothing Then Err.Raise 5, , "Hasil settings kosong."
    result.ValidateForQueue
    pItems.Add result
    pList.AddItem result.Summary
    pList.ListIndex = pItems.Count - 1
    SyncSelectedLayers
End Sub

Public Sub MRBehaviorValue(ByVal target As String, ByVal value As Variant)
    If LCase$(target) = "lbxsettinglists.index" Then
        If value < 1 Or value > pItems.Count Or Fix(value) <> value Then _
            Err.Raise 5, , "Index " & CStr(value) & " tidak tersedia; jumlah item " & CStr(pItems.Count) & "."
        pList.ListIndex = CLng(value) - 1
        SyncSelectedLayers
        Exit Sub
    End If
    If pList.ListIndex < 0 Then Err.Raise 5, , "Pilih lbxSettingLists.Index sebelum mengubah layer."
    Select Case LCase$(target)
        Case "chklayer1": chkLayer1.Value = CBool(value)
        Case "chklayer2": chkLayer2.Value = CBool(value)
        Case "chklayer3": chkLayer3.Value = CBool(value)
        Case Else: Err.Raise 5, , "Target menu tidak terdaftar: " & target
    End Select
    SaveSelectedLayers
End Sub

Public Sub MRBehaviorExport()
    Dim number As Long, description As String
    On Error GoTo Failed
    If pExporting Then Err.Raise 5, , "Export sedang berjalan."
    pMRExportAction = True
    cmdExport_Click
    pMRExportAction = False
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    pMRExportAction = False
    Err.Raise number, "ExportRelatedMenu.MRBehaviorExport", description
End Sub

Public Sub MRBehaviorClose()
    If pExporting Then Err.Raise 5, , "Tidak dapat menutup menu saat export."
    cmdClose_Click
End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
