Option Explicit

' Class Module (Name) = ERBehaviorSession. Owned by the main form while it is open.
Private pMenu As ExportRelatedMenu
Private pEditor As ExportRelatedSettings
Private pObserver As Object
Private pToken As String
Private pExecuting As Boolean
Private pInterrupted As Boolean
Private pExpectedClose As Boolean
Private pMainClosing As Boolean
Private pDefaults As Object
Private pFormatEditor As Object

Public Sub Start(ByVal script As String, ByVal observer As Object, ByVal token As String)
    Dim contract As New ERBehaviorContract, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, existing As Object, value As Variant
    Dim operation As String, number As Long, description As String
    On Error GoTo Failed
    Set block = contract.Validate(script)
    operation = "Membaca snapshot Default registry"
    SnapshotDefaults block, contract
    operation = "Memeriksa form yang sudah terbuka"
    For Each existing In VBA.UserForms
        If TypeName(existing) = "ExportRelatedMenu" Or TypeName(existing) = "ExportRelatedSettings" Then _
            Err.Raise 5, , "Tutup form Export Related yang sudah terbuka sebelum menjalankan behavior."
    Next existing
    Set existing = Nothing
    Set pObserver = observer: pToken = token
    pExecuting = True
    operation = "Membuka ExportRelatedMenu"
    Set pMenu = New ExportRelatedMenu
    pMenu.MRBindRunner Me, token
    pMenu.MRSetBehaviorMode True
    pMenu.Show vbModeless
    For Each ins In block.Instructions
        CheckInterrupted
        operation = ins.FormName & " / " & ins.Target & " (posisi " & CStr(ins.Position) & ")"
        If ins.Kind = "stage" Then
            Select Case LCase$(ins.FormName)

                Case "exportrelatedmenu"
                    If Not pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "Format Settings belum diselesaikan."
                    End If
                    If Not pEditor Is Nothing Then
                        Err.Raise 5, , _
                            "ExportRelatedSettings belum diselesaikan."
                    End If
                Case "exportrelatedsettings"
                    If pEditor Is Nothing Then
                        Err.Raise 5, , _
                            "Editor belum dibuka melalui @cmdAddSetting."
                    End If
                    If Not pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "Format Settings belum diselesaikan."
                    End If

                Case "jpegsettings"
                    If pEditor Is Nothing Then
                        Err.Raise 5, , _
                            "ExportRelatedSettings belum dibuka."
                    End If
                    If pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "JPEGSettings belum dibuka melalui @cmdFormatSettings."
                    End If
                    If LCase$(TypeName(pFormatEditor)) <> "jpegsettings" Then
                        Err.Raise 5, , _
                            "Format Settings aktif bukan JPEGSettings."
                    End If
            End Select
            DoEvents
            CheckInterrupted
        ElseIf ins.Kind = "action" Then
            ExecuteAction ins
        ElseIf ins.Kind <> "skip" Then
            value = ins.Value
            If ins.Kind = "default" Then value = pDefaults(LCase$(ins.Target))
            Select Case LCase$(ins.FormName)

                Case "exportrelatedsettings"
                    pEditor.MRBehaviorValue ins.Target, value

                Case "jpegsettings"
                    If pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "JPEGSettings belum dibuka melalui @cmdFormatSettings."
                    End If

                    pFormatEditor.MRBehaviorValue _
                        ins.Target, value

                Case Else
                    pMenu.MRBehaviorValue ins.Target, value

            End Select
        End If
    Next ins
    CheckInterrupted
    pExecuting = False
    If pEditor Is Nothing Then
        If Not pMainClosing Then pMenu.MRSetBehaviorMode False
        ' The main form owns this session. Do not retain it and delay Terminate.
        Set pMenu = Nothing
    End If
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    pExecuting = False
    AbortSession
    Err.Raise number, "ERBehaviorSession", operation & vbCrLf & description
End Sub

Private Sub SnapshotDefaults(ByVal block As MRBehaviorBlock, ByVal contract As ERBehaviorContract)
    Dim ins As MRBehaviorInstruction, key As String, raw As String, target As String
    Dim fso As Object
    Set pDefaults = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each ins In block.Instructions
        If ins.Kind = "default" Then
            target = LCase$(ins.Target)
            If Not pDefaults.Exists(target) Then
                key = contract.DefaultKey(ins.Target)
                raw = GetSetting("RinCorelMacros", "ExportRelatedMacro", key, ChrW(1))
                If raw = ChrW(1) Then Err.Raise 5, , "Default registry belum tersedia: " & key
                Select Case target
                    Case "txbdirectory"
                        If Len(raw) = 0 Then Err.Raise 5, , "Default LastDirectory kosong."
                        If Not fso.FolderExists(raw) Then Err.Raise 76, , "Folder Default LastDirectory tidak ditemukan: " & raw
                        pDefaults.Add target, raw
                    Case "cmbexformat"
                        Select Case LCase$(raw)
                            Case ".pdf", ".dxf", ".png", ".jpg"
                            Case Else: Err.Raise 5, , "Default LastExportFormat tidak valid: " & raw
                        End Select
                        pDefaults.Add target, LCase$(raw)
                    Case "chkparentdirectory"
                        If LCase$(raw) <> "true" And LCase$(raw) <> "false" Then Err.Raise 5, , "Default UseParentDirectory harus True/False."
                        pDefaults.Add target, (LCase$(raw) = "true")
                End Select
            End If
        End If
    Next ins
End Sub

Private Sub ExecuteAction(ByVal ins As MRBehaviorInstruction)
    Dim formName As String
    Dim target As String

    formName = LCase$(ins.FormName)
    target = LCase$(ins.Target)

    Select Case formName

        Case "exportrelatedmenu"
            Select Case target
                Case "cmdaddsetting"
                    Set pEditor = pMenu.MRCreateBehaviorEditor()
                    pEditor.MRSetBehaviorObserver Me
                    pMenu.Enabled = False
                    pEditor.Show vbModeless
                Case "cmdexport"
                    pMenu.MRBehaviorExport
                Case "cmdclose"
                    pExpectedClose = True
                    pMenu.MRBehaviorClose
            End Select

        Case "exportrelatedsettings"
            Select Case target
                Case "cmdformatsettings"
                    OpenFormatEditor
                Case "cmdsave", "cmdcancel"
                    If Not pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "Selesaikan Format Settings sebelum menutup ExportRelatedSettings."
                    End If

                    pEditor.MRBehaviorAction ins.Target
                    FinishEditor (target = "cmdsave")
            End Select

        Case "jpegsettings"
            Select Case target
                Case "cmdsave", "cmdclose"
                    If pFormatEditor Is Nothing Then
                        Err.Raise 5, , _
                            "JPEGSettings belum dibuka."
                    End If

                    pFormatEditor.MRBehaviorAction ins.Target
                    FinishFormatEditor

            End Select
    End Select

    CheckInterrupted
End Sub

Private Sub OpenFormatEditor()

    If pEditor Is Nothing Then
        Err.Raise 5, , _
            "ExportRelatedSettings belum dibuka."
    End If

    If Not pFormatEditor Is Nothing Then
        Err.Raise 5, , _
            "Format Settings sudah terbuka."
    End If

    Set pFormatEditor = _
        pEditor.MRCreateBehaviorFormatEditor()

    If pFormatEditor Is Nothing Then
        Err.Raise 91, , _
            "Format editor tidak berhasil dibuat."
    End If

    If LCase$(TypeName(pFormatEditor)) <> "jpegsettings" Then
        Err.Raise 5, , _
            "Format editor MacroBehavior belum didukung: " & _
            TypeName(pFormatEditor)
    End If

    pFormatEditor.MRSetBehaviorObserver Me

    pEditor.Enabled = False
    pFormatEditor.Show vbModeless
End Sub

Private Sub FinishFormatEditor()

    If pFormatEditor Is Nothing Then
        Err.Raise 5, , _
            "Format editor tidak tersedia."
    End If

    pFormatEditor.MRDetachBehavior
    Unload pFormatEditor
    Set pFormatEditor = Nothing

    If Not pEditor Is Nothing Then
        pEditor.Enabled = True
    End If

End Sub

Public Sub FormatEditorFinished(ByVal saved As Boolean)
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    If pExecuting Then
        pInterrupted = True
        FinishFormatEditor
    Else
        FinishFormatEditor
    End If

    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description

    FormatEditorFailed errorNumber, errorDescription
End Sub

Public Sub FormatEditorFailed( _
        ByVal number As Long, _
        ByVal description As String)

    Dim observer As Object
    Dim token As String

    If pExecuting Then
        pInterrupted = True
        Exit Sub
    End If

    Set observer = pObserver
    token = pToken

    AbortSession

    If Not observer Is Nothing Then
        CallByName observer, _
            "BehaviorFailed", _
            VbMethod, _
            token, _
            number, _
            description
    End If

End Sub

Private Sub FinishEditor(ByVal saved As Boolean)
    Dim result As ExportSettingItem
    If pEditor Is Nothing Then Err.Raise 5, , "Editor settings tidak tersedia."
    If saved Then
        Set result = pEditor.QueueResult
        If result Is Nothing Then Err.Raise 5, , "@cmdSave tidak menghasilkan item antrean."
        pMenu.MRCommitBehaviorItem result
    End If
    pEditor.MRDetachBehavior
    Unload pEditor
    Set pEditor = Nothing
    pMenu.Enabled = True
End Sub

' Called by manual Save/Cancel/X in an editor left open by the final stage.
Public Sub EditorFinished(ByVal saved As Boolean)
    Dim number As Long, description As String
    On Error GoTo Failed
    If pExecuting Then
        pInterrupted = True
        FinishEditor False
    Else
        FinishEditor saved
        pMenu.MRSetBehaviorMode False
        Set pMenu = Nothing
    End If
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    EditorFailed number, description
End Sub

Public Sub EditorFailed(ByVal number As Long, ByVal description As String)
    Dim observer As Object, token As String
    If pExecuting Then
        pInterrupted = True
        Exit Sub
    End If
    Set observer = pObserver: token = pToken
    AbortSession
    If Not observer Is Nothing Then CallByName observer, "BehaviorFailed", VbMethod, token, number, description
End Sub

Public Sub MainClosing()
    pMainClosing = True
    If pExecuting And Not pExpectedClose Then pInterrupted = True
End Sub

Public Sub MacroUnloaded(ByVal token As String)
    Dim observer As Object
    If token <> pToken Then Exit Sub
    Set observer = pObserver
    Set pObserver = Nothing
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
End Sub

Private Sub CheckInterrupted()
    If pInterrupted Then Err.Raise 5, "ERBehaviorSession", "Behavior diinterupsi: form ditutup atau editor diselesaikan sebelum instruksi selesai."
End Sub

Private Sub AbortSession()

    On Error Resume Next
    If Not pFormatEditor Is Nothing Then
        pFormatEditor.MRDetachBehavior
        Unload pFormatEditor
    End If

    Set pFormatEditor = Nothing
    If Not pEditor Is Nothing Then
        pEditor.MRDetachBehavior
        Unload pEditor
    End If

    Set pEditor = Nothing
    If Not pMenu Is Nothing Then
        pMenu.MRDetachRunner
        If Not pMainClosing Then
            pMenu.Enabled = True
            pMenu.MRSetBehaviorMode False
        End If
    End If

    Set pMenu = Nothing
    Set pObserver = Nothing
End Sub
