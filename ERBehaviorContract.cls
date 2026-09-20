Option Explicit

' Class Module (Name) = ERBehaviorContract. Pure validation; no host/UI/registry access.
Private pText As String
Private pAt As Long

Public Function Validate(ByVal script As String) As MRBehaviorBlock
    Dim parser As New MRBehaviorParser, blocks As Collection, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, form As String
    pText = script: pAt = 1
    Set blocks = parser.Parse(script, Array("ExportRelated"))
    Set block = blocks(1)
    If block.Instructions.Count = 0 Then Fail "Blok behavior diperlukan."
    For Each ins In block.Instructions
        pAt = ins.Position
        form = LCase$(ins.FormName)
        If form <> "exportrelatedmenu" And form <> "exportrelatedsettings" Then _
            Fail "UserForm tidak terdaftar: " & ins.FormName
        If ins.Kind <> "stage" Then ValidateInstruction ins
    Next ins
    ValidateFlow block
    Set Validate = block
End Function

Private Sub ValidateInstruction(ByVal ins As MRBehaviorInstruction)
    Dim expected As String, key As String
    key = LCase$(ins.Target)
    If ins.Kind = "action" Then
        If LCase$(ins.FormName) = "exportrelatedmenu" Then
            If key = "cmdaddsetting" Or key = "cmdexport" Or key = "cmdclose" Then Exit Sub
        Else
            If key = "cmdsave" Or key = "cmdcancel" Then Exit Sub
        End If
        Fail "Action tidak terdaftar: " & ins.FormName & ".@" & ins.Target
    End If
    If LCase$(ins.FormName) = "exportrelatedsettings" Then
        Select Case key
            Case "txbname", "txbpage", "txbdirectory", "cmbexformat": expected = "string"
            Case "chkparentdirectory": expected = "boolean"
        End Select
    Else
        Select Case key
            Case "lbxsettinglists.index": expected = "number"
            Case "chklayer1", "chklayer2", "chklayer3": expected = "boolean"
        End Select
    End If
    If Len(expected) = 0 Then Fail "Target tidak terdaftar: " & ins.FormName & "." & ins.Target
    If ins.Kind = "skip" Then Exit Sub
    If ins.Kind = "default" Then
        If Len(DefaultKey(ins.Target)) = 0 Or LCase$(ins.FormName) <> "exportrelatedsettings" Then _
            Fail "Target tidak mempunyai Default registry: " & ins.Target
        Exit Sub
    End If
    If ins.Kind <> expected Then Fail "Tipe nilai " & ins.Target & " harus " & expected & "."
    If key = "lbxsettinglists.index" Then
        If ins.Value < 1 Or ins.Value > 2147483647# Then Fail "Index harus bilangan bulat positif mulai dari 1."
        If Fix(ins.Value) <> ins.Value Then Fail "Index tidak boleh pecahan."
    End If
    If key = "cmbexformat" Then
        Select Case LCase$(CStr(ins.Value))
            Case ".pdf", ".dxf", ".png", ".jpg"
            Case Else: Fail "Format harus .pdf, .dxf, .png, atau .jpg."
        End Select
    End If
End Sub

Public Function DefaultKey(ByVal target As String) As String
    Select Case LCase$(target)
        Case "txbdirectory": DefaultKey = "LastDirectory"
        Case "cmbexformat": DefaultKey = "LastExportFormat"
        Case "chkparentdirectory": DefaultKey = "UseParentDirectory"
    End Select
End Function

Private Sub ValidateFlow(ByVal block As MRBehaviorBlock)
    Dim ins As MRBehaviorInstruction, editorOpen As Boolean, closed As Boolean
    Dim resumeAt As Long, itemCount As Long
    resumeAt = pAt
    For Each ins In block.Instructions
        pAt = ins.Position
        If closed Then Fail "Tidak boleh ada tahap/instruksi setelah @cmdClose."
        If LCase$(ins.FormName) = "exportrelatedsettings" Then
            If Not editorOpen Then Fail "ExportRelatedSettings belum dibuka; jalankan @cmdAddSetting dahulu."
        Else
            If editorOpen Then Fail "Selesaikan editor melalui @cmdSave atau @cmdCancel sebelum kembali ke menu."
        End If
        If LCase$(ins.Target) = "lbxsettinglists.index" And ins.Kind = "number" Then
            If ins.Value > itemCount Then Fail "Index di luar item yang dibentuk script; jumlah item: " & CStr(itemCount)
        End If
        If Left$(LCase$(ins.Target), 8) = "chklayer" And ins.Kind <> "skip" Then
            If itemCount = 0 Then Fail "Belum ada item untuk pengaturan layer."
        End If
        If ins.Kind = "action" Then
            Select Case LCase$(ins.Target)
                Case "cmdaddsetting": editorOpen = True
                Case "cmdsave": editorOpen = False: itemCount = itemCount + 1
                Case "cmdcancel": editorOpen = False
                Case "cmdexport"
                    If itemCount = 0 Then Fail "Antrean export kosong; belum ada @cmdSave."
                Case "cmdclose": closed = True
            End Select
        End If
    Next ins
    pAt = resumeAt
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long, lastBreak As Long
    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    lastBreak = InStrRev(prefix, vbLf)
    column = pAt - lastBreak
    Err.Raise 5, "ERBehaviorContract", "Semantic ExportRelated baris " & CStr(line) & ", kolom " & CStr(column) & ": " & message
End Sub
