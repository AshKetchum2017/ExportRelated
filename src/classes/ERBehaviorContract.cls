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
        If form <> "exportrelatedmenu" And form <> "exportrelatedsettings" And form <> "jpegsettings" Then _
            Fail "UserForm tidak terdaftar: " & ins.FormName
        If ins.Kind <> "stage" Then ValidateInstruction ins
    Next ins
    ValidateFlow block
    Set Validate = block
End Function

Private Sub ValidateInstruction(ByVal ins As MRBehaviorInstruction)
    Dim expected As String
    Dim key As String

    key = LCase$(ins.Target)

    ' Validasi action.
    If ins.Kind = "action" Then
        Select Case LCase$(ins.FormName)
            Case "exportrelatedmenu"
                Select Case key
                    Case "cmdaddsetting", _
                         "cmdexport", _
                         "cmdclose"
                        Exit Sub
                End Select
            Case "exportrelatedsettings"
                Select Case key
                    Case "cmdsave", _
                         "cmdcancel", _
                         "cmdformatsettings"
                        Exit Sub
                End Select
            Case "jpegsettings"
                Select Case key
                    Case "cmdsave", _
                         "cmdclose"
                        Exit Sub
                End Select
        End Select
        Fail "Action tidak terdaftar: " & _
            ins.FormName & ".@" & ins.Target
    End If

    ' Tentukan tipe value yang diharapkan untuk setiap target.
    Select Case LCase$(ins.FormName)
        Case "exportrelatedsettings"
            Select Case key
                Case "txbname", _
                     "txbpage", _
                     "txbdirectory", _
                     "cmbexformat"
                    expected = "string"
                Case "chkparentdirectory"
                    expected = "boolean"
            End Select
        Case "exportrelatedmenu"
            Select Case key
                Case "lbxsettinglists.index"
                    expected = "number"
                Case "chklayer1", _
                     "chklayer2", _
                     "chklayer3"
                    expected = "boolean"
            End Select
        Case "jpegsettings"
            Select Case key
                Case "cmbcolormode"
                    expected = "string"
            End Select
    End Select

    If Len(expected) = 0 Then
        Fail "Target tidak terdaftar: " & _
            ins.FormName & "." & ins.Target
    End If

    If ins.Kind = "skip" Then Exit Sub
    If ins.Kind = "default" Then
        If Len(DefaultKey(ins.Target)) = 0 Or _
           LCase$(ins.FormName) <> "exportrelatedsettings" Then

            Fail _
                "Target tidak mempunyai Default registry: " & _
                ins.Target
        End If

        Exit Sub
    End If

    If ins.Kind <> expected Then
        Fail _
            "Tipe nilai " & _
            ins.Target & _
            " harus " & expected & "."
    End If

    ' Validasi nilai khusus.
    If key = "lbxsettinglists.index" Then
        If ins.Value < 1 Or _
           ins.Value > 2147483647# Then
            Fail _
                "Index harus bilangan bulat positif mulai dari 1."
        End If
        If Fix(ins.Value) <> ins.Value Then
            Fail "Index tidak boleh pecahan."
        End If
    End If

    If key = "cmbexformat" Then
        Select Case LCase$(CStr(ins.Value))
            Case ".pdf", ".dxf", ".png", ".jpg"

            Case Else
                Fail _
                    "Format harus .pdf, .dxf, .png, atau .jpg."
        End Select
    End If

    If LCase$(ins.FormName) = "jpegsettings" And _
       key = "cmbcolormode" Then

        Select Case LCase$(CStr(ins.Value))
            Case "grayscale", "rgb", "cmyk"
            Case Else
                Fail _
                    "Color mode JPEG harus grayscale, rgb, atau cmyk."
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
    Dim ins As MRBehaviorInstruction
    Dim editorOpen As Boolean
    Dim formatOpen As Boolean
    Dim closed As Boolean
    Dim resumeAt As Long
    Dim itemCount As Long
    Dim formName As String
    Dim target As String

    resumeAt = pAt
    For Each ins In block.Instructions
        pAt = ins.Position
        formName = LCase$(ins.FormName)
        target = LCase$(ins.Target)

        If closed Then
            Fail _
                "Tidak boleh ada tahap/instruksi setelah @cmdClose."
        End If
        Select Case formName
            Case "exportrelatedmenu"
                If formatOpen Then
                    Fail _
                        "Selesaikan Format Settings sebelum kembali ke menu."
                End If
                If editorOpen Then
                    Fail _
                        "Selesaikan ExportRelatedSettings sebelum kembali ke menu."
                End If
            Case "exportrelatedsettings"
                If Not editorOpen Then
                    Fail _
                        "ExportRelatedSettings belum dibuka; " & _
                        "jalankan @cmdAddSetting dahulu."
                End If
                If formatOpen Then
                    Fail _
                        "Selesaikan Format Settings sebelum kembali ke ExportRelatedSettings."
                End If
            Case "jpegsettings"
                If Not editorOpen Then
                    Fail _
                        "ExportRelatedSettings belum dibuka."
                End If
                If Not formatOpen Then
                    Fail _
                        "JPEGSettings belum dibuka; " & _
                        "jalankan @cmdFormatSettings dahulu."
                End If
        End Select
        If target = "lbxsettinglists.index" And _
           ins.Kind = "number" Then

            If ins.Value > itemCount Then
                Fail _
                    "Index di luar item yang dibentuk script; jumlah item: " & _
                    CStr(itemCount)
            End If
        End If
        If Left$(target, 8) = "chklayer" And _
           ins.Kind <> "skip" Then

            If itemCount = 0 Then
                Fail _
                    "Belum ada item untuk pengaturan layer."
            End If
        End If
        If ins.Kind = "action" Then
            Select Case formName & "." & target
                Case "exportrelatedmenu.cmdaddsetting"
                    editorOpen = True
                Case "exportrelatedsettings.cmdformatsettings"
                    formatOpen = True
                Case "jpegsettings.cmdsave", _
                     "jpegsettings.cmdclose"
                    formatOpen = False
                Case "exportrelatedsettings.cmdsave"
                    editorOpen = False
                    itemCount = itemCount + 1
                Case "exportrelatedsettings.cmdcancel"
                    editorOpen = False
                Case "exportrelatedmenu.cmdexport"
                    If itemCount = 0 Then
                        Fail _
                            "Antrean export kosong; belum ada @cmdSave."
                    End If
                Case "exportrelatedmenu.cmdclose"
                    closed = True
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
