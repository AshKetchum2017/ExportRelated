Option Explicit

Private pDirectory As String
Private pFormat As String
Private pTemplateText As String
Private pPageText As String
Private pUseParentDirectory As Boolean

Public Property Get UseParentDirectory() As Boolean
    UseParentDirectory = pUseParentDirectory
End Property

Public Property Let UseParentDirectory(ByVal value As Boolean)
    pUseParentDirectory = value
End Property

Public Function DocumentDirectory(ByVal doc As Document) As String
    If doc Is Nothing Then Exit Function
    If Len(doc.FileName) = 0 Then Exit Function
    DocumentDirectory = doc.FilePath
End Function

Public Function ResolveDirectory(ByVal doc As Document) As String
    If Not pUseParentDirectory Then
        ResolveDirectory = pDirectory
        Exit Function
    End If
    ResolveDirectory = DocumentDirectory(doc)
    If Len(ResolveDirectory) = 0 Then
        Err.Raise 5, "ExportSettings", "ParentDirectory aktif. Simpan dokumen sumber terlebih dahulu atau matikan ParentDirectory."
    End If
End Function

Public Property Get Directory() As String
    Directory = pDirectory
End Property

Public Property Let Directory(ByVal value As String)
    pDirectory = Trim$(value)
End Property

Public Property Get FormatText() As String
    FormatText = pFormat
End Property

Public Property Let FormatText(ByVal value As String)
    pFormat = Trim$(value)
End Property

Public Property Get TemplateText() As String
    TemplateText = pTemplateText
End Property

Public Property Let TemplateText(ByVal value As String)
    pTemplateText = Trim$(value)
End Property

Public Property Get PageText() As String
    PageText = pPageText
End Property

Public Property Let PageText(ByVal value As String)
    pPageText = Trim$(value)
End Property

Public Sub LoadFromForm(ByVal frm As Object)
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ReadFailed
    operation = "frm.ERInputText(txbDirectory)"
    Me.Directory = frm.ERInputText(frm.txbDirectory)
    operation = "frm.chkParentDirectory.Value"
    Me.UseParentDirectory = CBool(frm.chkParentDirectory.Value)
    operation = "frm.cmbExFormat.Value"
    Me.FormatText = frm.cmbExFormat.Value
    operation = "frm.ERInputText(txbName)"
    Me.TemplateText = frm.ERInputText(frm.txbName)
    operation = "frm.ERInputText(txbPage)"
    Me.PageText = frm.ERInputText(frm.txbPage)
    Exit Sub

ReadFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "ExportSettings.LoadFromForm", _
        operation & " [" & TypeName(frm) & "]" & vbCrLf & _
        "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Public Function IsValidDirectory() As Boolean
    IsValidDirectory = (Len(Me.Directory) > 0) And (Len(Dir$(Me.Directory, vbDirectory)) > 0)
End Function
