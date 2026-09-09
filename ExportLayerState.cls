Option Explicit

' Snapshot print/export per layer; visibility, editability, dan urutan tidak diubah.
Private pLayers As Collection
Private pOriginal As Collection
Private pDesired As Collection

Private Sub Class_Initialize()
    Set pLayers = New Collection
    Set pOriginal = New Collection
    Set pDesired = New Collection
End Sub

Public Sub ValidateSelection(ByVal item As ExportSettingItem, ByVal groups As Collection)
    Dim targetPages As Collection
    Dim target As Page
    Dim lyr As Layer
    Dim layerNumber As Long
    Dim matchCount As Long
    Dim layerName As String

    Set targetPages = ResolvePages(item.SourceDocument, groups)
    For Each target In targetPages
        If Not (item.Layer1 Or item.Layer2 Or item.Layer3) Then
            If Not HasPrintableAutomaticLayer(item.SourceDocument, target) Then
                Err.Raise 5, "ExportLayerState", "Page " & CStr(target.Index) & _
                    ": tidak ada layer dipilih, layer otomatis lokal, atau master layer dengan print/export aktif."
            End If
        End If
        For layerNumber = 1 To 3
            layerName = "Layer " & CStr(layerNumber)
            If IsSelectedName(layerName, item) Then
                matchCount = 0
                For Each lyr In target.Layers
                    If Not lyr.Master And Not lyr.IsSpecialLayer Then
                        If StrComp(lyr.Name, layerName, vbBinaryCompare) = 0 Then matchCount = matchCount + 1
                    End If
                Next lyr
                If matchCount <> 1 Then
                    Err.Raise 5, "ExportLayerState", "Page " & CStr(target.Index) & ": layer " & layerName & _
                        " harus tersedia tepat satu kali sebagai layer lokal (ditemukan " & CStr(matchCount) & ")."
                End If
            End If
        Next layerNumber
    Next target
End Sub

Public Sub Apply(ByVal item As ExportSettingItem, ByVal groups As Collection)
    Dim targetPages As Collection
    Dim target As Page
    Dim lyr As Layer
    Dim layerIndex As Long

    If pLayers.Count > 0 Then Err.Raise 5, "ExportLayerState", "State layer sebelumnya belum dipulihkan."
    ValidateSelection item, groups
    Set targetPages = ResolvePages(item.SourceDocument, groups)
    ' Master layer mengikuti flag pengguna; hanya layer lokal yang perlu diubah.
    For Each target In targetPages
        For Each lyr In target.Layers
            Capture lyr, Not lyr.Master And Not lyr.IsSpecialLayer And IsSelectedName(lyr.Name, item)
        Next lyr
    Next target
    For layerIndex = 1 To pLayers.Count
        Set lyr = pLayers.Item(layerIndex)
        If lyr.Printable <> CBool(pDesired.Item(layerIndex)) Then lyr.Printable = CBool(pDesired.Item(layerIndex))
    Next layerIndex
End Sub

Private Sub Capture(ByVal lyr As Layer, ByVal desired As Boolean)
    Dim existing As Layer
    ' Semua master layer dan empat layer otomatis lokal mengikuti flag pengguna.
    ' Tidak disentuh saat Apply maupun Restore; jangan mengaktifkan layer yang disabled.
    If lyr.Master Then Exit Sub
    If IsAutomaticLayerName(lyr.Name) Then Exit Sub
    For Each existing In pLayers
        If existing Is lyr Then Exit Sub
    Next existing
    pOriginal.Add CBool(lyr.Printable)
    pDesired.Add desired
    pLayers.Add lyr
End Sub

' Coba pulihkan setiap layer, walaupun satu setter gagal. Kembalikan rincian error.
Public Function Restore() As String
    Dim lyr As Layer
    Dim layerIndex As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error Resume Next
    For layerIndex = 1 To pLayers.Count
        Set lyr = Nothing
        Err.Clear
        Set lyr = pLayers.Item(layerIndex)
        If Not lyr Is Nothing Then
            If lyr.Printable <> CBool(pOriginal.Item(layerIndex)) Then lyr.Printable = CBool(pOriginal.Item(layerIndex))
        End If
        errorNumber = Err.Number
        errorDescription = Err.Description
        If errorNumber <> 0 Then
            Restore = Restore & vbCrLf & "Restore layer #" & CStr(layerIndex) & " (" & CStr(errorNumber) & "): " & errorDescription
        End If
    Next layerIndex
    On Error GoTo 0
    If Len(Restore) = 0 Then
        Set pLayers = New Collection
        Set pOriginal = New Collection
        Set pDesired = New Collection
    End If
End Function

Private Function IsSelectedName(ByVal layerName As String, ByVal item As ExportSettingItem) As Boolean
    Select Case layerName
        Case "Layer 1": IsSelectedName = item.Layer1
        Case "Layer 2": IsSelectedName = item.Layer2
        Case "Layer 3": IsSelectedName = item.Layer3
    End Select
End Function

Private Function IsAutomaticLayerName(ByVal layerName As String) As Boolean
    Select Case layerName
        Case "scpro2_printmargin", "scpro2_printonly", "scpro2_regmarks", "Regmark"
            IsAutomaticLayerName = True
    End Select
End Function

Private Function HasPrintableAutomaticLayer(ByVal doc As Document, ByVal target As Page) As Boolean
    Dim lyr As Layer
    For Each lyr In target.Layers
        If lyr.Master Or IsAutomaticLayerName(lyr.Name) Then
            If lyr.Printable Then
                HasPrintableAutomaticLayer = True
                Exit Function
            End If
        End If
    Next lyr
    For Each lyr In doc.MasterPage.Layers
        If lyr.Printable Then
            HasPrintableAutomaticLayer = True
            Exit Function
        End If
    Next lyr
End Function

Private Function ResolvePages(ByVal doc As Document, ByVal groups As Collection) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim group As Variant
    Dim groupIndex As Long
    Dim pageIndex As Long
    Dim pageNumber As Long

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    For groupIndex = 1 To groups.Count
        group = groups.Item(groupIndex)
        For pageIndex = LBound(group) To UBound(group)
            pageNumber = CLng(group(pageIndex))
            If Not seen.Exists(pageNumber) Then
                seen.Add pageNumber, True
                result.Add doc.Pages(pageNumber)
            End If
        Next pageIndex
    Next groupIndex
    Set ResolvePages = result
End Function
