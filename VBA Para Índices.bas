Dim Extension
    Dim I As Long
    Dim xRg As Range
    Dim xStr As String
    Dim xFd As FileDialog
    Dim xFdItem As Variant
    Dim xFileName As String
    Dim xFileNum As Long
    Dim xFileKB As Long
    Dim RegExp As Object
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    If xFd.Show = -1 Then
        xFdItem = xFd.SelectedItems(1) & Application.PathSeparator
        xFileName = Dir(xFdItem & "*.pdf", vbDirectory)
        Set xRg = Range("A1")
        Range("A:C").ClearContents
        Range("A1:K1").Font.Bold = True
        
        xRg = "Nombre Documento"
        xRg.Offset(0, 1) = "Fecha Creación Documento"
        xRg.Offset(0, 2) = "Fecha Incorporación Expediete"
        xRg.Offset(0, 3) = "Orden Documento"
        xRg.Offset(0, 4) = "Número Páginas"
        xRg.Offset(0, 5) = "Página Inicio"
        xRg.Offset(0, 6) = "Página Fin"
        xRg.Offset(0, 7) = "Formato"
        xRg.Offset(0, 8) = "Tamaño"
        xRg.Offset(0, 9) = "Origen"
        xRg.Offset(0, 10) = "Observaciones"
        
        I = 2
        xStr = ""
        Do While xFileName <> ""
            Cells(I, 1) = Replace(xFileName, ".pdf", "")
            Set RegExp = CreateObject("VBscript.RegExp")
            RegExp.Global = True
            RegExp.Pattern = "/Type\s*/Page[^s]"
            xFileNum = FreeFile
            Open (xFdItem & xFileName) For Binary As #xFileNum
                xStr = Space(LOF(xFileNum))
                Get #xFileNum, , xStr
            Close #xFileNum
            Cells(I, 2) = Now
            Cells(I, 3) = Now
            Cells(I, 4) = I - 1
            Cells(I, 5) = RegExp.Execute(xStr).Count
            If I = 2 Then
                Cells(I, 6) = 1
                Cells(I, 7) = Cells(I, 5)
            Else
                Cells(I, 6) = Cells(I - 1, 7) + 1
                Cells(I, 7) = Cells(I, 6) + Cells(I, 5) - 1
            End If
            
            Largura = InStrRev(xFileName, ".")
            Largura = Len(xFileName) - Largura
            Extension = Right(xFileName, Largura)
            Cells(I, 8) = UCase(Extension)
            Cells(I, 9) = Round(FileLen(xFdItem & xFileName) / 1024) & "KB"
            Cells(I, 10) = "DIGITAL"
            
            I = I + 1
            xFileName = Dir
            
        Loop
        Columns("A:C").AutoFit
        Columns("B:C").VerticalAlignment = xlCenter
        Columns("B:C").HorizontalAlignment = xlCenter
        Cells(1, 1).VerticalAlignment = xlCenter
        Cells(1, 1).HorizontalAlignment = xlCenter
        
    End If
End Sub
