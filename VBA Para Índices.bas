Attribute VB_Name = "M�dulo2"
Sub Macro2Datos�ndice()
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
        Range("A1:C1").Font.Bold = True
        xRg = "Nombre Archivos"
        xRg.Offset(0, 1) = "P�ginas"
        xRg.Offset(0, 2) = "Tama�o"
        I = 2
        xStr = ""
        Do While xFileName <> ""
            Cells(I, 1) = xFileName
            Set RegExp = CreateObject("VBscript.RegExp")
            RegExp.Global = True
            RegExp.Pattern = "/Type\s*/Page[^s]"
            xFileNum = FreeFile
            Open (xFdItem & xFileName) For Binary As #xFileNum
                xStr = Space(LOF(xFileNum))
                Get #xFileNum, , xStr
            Close #xFileNum
            Cells(I, 2) = RegExp.Execute(xStr).Count
        
            Cells(I, 3) = Round(FileLen(xFdItem & xFileName) / 1024) & "KB"
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
