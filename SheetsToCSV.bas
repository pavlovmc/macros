Sub SheetsToFiles()
    Dim FilesToOpen
    Dim x As Integer
    
    Dim xDir As String
    Dim folder As FileDialog
    Set folder = Application.FileDialog(msoFileDialogFolderPicker)
    If folder.Show <> -1 Then Exit Sub
    xDir = folder.SelectedItems(1)

    Application.ScreenUpdating = False  'îòêëþ÷àåì îáíîâëåíèå ýêðàíà äëÿ ñêîðîñòè
    
    'âûçûâàåì äèàëîã âûáîðà ôàéëîâ äëÿ èìïîðòà
    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="All files (*.*), *.*", _
      MultiSelect:=True, Title:="Ôàéëû äëÿ êîíâåðòàöèè â csv")

    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "Íå âûáðàíî íè îäíîãî ôàéëà!"
        Exit Sub
    End If
    
    'ïðîõîäèì ïî âñåì âûáðàííûì ôàéëàì
    x = 1
    Dim fl As String
    While x <= UBound(FilesToOpen)
        Set importWB = Workbooks.Open(Filename:=FilesToOpen(x), UpdateLinks:=0)
        fl = importWB.Name
        For Each xWs In Sheets()
            xWs.SaveAs xDir & "\" & fl & "_" & xWs.Name & ".csv", xlCSV
            Next
        'Sheets().Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        importWB.Close savechanges:=False
        x = x + 1
    Wend

    Application.ScreenUpdating = True
End Sub
