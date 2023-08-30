Sub AggregateData()

    Dim baseFolderPaths As Variant
    Dim wsAggregate As Worksheet
    Dim lastRow As Long
    Dim targetFile As String
    Dim doc As Object
    Dim docText As String
    Dim lines As Variant
    Dim i As Long
    
    ' ダイアログを表示して複数のフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True ' 複数選択を許可
        If .Show = -1 Then ' フォルダが選択された場合
            baseFolderPaths = .SelectedItems
        Else ' キャンセルされた場合
            Exit Sub
        End If
    End With
    
    ' 集計シートを指定（必要に応じて変更）
    Set wsAggregate = ThisWorkbook.Sheets("集計シート名")
    
    ' 集計シートで最終行を取得
    lastRow = wsAggregate.Cells(wsAggregate.Rows.Count, "A").End(xlUp).Row + 1
    
    ' 選択された複数のフォルダを順番に処理
    For Each baseFolderPath In baseFolderPaths
        targetFile = Dir(baseFolderPath & "\*.doc")
        Do While targetFile <> ""
            Set doc = CreateObject("Word.Application")
            doc.Documents.Open baseFolderPath & "\" & targetFile
            docText = doc.ActiveDocument.Content.Text
            lines = Split(docText, vbCrLf)
            For i = LBound(lines) To UBound(lines)
                If InStr(lines(i), "CEV-") > 0 Then
                    wsAggregate.Cells(lastRow, 1).Value = targetFile
                    wsAggregate.Cells(lastRow, 2).Value = lines(i)
                    lastRow = lastRow + 1
                End If
            Next i
            doc.Quit
            Set doc = Nothing
            targetFile = Dir
        Loop
    Next baseFolderPath
End Sub
