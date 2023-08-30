Sub AggregateData()

    Dim baseFolderPath As String
    Dim wsAggregate As Worksheet
    Dim lastRow As Long
    Dim targetFile As String
    Dim doc As Object
    Dim docText As String
    Dim lines As Variant ' 配列として宣言
    Dim i As Long ' インデックス
    
    ' ダイアログを表示してフォルダを選択
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' フォルダが選択された場合
            baseFolderPath = .SelectedItems(1) & "\"
        Else ' キャンセルされた場合
            Exit Sub
        End If
    End With
    
    ' 集計シートを指定（必要に応じて変更）
    Set wsAggregate = ThisWorkbook.Sheets("集計シート名")
    
    ' 集計シートで最終行を取得
    lastRow = wsAggregate.Cells(wsAggregate.Rows.Count, "A").End(xlUp).Row + 1
    
    ' フォルダ内のファイルを繰り返し処理
    targetFile = Dir(baseFolderPath & "*.doc")
    Do While targetFile <> ""
        ' ワードファイルを開く
        Set doc = CreateObject("Word.Application")
        doc.Documents.Open baseFolderPath & targetFile
        
        ' 脆弱性調査シート内のテキストを取得
        docText = doc.ActiveDocument.Content.Text
        
        ' "CEV-"を含む行を抽出して配列に格納
        lines = Split(docText, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            If InStr(lines(i), "CEV-") > 0 Then
                ' 集計シートに追加
                wsAggregate.Cells(lastRow, 1).Value = targetFile
                wsAggregate.Cells(lastRow, 2).Value = lines(i)
                lastRow = lastRow + 1
            End If
        Next i
        
        ' ワードファイルを閉じる
        doc.Quit
        Set doc = Nothing
        
        ' 次のファイルを取得
        targetFile = Dir
    Loop
End Sub
