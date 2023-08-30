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
---------------

Sub AggregateData()
    Dim wsSettings As Worksheet
    Dim wsAggregate As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim targetFile As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim row As Long
    Dim destRow As Long
    
    ' 設定シートと集計シートを取得
    Set wsSettings = ThisWorkbook.Sheets("設定シート")
    Set wsAggregate = ThisWorkbook.Sheets("集計シート")
    
    ' 設定シートの最終行を取得
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "A").End(xlUp).Row
    
    ' フォルダごとにループ
    For row = 2 To lastRow ' 2行目から開始（1行目はヘッダ）
        folderPath = wsSettings.Cells(row, 1).Value
        destRow = wsAggregate.Cells(wsAggregate.Rows.Count, "A").End(xlUp).Row + 1
        
        ' 集計シートへフォルダ名を記入
        wsAggregate.Cells(destRow, 1).Value = folderPath
        
        ' 該当ワードファイルを開く
        fileName = "脆弱性調査シート.doc"
        targetFile = folderPath & "\" & fileName
        
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        Set wordDoc = wordApp.Documents.Open(targetFile)
        
        ' ワードファイル内のテキストを検索して集計シートへコピー
        For Each para In wordDoc.Paragraphs
            If InStr(para.Range.Text, "CEV-") > 0 Then
                wsAggregate.Cells(destRow, 2).Value = wsAggregate.Cells(destRow, 2).Value & para.Range.Text
            End If
        Next para
        
        ' 対象システム列を追記
        For Each table In wordDoc.Tables
            For Each cell In table.Columns(2).Cells
                If InStr(cell.Range.Text, "対象システム") > 0 Then
                    wsAggregate.Cells(destRow, 3).Value = cell.Next.Text
                    Exit For
                End If
            Next cell
        Next table
        
        ' ワード関連のオブジェクトを解放
        wordDoc.Close SaveChanges:=False
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
    Next row
End Sub
