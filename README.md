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
    Dim table As Object
    
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
            For r = 1 To table.Rows.Count
                For c = 1 To table.Columns.Count
                    If InStr(table.Cell(r, c).Range.Text, "対象システム") > 0 Then
                        wsAggregate.Cells(destRow, 3).Value = table.Cell(r, c + 1).Range.Text
                        Exit For
                    End If
                Next c
            Next r
        Next table
        
        ' ワード関連のオブジェクトを解放
        wordDoc.Close SaveChanges:=False
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
    Next row
End Sub
