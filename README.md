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
        
        ' 該当ワードファイルを開く
        fileName = "脆弱性調査シート.doc"
        targetFile = folderPath & "\" & fileName
        
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        Set wordDoc = wordApp.Documents.Open(targetFile)
        
        ' ワードファイル内のテキストを検索して集計シートへコピー
        Dim cevText As String
        For Each para In wordDoc.Paragraphs
            If InStr(para.Range.Text, "CEV-") > 0 Then
                cevText = cevText & para.Range.Text & vbNewLine
            End If
        Next para
        
        ' テーブルを検索して「対象システム」列を取得
        Dim systemText As String
        For Each table In wordDoc.Tables
            For r = 1 To table.Rows.Count
                For c = 1 To table.Columns.Count
                    If InStr(table.Cell(r, c).Range.Text, "対象システム") > 0 Then
                        systemText = table.Cell(r, c + 1).Range.Text
                        Exit For
                    End If
                Next c
            Next r
        Next table
        
        ' 集計シートへデータを追記
        wsAggregate.Cells(destRow, 1).Value = folderPath
        wsAggregate.Cells(destRow, 2).Value = cevText
        wsAggregate.Cells(destRow, 3).Value = systemText
        
        ' ワード関連のオブジェクトを解放
        wordDoc.Close SaveChanges:=False
        wordApp.Quit
        Set wordDoc = Nothing
        Set wordApp = Nothing
    Next row
End Sub



CMd
Sub ExecuteCommands()

    Dim cmdCell As Range
    Dim cmd As String
    Dim objShell As Object
    
    ' コマンド実行用のシェルオブジェクトを作成
    Set objShell = CreateObject("WScript.Shell")
    
    ' A列の各セルに対してコマンドを実行
    For Each cmdCell In ThisWorkbook.Sheets("シート名").Range("A1:A" & ThisWorkbook.Sheets("シート名").Cells(Rows.Count, 1).End(xlUp).Row)
        cmd = cmdCell.Value
        If cmd <> "" Then
            ' コマンドを実行
            objShell.Run "cmd /c " & cmd, 1, True
        End If
    Next cmdCell
    Set objShell = Nothing
End Sub
