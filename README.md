Sub メッセージ確認()

    Dim wsMessage As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastRowMessage As Long
    Dim lastRowTemplate As Long
    Dim i As Long, j As Long
    
    ' ワークシートを設定
    Set wsMessage = ThisWorkbook.Sheets("メッセージ確認シート")
    Set wsTemplate = ThisWorkbook.Sheets("templateシート")
    
    ' メッセージ確認シートの最終行を取得
    lastRowMessage = wsMessage.Cells(wsMessage.Rows.Count, "B").End(xlUp).Row
    
    ' templateシートの最終行を取得
    lastRowTemplate = wsTemplate.Cells(wsTemplate.Rows.Count, "B").End(xlUp).Row
    
    ' メッセージ確認シートのB列からホスト名を取得し、templateシートで一致する行を探す
    For i = 2 To lastRowMessage ' B列の2行目から始める（1行目はヘッダー）
        Dim hostName As String
        hostName = wsMessage.Cells(i, "B").Value
        
        ' メッセージ確認シートのホスト名が未入力の場合、ループを終了
        If hostName = "" Then Exit For
        
        For j = 2 To lastRowTemplate ' templateシートの2行目から始める（1行目はヘッダー）
            ' ホスト名が一致する場合
            If hostName = wsTemplate.Cells(j, "B").Value Then
                ' メッセージ確認シートのC列にメッセージをコピー
                wsMessage.Cells(i, "C").Value = wsTemplate.Cells(j, "C").Value
                Exit For ' 一致したらループを終了
            End If
        Next j
    Next i
    
    ' マクロの終了メッセージ
    MsgBox "処理が完了しました。"
    
End Sub