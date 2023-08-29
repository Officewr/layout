す


Sub フォルダ一覧取得()

    Dim フォルダパス As String
    Dim 出力シート As Worksheet
    Dim フォルダ一覧() As Variant
    Dim ファイル As Object
    Dim サブフォルダ As Object
    Dim 行番号 As Long
    
    ' フォルダのパスを指定
    フォルダパス = "フォルダのパス"
    
    ' 結果を出力するシートを選択（新しいシートを作成する場合は Sheets.Add）
    Set 出力シート = ThisWorkbook.Sheets("シート名")
    
    ' 初期化
    行番号 = 1
    
    ' フォルダとサブフォルダの一覧を取得
    Set フォルダ一覧 = GetFolders(フォルダパス)
    
    ' ヘッダを出力
    出力シート.Cells(行番号, 1).Value = "フォルダ名"
    出力シート.Cells(行番号, 2).Value = "サイズ"
    行番号 = 行番号 + 1
    
    ' フォルダの一覧を出力
    For Each フォルダ In フォルダ一覧
        出力シート.Cells(行番号, 1).Value = フォルダ.Path
        出力シート.Cells(行番号, 2).Value = フォルダ.Size
        行番号 = 行番号 + 1
    Next フォルダ
    
    ' フォルダとサブフォルダの一覧を再帰的に取得する関数
    Function GetFolders(ByVal パス As String) As Variant
        Dim フォルダ As Object
        Dim サブフォルダ As Object
        Dim フォルダ一覧 As Collection
        
        Set フォルダ一覧 = New Collection
        フォルダ一覧.Add CreateObject("Scripting.FileSystemObject").GetFolder(パス)
        
        For Each フォルダ In フォルダ一覧
            For Each サブフォルダ In フォルダ.SubFolders
                フォルダ一覧.Add サブフォルダ
            Next サブフォルダ
        Next フォルダ
        
        Set GetFolders = フォルダ一覧
    End Function
    
End Sub



20230829
Sub マクロを実行()
    Dim mainFolder As String
    Dim ws As Worksheet
    Dim folderName As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim table As Object
    Dim newRow As Long
    
    ' シート名 "Sheet1" を集計シート名に適宜変更
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 格納先フォルダを取得
    mainFolder = Range("A1").Value
    
    ' フォルダ名行
    newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    folderName = Mid(mainFolder, InStrRev(mainFolder, "\") + 1)
    ws.Cells(newRow, 1).Value = folderName
    
    ' Wordアプリケーションを起動
    Set wordApp = CreateObject("Word.Application")
    
    ' フォルダ内のファイルを処理
    fileName = Dir(mainFolder & "\*.docx")
    Do While fileName <> ""
        ' Wordドキュメントを開く
        Set wordDoc = wordApp.Documents.Open(mainFolder & "\" & fileName)
        
        ' DAVコピー
        For Each para In wordDoc.Paragraphs
            If InStr(para.Range.Text, "DAV") > 0 Then
                ws.Cells(newRow, 2).Value = para.Range.Text
                Exit For
            End If
        Next para
        
        ' 表コピー
        For Each table In wordDoc.Tables
            ws.Cells(newRow, 3).Value = table.Columns(1).Cells(2).Range.Text
            Exit For
        Next table
        
        ' 次のファイルへ
        wordDoc.Close
        fileName = Dir
    Loop
    
    ' Wordアプリケーションを終了
    wordApp.Quit
    
    MsgBox "処理が完了しました。"
End Sub



Sub マクロを実行()

    Dim ws As Worksheet
    Dim folderRange As Range
    Dim mainFolder As String
    Dim folderCell As Range
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim table As Object
    Dim newRow As Long
    
    ' シート名 "Sheet1" を集計シート名に適宜変更
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' フォルダ一覧の範囲を取得（A1からA20まで）
    Set folderRange = ws.Range("A1:A20")
    
    ' Wordアプリケーションを起動
    Set wordApp = CreateObject("Word.Application")
    
    ' 各フォルダの内容を処理
    For Each folderCell In folderRange
        mainFolder = folderCell.Value
        If mainFolder <> "" Then
            ' フォルダ名行
            newRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(newRow, 1).Value = Mid(mainFolder, InStrRev(mainFolder, "\") + 1)
            
            ' フォルダ内のファイルを処理
            fileName = Dir(mainFolder & "\*.docx")
            Do While fileName <> ""
                ' Wordドキュメントを開く
                Set wordDoc = wordApp.Documents.Open(mainFolder & "\" & fileName)
                
                ' DAVコピー
                For Each para In wordDoc.Paragraphs
                    If InStr(para.Range.Text, "DAV") > 0 Then
                        ws.Cells(newRow, 2).Value = para.Range.Text
                        Exit For
                    End If
                Next para
                
                ' 表コピー
                For Each table In wordDoc.Tables
                    ws.Cells(newRow, 3).Value = table.Columns(1).Cells(2).Range.Text
                    Exit For
                Next table
                
                ' 次のファイルへ
                wordDoc.Close
                fileName = Dir
            Loop
        End If
    Next folderCell
    
    ' Wordアプリケーションを終了
    wordApp.Quit
    
    MsgBox "処理が完了しました。"
End Sub

