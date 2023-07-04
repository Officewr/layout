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
