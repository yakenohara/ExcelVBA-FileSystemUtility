Attribute VB_Name = "FileSystem"
'
'親ディレクトリのパス名を返す
'
Public Function getParentPath(ByVal path As String) As Variant
    
    If Not (isEnablePath(path)) Then '無効なパスの場合
        getParentPath = CVErr(xlErrValue)
        Exit Function
    
    End If
    
    getParentPath = CreateObject("Scripting.FileSystemObject").getParentFolderName(path)
    
End Function

'
'相対パス指定を含んだパス名から、フルパスを返す
'
'無効なパス指定の場合は、#VALUE!を返す
'
Public Function getAbusolutePath(ByVal currentPath As String) As Variant
    
    Dim tmpStr As String
    
    getAbusolutePath = CVErr(xlErrValue) '#VALUE!を設定(仮)
    
    If changePath(currentPath) Then  'カレントパスを移動&成功した場合
        
        tmpStr = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") '絶対パス名を取得
        
        If isEnablePath(tmpStr) Then '有効パスの場合
            getAbusolutePath = tmpStr '絶対パスを格納
            
        End If
        
    End If
    
End Function

'
'ファイル名かフォルダ名を抽出する
'
Public Function getLastNameInPath(ByVal path As String) As Variant
    
    getLastNameInPath = Mid(path, InStrRev(path, "\") + 1)
    
End Function

'
'ファイルサイズ[Bytes]を返す
'
'無効パスの場合は、#VALUE! を返す
'
Public Function getFileSize(fileNamePath As String) As Variant

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '有効パス確認
    If Not (fso.FileExists(fileNamePath)) Then '無効パスの場合
        getFileSize = CVErr(xlErrValue) '#VALUE!を返す
        Exit Function
        
    End If
    
    getFileSize = fso.GetFile(fileNamePath).Size
    
    Set fso = Nothing
    
End Function

'
'ディレクトリかどうか(TRUE/FALSE)を返す
'
Public Function isDirectory(ByVal path As String) As Variant
    
    isDirectory = CreateObject("Scripting.FileSystemObject").FolderExists(path) 'FolderExistsメソッドの戻り値をそのまま返す
    
End Function

'
'ファイルかどうか(TRUE/FALSE)を返す
'
Public Function isFile(ByVal path As String) As Variant
    
    isFile = CreateObject("Scripting.FileSystemObject").FileExists(path) 'FolderExistsメソッドの戻り値をそのまま返す
    
End Function


'
'有効パスかどうか(TRUE/FALSE)を返す
'
Public Function isEnablePath(ByVal path As String) As Variant
    
    Dim isDirectory As Boolean
    Dim isFile As Boolean
    
    isDirectory = CreateObject("Scripting.FileSystemObject").FolderExists(path) 'FolderExistsメソッドの戻り値をそのまま返す
    isFile = CreateObject("Scripting.FileSystemObject").FileExists(path) 'FileExistsメソッドの戻り値をそのまま返す
    
    isEnablePath = (isDirectory Or isFile)
    
End Function

'
'指定パスに移動する
'
'移動に成功した場合は'TRUE', 失敗した場合は'FALSE'を返す
'
Private Function changePath(FolderName As String) As Variant '引数はフルパス
    
    changePath = True '成功を格納(仮)
    
    On Error GoTo ERR
    
    'カレントディレクトリを変更
    If Left(FolderName, 2) = "\\" Then '最初の2文字が\\の場合（ネットワークの場合）
        'WSHでディレクトリを変更
        CreateObject("WScript.Shell").CurrentDirectory = FolderName
        
    Else 'ローカルドライブの場合
        'ChDriveとChDirでカレントドライブとカレントディレクトリを変更
        ChDrive Left(FolderName, 1)
        ChDir FolderName
        
    End If
        
    On Error GoTo 0 '"Goto ERR"状態をクリア
    Exit Function
    
ERR:
    changePath = False '失敗を格納
    Resume Next 'エラーが出ても無視して次の行へ

End Function

