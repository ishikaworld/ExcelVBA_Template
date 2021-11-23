    '-----------------------------------------------------------------------
' Seleniumのchromedriver.exeを自動で更新するスクリプト
' 更新日:2020-04-13
' ★SeleniumBasicやPython環境など、複数のChromedriverを一度に更新できるようにしました。
' ★外部アプリからの呼び出しに際してメッセージポップアップをオフにできるようにしました。
'-----------------------------------------------------------------------
Call update_driver
Sub update_driver()
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set objShell = CreateObject("WScript.Shell")
    Set objShell = WScript.CreateObject("WScript.Shell")
    If destDir = "" Then
        appdata = objShell.ExpandEnvironmentStrings("%APPDATA%")            'C:\Users\TEST\AppData\Roaming
        appdata = fso.GetParentFolderName(appdata)                          'C:\Users\TEST\AppData
        appdata = appdata & "\Local\SeleniumBasic"                          'C:\Users\TEST\AppData\Local\SeleniumBasic
    End If
    '-------------------------------------------------------------
    'USER SETTING
    chromeDir = "C:\Program Files (x86)\Google\Chrome\Application"      'chrome.exeのあるディレクトリ(GoogleChromeのフォルダ)
    'seleniumDirs = Array(appdata, "C:\Program Files\SeleniumBasic")     '★chromedriver.exeのディレクトリ。複数可(Seleniumのフォルダ)
    seleniumDirs = Array(appdata)
    
    '-------------------------------------------------------------
    For Each seleniumDir In seleniumDirs
        If myDir(seleniumDir) <> "" Then
            version = chkVersion(chromeDir, seleniumDir, msgbox_is)
            Call chromedriverDL(version, seleniumDir, msgbox_is)
        End If
    Next
End Sub
'=======================================================================
'chromedriverをダウンロードする関数
'version : ダウンロードしたいバージョン。例)「78.0.3904.105」
'destDir : 新しいchromedriverを置くディレクトリ。
Sub chromedriverDL(version, destDir, msgbox_is)
    '例)「78.0.3904.105」= [メジャー].[マイナー].[ビルド].[リビジョン]どこまで一致させるか
    If version = "" Then Exit Sub
    dots = Split(version, ".")
    version = Left(version, Len(version) - Len(dots(UBound(dots))) - 1)     '"ビルドまで一致"指定：(互換性高。入手性低)
    version = Left(version, Len(version) - Len(dots(UBound(dots))) - 1)     '"マイナーまで一致"指定：(互換性低。入手性高)※この行をコメントアウトで"ビルドまで一致"に変更。

    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Set httpReq = CreateObject("MSXML.XMLHTTPRequest")
    End If

    Const zipName = "chromedriver_win32.zip"
    workDir = objShell.CurrentDirectory & "\" & zipName         'このvbsの場所を作業フォルダに
    'workDir = ThisWorkbook.path & "\" & zipName

    httpReq.Open "GET", "https://chromedriver.storage.googleapis.com", False
    httpReq.Send

    txtS = InStr(httpReq.responseText, version)
    txtE = InStr(txtS, httpReq.responseText, "/chromedriver_linux64.zip</Key>")
    version = Mid(httpReq.responseText, txtS, txtE - txtS)

    URL = "https://chromedriver.storage.googleapis.com/" & version & "/chromedriver_win32.zip"

    httpReq.Open "GET", URL, False
    httpReq.Send

    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.write httpReq.responseBody
    objStream.SaveToFile workDir, 2
    objStream.Close

    Call unzip(workDir, destDir)
    fso.DeleteFile workDir

    Set objStream = Nothing
    Set httpReq = Nothing
    '----------------------------------------------
    '更新した履歴(バージョン)をテキストに記録する
    FileName = destDir & "\version.txt"
    Set tso = fso.OpenTextFile(FileName, 2, True)
    tso.write (version)
    tso.Close
    '----------------------------------------------
    If msgbox_is = True Then MsgBox "updated chromedriver : " & version
End Sub
'=======================================================================
'.zipファイルを解凍する関数
'sourcePath : 解凍したいzipのファイルパス
'destPath : 展開したファイルを置くディレクトリ
Sub unzip(sourcePath, destDir)
    Const FOF_SILENT = &H4                '進捗ダイアログを表示しない。
    Const FOF_NOCONFIRMATION = &H10       '上書き確認ダイアログを表示しない（[すべて上書き]と同じ）。

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(destDir)
    Set FilesInZip = objShell.Namespace(sourcePath).items

    '解凍
    If (Not objFolder Is Nothing) Then
       objFolder.CopyHere FilesInZip, FOF_NOCONFIRMATION + FOF_SILENT
    End If
End Sub
'=======================================================================
'インストールされているGoogleChromeのバージョンをチェックする
'chromeDir : chrome.exeのあるディレクトリ。(例)"C:\Program Files (x86)\Google\Chrome\Application"
'destDir : 新しいchromedriverを置くディレクトリ。
Function chkVersion(chromeDir, destDir, msgbox_is)
    '-------------------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getFolder(chromeDir)

    'サブフォルダ一覧
    version = ""
    For Each subfolder In folder.subfolders
        dots = Split(subfolder.name, ".")
        If UBound(dots) > 2 Then
            version = subfolder.name
            Exit For
        End If
    Next

    'エラーチェック
    If version = "" Then
        MsgBox "現在のChromeのversionが取得できませんでした。" & vbCrLf & "終了します。"
        WScript.Quit -1
    End If
    '-------------------------------------------------------------
    'このvbsによって最近アップデートしたchromedriverのバージョンと
    'このPCのGoogleChromeのバージョンを比較して、そもそもアップデートが必要なのか判断する
    FileName = destDir & "\version.txt"        'C:\Users\TEST\AppData\Local\SeleniumBasic\version.txt

    If fso.FileExists(FileName) Then
        Set fp = fso.OpenTextFile(FileName, 1)
        curVersion = fp.ReadAll
        fp.Close

        'メジャー.マイナー.ビルド.リビジョン
        dots = Split(curVersion, ".")
        If UBound(dots) > 2 Then
            buildver = Left(curVersion, Len(curVersion) - Len(dots(UBound(dots))) - 1)
            minorver = Left(buildver, Len(buildver) - Len(dots(UBound(dots) - 1)) - 1)

            If InStr(version, buildver) > 0 Then
                If msgbox_is Then
                    MsgBox destDir & "の" & vbCrLf & "chromedriverはビルドバージョンまで同じです" & vbCrLf & "アップデートの必要はありません"
                    chkVersion = ""
                    Exit Function
                End If
            ElseIf InStr(version, minorver) > 0 Then
                If msgbox_is Then
                    MsgBox destDir & "の" & vbCrLf & "chromedriverはマイナーバージョンまで同じです" & vbCrLf & "アップデートの必要はありません"
                    chkVersion = ""
                    Exit Function
                End If
            Else
                If msgbox_is = True Then MsgBox "最適なChromedriverをダウンロードする必要があります:" & version
            End If
        End If
    Else
        Set tso = fso.OpenTextFile(FileName, 2, True)
        tso.write ("0.0.0.0")
        tso.Close
    End If
    '-------------------------------------------------------------
    chkVersion = version
End Function
'=======================================================================
'DIR関数(パスの存在判定をする:VBAにあってVBSにない関数)
'path : 存在やファイル名をチェックしたいパス
Function myDir(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then
        myDir = path
    Else
        myDir = ""
    End If
    Set fso = Nothing
End Function
'=======================================================================