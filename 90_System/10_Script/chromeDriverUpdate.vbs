    '-----------------------------------------------------------------------
' Selenium��chromedriver.exe�������ōX�V����X�N���v�g
' �X�V��:2020-04-13
' ��SeleniumBasic��Python���ȂǁA������Chromedriver����x�ɍX�V�ł���悤�ɂ��܂����B
' ���O���A�v������̌Ăяo���ɍۂ��ă��b�Z�[�W�|�b�v�A�b�v���I�t�ɂł���悤�ɂ��܂����B
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
    chromeDir = "C:\Program Files (x86)\Google\Chrome\Application"      'chrome.exe�̂���f�B���N�g��(GoogleChrome�̃t�H���_)
    'seleniumDirs = Array(appdata, "C:\Program Files\SeleniumBasic")     '��chromedriver.exe�̃f�B���N�g���B������(Selenium�̃t�H���_)
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
'chromedriver���_�E�����[�h����֐�
'version : �_�E�����[�h�������o�[�W�����B��)�u78.0.3904.105�v
'destDir : �V����chromedriver��u���f�B���N�g���B
Sub chromedriverDL(version, destDir, msgbox_is)
    '��)�u78.0.3904.105�v= [���W���[].[�}�C�i�[].[�r���h].[���r�W����]�ǂ��܂ň�v�����邩
    If version = "" Then Exit Sub
    dots = Split(version, ".")
    version = Left(version, Len(version) - Len(dots(UBound(dots))) - 1)     '"�r���h�܂ň�v"�w��F(�݊������B���萫��)
    version = Left(version, Len(version) - Len(dots(UBound(dots))) - 1)     '"�}�C�i�[�܂ň�v"�w��F(�݊�����B���萫��)�����̍s���R�����g�A�E�g��"�r���h�܂ň�v"�ɕύX�B

    Set objShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set httpReq = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Set httpReq = CreateObject("MSXML.XMLHTTPRequest")
    End If

    Const zipName = "chromedriver_win32.zip"
    workDir = objShell.CurrentDirectory & "\" & zipName         '����vbs�̏ꏊ����ƃt�H���_��
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
    '�X�V��������(�o�[�W����)���e�L�X�g�ɋL�^����
    FileName = destDir & "\version.txt"
    Set tso = fso.OpenTextFile(FileName, 2, True)
    tso.write (version)
    tso.Close
    '----------------------------------------------
    If msgbox_is = True Then MsgBox "updated chromedriver : " & version
End Sub
'=======================================================================
'.zip�t�@�C�����𓀂���֐�
'sourcePath : �𓀂�����zip�̃t�@�C���p�X
'destPath : �W�J�����t�@�C����u���f�B���N�g��
Sub unzip(sourcePath, destDir)
    Const FOF_SILENT = &H4                '�i���_�C�A���O��\�����Ȃ��B
    Const FOF_NOCONFIRMATION = &H10       '�㏑���m�F�_�C�A���O��\�����Ȃ��i[���ׂď㏑��]�Ɠ����j�B

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(destDir)
    Set FilesInZip = objShell.Namespace(sourcePath).items

    '��
    If (Not objFolder Is Nothing) Then
       objFolder.CopyHere FilesInZip, FOF_NOCONFIRMATION + FOF_SILENT
    End If
End Sub
'=======================================================================
'�C���X�g�[������Ă���GoogleChrome�̃o�[�W�������`�F�b�N����
'chromeDir : chrome.exe�̂���f�B���N�g���B(��)"C:\Program Files (x86)\Google\Chrome\Application"
'destDir : �V����chromedriver��u���f�B���N�g���B
Function chkVersion(chromeDir, destDir, msgbox_is)
    '-------------------------------------------------------------
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getFolder(chromeDir)

    '�T�u�t�H���_�ꗗ
    version = ""
    For Each subfolder In folder.subfolders
        dots = Split(subfolder.name, ".")
        If UBound(dots) > 2 Then
            version = subfolder.name
            Exit For
        End If
    Next

    '�G���[�`�F�b�N
    If version = "" Then
        MsgBox "���݂�Chrome��version���擾�ł��܂���ł����B" & vbCrLf & "�I�����܂��B"
        WScript.Quit -1
    End If
    '-------------------------------------------------------------
    '����vbs�ɂ���čŋ߃A�b�v�f�[�g����chromedriver�̃o�[�W������
    '����PC��GoogleChrome�̃o�[�W�������r���āA���������A�b�v�f�[�g���K�v�Ȃ̂����f����
    FileName = destDir & "\version.txt"        'C:\Users\TEST\AppData\Local\SeleniumBasic\version.txt

    If fso.FileExists(FileName) Then
        Set fp = fso.OpenTextFile(FileName, 1)
        curVersion = fp.ReadAll
        fp.Close

        '���W���[.�}�C�i�[.�r���h.���r�W����
        dots = Split(curVersion, ".")
        If UBound(dots) > 2 Then
            buildver = Left(curVersion, Len(curVersion) - Len(dots(UBound(dots))) - 1)
            minorver = Left(buildver, Len(buildver) - Len(dots(UBound(dots) - 1)) - 1)

            If InStr(version, buildver) > 0 Then
                If msgbox_is Then
                    MsgBox destDir & "��" & vbCrLf & "chromedriver�̓r���h�o�[�W�����܂œ����ł�" & vbCrLf & "�A�b�v�f�[�g�̕K�v�͂���܂���"
                    chkVersion = ""
                    Exit Function
                End If
            ElseIf InStr(version, minorver) > 0 Then
                If msgbox_is Then
                    MsgBox destDir & "��" & vbCrLf & "chromedriver�̓}�C�i�[�o�[�W�����܂œ����ł�" & vbCrLf & "�A�b�v�f�[�g�̕K�v�͂���܂���"
                    chkVersion = ""
                    Exit Function
                End If
            Else
                If msgbox_is = True Then MsgBox "�œK��Chromedriver���_�E�����[�h����K�v������܂�:" & version
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
'DIR�֐�(�p�X�̑��ݔ��������:VBA�ɂ�����VBS�ɂȂ��֐�)
'path : ���݂�t�@�C�������`�F�b�N�������p�X
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