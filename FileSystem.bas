Attribute VB_Name = "FileSystem"
'
'�e�f�B���N�g���̃p�X����Ԃ�
'
Public Function getParentPath(ByVal path As String) As Variant
    
    If Not (isEnablePath(path)) Then '�����ȃp�X�̏ꍇ
        getParentPath = CVErr(xlErrValue)
        Exit Function
    
    End If
    
    getParentPath = CreateObject("Scripting.FileSystemObject").getParentFolderName(path)
    
End Function

'
'���΃p�X�w����܂񂾃p�X������A�t���p�X��Ԃ�
'
'�����ȃp�X�w��̏ꍇ�́A#VALUE!��Ԃ�
'
Public Function getAbusolutePath(ByVal currentPath As String) As Variant
    
    Dim tmpStr As String
    
    getAbusolutePath = CVErr(xlErrValue) '#VALUE!��ݒ�(��)
    
    If changePath(currentPath) Then  '�J�����g�p�X���ړ�&���������ꍇ
        
        tmpStr = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") '��΃p�X�����擾
        
        If isEnablePath(tmpStr) Then '�L���p�X�̏ꍇ
            getAbusolutePath = tmpStr '��΃p�X���i�[
            
        End If
        
    End If
    
End Function

'
'�t�@�C�������t�H���_���𒊏o����
'
Public Function getLastNameInPath(ByVal path As String) As Variant
    
    getLastNameInPath = Mid(path, InStrRev(path, "\") + 1)
    
End Function

'
'�t�@�C���T�C�Y[Bytes]��Ԃ�
'
'�����p�X�̏ꍇ�́A#VALUE! ��Ԃ�
'
Public Function getFileSize(fileNamePath As String) As Variant

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '�L���p�X�m�F
    If Not (fso.FileExists(fileNamePath)) Then '�����p�X�̏ꍇ
        getFileSize = CVErr(xlErrValue) '#VALUE!��Ԃ�
        Exit Function
        
    End If
    
    getFileSize = fso.GetFile(fileNamePath).Size
    
    Set fso = Nothing
    
End Function

'
'�f�B���N�g�����ǂ���(TRUE/FALSE)��Ԃ�
'
Public Function isDirectory(ByVal path As String) As Variant
    
    isDirectory = CreateObject("Scripting.FileSystemObject").FolderExists(path) 'FolderExists���\�b�h�̖߂�l�����̂܂ܕԂ�
    
End Function

'
'�t�@�C�����ǂ���(TRUE/FALSE)��Ԃ�
'
Public Function isFile(ByVal path As String) As Variant
    
    isFile = CreateObject("Scripting.FileSystemObject").FileExists(path) 'FolderExists���\�b�h�̖߂�l�����̂܂ܕԂ�
    
End Function


'
'�L���p�X���ǂ���(TRUE/FALSE)��Ԃ�
'
Public Function isEnablePath(ByVal path As String) As Variant
    
    Dim isDirectory As Boolean
    Dim isFile As Boolean
    
    isDirectory = CreateObject("Scripting.FileSystemObject").FolderExists(path) 'FolderExists���\�b�h�̖߂�l�����̂܂ܕԂ�
    isFile = CreateObject("Scripting.FileSystemObject").FileExists(path) 'FileExists���\�b�h�̖߂�l�����̂܂ܕԂ�
    
    isEnablePath = (isDirectory Or isFile)
    
End Function

'
'�w��p�X�Ɉړ�����
'
'�ړ��ɐ��������ꍇ��'TRUE', ���s�����ꍇ��'FALSE'��Ԃ�
'
Private Function changePath(FolderName As String) As Variant '�����̓t���p�X
    
    changePath = True '�������i�[(��)
    
    On Error GoTo ERR
    
    '�J�����g�f�B���N�g����ύX
    If Left(FolderName, 2) = "\\" Then '�ŏ���2������\\�̏ꍇ�i�l�b�g���[�N�̏ꍇ�j
        'WSH�Ńf�B���N�g����ύX
        CreateObject("WScript.Shell").CurrentDirectory = FolderName
        
    Else '���[�J���h���C�u�̏ꍇ
        'ChDrive��ChDir�ŃJ�����g�h���C�u�ƃJ�����g�f�B���N�g����ύX
        ChDrive Left(FolderName, 1)
        ChDir FolderName
        
    End If
        
    On Error GoTo 0 '"Goto ERR"��Ԃ��N���A
    Exit Function
    
ERR:
    changePath = False '���s���i�[
    Resume Next '�G���[���o�Ă��������Ď��̍s��

End Function

