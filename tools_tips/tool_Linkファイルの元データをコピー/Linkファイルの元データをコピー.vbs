'*************************************************************
' link�t�@�C���̌��f�[�^���R�s�[����
'   �Ώۂ̓t�H���_��t�@�C����Link�t�@�C���Ɠ����t�H���_�ɃR�s�[
'
' 2020/03/21 v2 shibuyano �t�H���_�R�s�[�@�\�̒ǉ�
' 2020/03/20 v1 shibuyano �t�@�C���R�s�[�@�\
'*************************************************************

Option Explicit

'--- �萔�̐錾 ---
Dim args              ' ����
Dim objFSO            ' "Scripting.FIlesystemObject"
' Dim objFile           ' �w��t�@�C���̃I�u�W�F�N�g
Dim objWsh            ' "WScript.Shell"
Dim ObjSc             '

Dim strTargetFile     ' �^�[�Q�b�g�t�@�C���̃t���p�X
Dim strTargetFilePath ' �^�[�Q�b�g�t�@�C�����ۑ�����Ă���t�H���_
Dim strLnkFile        ' Link�t�@�C���̃t���p�X
Dim strLnkFilePath    ' Link�t�@�C�����ۑ�����Ă���t�H���_
Dim strTargetFileInLnkFolder

Dim strMessage
Dim arg              ' for loop
Dim Rt
Dim ExplorerOpen

' �O�����P�F�������Ȃ��ꍇ�͏I��
Set args = WScript.Arguments
If args.Count < 1 Then
  MsgBox "Drag and Dropt msg file to This vbscript.", vbExclamation + vbSystemModal
  WScript.Quit
End If

' Link�t�@�C���� 5 �������͏I�����Explore���J��
If args.Count > 5 Then
    ExplorerOpen = False
Else
    ExplorerOpen = True
End If

'--- OBJ�錾 ---
Set objFSO = CreateObject("Scripting.FIlesystemObject")
strMessage = ""   ' �G���[���b�Z�[�W�̏�����

'Drop���ꂽ�t�@�C�������ɏ���
For Each arg In WScript.Arguments
    ' arg �͏�������t�@�C���̃t���p�X

    ' �g���q lnk ���C�t�@�C��/�t�H���_�����݂��Ă��邱�ƁB�t�@�C�������݂��Ȃ��ꍇ�̓X���[
    If ( (objFSO.FileExists(arg) Or objFSO.FolderExists(arg)) And (StrComp(objFSO.GetExtensionName(arg), "lnk" ,1 ) = 0)) Then
        ' WScript.Echo "�����N��̃t�@�C�����R�s�[���܂��B"
        ' Shortcut �t�@�C������targetfile�����擾
        Set objWsh = WScript.CreateObject("WScript.Shell")
        Set ObjSc = objWsh.CreateShortcut(arg)                      ' arg �œn���ꂽ�V���[�g�J�b�g���擾

        strTargetFile = ObjSc.TargetPath                            ' �����N��̃t���p�X
        strTargetFilePath = objFSO.GetParentFolderName(strTargetFile)
        strLnkFile = ObjSc.FullName                                 ' �����N�t�@�C���̃p�X
        strLnkFilePath = objFSO.GetParentFolderName(strLnkFile)     ' �����N�t�@�C�����ۑ�����Ă���t�H���_

        ' �R�s�[��t�@�C����FullPath���쐬�F �����N�t�@�C���̃t�H���_path + �����N��t�@�C����
        strTargetFileInLnkFolder = strLnkFilePath  & mid(strTargetFile, Len(strTargetFilePath)+1)

        ' �f�o�b�O�p
        ' WScript.Echo strTargetFileInLnkFolder    ' �f�o�b�O�p
        ' strMessage = "�����N��̃t���p�X �@�@�F" & strTargetFile & vbCrLf & _
        '              "�����N��̐e�t�H���_�@ �F" & strTargetFilePath & vbCrLf & _
        '              "Lnk�t�@�C���̃t���p�X�@�F" & strLnkFile & vbCrLf & _
        '              "Lnk�t�@�C���̐e�t�H���_�F" & strLnkFilePath & vbCrLf
        ' WScript.Echo strMessage

        ' �����N��t�@�C���̑��݊m�F
        if objFSO.FileExists(strTargetFile) = True Then
            ' �^�[�Q�b�g�� ***�t�@�C��*** ������ꍇ�C�t�@�C�����R�s�[
            Rt = objFSO.CopyFile(strTargetFile, strLnkFilePath & "/" , False)
            if ExplorerOpen then
                Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFile & chr(34),  1, True )
                ' Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFileInLnkFolder & chr(34),  1, True )
            End If

        Elseif objFSO.FolderExists(strTargetFile) = True then
            ' �^�[�Q�b�g�� ***�t�H���_*** ������ꍇ.�t�H���_���R�s�[
            Rt = objFSO.CopyFolder(strTargetFile, strLnkFilePath & "/")
            if ExplorerOpen then
                Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFile & chr(34),  1, True )
                ' Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFileInLnkFolder & chr(34),  1, True )
            End If
        Else
            ' �^�[�Q�b�g�悪�Ȃ��ꍇ�̃��b�Z�[�W�쐬
            strMessage = strMessage & vbCrLf & _
                        "���^�[�Q�b�g�悪����܂���" & vbCrLf & _
                        "  �����N��̃t���p�X �@�@�F" & strTargetFile & vbCrLf & _
                        "  Lnk�t�@�C���̃t���p�X�@�F" & strLnkFile

        End If
        ' �I�u�W�F�N�g�̊J��
        set ObjSc = nothing
    Else
        strMessage = strMessage & vbCrLf & _
                    "��Drop���ꂽ�t�@�C����link�t�@�C���ł͂Ȃ��B�܂��́C�����N��̃t�@�C���^�t�H���_�����݂��܂���B" & _
                    "  Lnk�t�@�C���̃t���p�X�@�F" & strLnkFile
    End If

Next

if strMessage <> "" then
    WScript.Echo strMessage
Else
    WScript.Echo args.Count & "�̃t�@�C���̃R�s�[���������܂���"
end if

' �I�u�W�F�N�g�̊J��
set objWsh = nothing
set objFSO = nothing

