'*************************************************************
' linkファイルの元データをコピーする
'   対象はフォルダやファイルをLinkファイルと同じフォルダにコピー
'
' 2020/03/21 v2 shibuyano フォルダコピー機能の追加
' 2020/03/20 v1 shibuyano ファイルコピー機能
'*************************************************************

Option Explicit

'--- 定数の宣言 ---
Dim args              ' 引数
Dim objFSO            ' "Scripting.FIlesystemObject"
' Dim objFile           ' 指定ファイルのオブジェクト
Dim objWsh            ' "WScript.Shell"
Dim ObjSc             '

Dim strTargetFile     ' ターゲットファイルのフルパス
Dim strTargetFilePath ' ターゲットファイルが保存されているフォルダ
Dim strLnkFile        ' Linkファイルのフルパス
Dim strLnkFilePath    ' Linkファイルが保存されているフォルダ
Dim strTargetFileInLnkFolder

Dim strMessage
Dim arg              ' for loop
Dim Rt
Dim ExplorerOpen

' 前処理１：引数がない場合は終了
Set args = WScript.Arguments
If args.Count < 1 Then
  MsgBox "Drag and Dropt msg file to This vbscript.", vbExclamation + vbSystemModal
  WScript.Quit
End If

' Linkファイルが 5 未満時は終了後にExploreを開く
If args.Count > 5 Then
    ExplorerOpen = False
Else
    ExplorerOpen = True
End If

'--- OBJ宣言 ---
Set objFSO = CreateObject("Scripting.FIlesystemObject")
strMessage = ""   ' エラーメッセージの初期化

'Dropされたファイルを順に処理
For Each arg In WScript.Arguments
    ' arg は処理するファイルのフルパス

    ' 拡張子 lnk かつ，ファイル/フォルダが存在していること。ファイルが存在しない場合はスルー
    If ( (objFSO.FileExists(arg) Or objFSO.FolderExists(arg)) And (StrComp(objFSO.GetExtensionName(arg), "lnk" ,1 ) = 0)) Then
        ' WScript.Echo "リンク先のファイルをコピーします。"
        ' Shortcut ファイル内のtargetfile情報を取得
        Set objWsh = WScript.CreateObject("WScript.Shell")
        Set ObjSc = objWsh.CreateShortcut(arg)                      ' arg で渡されたショートカットを取得

        strTargetFile = ObjSc.TargetPath                            ' リンク先のフルパス
        strTargetFilePath = objFSO.GetParentFolderName(strTargetFile)
        strLnkFile = ObjSc.FullName                                 ' リンクファイルのパス
        strLnkFilePath = objFSO.GetParentFolderName(strLnkFile)     ' リンクファイルが保存されているフォルダ

        ' コピー先ファイルのFullPathを作成： リンクファイルのフォルダpath + リンク先ファイル名
        strTargetFileInLnkFolder = strLnkFilePath  & mid(strTargetFile, Len(strTargetFilePath)+1)

        ' デバッグ用
        ' WScript.Echo strTargetFileInLnkFolder    ' デバッグ用
        ' strMessage = "リンク先のフルパス 　　：" & strTargetFile & vbCrLf & _
        '              "リンク先の親フォルダ　 ：" & strTargetFilePath & vbCrLf & _
        '              "Lnkファイルのフルパス　：" & strLnkFile & vbCrLf & _
        '              "Lnkファイルの親フォルダ：" & strLnkFilePath & vbCrLf
        ' WScript.Echo strMessage

        ' リンク先ファイルの存在確認
        if objFSO.FileExists(strTargetFile) = True Then
            ' ターゲットの ***ファイル*** がある場合，ファイルをコピー
            Rt = objFSO.CopyFile(strTargetFile, strLnkFilePath & "/" , False)
            if ExplorerOpen then
                Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFile & chr(34),  1, True )
                ' Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFileInLnkFolder & chr(34),  1, True )
            End If

        Elseif objFSO.FolderExists(strTargetFile) = True then
            ' ターゲットの ***フォルダ*** がある場合.フォルダをコピー
            Rt = objFSO.CopyFolder(strTargetFile, strLnkFilePath & "/")
            if ExplorerOpen then
                Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFile & chr(34),  1, True )
                ' Rt = objWsh.Run("explorer.exe /select," & chr(34) & strTargetFileInLnkFolder & chr(34),  1, True )
            End If
        Else
            ' ターゲット先がない場合のメッセージ作成
            strMessage = strMessage & vbCrLf & _
                        "★ターゲット先がありません" & vbCrLf & _
                        "  リンク先のフルパス 　　：" & strTargetFile & vbCrLf & _
                        "  Lnkファイルのフルパス　：" & strLnkFile

        End If
        ' オブジェクトの開放
        set ObjSc = nothing
    Else
        strMessage = strMessage & vbCrLf & _
                    "★Dropされたファイルがlinkファイルではない。または，リンク先のファイル／フォルダが存在しません。" & _
                    "  Lnkファイルのフルパス　：" & strLnkFile
    End If

Next

if strMessage <> "" then
    WScript.Echo strMessage
Else
    WScript.Echo args.Count & "個のファイルのコピーが完了しました"
end if

' オブジェクトの開放
set objWsh = nothing
set objFSO = nothing

