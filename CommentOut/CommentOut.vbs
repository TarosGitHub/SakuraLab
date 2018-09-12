'**
'* 選択文字列をコメントアウトする
'*
Option Explicit

'**
'* @fn        CommentOutInBlock
'* @brief     選択文字列をコメントアウト
'* @param     なし
'* @return    なし
'* @details   選択している文字列をコメントアウトする。<br/>
'*            文字列を選択していない場合はコメントブロックを作成し、その中央にカーソルを移動する。
'*
Function CommentOutInBlock()
	Dim slct_str	' 選択文字列
	Dim file_ext	' ファイル拡張子
	Dim cmnt_str	' コメントアウト後の文字列
	
	' 選択文字列取得
	slct_str = GetSelectedString()
	
	' ファイル拡張子取得
	file_ext = GetFilenameExtension()
	
	' コメントアウト後の文字列取得
	cmnt_str = GetCommentOutStr(slct_str, file_ext)
	
	If cmnt_str = "" Then
		Exit Function
	End If
	
	' コメントアウト
	InsText(cmnt_str)
	
	' 文字列を選択していない場合、中央にカーソルを移動
	If slct_str = "" Then
		MoveCursor(file_ext)
	End If
End Function

'**
'* @fn        GetFilenameExtension
'* @brief     ファイル拡張子取得
'* @param     なし
'* @return    ファイル拡張子
'* @details   ファイル拡張子を取得する。
'*
Function GetFilenameExtension()
	Dim file_sys_obj	' ファイルシステムオブジェクト
	Dim file_path		' ファイルフルパス
	Dim file_ext		' ファイル拡張子
	
	' ファイルシステムオブジェクト作成
	Set file_sys_obj = CreateObject("Scripting.FileSystemObject")
	
	' ファイルフルパス取得
	file_path = ExpandParameter("$F")
	
	' ファイル拡張子取得
	file_ext = file_sys_obj.GetExtensionName(file_path)
	
	Set file_sys_obj = Nothing
	GetFilenameExtension = file_ext
End Function

'**
'* @fn        GetCommentOutStr
'* @brief     文字列をコメントアウト
'* @param     str 文字列
'* @param     file_ext ファイル拡張子
'* @return    コメントアウト後の文字列
'* @details   文字列をコメントアウトし、その文字列を返す。
'*
Function GetCommentOutStr(str, file_ext)
	Dim cout_str	' コメントアウト後の文字列
	
	Select Case file_ext
		Case "c", "cpp", "h"
			cout_str = "/* " & str & " */"
			
		Case "xml", "html"
			cout_str = "<!-- " & str & " -->"
			
		Case Else
			cout_str = ""
	End Select
	
	GetCommentOutStr = cout_str
End Function

'**
'* @fn        MoveCursor
'* @brief     カーソルの移動
'* @param     file_ext ファイル拡張子
'* @return    なし
'* @details   カーソルを移動する。ファイル拡張子によって移動距離を変える。
'*
Function MoveCursor(file_ext)
	Dim i
	
	Select Case file_ext
		Case "c", "cpp", "h"
			For i = 1 To 3
				Editor.Left()
			Next
			
		Case "xml", "html"
			For i = 1 To 4
				Editor.Left()
			Next
			
		Case Else
			' DO NOTHING
	End Select
End Function

'******************************
'* Call Main Routine
'******************************
Call CommentOutInBlock()

