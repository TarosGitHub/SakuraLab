'**
'* 選択したフォルダ配下をサクラエディタのgrep機能で検索する
'*
Option Explicit

' PCにインストールされているサクラエディタ(sakura.exe)のパス
Const APP_SAKURA_PATH = """C:\Program Files\sakura\sakura.exe"""

' GREPモードで立ち上げたサクラエディタのダイアログの[ファイル]に指定されるファイル名
Const APP_SAKURA_GREP_FILE = """*.c;*.cpp;*.h"""

' GREPモードで立ち上げたサクラエディタのダイアログの設定オプション
' サクラエディタのヘルプ「コマンドラインオプション」の「-GOPTのオプション」を参照
Const APP_SAKURA_GREP_OPT = "SLPW"

' 立ち上げたサクラエディタウィンドウの状態
Const WNDW_HIDE               = 0  ' ウィンドウを非表示
Const WNDW_NORMAL_FOCUS       = 1  ' 通常のウィンドウ、かつ最前面のウィンドウ
Const WNDW_MINIMAZED_FORCUS   = 2  ' 最小化、かつ最前面ウィンドウ
Const WNDW_MAXMIZED_FOCUS     = 3  ' 最大化、かつ最前面ウィンドウ
Const WNDW_NORMAL_NO_FOCUS    = 4  ' 通常のウィンドウ、ただし最前面にならない
Const WNDW_MINIMIZED_NO_FOCUS = 6  ' 最小化、ただし最前面にならない

'**
'* サクラでGrepのメイン処理。選択したフォルダのフォルダ名を取得し、それを引数に関数sakura_de_grepをコールし、後処理を行う。
'* サクラでGrepに渡されたフォルダ名が存在しない場合はエラーダイアログを表示する。
'*
'* @return 常に正常(0)を返す。
'*
Function sakura_de_grep_main()
	Dim obj_args    ' WshArgumentsオブジェクト(引数のコレクション)

	' 前処理
	Set obj_args = WScript.Arguments

	' メイン処理
	If 0 = obj_args.Count Then
		' 「サクラでGrep.vbs」に渡されたフォルダ名が存在しない場合、エラーダイアログを表示
		Call MsgBox("Error: Please select a folder.", 0, "サクラでGrep")
	Else
		' サクラエディタをGREPモードで立ち上げる
		Call sakura_de_grep(obj_args)
	End If

	' 後処理
	Set obj_args = Nothing

	sakura_de_grep_main = 0
End Function

'**
'* 指定されたフォルダ名が設定された状態で、GREPモードでサクラエディタを立ち上げる。
'*
'* @param obj_args [送る]で「サクラでgrep.vbs」に渡されたフォルダ名
'* @return 常に正常(0)を返す
'*
Function sakura_de_grep(ByRef obj_args)
	Dim app         ' 実行するアプリケーションとそのオプション
	Dim fld_name    ' フォルダ名

	' フォルダ名の設定
	fld_name = """" & obj_args(0) & """"

	' サクラエディタのオプションの設定
	app = APP_SAKURA_PATH _
	        & " -GREPMODE" _
	        & " -GFILE=" & APP_SAKURA_GREP_FILE _
	        & " -GFOLDER=" & fld_name _
	        & " -GOPT=" & APP_SAKURA_GREP_OPT _
	        & " -GREPDLG"

	' サクラエディタの実行
	Call CreateObject("WScript.Shell").Run(app, WNDW_NORMAL_FOCUS, False)

	sakura_de_grep = 0
End Function

'******************************
'* Call Main Routine
'******************************
Call sakura_de_grep_main()
