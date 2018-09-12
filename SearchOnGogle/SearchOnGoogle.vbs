'**
'* 文字列をGoogleで検索する
'*
Option Explicit

'**
'* @fn        search_on_google
'* @brief     文字列をインターネット(Google)で検索
'* @param     なし
'* @return    なし
'* @details   選択している文字列をインターネットで検索する。<br/>
'*            選択していない場合はカーソル位置の文字列を検索する。
'*
Function search_on_google()
	Dim shell_obj		' Shell.Application
	Dim slct_str		' 選択文字列
	Dim srch_str		' 検索文字列
	
	Set shell_obj = CreateObject("Shell.Application")
	
	' 選択文字列の取得
	slct_str = get_slct_str()
	
	' 検索文字列の作成
	srch_str = "http://www.google.co.jp/search?" & _
				"hl=ja&" & _
				"q=" & slct_str & "&" & _
				"lr=lang_ja&" & _
				"gws_rd=ssl"
	
	' アプリケーション実行
	shell_obj.ShellExecute(srch_str)
	
	Set shell_obj = Nothing
End Function

'**
'* @fn        get_slct_str
'* @brief     単語を選択し、文字列を取得
'* @param     なし
'* @return    なし
'* @details   単語を選択し、文字列を取得。
'*
Function get_slct_str()
	Dim ret_selt_str	' 戻り値
	
	' 選択文字列の取得
	If IsTextSelected Then
		' 文字列が既に選択されている場合
		ret_selt_str = GetSelectedString()
	Else
		' 文字列がまだ選択されていない場合
		SelectWord()
		ret_selt_str = GetSelectedString()
	End If
	
	' return
	get_slct_str = ret_selt_str
End Function

'******************************
'* Call Main Routine
'******************************
Call search_on_google()

