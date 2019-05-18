Option Explicit

Dim CONST_THISFOLDER
Dim CONST_PARENTFOLDER
Dim CONST_DATESTAMP
Dim CONST_TIMESTAMP
Dim CONST_WIN32EXCEL
Const WIN32EXCELNAME = "user32dll.xlsm"
CONST_THISFOLDER = Left(Wscript.ScriptFullName, InstrRev(Wscript.ScriptFullName, "\" )-1)
CONST_PARENTFOLDER = Left(CONST_THISFOLDER, InstrRev(CONST_THISFOLDER, "\" )-1)
CONST_DATESTAMP = convertTimeFormatToString(Now(), "yyyymmdd")
CONST_TIMESTAMP = convertTimeFormatToString(Now(), "yyyymmdd_hhmiss")
CONST_WIN32EXCEL = CONST_THISFOLDER & "\" & WIN32EXCELNAME

Const OUTDIR = "output"



Dim g_shell
Dim g_fso
Dim g_propDic
Dim g_workDirPath
Set g_shell = createObject("Wscript.Shell")
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_propDic = readPropetiesFile(CONST_PARENTFOLDER & "\設定.ini", "=", "#")
Dim g_userCmdDirPath, g_entryCmdFilePath
Dim g_errFlg
Dim g_traceLog
Dim g_onlyFirstMsg_showedList
Dim g_intervalCmd_counter
Dim g_userInfo
DIm g_eh
g_entryCmdFilePath = Wscript.Arguments(0)
g_userCmdDirPath = CONST_PARENTFOLDER & "\" & g_propDic("ユーザコマンドフォルダ")
g_userInfo = getUserInfoStr
g_errFlg = False
Set g_onlyFirstMsg_showedList = CreateObject("Scripting.Dictionary")
Set g_intervalCmd_counter = CreateObject("Scripting.Dictionary")
Set g_eh = New ExcelHandler

On Error Resume Next
Set g_traceLog = g_fso.OpenTextFile(CONST_PARENTFOLDER & "\log\trace_" & g_userInfo & "_" & CONST_DATESTAMP & ".log", 8, True)
If Err.Number <> 0 Then
	Msgbox "エラーです。" & vbNewLine & "複数のプロセスで同時に実行されているため、ログの書込に失敗しました。"
	g_errFlg = True
End If
On Error Goto 0

If Not g_errFlg Then
	Call makeWorkDir("")

	If Not g_fso.FolderExists(g_workDirPath) Then
		Msgbox "作業フォルダが存在しません。"
		g_errFlg = True
	End If
	
	Dim gCnt_all
	Dim gCnt_cmd

	gCnt_all = 0

	If Not g_errFlg Then
		writeLog("開始しました")
		Call main()
	End If
	Call finalizeAutoWorkDir()
	
	Call printBill
	g_traceLog.Close
End If

Set g_onlyFirstMsg_showedList = Nothing
Set g_intervalCmd_counter = Nothing
Set g_shell = Nothing
Set g_fso = Nothing
Set g_traceLog = Nothing
Set g_eh = Nothing


'メイン関数
Private Sub main()
	echoAndLog "コマンドセット:" & g_entryCmdFilePath
	Dim param_userInput
	param_userInput = receiveUserInput(g_entryCmdFilePath)
	Call executeCmdFile(g_entryCmdFilePath, param_userInput)
End Sub

'コマンドファイルの実行
Private Sub executeCmdFile(ByVal p_cmdFilePath, ByRef p_paramAryForFile)
	Dim cmdFile
	Set cmdFile = g_fso.OpenTextFile(p_cmdFilePath, 1)
	
	Dim lineNo
	lineNo = 0
	Do while Not cmdFile.AtEndOfStream
		Dim cmdLine
		cmdLine = cmdFile.ReadLine
		lineNo = lineNo + 1
		If cmdLine = "" OR Left(cmdLine, 1) = "#" OR Left(cmdLine, 1) = "!" Then
			'コメント行なので省略
		ElseIf Instr(cmdLine, Chr(9)) = 0 And Instr(cmdLine, " ") > 0 Then
			Msgbox "コマンドがTSV形式ではありません。無視します。" & vbNewLine & cmdLine
		Else
			'1要素のみの場合に無理やりTSVにする
			If Instr(cmdLine, Chr(9)) = 0 Then
				cmdLine = cmdLine & Chr(9)
			End If
			'パラメータ埋込み
			cmdLine = setUserCmdParam(cmdLine, p_paramAryForFile)
			'分解
			Dim cmdName, str_cmdElem, str_cmdParam
			Dim isLoop
			Dim roopCnt
			str_cmdElem = Left(cmdLine, Instr(cmdLine, Chr(9)) - 1)
			str_cmdParam = Right(cmdLine, Len(cmdLine) - Instr(cmdLine, Chr(9)))
			isLoop = False
			'同一コマンドを複数回実行する場合
			If Instr(str_cmdElem, ":") > 0 Then
				cmdName = Left(str_cmdElem, Instr(str_cmdElem, ":") - 1)
				roopCnt = Right(str_cmdElem, Len(str_cmdElem) - Instr(str_cmdElem, ":"))
				If roopCnt = "" Then
					roopCnt = 1
				Else
					roopCnt = Clng(roopCnt)
				End If
				isLoop = True
			'一定回数置きに実行するコマンド
			ElseIf Instr(str_cmdElem, "@") > 0 Then
				Dim span
				Dim cmdKey
				cmdName = Left(str_cmdElem, Instr(str_cmdElem, "@") - 1)
				span = Right(str_cmdElem, Len(str_cmdElem) - Instr(str_cmdElem, "@"))
				Dim intervalKey
				intervalKey = p_cmdFilePath & ":" & lineNo
				If g_intervalCmd_counter.Exists(intervalKey) Then
					g_intervalCmd_counter(intervalKey) = g_intervalCmd_counter(intervalKey) + 1
				Else
					g_intervalCmd_counter(intervalKey) = 1
				End If
				'実行タイミングかどうかを判定して実行する(roopCnt = 1 に設定)
				If g_intervalCmd_counter(intervalKey) Mod CInt(span) = 0 Then
					roopCnt = 1
				Else
					roopCnt = 0
				End If
			Else
				cmdName = str_cmdElem
				roopCnt = 1
			End If
			Dim i
			Dim paramAry
			paramAry = split(str_cmdParam, Chr(9))
			For i=1 To roopCnt
				If roopCnt > 1 Then
					echoAndLog "繰り返し実行(cmdName) " & i & "/" &  roopCnt
				End If
				If isLoop Then
					'明示的にループが指定されたときのみ
					g_propDic("ループインデックス") = i  '多重ループだと内部優先
				End If
				Call executeOneCmd(cmdName, paramAry, p_cmdFilePath)
			Next
		End If
	Loop
	cmdFile.Close
End Sub

'コマンド実行
Private Sub executeOneCmd(ByVal cmdName, ByRef paramAry, ByVal p_cmdFile_callFrom)
	'ユーザコマンドか否かかの判断、ユーザコマンドファイルパスの生成
	Dim userCmdPath
	userCmdPath = getUserCmdPath(cmdName)
	If userCmdPath <> "" Then
		echoAndLog "ユーザコマンドを実行します:" & userCmdPath
		If p_cmdFile_callFrom = userCmdPath Then
			Msgbox "ユーザコマンドが自分自身を実行しようとしました。" & vbNewLine & userCmdPath & "このコマンドをスキップします。"
		Else
			Call executeCmdFile(userCmdPath, paramAry)
		End If
	Else
		Call executeInternalCmd(cmdName, paramAry)
	End If
End Sub

'ユーザコマンドのファイルパス取得(見つからなければ空文字列を返す)
Private Function getUserCmdPath(ByVal p_cmdName)
	Dim path
	path = g_userCmdDirPath & "\" & p_cmdName
	If Not g_fso.FileExists(path) Then
		path = path & ".txt"
		If Not g_fso.FileExists(path) Then
			getUserCmdPath = ""
			Exit Function
		End If
	End If
	getUserCmdPath = path
	Exit Function
End Function

'内部コマンド実行する
Private Sub executeInternalCmd(ByVal p_cmdName, ByRef p_paramAry)
	gCnt_all = gCnt_all + 1
	Dim tmp
	'コマンドファイルの詳細部の #記号部を、paramファイルから受け取ったパラメータに置換する
	
	echoAndLog "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+"
	echoAndLog gCnt_all & ":" & p_cmdName
	echoAndLog "param:" & Join(p_paramAry, " ")
	Dim batPath
	select case LCase(p_cmdName)
		'標準出力
		case LCase("echo")
			Wscript.Echo p_paramAry(0)
		'警告(ダイアログにYESを押すと作業終了)
		case LCase("alert")
			If Msgbox(p_paramAry(0), vbYesNo) = vbYes Then
				Wscript.Quit
			End If
		'逆警告(ダイアログにNOを押すと作業終了)
		case LCase("noalert")
			If Msgbox(p_paramAry(0), vbYesNo) = vbNo Then
				Wscript.Quit
			End If
		'ダイアログ表示(1回の実行で1度だけ表示するメッセージ)
		case LCase("msgboxOnlyFirst")
			If Not g_onlyFirstMsg_showedList.Exists(p_paramAry(0)) Then
				msgbox replace(p_paramAry(0), "\n", vbNewLine)
				Wscript.Sleep(3000)
				g_onlyFirstMsg_showedList(p_paramAry(0)) = "dummyValue"
			End If
		'ダイアログ表示
		case LCase("msgbox")
			msgbox replace(p_paramAry(0), "\n", vbNewLine)
		'ポップアップ表示
		case LCase("Popup")
			Call myPopup(p_paramAry)
		'キー押下(特殊キーOK)
		case LCase("sendkey")
			If UBound(p_paramAry) >= 0 Then
				If p_paramAry(0) <> "" Then
					Call g_shell.SendKeys(p_paramAry(0))
				End If
			End If
		'キー押下(特殊キー不可の代わりに特殊文字のエスケープを自動で行う)
		case LCase("sendrawkey")
			tmp = p_paramAry(0)
			tmp = Replace(tmp, "{", "エスケープ_開始大カッコ")
			tmp = Replace(tmp, "}", "エスケープ_終了大カッコ")
			tmp = Replace(tmp, "^", "{^}")
			tmp = Replace(tmp, "%", "{%}")
			tmp = Replace(tmp, "+", "{+}")
			tmp = Replace(tmp, "~", "{~}")
			tmp = Replace(tmp, "[", "{[}")
			tmp = Replace(tmp, "]", "{]}")
			tmp = Replace(tmp, "エスケープ_開始大カッコ","{{}")
			tmp = Replace(tmp, "エスケープ_終了大カッコ","{}}")
			If tmp <> "" Then
				Call g_shell.SendKeys(tmp)
			End If
		'コマンド実行
		case LCase("cmd")
			Call g_shell.Run(Join(p_paramAry, " "), 0, True)
		'文字列をクリップボードに格納
		case LCase("clip")
			tmp = Join(p_paramAry, " ")
			tmp = escapeCommand(tmp)
			If Trim(tmp) = "" Then
				Wscript.Echo "clipコマンドに失敗しました。かわりに-記号をクリップボードに格納します。"
				tmp = "-"
			End If
			Call g_shell.Run("cmd.exe /c ECHO " & tmp & "|clip", 0, True)
			Wscript.Echo "コピーしました"
		'作業フォルダ再作成
		case LCase("renewWorkDir")
			Call finalizeAutoWorkDir
			If UBound(p_paramAry) = -1 Then
				Call makeWorkDir("")
			Else
				Call makeWorkDir(p_paramAry(0))
			End If
		'指定座標をマウスでクリック
		case LCase("click")
			Call g_eh.mouseClick(p_paramAry(0), p_paramAry(1))
		case LCase("doubleclick")
			Call g_eh.mouseDoubleClick(p_paramAry(0), p_paramAry(1))
		'指定座標をマウスでドラッグ
		case LCase("dragdrop")
			Call g_eh.mouseDrag(p_paramAry(0), p_paramAry(1), p_paramAry(2), p_paramAry(3))
		'画面全体のハードコピーをクリップボードへ
		case LCase("clipAllScreen")
			Call g_eh.clipAllScreen()
		'画面全体のハードコピーをクリップボードへ
		case LCase("clipActiveWindow")
			Call g_eh.clipActiveWindow()
		case LCase("saveAllScreenAsFile")
			Call g_eh.saveAllScreenAsFile()
		'クリップボードに保存されたテキストをテキストファイル(workフォルダ配下)に保存する
		case LCase("savaClipboardText")
			Call g_shell.Run("notepad", 1, False)
			'背後に表示されるケースがあるので念のため
			Call g_shell.Popup("メモ帳をアクティブにしてください。2秒後に閉じます。", 3, "アクティブ！！！")
			Wscript.Sleep(1500)
			Call g_shell.SendKeys("^v")
			Wscript.Sleep(1000)
			Call g_shell.SendKeys("%fa")
			Wscript.Sleep(1000)
			Call g_shell.Run("cmd.exe /c ECHO " & g_workDirPath & "\" & convertTimeFormatToString(Now(), "yyyymmdd_hhmiss") & ".txt" & "|clip", 0, True)
			Wscript.Sleep(100)
			'Msgbox "メモ帳のファイルパス入力欄がアクティブになった後、このメッセージボックスを閉じてください"
			Wscript.Sleep(500)
			Call g_shell.SendKeys("^v")
			Call g_shell.SendKeys("{ENTER}")
			Wscript.Sleep(1000)
			Call g_shell.SendKeys("%fx")
			Wscript.Sleep(100)
		'一定時間停止
		case LCase("sleep")
			Dim sleepMs
			sleepMs = Clng(p_paramAry(0))
			If LCase(g_propDic("スリープの代わりにポップアップ")) = "on" And sleepMs >= 1000 Then
				Wscript.Sleep 400
				Msgbox p_paramAry(0) & "msスリープ"
				Wscript.Sleep 400
			Else
				Wscript.Sleep(sleepMs)
			End If
		'変数をセット
		case LCase("setVar")
			g_propDic(p_paramAry(0)) = p_paramAry(1)
		'変数に加算
		case LCase("addVar")
			g_propDic(p_paramAry(0)) = Clng(g_propDic(p_paramAry(0))) + Clng(p_paramAry(1))
	end select
End Sub

'ユーザコマンドファイルのパラメータ(文字列形式cmdDetail)の #記号部をCmdSetファイルから受け取ったパラメータ(userCmdParamAry)に置換する
Private Function setUserCmdParam(ByVal cmdDetail, ByRef userCmdParamAry)
	Dim idx
	idx = 0
	Do while(idx <= UBound(userCmdParamAry))
		cmdDetail = Replace(cmdDetail, "#" & idx & "#", userCmdParamAry(idx))
		idx = idx + 1
	Loop
	'省略された引数は削除
	For idx = 0 To 20
		cmdDetail = Replace(cmdDetail, "#" & idx & "#", "")
	Next
	'共通変数の設定
	Dim tmpKey
	For Each tmpKey In g_propDic.Keys
		cmdDetail = Replace(cmdDetail, "#" & tmpKey & "#", g_propDic(tmpKey))
	Next
	setUserCmdParam = cmdDetail
End Function

'コマンド文字列の特殊文字をエスケープ
Private Function escapeCommand(ByVal p_cmd)
	p_cmd = replace(p_cmd, ">", "＞")
	p_cmd = replace(p_cmd, "<", "＜")
	p_cmd = replace(p_cmd, "|", "｜")
	p_cmd = replace(p_cmd, "%", "％")
	escapeCommand = p_cmd
End Function

'ユーザ入力を受け取る
Private Function receiveUserInput(ByVal p_cmdFilePath)
	Dim cFile
	Dim tmp
	Dim buf
	buf = ""
	Set cFile = g_fso.OpenTextFile(p_cmdFilePath, 1)
	Do while Not cFile.AtEndOfStream
		tmp = cFile.readLine
		If Left(tmp, 1) = "!" Then
			buf = buf & InputBox(Right(tmp, Len(tmp)-1), "パラメータ入力") & ","
		End If
	Loop
	cFile.Close
	Set cFile = Nothing
	receiveUserInput = split(buf, ",")
End Function

'ポップアップメッセージ
Private Function myPopup(p_paramAry)
	Dim msg, sleepTime, title, subMsg
	sleepTime = CInt(p_paramAry(1))
	msg = Cstr(p_paramAry(0)) & vbNewLine & vbNewLine & "「OK」：続行(" & sleepTime & "秒後に自動でOK押下)" & vbNewLine & "「ｷｬﾝｾﾙ」：待機"
	IF UBound(p_paramAry) >= 2 Then
		title = Cstr(p_paramAry(2)) & "(" & sleepTime & "秒で閉じます)"
		'待機用メッセージを取得
		If UBound(p_paramAry) = 3 Then
			subMsg = Cstr(p_paramAry(3))
		Else
			subMsg = "待機完了後、「OK」を押してください。"
		End If
	Else
		title = sleepTime & "秒で閉じます"
	End If
	Dim pushVal
	pushVal = g_shell.Popup(msg, sleepTime, title, 1)
	'キャンセル押下の場合はMsgbox(OK押下あるいは自動終了時はスキップ)
	If pushVal = vbCancel Then
		Msgbox(subMsg)
	End If
End Function

'---------------------------------------------------------
'日付オブジェクトを文字列に変換(VBScriptにはDateFormat系の関数がないので自前で実装)
'	使用例：convertTimeFormatToString(Now(), "yyyymmddhhmiss")
'---------------------------------------------------------
function convertTimeFormatToString(ByVal p_timeObj, ByVal p_formatString)
	Dim ret
	ret = p_formatString
	ret = Replace(ret, "yyyy", Right("0000" & Year(p_timeObj), 4))
	ret = Replace(ret, "mm", Right("0000" & Month(p_timeObj), 2))
	ret = Replace(ret, "dd", Right("0000" & Day(p_timeObj), 2))
	ret = Replace(ret, "hh", Right("0000" & Hour(p_timeObj), 2))
	ret = Replace(ret, "mi", Right("0000" & Minute(p_timeObj), 2))
	ret = Replace(ret, "ss", Right("0000" & Second(p_timeObj), 2))
	convertTimeFormatToString = ret
End Function

Sub echoAndLog(ByVal p_str)
	Wscript.Echo p_str
	writeLog(p_str)
End Sub

Sub writeLog(ByVal p_str)
	g_traceLog.writeLine(Now & "	" & p_str)
End Sub

'----------------------------------------------------------
'作業フォルダ管理
'----------------------------------------------------------
'初期処理
Private Sub makeWorkDir(Byval p_apdStr)
	g_workDirPath = CONST_PARENTFOLDER & "\" & OUTDIR & "\" & g_userInfo & "_" & convertTimeFormatToString(Now(), "yyyymmdd_hhmiss")
	If p_apdStr <> "" Then
		g_workDirPath = g_workDirPath & "_" & p_apdStr
	End If
	g_fso.createFolder(g_workDirPath)
	Wscript.Sleep 100
End Sub
'最終処理
Private Sub finalizeAutoWorkDir()
	Dim wDir
	Set wDir = g_fso.getFolder(g_workDirPath)
	If wDir.Files.Count = 0  And wDir.SubFolders.Count = 0 Then
		echoAndLog("現在の作業フォルダが空なので削除します。")
		Wscript.Echo g_workDirPath
		g_fso.deleteFolder(g_workDirPath)
	End If
	g_workDirPath = ""
End Sub

'----------------------------------------------------------
'プロパティファイル読み込み
'----------------------------------------------------------
'プロパティファイルを読み込み、Dictionary型で返す
'	使用例：Set dicObj_propFile = readPropetiesFile("C:\Users\admin\Desktop\test.properties", "=", "#")
Function readPropetiesFile(ByVal p_filePath, ByVal p_delimiter, ByVal p_mark_oneLineComment)
	Dim dicObj_propFile
	Set dicObj_propFile = CreateObject("Scripting.Dictionary")
	Dim fileLinesAry
	Dim oneLine
	Dim pFile
	Set pFile =g_fso.OpenTextFile(p_filePath, 1)
	fileLinesAry = split(pFile.ReadAll, vbNewLine)
	pFile.Close
	For Each oneLine In fileLinesAry
		'空行でもコメント行でなければ読む
		If Len(Trim(oneLine)) <> 0 And Left(oneLine, Len(p_mark_oneLineComment)) <> p_mark_oneLineComment Then
			Call dicObj_propFile.Add(Left(oneLine, Instr(oneLine, p_delimiter)-1), Right(oneLine, Len(oneLine) - Instr(oneLine, p_delimiter) - Len(p_delimiter)+1))
		End If
	Next
	Set readPropetiesFile = dicObj_propFile
End Function

'----------------------------------------------------
'自身の情報を取得
'----------------------------------------------------
Private Function getUserInfoStr()
	Dim shell
	Dim shellResult
	Dim whoAmI
	Set shell = CreateObject("Wscript.Shell")
	whoAmI = shell.expandEnvironmentStrings("%username% %computername%")
	Dim neonId
	neonId = shell.expandEnvironmentStrings("%username%")
	Set shellResult = shell.Exec("cmd.exe /c @ECHO OFF&for /F ""tokens=3,4"" %I in ('net user " & neonId & " /domain ^|find ""ネーム""') do  echo %I%J")
	whoAmI = whoAmI & " " & shellResult.StdOut.ReadLine()
	Set shellResult = Nothing
	Set shell = Nothing
	getUserInfoStr = Replace(whoAmI, " ", "_")
End Function

'-----------------------------------------------------
'ツール使用料の請求書を作成・印刷する
'-----------------------------------------------------
Private Sub printBill()
	Dim gmm
	gmm = g_propDic("GIVE_ME_MONEY")
	If gmm = 0 Then
		Exit Sub
	End If
	If IsNumeric(gmm) Then
		If gmm < 0 Then
			echoAndLog("使用料の不正な編集を検知しました。")
			echoAndLog("「" & g_userInfo & "」をブラックリストに登録します。")
			gmm = 1000
		End If
		echoAndLog("+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-")
		echoAndLog("使用者：" & Replace(g_userInfo, "_", " "))
		echoAndLog("使用料合計: \" & FormatNumber(gCnt_all * gmm, 0, 0, 0, -1))
		echoAndLog("コマンド数合計:" & gCnt_all)
		echoAndLog("使用料/1コマンド: \" & FormatNumber(gmm, 0, 0, 0, -1))
		echoAndLog("支払期日:" & DateAdd("d", 3, Date) & " 18:00")
		echoAndLog("支払先:瀬藤")
		echoAndLog("+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-")
		'Dim excelApp
		'Dim wb
		'Set excelApp = CreateObject("Excel.Application")
		'excelApp.Visible=True
		'Set wb = excelApp.WorkBooks.Add
	End If
End Sub

'Excelを扱うクラス
Class ExcelHandler
	Dim win32wb
	
	Dim available
	Dim Excel
	
	'コンストラクタ
	Private Sub Class_Initialize()
		available = False
	End Sub
	
	'ターミネータ
	Private Sub Class_Terminate()
		If Not available Then
			Exit Sub
		End If
		win32wb.Close(False)
		Excel.Quit
	End Sub
	
	'初期化(Excel起動、ワークブックを開く)
	Private Sub init()
		If available Then
			Exit Sub
		End If
		Set Excel = WScript.CreateObject("Excel.Application")
		Excel.Visible = False
		Set win32wb = Excel.Workbooks.Open(CONST_WIN32EXCEL,,,True)
		available = True
	End SUb
	
	'マウスクリック
	Public Sub mouseClick(Byval x, ByVal y)
		init
		Call Excel.Run(WIN32EXCELNAME & "!Click(" & x & "," & y & ")")
	End Sub
	Public Sub mouseDoubleClick(Byval x, ByVal y)
		init
		Call Excel.Run(WIN32EXCELNAME & "!DoubleClick(" & x & "," & y & ")")
	End Sub
	
	'マウスドラッグ
	Public Sub mouseDrag(Byval bx, ByVal by, Byval ax, ByVal ay)
		init
		Call Excel.Run(WIN32EXCELNAME & "!Drag(" & bx & "," & by & "," & ax & "," & ay & ")")
	End Sub
	
	'スクリーンショットをクリップボードへ（全画面）
	Public Sub clipAllScreen()
		init
		Call Excel.Run(WIN32EXCELNAME & "!clipAllScreen()")
	End Sub
	'スクリーンショットをクリップボードへ（アクティブウィンドウ）
	Public Sub clipActiveWindow()
		init
		Call Excel.Run(WIN32EXCELNAME & "!clipActiveWindow()")
	End Sub
	'スクリーンショットをファイル保存
	Public Sub saveAllScreenAsFile()
		init
		Dim picDir
		Dim bfPicCnt
		Set picDir = g_fso.GetFolder(g_shell.ExpandEnvironmentStrings("%userprofile%") & "\Pictures\Screenshots")
		bfPicCnt = picDir.Files.Count
		Call Excel.Run(WIN32EXCELNAME & "!saveAllScreenAsFile()")
		'ファイルの保存が完了するまで待つ
		Do while picDir.Files.Count = bfPicCnt
			Wscript.Sleep 100
		Loop
		'ファイルを取得(pictureフォルダ内の更新日付最大が、今回のスクリーンショットのファイルであると判定)
		Dim picFile
		Dim targetFile
		For Each picFile In picDir.Files
			If Not IsObject(targetFile) Then
				Set targetFile = picFile
			ElseIf picFile.DateLastModified > targetFile.DateLastModified Then
				Set targetFile = picFile
			End If
		Next
		Call g_fso.CopyFile(targetFile, g_workDirPath & "\ss_" & convertTimeFormatToString(Now(), "yyyymmdd_hhmiss")  & "." & g_fso.GetExtensionName(targetFile))
	End Sub
	
	
End Class


