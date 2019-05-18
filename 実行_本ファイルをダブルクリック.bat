@echo off
setlocal enabledelayedexpansion

ECHO +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
ECHO 自動操作
ECHO 　パラメータファイルに従ってWindowsを自動操作します
ECHO +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-


REM 設定ファイル
SET CURRENT=%~d0%~p0
%~d0
CD %CURRENT%

SET CMD_SET_DIR=%CURRENT%cmd
CALL :choiceCmdSet

ECHO "%CMDSETFILE%"

REM scriptフォルダ内のvbsを実行
cscript %~d0%~p0script\core.vbs "%CMDSETFILE%" //nologo

ECHO 終了します。
pause >nul
EXIT /B


REM +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
REM フォルダ配下のファイルから選択
REM +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
:choiceCmdSet
SET CHOICE_IDX=1

ECHO +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
ECHO 実行するコマンドセットを以下フォルダから選択します。
ECHO   %CMD_SET_DIR%
ECHO +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
for /F %%I in ('dir /b /a-d %CMD_SET_DIR%') do (
	echo !CHOICE_IDX!:%%I
	SET CHOICE!CHOICE_IDX!=%%I
	SET /a CHOICE_IDX=!CHOICE_IDX!+1
)
SET /p USER_CHOICE=＞＞
IF [!CHOICE%USER_CHOICE%!]==[] (
	ECHO 不正な選択です
	GOTO choiceCmdSet
)
SET CMDSETFILE=%CMD_SET_DIR%\!CHOICE%USER_CHOICE%!

EXIT /B
