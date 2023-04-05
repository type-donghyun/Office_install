@ECHO OFF
::_____________________________________________________________________________________________________________________________________________________________
:: 관리자 권한 요청

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

IF %errorlevel% neq 0 (
	GOTO UACPrompt
) ELSE (
	GOTO gotAdmin
)

:UACPrompt
	ECHO SET UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	ECHO UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

	"%temp%\getadmin.vbs"
	EXIT /B

:gotAdmin
	IF EXIST "%temp%\getadmin.vbs" (
		DEL "%temp%\getadmin.vbs"
	)
	PUSHD "%CD%"
	CD /D "%~dp0"

::_____________________________________________________________________________________________________________________________________________________________

CHCP 65001 >nul
TITLE Office 365 설치

ECHO Microsoft Office 365 설치 프로세스입니다.
ECHO 설치방법을 선택해주세요.
ECHO.
:_start
CHOICE /c 123 /n /m "[1] 권장 설치 [2] 모두 설치 [3] 사용자 지정 설치 [4] 종료"
CLS
IF %errorlevel% equ 1 (
	SET mode=recommendedinstall
	GOTO _recommendedinstall
) ELSE IF %errorlevel% equ 2 (
	SET mode=allinstall
	GOTO _allinstall
) ELSE IF %errorlevel% equ 3 (
	SET mode=custominstall
	GOTO _custominstall
) ELSE IF %errorlevel% equ 4 (
	EXIT
)

:_recommendedinstall
SET Access=exclude
SET Excel=include
SET Groove=exclude
SET Lync=exclude
SET OneDrive=exclude
SET OneNote=exclude
SET Outlook=exclude
SET PowerPoint=include
SET Publisher=exclude
SET Teams=exclude
SET Word=include
SET Bing=exclude
GOTO _check

:_allinstall
SET Access=include
SET Excel=include
SET Groove=include
SET Lync=include
SET OneDrive=include
SET OneNote=include
SET Outlook=include
SET PowerPoint=include
SET Publisher=include
SET Teams=include
SET Word=include
SET Bing=include
GOTO _check

:_custominstall
CHOICE /c 12 /n /m "Access [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Access=include
) ELSE IF %errorlevel% equ 2 (
	 SET Access=exclude
)
CLS

CHOICE /c 12 /n /m "Excel [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Excel=include
) ELSE IF %errorlevel% equ 2 (
	 SET Excel=exclude
)
CLS

CHOICE /c 12 /n /m "Groove [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Groove=include
) ELSE IF %errorlevel% equ 2 (
	 SET Groove=exclude
)
CLS

CHOICE /c 12 /n /m "Lync [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Lync=include
) ELSE IF %errorlevel% equ 2 (
	 SET Lync=exclude
)
CLS

CHOICE /c 12 /n /m "OneDrive [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET OneDrive=include
) ELSE IF %errorlevel% equ 2 (
	 SET OneDrive=exclude
)
CLS

CHOICE /c 12 /n /m "OneNote [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET OneNote=include
) ELSE IF %errorlevel% equ 2 (
	 SET OneNote=exclude
)
CLS

CHOICE /c 12 /n /m "Outlook [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Outlook=include
) ELSE IF %errorlevel% equ 2 (
	 SET Outlook=exclude
)
CLS

CHOICE /c 12 /n /m "PowerPoint [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET PowerPoint=include
) ELSE IF %errorlevel% equ 2 (
	 SET PowerPoint=exclude
)
CLS

CHOICE /c 12 /n /m "Publisher [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Publisher=include
) ELSE IF %errorlevel% equ 2 (
	 SET Publisher=exclude
)
CLS

CHOICE /c 12 /n /m "Teams [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Teams=include
) ELSE IF %errorlevel% equ 2 (
	 SET Teams=exclude
)
CLS

CHOICE /c 12 /n /m "Word [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Word=include
) ELSE IF %errorlevel% equ 2 (
	 SET Word=exclude
)
CLS

CHOICE /c 12 /n /m "Bing [1] 포함 [2] 제외"
IF %errorlevel% equ 1 (
	 SET Bing=include
) ELSE IF %errorlevel% equ 2 (
	 SET Bing=exclude
)
CLS

:_check
IF %Access% equ include (
	ECHO ● Access 포함
) ELSE IF %Access% equ exclude (
	ECHO ○ Access 제외
)

IF %Excel% equ include (
	ECHO ● Excel 포함
) ELSE IF %Excel% equ exclude (
	ECHO ○ Excel 제외
)

IF %Groove% equ include (
	ECHO ● Groove 포함
) ELSE IF %Groove% equ exclude (
	ECHO ○ Groove 제외
)

IF %Lync% equ include (
	ECHO ● Lync 포함
) ELSE IF %Lync% equ exclude (
	ECHO ○ Lync 제외
)

IF %OneDrive% equ include (
	ECHO ● OneDrive 포함
) ELSE IF %OneDrive% equ exclude (
	ECHO ○ OneDrive 제외
)

IF %OneNote% equ include (
	ECHO ● OneNote 포함
) ELSE IF %OneNote% equ exclude (
	ECHO ○ OneNote 제외
)

IF %Outlook% equ include (
	ECHO ● Outlook 포함
) ELSE IF %Outlook% equ exclude (
	ECHO ○ Outlook 제외
)

IF %PowerPoint% equ include (
	ECHO ● PowerPoint 포함
) ELSE IF %PowerPoint% equ exclude (
	ECHO ○ PowerPoint 제외
)

IF %Publisher% equ include (
	ECHO ● Publisher 포함
) ELSE IF %Publisher% equ exclude (
	ECHO ○ Publisher 제외
)

IF %Teams% equ include (
	ECHO ● Teams 포함
) ELSE IF %Teams% equ exclude (
	ECHO ○ Teams 제외
)

IF %Word% equ include (
	ECHO ● Word 포함
) ELSE IF %Word% equ exclude (
	ECHO ○ Word 제외
)

IF %Bing% equ include (
	ECHO ● Bing 포함
) ELSE IF %Bing% equ exclude (
	ECHO ○ Bing 제외
)

IF %mode% equ recommendedinstall (
	GOTO _1
) ELSE IF %mode% equ allinstall (
	GOTO _2
) ELSE IF %mode% equ custominstall (
	GOTO _3
)

:_1
CHOICE /c 12 /n /m "[1] 계속 설치 [2] 처음으로"
CLS
IF %errorlevel% equ 1 (
	GOTO _configuration
) ELSE IF %errorlevel% equ 2 (
	GOTO _start
)

:_2
CHOICE /c 12 /n /m "[1] 계속 설치 [2] 처음으로"
CLS
IF %errorlevel% equ 1 (
	GOTO _configuration
) ELSE IF %errorlevel% equ 2 (
	GOTO _start
)

:_3
CHOICE /c 123 /n /m "[1] 계속 설치 [2] 설정 변경 [3] 처음으로"
CLS
IF %errorlevel% equ 1 (
	GOTO _configuration
) ELSE IF %errorlevel% equ 2 (
	GOTO _custominstall
) ELSE IF %errorlevel% equ 3 (
	GOTO _start
)

:_configuration
DEL data\configuration.xml2 >nul 2>&1

ECHO ^<Configuration ID="2a0cb094-969c-4188-beb4-fa406759387e"^>>>configuration.xml
ECHO   ^<Add OfficeClientEdition="64" Channel="Current"^>>>configuration.xml
ECHO     ^<Product ID="O365ProPlusRetail"^>>>configuration.xml
ECHO       ^<Language ID="ko-kr" /^>>>configuration.xml

IF %Access% equ exclude (
	ECHO       ^<ExcludeApp ID="Access" /^>>>configuration.xml
)

IF %Excel% equ exclude (
	ECHO       ^<ExcludeApp ID="Excel" /^>>>configuration.xml
)

IF %Groove% equ exclude (
	ECHO       ^<ExcludeApp ID="Groove" /^>>>configuration.xml
)

IF %Lync% equ exclude (
	ECHO       ^<ExcludeApp ID="Lync" /^>>>configuration.xml
)

IF %OneDrive% equ "exclude" (
	ECHO       ^<ExcludeApp ID="OneDrive" /^>>>configuration.xml
)

IF %OneNote% equ exclude (
	ECHO       ^<ExcludeApp ID="OneNote" /^>>>configuration.xml
)

IF %Outlook% equ exclude (
	ECHO       ^<ExcludeApp ID="Outlook" /^>>>configuration.xml
)

IF %PowerPoint% equ exclude (
	ECHO       ^<ExcludeApp ID="PowerPoint" /^>>>configuration.xml
)

IF %Publisher% equ exclude (
	ECHO       ^<ExcludeApp ID="Publisher" /^>>>configuration.xml
)

IF %Teams% equ exclude (
	ECHO       ^<ExcludeApp ID="Teams" /^>>>configuration.xml
)

IF %Word% equ exclude (
	ECHO       ^<ExcludeApp ID="Word" /^>>>configuration.xml
)

IF %Bing% equ exclude (
	ECHO       ^<ExcludeApp ID="Bing" /^>>>configuration.xml
)

ECHO     ^</Product^>>>configuration.xml
ECHO   ^</Add^>>>configuration.xml
ECHO   ^<Updates Enabled="TRUE" /^>>>configuration.xml
ECHO   ^<RemoveMSI /^>>>configuration.xml
ECHO ^</Configuration^>>>configuration.xml

IF NOT EXIST temp (
	MKDIR temp >nul
)

MOVE /y configuration.xml temp >nul 2>&1

CD TEMP >nul 2>&1

ECHO 설치 파일 다운로드 중...
setup.exe /download configuration.xml

CLS
ECHO 설치 중...
setup.exe /configure configuration.xml
