#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icon\gear.ico
#AutoIt3Wrapper_Outfile=W10Configurator.exe
#AutoIt3Wrapper_Res_Fileversion=0.2.4.50
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=W10Configurator
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=Made By|Williamas Kumeliukas
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;*****************************************
;*  W10Configurator.au3 by Williamas Kumeliukas  *
;*****************************************

#REGION Includes <<<<<<<<
#include <Constants.au3>
#include <Date.au3>
#include <StaticConstants.au3>
#include <WinNet.au3>
#include <Inet.au3>
#include <GUIConstantsEx.au3>
#Include <GuiButton.au3>
#include <TreeViewConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <AutoItConstants.au3>
#include <Array.au3>
#include <FontConstants.au3>
#include <GuiRichEdit.au3>
#include <GuiTab.au3>
#include <GuiListView.au3>
#include <String.au3>
#include <StringConstants.au3>
#include <WinAPIShellEx.au3>
#include <Misc.au3>
#include <Process.au3>
#include <EventLog.au3>
#include <APIDiagConstants.au3>
#EndRegion Includes <<<<<<

Opt("GUIResizeMode", $GUI_DOCKAUTO)

HotKeySet("{NUMPAD0}", "_Exit")



#REGION DECLARE VARIABLES  FOR LATER USE
Global $ConfigDir =  @ScriptDir ;temporary
Global $nas, $cw, $cwe,  $c, $cmdfile, $nas, $sRemoteName
Global $regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation", $GUI
Global  $iniFile = $ConfigDir & "\config.ini", $iRead, $isAdmin, $iniOpen = @ScriptDir & "\config.ini"
Global $console
Global $tree
Global $info
Global $sInfos
Global $compname
Global $supportHours
Global $supportPhone
Global $supportUrl
Global $Manufacturer
Global $model
Global $OEMLogo
Global $oemfile
Global $cLogo
Global $wulv
Global $wul
Global $wup
Global $Go
Global $sServer, $sShare,$letter, $sLocalName = $letter & ":"
Global $sServerShare =  $sServer & $sShare
Global $sUserName = "CHANGEME", $sPassword = "CHANGEME"
Global $sResult, $aResults, $sSaverActive
Global $syspin = "syspin.exe"
Global $ColNeeded
Global $Host = @ComputerName
Global $sBannedList
Global $BannedList 
Global $label[10]
Global $task[15]
Global $sRZVersion ;Local RZGet Version
Global $sORZVersion ;Online RZGet Version
#EndRegion DECLARE VARIABLES FOR LATER USE


GUI()

Func c($cw) ;INFORMATIVE MESSAGES IN CONSOLE
$c = 0
ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> " &  $cw)

EndFunc

Func cw($cwe) ; @@@@ CONSOLE WARNINGS / ERRORS MESSAGE @@@@@
$c = 1
ConsoleWriteGUI($console,@CRLF & _NowTime() & " =-> ERROR: " &  $cwe)
EndFunc

Func ConsoleWriteGUI(Const ByRef $hConsole, Const $sTxt); READ C($cw) & cw($cw) ($cwe_ soon)
	_GUICtrlRichEdit_GotoCharPos($Console, -1)
	If $c = 1 Then
		 _GUICtrlRichEdit_SetCharColor($Console, 0x0000FF)
		 _GUICtrlRichEdit_InsertText($Console,  $sTxt)
		 Else
			 _GUICtrlRichEdit_SetCharColor($Console, 0xFFFFFF)
			 _GUICtrlRichEdit_InsertText($Console,  $sTxt)
	EndIf

	_GUICtrlRichEdit_SetFont($Console, 12, "Consolas")

	_GUICtrlRichEdit_HideSelection($Console,True)
EndFunc


#Region INI READ/WRITE
Func ini($section, $key, $value, $iniread = True) ;Read or edit value in config.ini

	;$iniOpen = FileOpen($iniFile, 2)
	;c("read state = " & $iniread)
	Switch $iniread

		Case False ; = overwrite

			return IniWrite($iniFile, $section, $key, $value)

		Case True ; = read only

			return IniRead($iniFile, $section, $key, $value)

	EndSwitch

EndFunc
#EndRegion INI READ/WRITE

Func selfoem() ;**** AutoIt version of oem() ****
	c("OEM installation started !")
	sleep(500)
	c("Installing OEM Logo...")
	_FileCopy($OEMLogo, "C:\Windows\System32")
	RegWrite( $regPath, "Manufacturer", "REG_SZ", $Manufacturer )
	c("Added Manufacturer registry key...")
	RegWrite( $regPath, "Model", "REG_SZ", $Model) ;
	RegWrite($regPath, "Logo", "REG_SZ", $OEMLogo)
	c("Added Logo (OEM Logo references) registry key...")
	Regwrite($regPath, "SupportHours", "REG_SZ", $supportHours)
	c("Added SupportHours registry key...")
	RegWrite($regPath, "SupportPhone", "REG_SZ", $supportPhone)
	c("Added SupportPhone registry key...")
	RegWrite($regPath, "SupportURL", "REG_SZ", "https://" & $supportUrl)
	c("Added SupportURL registry key...")
	sleep(500)
	c("OEM installation done !")

EndFunc 	   ;**** AutoIt version of oem() ****


#Region RZGET <<<  Work flawless (RZGet is another solution like Ninite which offers a variety of softwares to download & install)
Func Rzget()
    Local $rzcmd

	If FileExists($ConfigDir & "\rzget.exe") Then ; IF RZGET IS IN CONFIG FOLDER @@@@
		_CheckRZGetVersion()
		$rzcmd = '@echo off' & @CRLF _
		& 'call :isAdmin' & @CRLF _
		& 'if %errorlevel% == 0 (' & @CRLF _
		& 'goto :run' & @CRLF _
		& ') else (' & @CRLF _
		& 'echo Requesting administrative privileges...' & @CRLF _
		& 'goto :UACPrompt' & @CRLF _
		& ')' & @CRLF _
		& 'exit /b' & @CRLF _
		& ':isAdmin' & @CRLF _
		& 'fsutil dirty query %systemdrive% >nul' & @CRLF _
		& 'exit /b' & @CRLF _
		& ':run' & @CRLF _
		& 'cmd /c ' & $ConfigDir & '\rzget.exe install "Google Chrome"' & @CRLF _
		& 'rzget.exe install "7-zip"' & @CRLF _
		& 'rzget.exe install "AdobeReader DC"' &  @CRLF _
		& 'rzget.exe install "Edge"' & @CRLF _
		& 'echo done.' & @CRLF _
		& 'exit /b' & @CRLF _
		& ':UACPrompt' & @CRLF _
		& 'echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"' & @CRLF _
		& 'echo UAC.ShellExecute "cmd.exe", "/c %~s0 %~1", "", "runas", 1 >> "%temp%\getadmin.vbs"' & @CRLF _
		& '"%temp%\getadmin.vbs"' & @CRLF _
		& 'del "%temp%\getadmin.vbs"' & @CRLF _
		& 'exit /B'
	Else ; RZGET.EXE WAS NOT FOUND IN CONFIG FOLDER @@@@
		c("ruckzuck executable not found...")
		c("")
	EndIf
	c("Creating Rzget process...")
	FileWrite($ConfigDir & "\rzget.bat", $rzcmd)
    c("waiting for Rzget to load...")
    ShellExecute($ConfigDir &  "\rzget.bat")
    c("Done !")

EndFunc

Func _GetOnlineRZVersion()
	local $Powershell = Run("powershell -Command (((curl 'https://ruckzuck.tools' -Usebasicparsing).Links | Select href)| Select-String -Pattern 'RZGet.exe')", @SystemDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
	$sORZVersion = ""
	While 1
		$sORZVersion &= StdoutRead($Powershell)
		if @error Then ExitLoop
	WEnd
	$sORZVersion = StringReplace(StringTrimRight(StringTrimLeft($sORZVersion,63),6),"/RZGet.exe}","")
	c("Online Version of RZGet is: " & $sORZVersion)
EndFunc

Func _CheckRZGetVersion()
	$sRZVersion = FileGetVersion( @ScriptDir & "\Ressources\Rzget.exe")
	Call(_GetOnlineRZVersion)
	If $sORZVersion <> $sRZVersion Then
		c("Updating RZGet Version")
		InetGet("https://github.com/rzander/ruckzuck/releases/download/" & $sORZVersion & "/Rzget.exe", @ScriptDir & "\Ressources\Rzget.exe",2,0)
	Else
		c("RZGet is up to date.")
	EndIf
EndFunc
#EndRegion RZGET <<< (working flawless)

#Region SCREENSAVER
Func ScreenSaver($Alive) ;False = reset to default & True to prevent/disable sleep/power-savings modes (AND screensaver)

	If $Alive = True Then
		Local $ssKeepAlive = DllCall( 'kernel32.dll', 'long', 'SetThreadExecutionState', 'long', 0x80000003)
		If @error Then
			Return SetError( 2, @error, 0x80000000)
			cw(@error)
		EndIf
	Return $ssKeepAlive[0]	; Previous state (typically 0x80000000 [-2147483648])
		c("Hibernate mode = " & $sSaverActive)
	ElseIf $Alive = False Then
		; Flag:	ES_CONTINUOUS (0x80000000) -> (default) -> used alone, it resets timers & allows regular sleep/power-savings mode
		Local $ssKeepAlive=DllCall('kernel32.dll','long','SetThreadExecutionState','long',0x80000000)
		If @error Then
			Return SetError(2,@error,0x80000000)
			cw(@error)
		EndIf
		Return $ssKeepAlive[0]	; Previous state
	EndIf
EndFunc
#EndRegion SCREENSAVER

Func _FileCopy($sRemoteName, $d)
	Local $FOF_RESPOND_YES = 16
	Local $FOF_SIMPLEPROGRESS = 256
	$oShell = ObjCreate("shell.application")
	$oShell.namespace($d).CopyHere($sRemoteName,$FOF_SIMPLEPROGRESS)
EndFunc   ;==>_FileCopy

	#Region NAS SCRIPT <<<<<<<<<
Func nas()

	c("Verifying if there is an existing connection to: >>> NAS-02 <<< ...")
	If _WinNet_GetConnection($sLocalName) = $sServerShare Then
		c("Connection aldready existing, skipping this step....")
	Else
		iF $sUserName or $sPassword = "CHANGEME" Then
		MsgBox(0,"", "Username or Password must be changed before attempting to connect NAS.")
		Exit
		EndIf
		c("Attributing letter: " & $letter)
		$sResult = _WinNet_AddConnection2( $sLocalName, $sServerShare, $sUserName, $sPassword, 1, 2) ;(1,2 = remember connection and user interact on error)
		c("Connected! Result = " & $sResult)
		If $sResult Then
			;USE _FILECOPY HERE TO OBTAIN
			_WinNet_CancelConnection2($sLocalName)
		Else
			c("Connection Failed... Inform WilliamasKumeliukas of this problem please.")
			Sleep(10 * 1000)
			Exit
		EndIf
	EndIf

EndFunc
#EndRegion NAS SCRIPT <<<<<<<<

#Region  GUI SCRIPT <<<<<<<< @@@@@@@@@@@@@@@@@@@@

Func GUI()

	 $GUI = GUICreate(@ScriptName & " - Version: " & FileGetVersion(@ScriptFullPath), 1043,820,-1,-1,-1,-1) ;740,632)
	GUISetIcon("Icon\gear.ico", $GUI)
	$tab = GUICtrlCreatetab(0,0,697,541,-1,-1)
	GuiCtrlSetState($tab, $GUI_ONTOP) ;2048
	GUICtrlCreateTabItem("System Info")
	GUICtrlCreateTabItem("Configuration")
	GUICtrlCreateTabItem("Windows Updater")
	GUICtrlCreateTabItem("")
_GUICtrlTab_SetCurFocus($tab,-1)
GUICtrlSetResizing(-1,70)

;GUI setup for Tasks
GUICtrlCreateGraphic(705, 45, 300, 455, BitOR($GUI_SS_DEFAULT_GRAPHIC,$SS_WHITEFRAME,$WS_BORDER))
GUICtrlSetColor(-1, 0x000000)
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlSetGraphic(-1, $GUI_GR_COLOR, 0x000000, 0xFFFFFF)
$task1 =  GuiCtrlCreateCheckbox("Windows Updates", 710, 48, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task2 =  GuiCtrlCreateCheckbox("rzget", 710, 70, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task3 =  GuiCtrlCreateCheckbox("Default = Chrome", 710, 92, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task4 =  GuiCtrlCreateCheckbox(".pdf/pdfxml = Reader", 710, 114, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task5 =  GuiCtrlCreateCheckbox("Install OEM + Logo", 710, 136, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task6 =  GuiCtrlCreateCheckbox("Disable hibernation", 710, 158, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task7 =  GuiCtrlCreateCheckbox("Add Office 2010 to C:\", 710, 180, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task8 =  GuiCtrlCreateCheckbox("Install Office 2010", 710, 202, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task9 =  GuiCtrlCreateCheckbox("ComputerName", 710, 224, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task10 = GuiCtrlCreateCheckbox("Add Office icons to taskbar", 710, 246, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)
$task11 = GuiCtrlCreateCheckbox("Select All", 710, 268, 260, 20, BitOR($TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES))
	GUICtrlSetFont(-1,13,400,0,"Lucida Bright", 5)
 	GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
	 GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$sRunTasks = GUICtrlCreateButton("Run", 910, 510, 80, 30)

$Console = _GUICtrlRichEdit_Create($GUI, "",0,559,1043,258, BitOR($ES_MULTILINE,$WS_VSCROLL,$ES_AUTOVSCROLL));9, 364, 719, 258
	_GUICtrlRichEdit_SetEventMask($Console,$ENM_LINK )
	_GUICtrlRichEdit_AutoDetectURL($Console, True)
	_GUICtrlRichEdit_SetCharColor($Console, 0xFFFFFF)
	_GUICtrlRichEdit_SetBkColor($Console, 0x000000)
	_GUICtrlRichEdit_SetReadOnly($Console, True)
	_GUICtrlRichEdit_HideSelection($Console,True)
	_GUICtrlRichEdit_SetFont($Console, 12, "Consolas")

$label[0] = GUICtrlCreateLabel("Select tasks to do:",780,23,150,18,$SS_CENTER,-1)
 GUICtrlSetFont(-1,12,800,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)



 GUISwitch($GUI,_GUICtrlTab_SetCurFocus($tab,1)&GUICtrlRead ($tab, 1))

$Go = GUICtrlCreateButton("Start", 600, 480, 80, 30, BitOR($BS_DEFPUSHBUTTON, $BS_CENTER))
	GUICtrlSetFont(-1,10,600,0,"Lucida Bright", 5)
	GUICtrlSetResizing($Go,$GUI_DOCKAUTO)

 $label[1] = GUICtrlCreateLabel("Manufacturer:",107,52, 105, -1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[1], $GUI_DOCKAUTO)
 $Manufacturer = GUICtrlCreateInput("",201,52,222,20,-1,$WS_EX_CLIENTEDGE)
GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[2] = GUICtrlCreateLabel("Model:",153,134, 105)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[2],$GUI_DOCKAUTO)
 $model = GUICtrlCreateInput("",201,132,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[3] = GUICtrlCreateLabel("Support Hours:", 100,194, 110, -1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $supportHours = GUICtrlCreateInput("",201,192,218,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

  $label[4] = GUICtrlCreateLabel("Support Website:",86,256,130,15,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $supportUrl = GUICtrlCreateInput("https://",201,255,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $OEMLogo = GUICtrlCreateInput("",201,319,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $label[5] = GUICtrlCreateLabel("Logo:",159,321,101,15,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $oemfile = GUICtrlCreateButton("...",440,319,44,24,-1,-1)
 GUICtrlSetFont(-1, 10, 400, 0, "MS sans serif", 2)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[6] = GUICtrlCreateLabel("Computer Name:",88,442,110,15,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $compname = GUICtrlCreateInput($Host,201,440,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 GUISwitch($GUI,_GUICtrlTab_SetCurFocus($tab,0)&GUICtrlRead ($tab, 1))

 Global $iEdit = GUICtrlCreateEdit( "", 5, 25, 677, 481, BitOR( $ES_AUTOVSCROLL, $ES_READONLY, $ES_MULTILINE, $ES_UPPERCASE, $WS_VSCROLL, $WS_HSCROLL ), -1 )

 GUISwitch($GUI,_GUICtrlTab_SetCurFocus($tab,2)&GUICtrlRead ($tab, 1))

 $wulv = GUICtrlCreatelistview("Descriptions|Status",17,192,647,334,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 _GUICtrlListView_SetColumnWidth($wulv, 0, 500)
GUICtrlSetBkColor(-1, $color_silver)
GUICtrlSetColor(-1, 0x75EE3B)
_GUICtrlListView_SetColumnWidth($wulv, 1, 140)
GUICtrlSetBkColor(-1, $color_black)
GUICtrlSetColor(-1, 0x75EE3B)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[7] = GUICtrlCreateLabel("Search not started yet",16,170, 150 , 15, $SS_LEFTNOWORDWRAP)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[7], $GUI_DOCKAUTO)
 $wup = GUICtrlCreateProgress(16,142,649,20,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[8] = GUICtrlCreateInput("", 16,90, 300, 40) ;User defined additions to banned list
 GUICtrlCreateLabel("Use the below box to exclude updates from being applied: (ex. Silverlight)", 16,55,310,40)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $wus = GUICtrlCreateButton("Search",500,80,100,30,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 GUISwitch($GUI,_GUICtrlTab_SetCurFocus($tab,1)&GUICtrlRead ($tab, 1))



_GUICtrlTab_SetCurFocus($tab,0)
;~ 	Global $console = GUICtrlCreateEdit("", 9, 364, 719, 258, BitOR ($ES_AUTOVSCROLL,$ES_READONLY,$ES_MULTILINE,$ES_UPPERCASE,$WS_VSCROLL,$WS_HSCROLL),-1)
;~ 	GUICtrlSetFont(-1,10,400,0,"Consolas")
;~ 	GUICtrlSetBkColor($console, 0x000000)
;~ 	GUICtrlSetColor($console, 0x00FF00)

	ConsoleWriteGUI($console, "=============== Win10 CONFIGURATOR CONSOLE===============    " &  _NowTime(3) )

 GUISetState(@SW_SHOW,$GUI)
;~  return $GUI

While 1

	Switch GUIGetMsg()

		Case $GUI_EVENT_CLOSE
			Exit

		Case $task11
			If GUICtrlRead($task11) == $GUI_CHECKED Then
				GUICtrlSetState($task1, $GUI_Checked)
				GUICtrlSetState($task2, $GUI_Checked)
				GUICtrlSetState($task3, $GUI_Checked)
				GUICtrlSetState($task4, $GUI_Checked)
				GUICtrlSetState($task5, $GUI_Checked)
				GUICtrlSetState($task6, $GUI_Checked)
				GUICtrlSetState($task7, $GUI_Checked)
				GUICtrlSetState($task8, $GUI_Checked)
				GUICtrlSetState($task9, $GUI_Checked)
				GUICtrlSetState($task10, $GUI_Checked)
			ElseIf GUICtrlRead($task11) == $GUI_UNCHECKED Then
				GUICtrlSetState($task1, $GUI_UnChecked)
				GUICtrlSetState($task2, $GUI_UnChecked)
				GUICtrlSetState($task3, $GUI_UnChecked)
				GUICtrlSetState($task4, $GUI_UnChecked)
				GUICtrlSetState($task5, $GUI_UnChecked)
				GUICtrlSetState($task6, $GUI_UnChecked)
				GUICtrlSetState($task7, $GUI_UnChecked)
				GUICtrlSetState($task8, $GUI_UnChecked)
				GUICtrlSetState($task9, $GUI_UnChecked)
				GUICtrlSetState($task10, $GUI_UnChecked)
			EndIf
		Case $sRunTasks
			c("Running Tasks")
			If GUICtrlRead($Task1) == $GUI_Checked Then
				c("Running Windows Updates Task")
				_PopulateNeeded($Host)
			Endif
			If GUICtrlRead($Task2) == $GUI_Checked Then
				c("Running RZGet")
				Rzget()
			EndIF
			If GUICtrlRead($Task3) == $GUI_Checked Then
				c("Setting Chrome as default browser")
				defaultBrowser()
			EndIF
			If GUICtrlRead($Task4) == $GUI_Checked Then
				c("Setting .PDF and .PDFXML to Adobe Reader")
				Sleep(10000)
			EndIF
			If GUICtrlRead($task5) == $GUI_Checked Then
				c("Installing OEM + Logo")
				selfoem()
			EndIF
			If GUICtrlRead($Task6) == $GUI_Checked Then
				c("Disabling Hibernation")
				;ScreenSaver()
			EndIF
			If GUICtrlRead($Task7) == $GUI_Checked Then
				c("Adding office 2010 to C:\")
				sleep(10000)
			EndIF
			If GUICtrlRead($Task8) == $GUI_Checked Then
				c("Installing MS Office 2010")
				Office2010()
			EndIF
			If GUICtrlRead($Task9) == $GUI_Checked Then
				c("Setting computer name")
				Sleep(10000)
				;_SetComputerName() ;Temporarily commented out to prevent changing of name
			EndIF
			If GUICtrlRead($Task10) == $GUI_Checked Then
				c("Adding Office icons to taskbar")
				Sleep(10000)
			Endif
		
		Case $Go
			c("Configuration started!")
			 ; CONFIGURATION PROCESS START HERE @@@@@@@@@@@@@@@@

			;WILL BE DONE WHEN EVERYTHING WILL BE TESTED AND WORKING
			_GetSystemInfo()

			 ; CONFIGURATION PROCESS END HERE @@@@@@@@@@@@@@@@@@
		Case $wus
			 _PopulateNeeded($Host)
	EndSwitch
Wend

EndFunc
#EndRegion GUI <<<<<<<<<<<<<<@@@@@@@@@@@@@

	#Region SetComputerName <<<<<<<<
Func _SetComputerName($sCmpName)

    Local $sLogonKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Local $sCtrlKey = "HKLM\SYSTEM\Current\ControlSet"
    Local $aRet

    If StringRegExp($sCmpName, '|/|:|*|?|"|<|>|.|,|~|!|@|#|$|%|^|&|(|)|;|{|}|_|=|+|[|]|x60' & "|'", 0) = 1 Then Return 0

;~     ; 5 = ComputerNamePhysicalDnsHostname
    $aRet = DllCall("Kernel32.dll", "BOOL", "SetComputerNameEx", "int", 5, "str", $sCmpName)
    If $aRet[0] = 0 Then Return SetError(1, 0, 0)
    RegWrite($sCtrlKey & "ControlComputernameActiveComputername", "ComputerName", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ControlComputernameComputername", "ComputerName", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ServicesTcpipParameters", "Hostname", "REG_SZ", $sCmpName)
    RegWrite($sCtrlKey & "ServicesTcpipParameters", "NV Hostname", "REG_SZ", $sCmpName)
    RegWrite($sLogonKey, "AltDefaultDomainName", "REG_SZ", $sCmpName)
    RegWrite($sLogonKey, "DefaultDomainName", "REG_SZ", $sCmpName)
    RegWrite("HKEY_USERS.DefaultSoftwareMicrosoftWindows MediaWMSDKGeneral", "Computername", "REG_SZ", $sCmpName)

    ; Set Global Environment Variable
    RegWrite($sCtrlKey & "ControlSession ManagerEnvironment", "Computername", "REG_SZ", $sCmpName)
		$aRet = DllCall("Kernel32.dll", "BOOL", "SetEnvironmentVariable", "str", "Computername", "str", $sCmpName)

    If $aRet[0] = 0 Then Return SetError(2, 0, 0)
		$iRet2 = DllCall("user32.dll", "lresult", "SendMessageTimeoutW", "hwnd", 0xffff, "dword", 0x001A, "ptr", 0, _
            "wstr", "Environment", "dword", 0x0002, "dword", 5000, "dword_ptr*", 0)

    If $iRet2[0] = 0 Then Return SetError(2, 0, 0)
    Return 1

EndFunc   ;==>_SetComputerName
#EndRegion SetComputerName <<<<<<<<

Func defaultBrowser()
	 ;HKEY_CURRENT_USER\SOFTWARE\RegisteredApplications
	 ;HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice (ProgId REG_SZ ChromeHTML)
	 ;
 EndFunc

 #Region 		 INSTALL OFFICE 2010

 Func Office2010() ;==> TO DO

$officePath = "CHANGEME"

 EndFunc

 #EndRegion INSTALL OFFICE 2010

Func _IsChecked($control)
    Return BitAnd(GUICtrlRead($control),$GUI_CHECKED) = $GUI_CHECKED
EndFunc


	#Region WINDOWS UPDATE

Func _CreateMsUpdateSession($strhost = @ComputerName)
	c("Creating a Windows Update Session...")
	$objsession = ObjCreate("Microsoft.Update.Session", $strhost)
	If Not IsObj($objsession) Then Return 0
	Return $objsession
EndFunc   ;_CreateMsUpdateSession

Func _CreateSearcher($objsession)
	c("Creating Searcher Session..."& @CRLF & "Searching for updates available...")
	If Not IsObj($objsession) Then Return -1
	Return $objsession.createupdatesearcher
EndFunc   ;_CreateSearcher

Func _FetchNeededData($Host)
	$objsearcher = _CreateSearcher(_CreateMsUpdateSession($Host))
	$colneeded = _GetNeededUpdates($objsearcher)
	$objsearcher = 0
	Dim $arrneeded[1][2]
	For $i = 0 To $colneeded.updates.count - 1
		If $i < $colneeded.updates.count - 1 Then ReDim $arrneeded[$i + 2][2]
		$update = $colneeded.updates.item($i)
		$arrneeded[$i][0] = $update.title
		$arrneeded[$i][1] = $update.description
	Next
	If Not IsArray($arrneeded) Then
		cw("Windows Updates Service seems to have encounted a problem with: " & $Host)
		Return 0
	EndIf
	Return $arrneeded
EndFunc   ;_FetchNeededData

Func _GetNeededUpdates($objsearcher)
	If Not IsObj($objsearcher) Then Return -5
	$colneeded = $objsearcher.search("IsInstalled=0 and Type='Software'")
	Return $colneeded
EndFunc   ;_GetNeededUpdates

Func _GetTotalHistoryCount($objsearcher)
	If Not IsObj($objsearcher) Then Return -2
	Return $objsearcher.gettotalhistorycount
EndFunc   ;_GetTotalHistoryCount

Func _UserBannedList()
	$sBannedList = GUICtrlRead($label[8]) ;Read what is input into text box
	If $sBannedList = "" Then
			$BannedList = StringSplit("Silverlight", ", ", $STR_ENTIRESPLIT) ;Just stick with Silverlight
		Else
			$BannedList = StringSplit("Silverlight, " & $sBannedList, ", ", $STR_ENTIRESPLIT) ;Use Silverlight and whatever else is in the text box
		EndIf
EndFunc ;Create Banned list

Func _PopulateNeeded($Host)

	GUICtrlSetData($wul,"Searching for updates available...")
		c("Searching for updates available...")
	_GUICtrlListView_DeleteAllItems(ControlGetHandle($Gui, "", $wulv))
	$arrNeeded = _FetchNeededData($Host)
	GUICtrlSetData($wul,"Retrieving results...")
	c("Retrieving results...")
	If IsArray($arrNeeded) And $arrNeeded[0][0] <> "" Then
		For $i = 0 To UBound($arrNeeded) - 1
			_GUICtrlListView_AddItem($wulv, $arrNeeded[$i][0])
			Call(_UserBannedList)
			$Dirty = False
			For $Check = 1 To $BannedList[0]
				If StringInStr($arrNeeded[$i][0], $BannedList[$Check]) Then $Dirty = True
			Next
			If $Dirty == False Then _GUICtrlListView_SetItemSelected($wulv, $i, True)
		Next
		$objsearcher = 0
		$arrNeeded = 0
		_UpdatesDownloadAndInstall()

	Else

		GUICtrlSetData($wul,"Your windows is up to date.")
		c("Your Windows is up to date.")
		Return 0
	EndIf
EndFunc   ;_PopulateNeeded



Func _UpdatesDownloadAndInstall()

	$selected = _GUICtrlListView_GetSelectedIndices($wulv, True)
	If $selected[0] = 0 Then
		c("Results: Your Windows seems up to date.")
		Return 0
	EndIf
	$objsearcher = _CreateMsUpdateSession($Host)
	For $x = 1 To $selected[0]
		$item = _GUICtrlListView_GetItemText($wulv, $selected[$x])
		For $i = 0 To $colneeded.updates.count - 1
			$update = $colneeded.updates.item($i)
			If $item = $update.title Then
				GUICtrlSetData($wul, "Downloaded : " & $x & " / Total : " & $selected[0] & " updates")
				c("Downloaded: " & $x & "  / Total: " & $selected[0] & " updates")
				Global $calculate = $x / $selected[0] * 100
				Global $rawpercent = Number($calculate,1)
				Global $percent = $rawpercent & "%"
				GUICtrlSetData($wup,$rawpercent)
;~ 				GUICtrlSetData($UpdatesPercent,$percent)
				_GUICtrlListView_SetItemText($wulv, $i, "Downloading...", 1)
				_GUICtrlListView_SetItemFocused($wulv, $i)
;~ 				_GUICtrlListView_EnsureVisible($wulv, $i)
				$updatestodownload = ObjCreate("Microsoft.Update.UpdateColl")
				$updatestodownload.add($update)
				$downloadsession = $objsearcher.createupdatedownloader()
				$downloadsession.updates = $updatestodownload
				$downloadsession.download
				_GUICtrlListView_SetItemText($wulv, $i, "Ready", 1)
			EndIf
		Next
	Next
	$rebootneeded = False
	GUICtrlSetData($wup,"0")
	c("All updates are downloaded, proceeding to install updates...")
	For $x = 1 To $selected[0]
		$item = _GUICtrlListView_GetItemText($wulv, $selected[$x])
		For $i = 0 To $colneeded.updates.count - 1
			$update = $colneeded.updates.item($i)
			If $item = $update.title And $update.isdownloaded Then
				GUICtrlSetData($wul, "Installed :" & $x & "  / Total :" & $selected[0] & " updates")
				c("Installed: " & $x & "  / Total: " & $selected[0] & " updates")
				$calculate = $x / $selected[0] * 100
				$rawpercent = Number($calculate,1)
				$percent = $rawpercent & "%"
				GUICtrlSetData($wup,$rawpercent)
;~ 				GUICtrlSetData($UpdatesPercent,$percent)
				_GUICtrlListView_SetItemText($wulv, $i, "Installing...", 1)
				_GUICtrlListView_SetItemFocused($wulv, $i)
;~ 				_GUICtrlListView_EnsureVisible($wulv, $i)
				$installsession = $objsearcher.createupdateinstaller()
				$updatestoinstall = ObjCreate("Microsoft.Update.UpdateColl")
				$updatestoinstall.add($update)
				$installsession.updates = $updatestoinstall
				$installresult = $installsession.install
				If $installresult.rebootrequired Then
					$rebootneeded = True
				EndIf
				_GUICtrlListView_SetItemText($wulv, $i, "Success", 1)
			EndIf
		Next
	Next
	If $rebootneeded Then
;~ 		MsgBox(64, "Reboot required", "A reboot is required to complete the installations, Rebooting in 10 seconds.", 10)
		c("A reboot is required to finish installing updates, ")
;~ 		Shutdown(2 + 4 + 16)
	Else
		_GUICtrlListView_DeleteAllItems($wulv)
		_PopulateNeeded($Host)
	EndIf
	GUICtrlSetData($wup, "0")
	$downloadsession = 0
	$updatestodownload = 0
	Return 0

EndFunc   ;_UpdatesDownloadAndInstall

	#EndRegion WINDOWS UPDATE



	Func _GetSystemInfo()
	Local $sReturn = "OK"

	; OS
	Dim $Obj_WMIService = ObjGet("winmgmts:\\" & "localhost" & "\root\cimv2")
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_OperatingSystem")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		$sInfos &= " OS : " & $Obj_Item.Caption & " " & @OSArch & " version " & $Obj_Item.Version & @CRLF
	Next

	; Computer
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_ComputerSystem")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		$sInfos &= " Manufacturer : " & $Obj_Item.Manufacturer & @CRLF
		$sInfos &= " Model : " & $Obj_Item.Model & @CRLF
		$sInfos &= " RAM : " & Round((($Obj_Item.TotalPhysicalMemory / 1024) / 1024), 0) & " Mo" & @CRLF
	Next

	; Processor
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_Processor")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		$sInfos &= " CPU: " & $Obj_Item.Name & @CRLF
		$sInfos &= " Socket : " & $Obj_Item.SocketDesignation & @CRLF
	Next

	; Graphic card
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_VideoController")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		$sInfos &= " Graphic card : " &$Obj_Item.Name & @CRLF &  @CRLF
	Next

	; HDD
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_DiskDrive")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		if $Obj_Item.MediaType = "Fixed hard disk media" Then
			$sInfos &= " HDD " & $Obj_Item.Index & " : " & $Obj_Item.Model & " - " & Round($Obj_Item.Size / 1000000000, 0) & " Go - Status " & $Obj_Item.Status & @CRLF
			If $Obj_Item.Status <> "OK" Then
				$sReturn = "HDD SMART Results " & $Obj_Item.Index & "  : " & $Obj_Item.Status
			EndIf
		EndIf
	Next

GUICtrlSetData($iEdit, $sInfos )

	Return $sReturn
EndFunc

Func _GetSmart()

	Dim $Obj_WMIService = ObjGet("winmgmts:\\" & "localhost" & "\root\wmi")
		Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from MSStorageDriver_FailurePredictData")
		Local $Obj_Item
		Local $aSmart
		Local $aSmartAttribute[]
		Local $i
		Local $bReturn = False
		Local $aHDDname[2]
		Local $sHDDname

		; List of important SMART values :
		; ID = 1 Rate of errors with read failures
		; ID = 5 Number of allocated sectors Nombre de secteur réalloués
		; ID = 9 Hours of being active
		; ID = 12 Number of boot
		; ID = 194 HDD temperature
		; ID = 197  unstable sectors  count
		; ID = 198  uncorrectable sectors count


		For $Obj_Item In $Obj_Services

			Local $aSmartFound = [5,9,12,194,197,198]
			Local $sMax
			Local $sPos
			Local $sMax

			if IsArray($Obj_Item.VendorSpecific) Then
				$bReturn = True
				$aHDDname = StringRegExp($Obj_Item.InstanceName, 'Ven_(.*)&Prod_(.*)\\',3)
				If(IsArray($aHDDname) And UBound($aHDDname) = 2) Then
					$sHDDname = $aHDDname[0] & " " & $aHDDname[1]
				Else
					$aHDDname = StringRegExp($Obj_Item.InstanceName, 'Disk([A-Za-z0-9]*)',3)
					If(IsArray($aHDDname)) Then
						$sHDDname = $aHDDname[0]
					Else
						$sHDDname = "Undefined"
					EndIf
				EndIf

				$aSmart = $Obj_Item.VendorSpecific

				$sMax = UBound($aSmart) - 1
				For $i=2 To $sMax Step 12
					If _ArrayBinarySearch($aSmartFound, $aSmart[$i]) <> -1 Then
						_ArrayDelete($aSmartFound, "0;" & $aSmart[$i])
						if($aSmart[$i]=9 Or $aSmart[$i]=12) Then
							; calcul POH
							$aSmartAttribute[$aSmart[$i]]=$aSmart[$i+6] * 256 + $aSmart[$i+5]
						Else
							$aSmartAttribute[$aSmart[$i]]=$aSmart[$i+5]
						EndIf

					EndIf
				Next
			EndIf

			$aResults[$sHDDname]=$aSmartAttribute
		Next

	Return $bReturn

EndFunc



Func _Exit()

Exit

EndFunc