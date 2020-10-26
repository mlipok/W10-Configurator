#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icon\gear.ico
#AutoIt3Wrapper_Outfile=W10Configurator.exe
#AutoIt3Wrapper_Res_Fileversion=0.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=W10Configurator
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=Made By|Williamas Kumeliukas
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;*****************************************
;W10Configurator.au3 by Williamas Kumeliukas
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
#EndRegion

HotKeySet("{NUMPAD0}", "_Exit")
Opt("GUIResizeMode", $GUI_DOCKAUTO)

#REGION DECLARE VARIABLES
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
Global $BannedList = StringSplit("Silverlight", "|")
Global $label[10]
Global $task[15]
#EndRegion DECLARE VARIABLES


Gui()


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

 $tree = GUICtrlCreateTreeView(700,60,318,481,BitOR($TVS_SINGLEEXPAND, $TVS_SHOWSELALWAYS,$TVS_FULLROWSELECT, $TVS_DISABLEDRAGDROP,$TVS_CHECKBOXES,$TVS_TRACKSELECT),$WS_EX_CLIENTEDGE);225,10,180,331
	GUICtrlSetFont(-1,15,400,0,"Lucida Bright")
	 $task[1] = GUICtrlCreateTreeViewItem("Windows Updates", $tree)
	 $task[2] = GUICtrlCreateTreeViewItem("rzget", $tree)
	 $task[3] = GUICtrlCreateTreeViewItem("Default = Chrome", $tree)
	 $task[4] = GUICtrlCreateTreeViewItem(".pdf/pdfxml = Reader", $tree)
	 $task[5] = GUICtrlCreateTreeViewItem("Install OEM + Logo", $tree)
	 $task[6] = GUICtrlCreateTreeViewItem("Disable hibernation", $tree)
	 $task[7] = GUICtrlCreateTreeViewItem("Add Office 2010 to C:\", $tree)
	 $task[8] = GUICtrlCreateTreeViewItem("Install Office 2010", $tree)
	 $task[9] = GUICtrlCreateTreeViewItem("ComputerName", $tree)
	 $task[10] = GUICtrlCreateTreeViewItem("Add Office icons to taskbar", $tree)
	 $task[11] = GUICtrlCreateTreeViewItem("Select All", $tree)

$Console = _GUICtrlRichEdit_Create($GUI, "",0,559,1043,258, BitOR($ES_MULTILINE,$WS_VSCROLL,$ES_AUTOVSCROLL));9, 364, 719, 258
	_GUICtrlRichEdit_SetEventMask($Console,$ENM_LINK )
	_GUICtrlRichEdit_AutoDetectURL($Console, True)
	_GUICtrlRichEdit_SetCharColor($Console, 0xFFFFFF)
	_GUICtrlRichEdit_SetBkColor($Console, 0x000000)
	_GUICtrlRichEdit_SetReadOnly($Console, True)
	_GUICtrlRichEdit_HideSelection($Console,True)
	_GUICtrlRichEdit_SetFont($Console, 12, "Consolas")

$label[0] = GUICtrlCreateLabel("Select tasks to do:",730,20,254,26,$SS_CENTER,-1)
 GUICtrlSetFont(-1,12,800,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)



 GUISwitch($GUI,_GUICtrlTab_SetCurFocus($tab,1)&GUICtrlRead ($tab, 1))

$Go = GUICtrlCreateButton("Start", 572, 30, 50, 30, BitOR($BS_DEFPUSHBUTTON, $BS_CENTER))
	GUICtrlSetFont(-1,12,600,0,"Lucida Bright", 5)
	GUICtrlSetResizing($Go,$GUI_DOCKAUTO)

 $label[1] = GUICtrlCreateLabel("Manufacturer:",107,52, 90, -1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[1], $GUI_DOCKAUTO)
 $Manufacturer = GUICtrlCreateInput("",201,52,222,20,-1,$WS_EX_CLIENTEDGE)
GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[2] = GUICtrlCreateLabel("Model:",155,132)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[2],$GUI_DOCKAUTO)
 $model = GUICtrlCreateInput("",201,132,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[3] = GUICtrlCreateLabel("Support Hours:", 123,195)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $supportHours = GUICtrlCreateInput("",201,192,218,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

  $label[4] = GUICtrlCreateLabel("Support Website:",115,255,101,15,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $supportUrl = GUICtrlCreateInput("https://",201,255,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $OEMLogo = GUICtrlCreateInput("",201,319,222,20,-1,$WS_EX_CLIENTEDGE)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $label[5] = GUICtrlCreateLabel("Logo:",127,324,101,15,-1,-1)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)
 $oemfile = GUICtrlCreateButton("...",440,319,44,24,-1,-1)
 GUICtrlSetFont(-1, 10, 400, 0, "MS sans serif", 2)
 GUICtrlSetResizing(-1,$GUI_DOCKAUTO)

 $label[6] = GUICtrlCreateLabel("Computer Name:",115,440,86,15,-1,-1)
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

 $label[7] = GUICtrlCreateLabel("Search not started yet",97,62, -1 , 15, $SS_LEFTNOWORDWRAP)
 GUICtrlSetFont(-1,10,400,0,"Lucida Bright", 5)
 GUICtrlSetResizing($label[7], $GUI_DOCKAUTO)
 $wup = GUICtrlCreateProgress(37,142,621,20,-1,-1)
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

		Case $task[11]
			If BitAND(GUICtrlRead($task[11]), $GUI_CHECKED) Then

				$task[0] = "10"
				c(StringLen($cwe))
				_ArrayDisplay($task)
				For $count = 1 To $task[0] Step 1
				GUICtrlSetState($task[$count], $GUI_CHECKED)
				Next
				c("All tasks are checked.")
				Else

				For $count = 1 To $task[0] Step 1
				GUICtrlSetState($task[$count], $GUI_UNCHECKED)
				Next
				cw("All tasks are unchecked.")

			EndIf

		Case $task[1]
			If BitAND(GUICtrlRead($task[1]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[1], 1))
			Else
				c("Task is unchecked: " & GUICtrlRead($task[1], 1))
			EndIf

		Case $task[2]
			If BitAND(GUICtrlRead($task[2] ), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[2] , 1))
			Else
				c("Task is unchecked: " & GUICtrlRead($task[2], 1) )
			EndIf

		Case $task[3]
			If BitAND(GUICtrlRead($task[3] ), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[3], 1) )
			Else
				c("Task is unchecked: " & GUICtrlRead($task[3], 1) )
			EndIf

		Case $task[4]
			If BitAND(GUICtrlRead($task[4] ), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[4], 1) )
			Else
				c("Task is unchecked: " & GUICtrlRead($task[4], 1) )
			EndIf

		Case $task[5]
			If BitAND(GUICtrlRead($task[5]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[5], 1) )
			Else
				c("Task is unchecked: " & GUICtrlRead($task[5] , 1))
			EndIf

		Case $task[6]
			If BitAND(GUICtrlRead($task[6] ), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[6] , 1))
			Else
				c("Task is unchecked: " & GUICtrlRead($task[6], 1) )
			EndIf

		Case $task[7]
			If BitAND(GUICtrlRead($task[7]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[7], 1))
			Else
				c("Task is unchecked: " & GUICtrlRead($task[7], 1))
			EndIf

		Case $task[8]
			If BitAND(GUICtrlRead($task[8]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlread($task[8], 1))
			Else
				c("Task is unchecked: " & GUICtrlRead($task[8], 1))
			EndIf

		Case $task[9]
			If BitAND(GUICtrlRead($task[9]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[9], 1) )
			Else
				c("Task is unchecked: " & GUICtrlRead($task[9], 1) )
			EndIf

		Case $task[10]
			If BitAND(GUICtrlRead($task[10]), $GUI_CHECKED) Then
				c("Task is checked: " & GUICtrlRead($task[10], 1) )
			Else
				c("Task is unchecked: " & GUICtrlRead($task[10], 1) )
			EndIf

		Case $Go
			c("Configuration started!")
			 ; CONFIGURATION PROCESS START HERE @@@@@@@@@@@@@@@@

			; TO DO WHEN EVERYTHING WILL BE TESTED AND WORKING

			 ; CONFIGURATION PROCESS END HERE @@@@@@@@@@@@@@@@@@
		Case $wus

	EndSwitch
Wend

EndFunc
#EndRegion GUI <<<<<<<<<<<<<<@@@@@@@@@@@@@

Func _IsChecked($control)
    Return BitAnd(GUICtrlRead($control),$GUI_CHECKED) = $GUI_CHECKED
EndFunc

 Func Office2010() ;==> TO DO

$officePath = "CHANGEME"

 EndFunc

 #EndRegion INSTALL OFFICE 2010


 Func defaultBrowser() ;~ TO DO
	 ;HKEY_CURRENT_USER\SOFTWARE\RegisteredApplications
	 ;HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice (ProgId REG_SZ ChromeHTML)
	 ;
 EndFunc



	#Region SetComputerName <<<<<<<<
Func _SetComputerName($sCmpName)

    Local $sLogonKey = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Local $sCtrlKey = "HKLM\SYSTEM\Current\ControlSet"
    Local $aRet

    If StringRegExp($sCmpName, '|/|:|*|?|"|<|>|.|,|~|!|@|#|$|%|^|&|(|)|;|{|}|_|=|+|[|]|x60' & "|'", 0) = 1 Then Return 0

    ; 5 = ComputerNamePhysicalDnsHostname
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





#Region SCREENSAVER
Func ScreenSaver($Alive) ;False to reset / True to prevent/disable sleep/power-savings modes (AND screensaver)

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
		Else
			cw("")
	EndIf
EndFunc
#EndRegion SCREENSAVER





Func _FileCopy($sRemoteName, $d)
	Local $FOF_RESPOND_YES = 16
	Local $FOF_SIMPLEPROGRESS = 256
	$oShell = ObjCreate("shell.application")
	$oShell.namespace($d).CopyHere($sRemoteName,$FOF_SIMPLEPROGRESS)
EndFunc   ;==>_FileCopy


Func _Exit()

	Exit

EndFunc
