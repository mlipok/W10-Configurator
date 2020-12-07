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

#Region Includes <<<<<<<<
#include <Constants.au3>
#include <Date.au3>
#include <StaticConstants.au3>
#include <WinNet.au3>
#include <Inet.au3>
#include <GUIConstantsEx.au3>
#include <GuiButton.au3>
#include <TreeViewConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <AutoItConstants.au3>
#include <Array.au3>
#include <WinAPI.au3>
#include <GuiListBox.au3>
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
#EndRegion Includes <<<<<<<<

Opt("GUIResizeMode", $GUI_DOCKAUTO)


#Region DECLARE VARIABLES  FOR LATER USE

Global  $ColCount, $ConfigDir = @ScriptDir & "\Ressources"
Global $hGUI, $hList, $hInput, $aSelected, $sChosen, $hUP, $hDOWN, $hENTER, $hESC
Global $sCurrInput = "", $aCurrSelected[2] = [-1, -1], $iCurrIndex = -1, $hListGUI = -1
Global $hListGUI, $hList, $hInput, $aSelected, $sChosen, $hUP, $hDOWN, $hENTER, $hESC
Global $sCurrInput = "", $aCurrSelected[2] = [-1, -1], $iCurrIndex = -1, $hListGUI = -1
Global $nas, $cw, $cwe, $c, $cmdfile, $nas, $sRemoteName
Global $regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation", $hGUI
Global $iniFile = $ConfigDir & "\config.ini", $iRead, $isAdmin, $iniOpen = @ScriptDir & "\config.ini"
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
Global $GUI
Global $Go
Global $sServer, $sShare, $letter, $sLocalName = $letter & ":"
Global $sServerShare = $sServer & $sShare
Global $sUserName = "CHANGEME", $sPassword = "CHANGEME"
Global $sResult, $aResults, $sSaverActive
Global $syspin = "syspin.exe"
Global $ColNeeded
Global $Host = @ComputerName
Global $sBannedList
Global $BannedList
Global $label[10]
Global $task[15]
Global $hInput2
Global $sRZVersion ;Local RZGet Version
Global $sORZVersion ;Online RZGet Version
Global $sRZCatalog ;RZ Catalog
Global $sRZGet = FileGetShortName("'" & @ScriptDir & "\Ressources\RZget.exe'")
#EndRegion DECLARE VARIABLES  FOR LATER USE


GUI()

#Region  GUI SCRIPT <<<<<<<< @@@@@@@@@@@@@@@@@@@@

Func GUI()

	Global $GUI = GUICreate(@ScriptName & " - Version: " & FileGetVersion(@ScriptFullPath), 1043, 820, -1, -1, -1, -1) ;740,632)
	GUISetIcon("Icon\gear.ico", $GUI)
	$tab = GUICtrlCreateTab(0, 0, 697, 541, -1, -1)
	GUICtrlSetState($tab, $GUI_ONTOP) ;2048
	GUICtrlCreateTabItem("System Info")
	GUICtrlCreateTabItem("Configuration")
	GUICtrlCreateTabItem("Windows Updater")
	GUICtrlCreateTabItem("")
	_GUICtrlTab_SetCurFocus($tab, -1)
	GUICtrlSetResizing(-1, 70)

	;Menu items
	local $FileMenu = GUICtrlCreateMenu("File")
	local $FileImp = GUICtrlCreateMenuItem("Import", $FileMenu)
	GUICtrlSetState(-1, $GUI_DEFBUTTON)
	local $FileExp = GUICtrlCreateMenuItem("Export",$FileMenu)
	GUICtrlCreateMenuItem("", $Filemenu, 2) ; create a separator line
	local $FileEx = GUICtrlCreateMenuItem("Exit",$FileMenu)
	local $HelpMenu = GUICtrlCreateMenu("Help")
	local $HelpInfo = GUICtrlCreateMenuItem("Configurator Help",$HelpMenu)
	local $HelpRep = GUICtrlCreateMenuItem("Report Issue",$HelpMenu)
	local $HelpGit = GUICtrlCreateMenuItem("GitHub",$HelpMenu)
	local $HelpAbout = GUICtrlCreateMenuItem("About",$HelpMenu)

	#Region GUI setup for Tasks
	GUICtrlCreateGraphic(705, 45, 300, 455, BitOR($GUI_SS_DEFAULT_GRAPHIC, $SS_WHITEFRAME, $WS_BORDER))
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetGraphic(-1, $GUI_GR_COLOR, 0x000000, 0xFFFFFF)
	$task[1] = GUICtrlCreateCheckbox("Windows Updater", 710, 48, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)
	$task[2] = GUICtrlCreateCheckbox("Install Softwares", 710, 70, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$hInput2 = GUICtrlCreateInput("", 710, 440, 280, 30)
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)
	GUICtrlSetState(-1, $GUI_DISABLE)

	$task[3] = GUICtrlCreateCheckbox("Default = Chrome", 710, 92, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[4] = GUICtrlCreateCheckbox(".pdf/pdfxml = Reader", 710, 114, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[5] = GUICtrlCreateCheckbox("Install OEM + Logo", 710, 136, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[6] = GUICtrlCreateCheckbox("Disable hibernation", 710, 158, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[7] = GUICtrlCreateCheckbox("Copy Office 2010 to C:\", 710, 180, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[8] = GUICtrlCreateCheckbox("Install Office 2010", 710, 202, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[9] = GUICtrlCreateCheckbox("ComputerName", 710, 224, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[10] = GUICtrlCreateCheckbox("Add Office icons to taskbar", 710, 246, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[11] = GUICtrlCreateCheckbox("Select All", 710, 268, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$sRunTasks = GUICtrlCreateButton("Run", 910, 510, 80, 30)

	$console = _GUICtrlRichEdit_Create($GUI, "", 0, 559, 1043, 258, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL)) ;9, 364, 719, 258
	_GUICtrlRichEdit_SetEventMask($console, $ENM_LINK)
	_GUICtrlRichEdit_AutoDetectURL($console, True)
	_GUICtrlRichEdit_SetCharColor($console, 0xFFFFFF)
	_GUICtrlRichEdit_SetBkColor($console, 0x000000)
	_GUICtrlRichEdit_SetReadOnly($console, True)
	_GUICtrlRichEdit_HideSelection($console, True)
	_GUICtrlRichEdit_SetFont($console, 12, "Consolas")

	$label[0] = GUICtrlCreateLabel("Select tasks to do:", 780, 23, 150, 18, $SS_CENTER, -1)
	GUICtrlSetFont(-1, 12, 800, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	ConsoleWriteGUI($console, "=============== Win10 CONFIGURATOR CONSOLE ===============")

	;Config Tab
	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 1) & GUICtrlRead($tab, 1))

GUICtrlCreateGroup("OEM",6,35,380,200)
	$label[1] = GUICtrlCreateLabel("Manufacturer:", 31, 52, 95)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[1], $GUI_DOCKAUTO)
	$Manufacturer = GUICtrlCreateInput("", 125, 50, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[2] = GUICtrlCreateLabel("Model:", 77, 90)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[2], $GUI_DOCKAUTO)
	$model = GUICtrlCreateInput("", 125, 88, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[3] = GUICtrlCreateLabel("Support Hours:", 24, 128, 95)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$supportHours = GUICtrlCreateInput("", 125, 126, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[4] = GUICtrlCreateLabel("Support Website:", 10, 166, 110)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$supportUrl = GUICtrlCreateInput("https://", 125, 164, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$OEMLogo = GUICtrlCreateInput("", 125, 202, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$label[5] = GUICtrlCreateLabel("Logo:", 83, 204, 101)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$oemfile = GUICtrlCreateButton("...", 350, 202, 30, 20, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "MS sans serif", 2)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

GUICtrlCreateGroup("Misc",6,235,380,200)
	$label[6] = GUICtrlCreateLabel("Computer Name:", 12, 252, 110, 15, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$compname = GUICtrlCreateInput($Host, 125, 250, 222, 20, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	;System Info tab
	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 0) & GUICtrlRead($tab, 1))

	Global $iEdit = GUICtrlCreateEdit("", 5, 25, 677, 481, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $ES_MULTILINE, $ES_UPPERCASE, $WS_VSCROLL, $WS_HSCROLL), -1)

	_GetSystemInfo()

	;Windows Updater tab
	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 2) & GUICtrlRead($tab, 1))

	$wulv = GUICtrlCreateListView("Descriptions|Status", 17, 192, 647, 334, -1, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	_GUICtrlListView_SetColumnWidth($wulv, 0, 500)
	GUICtrlSetBkColor(-1, $color_silver)
	GUICtrlSetColor(-1, 0x75EE3B)
	_GUICtrlListView_SetColumnWidth($wulv, 1, 140)
	GUICtrlSetBkColor(-1, $color_black)
	GUICtrlSetColor(-1, 0x75EE3B)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[7] = GUICtrlCreateLabel("Search not started yet", 16, 170, 150, 15, $SS_LEFTNOWORDWRAP)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[7], $GUI_DOCKAUTO)
	$wup = GUICtrlCreateProgress(16, 142, 649, 20, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[8] = GUICtrlCreateInput("", 16, 90, 300, 40) ;User defined additions to banned list
	GUICtrlCreateLabel("Use the below box to exclude updates from being applied: (ex. Silverlight)", 16, 55, 310, 40)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$wus = GUICtrlCreateButton("Search", 500, 80, 100, 30, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 1) & GUICtrlRead($tab, 1))

	_GUICtrlTab_SetCurFocus($tab, 0)

	GUISetState(@SW_SHOW, $GUI)
;~  return $GUI
	#EndRegion GUI setup for Tasks

	While 1

		Switch GUIGetMsg()

			Case $GUI_EVENT_CLOSE
				Exit
			Case $FileEx
				Exit
			Case $FileImp
				_ImportIni()
			Case $FileExp
				_Exportini()
			Case $HelpAbout
				MsgBox(8256, "About","Windows 10 Configurator was created as a project to help auotmate some repitive tasks being used by Network Administrators, Helpdesk personnel, etc.")
			Case $HelpGit
				ShellExecute("https://github.com/Jivaross/W10-Configurator")
			Case $HelpInfo
				MsgBox(8224, "Configurator Help","Insert halp Text here.")
			Case $HelpRep
				ShellExecute("https://github.com/Jivaross/W10-Configurator/issues")
			Case $task[11]
				If GUICtrlRead($task[11]) == $GUI_CHECKED Then
					For $i = 10 To 1 Step -1
						GUICtrlSetState($task[$i], $GUI_CHECKED)
					Next
					c("All tasks are checked.")
				ElseIf GUICtrlRead($task[11]) == $GUI_UNCHECKED Then
					For $i = 10 To 1 Step -1
						GUICtrlSetState($task[$i], $GUI_UNCHECKED)
					Next
					c("All tasks are unckecked.")
				EndIf

			Case $task[2]
				If GUICtrlRead($task[2]) == $GUI_Checked Then
					;If GUICtrlGetState($hInput2) = "144" Then GUICtrlSetState($hInput2, $GUI_ENABLE)
					;If FileExists($ConfigDir & "\rzget.bat") Then FileDelete($configDir & "\rzget.bat")
					_RZCatalog()

				Else
					;If GUICtrlGetState($hInput2) = "80" Then GUICtrlSetState($hInput2, $GUI_DISABLE)
					GUICtrlSetData($hInput2, "") ; input emptied
				EndIf


			Case $sRunTasks


				c("Configuration started !")

				If GUICtrlRead($task[2]) == $GUI_Checked Then

					Rzget()
					GUICtrlSetState($task[2], $GUI_UNCHECKED)
					GUICtrlSetData($hInput2, "") ; emptied
				EndIf

				If GUICtrlRead($task[3]) == $GUI_Checked Then
					c("Setting Chrome as default browser..")
					defaultBrowser()
				EndIf

				If GUICtrlRead($task[4]) == $GUI_Checked Then
					c("Setting .PDF and .PDFXML to Adobe Reader...")
				EndIf

				If GUICtrlRead($task[5]) == $GUI_Checked Then
					c("Installing OEM + Logo...")
					selfoem()
				EndIf

				If GUICtrlRead($task[6]) == $GUI_Checked Then
					c("Disabling Hibernation...")
					;ScreenSaver(True) ;True to disable / false to reset
				EndIf

				If GUICtrlRead($task[7]) == $GUI_Checked Then
					c("Adding office 2010 to C:\...")
				EndIf

				If GUICtrlRead($task[8]) == $GUI_Checked Then
					c("Installing MS Office 2010...")
					Office2010()
				EndIf

				If GUICtrlRead($task[9]) == $GUI_Checked Then
					c("Changing computer name...")
					;_SetComputerName() ;Temporarily commented out to prevent changing of name
				EndIf

				If GUICtrlRead($task[1]) == $GUI_Checked Then
					c("Running Windows Updater...")
					_PopulateNeeded($Host)
				EndIf

				If GUICtrlRead($task[10]) == $GUI_Checked Then
					c("Adding Office icons to taskbar...")
				EndIf
				c("All tasks are completed!")

			Case $wus
				_PopulateNeeded($Host)

			Case $oemfile
				$cLogo = FileOpenDialog(@ScriptName, @DesktopDir, "BMP files (*.bmp)", $FD_FILEMUSTEXIST)
				GUICtrlSetData($OEMLogo, $cLogo)
		EndSwitch
	WEnd

EndFunc   ;==>GUI
#EndRegion  GUI SCRIPT <<<<<<<< @@@@@@@@@@@@@@@@@@@@

Func _Main()


	$hGUI = GUICreate("Software Selector", 300, 100)
	GUICtrlCreateLabel("Start to type letters and I'll suggest softwares:", 10, 10, 280, 20)
	$hInput = GUICtrlCreateInput("", 10, 40, 280, 20)
	$hButton = GUICtrlCreateButton("Choose", 250, 70)
	GUISetState(@SW_SHOW, $hGUI)

	$hUP = GUICtrlCreateDummy()
	$hDOWN = GUICtrlCreateDummy()
	$hENTER = GUICtrlCreateDummy()
	$hESC = GUICtrlCreateDummy()
	Dim $AccelKeys[4][2] = [["{UP}", $hUP], ["{DOWN}", $hDOWN], ["{ENTER}", $hENTER], ["{ESC}", $hESC]]
	GUISetAccelerators($AccelKeys)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $hButton
				GUICtrlSetData($hInput2, GUICtrlRead($hInput))
				ExitLoop
			Case $hESC
				If $hListGUI <> -1 Then ; List is visible.
					GUIDelete($hListGUI)
					$hListGUI = -1
				Else
					ExitLoop
				EndIf

			Case $hUP
				If $hListGUI <> -1 Then ; List is visible.
					$iCurrIndex -= 1
					If $iCurrIndex < 0 Then
						$iCurrIndex = 0
					EndIf
					_GUICtrlListBox_SetCurSel($hList, $iCurrIndex)
				EndIf

			Case $hDOWN
				If $hListGUI <> -1 Then ; List is visible.
					$iCurrIndex += 1
					If $iCurrIndex > _GUICtrlListBox_GetCount($hList) - 1 Then
						$iCurrIndex = _GUICtrlListBox_GetCount($hList) - 1
					EndIf
					_GUICtrlListBox_SetCurSel($hList, $iCurrIndex)
				EndIf

			Case $hENTER
				If $hListGUI <> -1 And $iCurrIndex <> -1 Then ; List is visible and a item is selected.
					$sChosen = _GUICtrlListBox_GetText($hList, $iCurrIndex)
				EndIf

			Case $hList
				$sChosen = GUICtrlRead($hList)
		EndSwitch

		Sleep(10)
		$aSelected = _GetSelectionPointers($hInput)
		If GUICtrlRead($hInput) <> $sCurrInput Or $aSelected[1] <> $aCurrSelected[1] Then ; Input content or pointer are changed.
			$sCurrInput = GUICtrlRead($hInput)
			$aCurrSelected = $aSelected ; Get pointers of the string to replace.
			$iCurrIndex = -1
			If $hListGUI <> -1 Then ; List is visible.
				GUIDelete($hListGUI)
				$hListGUI = -1
			EndIf
			$hList = _PopupSelector($hGUI, $hListGUI, _CheckInputText($sCurrInput, $aCurrSelected)) ; ByRef $hListGUI, $aCurrSelected.
		EndIf
		If $sChosen <> "" Then
			GUICtrlSendMsg($hInput, 0x00B1, $aCurrSelected[0], $aCurrSelected[1]) ; $EM_SETSEL.
			_InsertText($hInput, "'" &$sChosen & "' ")
			$sCurrInput = GUICtrlRead($hInput)
			GUIDelete($hListGUI)
			$hListGUI = -1
			$sChosen = ""
		EndIf
	WEnd
	GUIDelete($hGUI)
EndFunc   ;==>_Main
Func _CheckInputText($sCurrInput, ByRef $aSelected)
	Local $sPartialData = ""
	If (IsArray($aSelected)) And ($aSelected[0] <= $aSelected[1]) Then
		Local $aSplit = StringSplit(StringLeft($sCurrInput, $aSelected[0]), " ")
		$aSelected[0] -= StringLen($aSplit[$aSplit[0]])
		If $aSplit[$aSplit[0]] <> "" Then
			For $A = 1 To $sRZCatalog[0] - 1
;~ 				c($sRZCatalog[$A])
				If StringLeft($sRZCatalog[$A], StringLen($aSplit[$aSplit[0]])) = $aSplit[$aSplit[0]] And $sRZCatalog[$A] <> $aSplit[$aSplit[0]] Then
					$sPartialData &= $sRZCatalog[$A] & "|"
				EndIf
			Next
		EndIf
	EndIf
	Return $sPartialData
EndFunc   ;==>_CheckInputText

Func _PopupSelector($hGUI, ByRef $hListGUI, $sCurr_List)

	Local $hList = -1
	If $sCurr_List = "" Then
		Return $hList
	EndIf
	$hListGUI = GUICreate("", 280, 360, 10, 62, $WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_MDICHILD), $hGUI)
	$hList = GUICtrlCreateList("", 0, 0, 280, 350, BitOR(0x00100000, 0x00200000))
	GUICtrlSetData($hList, $sCurr_List)
	GUISetControlsVisible($hListGUI) ; To Make Control Visible And Window Invisible.
	GUISetState(@SW_SHOWNOACTIVATE, $hListGUI)
	Return $hList
EndFunc   ;==>_PopupSelector

Func _InsertText(ByRef $hEdit, $sString)

	Local $aSelected = _GetSelectionPointers($hEdit)
	GUICtrlSetData($hEdit, StringLeft(GUICtrlRead($hEdit), $aSelected[0]) & $sString & StringTrimLeft(GUICtrlRead($hEdit), $aSelected[1]))
	Local $iCursorPlace = StringLen(StringLeft(GUICtrlRead($hEdit), $aSelected[0]) & $sString)
	GUICtrlSendMsg($hEdit, 0x00B1, $iCursorPlace, $iCursorPlace) ; $EM_SETSEL.
EndFunc   ;==>_InsertText

Func _GetSelectionPointers($hEdit)
	Local $aReturn[2] = [0, 0]
	Local $aSelected = GUICtrlRecvMsg($hEdit, 0x00B0) ; $EM_GETSEL.
	If IsArray($aSelected) Then
		$aReturn[0] = $aSelected[0]
		$aReturn[1] = $aSelected[1]
	EndIf
	Return $aReturn
EndFunc   ;==>_GetSelectionPointers

Func GUISetControlsVisible($hWnd) ; By Melba23.
	Local $aControlGetPos = 0, $hCreateRect = 0, $hRectRgn = _WinAPI_CreateRectRgn(0, 0, 0, 0)
	Local $iLastControlID = _WinAPI_GetDlgCtrlID(GUICtrlGetHandle(-1))
	For $i = 3 To $iLastControlID
		$aControlGetPos = ControlGetPos($hWnd, '', $i)
		If IsArray($aControlGetPos) = 0 Then ContinueLoop
		$hCreateRect = _WinAPI_CreateRectRgn($aControlGetPos[0], $aControlGetPos[1], $aControlGetPos[0] + $aControlGetPos[2], $aControlGetPos[1] + $aControlGetPos[3])
		_WinAPI_CombineRgn($hRectRgn, $hCreateRect, $hRectRgn, 2)
		_WinAPI_DeleteObject($hCreateRect)
	Next
	_WinAPI_SetWindowRgn($hWnd, $hRectRgn, True)
	_WinAPI_DeleteObject($hRectRgn)
EndFunc   ;==>GUISetControlsVisible


Func c($cw) ;INFORMATIVE MESSAGES IN CONSOLE
	$c = 0
	ConsoleWrite($cw & @CRLF)
	ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> " & $cw)

EndFunc   ;==>c

Func cw($cwe) ; @@@@ CONSOLE WARNINGS / ERRORS MESSAGE @@@@@
	$c = 1
	ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> ERROR: " & $cwe)
EndFunc   ;==>cw

Func ConsoleWriteGUI(Const ByRef $hConsole, Const $sTxt) ; READ C($cw) & cw($cw) ($cwe_ soon)
	_GUICtrlRichEdit_GotoCharPos($console, -1)
	If $c = 1 Then
		_GUICtrlRichEdit_SetCharColor($console, 0x0000FF)
		_GUICtrlRichEdit_InsertText($console, $sTxt)
	Else
		_GUICtrlRichEdit_SetCharColor($console, 0xFFFFFF)
		_GUICtrlRichEdit_InsertText($console, $sTxt)
	EndIf

	_GUICtrlRichEdit_SetFont($console, 12, "Consolas")

	_GUICtrlRichEdit_HideSelection($console, True)
EndFunc   ;==>ConsoleWriteGUI


#Region INI READ/WRITE
Func ini($section, $key, $value, $iniread = True) ;Read or edit value in config.ini

	;$iniOpen = FileOpen($iniFile, 2)
	;c("read state = " & $iniread)
	Switch $iniread

		Case False ; = overwrite

			Return IniWrite($iniFile, $section, $key, $value)

		Case True ; = read only

			Return IniRead($iniFile, $section, $key, $value)

	EndSwitch

EndFunc   ;==>ini
Func _ImportIni()
	C("Importing settings...")
	$iniFile = FileOpenDialog("Choose File...", @WorkingDir, "Config files (*.ini)")
	;Set values of check mark boxes, text boxes, etc. based on sections read within ini file array using inireadsection (creates 2d array) or iniread (reads single key value)
EndFunc
Func _ExportIni()
	c("Exporting settings...")
	;Add in ini file export items.
EndFunc
#EndRegion INI READ/WRITE

Func selfoem() ;**** AutoIt version of oem() ****
	c("" & @CRLF & "OEM installation started !")
	c("Installing OEM Logo...")
	_FileCopy($OEMLogo, "C:\Windows\System32")
	RegWrite($regPath, "Manufacturer", "REG_SZ", $Manufacturer)
	c("Added Manufacturer registry key...")
	RegWrite($regPath, "Model", "REG_SZ", $model)
	c("Added Model registry key...")
	RegWrite($regPath, "Logo", "REG_SZ", $OEMLogo)
	c("Added Logo (OEM Logo references) registry key...")
	RegWrite($regPath, "SupportHours", "REG_SZ", $supportHours)
	c("Added SupportHours registry key...")
	RegWrite($regPath, "SupportPhone", "REG_SZ", $supportPhone)
	c("Added SupportPhone registry key...")
	RegWrite($regPath, "SupportURL", "REG_SZ",  $supportUrl)
	c("Added SupportURL registry key...")
	c("OEM installation applied !")

EndFunc   ;==>selfoem


#Region RZGET <<<  Work flawless (RZGet is another solution like Ninite which offers a variety of softwares to download & install)
Func Rzget()
	Local $rzcmd

	If FileExists($ConfigDir & "\rzget.exe") Then ; IF RZGET IS IN CONFIG FOLDER @@@@
		_CheckRZGetVersion()
		c("Installing "&GUICtrlRead($hInput2))
		RunWait('powershell -Command "(& '&$sRZGet&' install '&GUICtrlRead($hInput2)&') | Out-String"', @SystemDir, @SW_SHOW, $STDOUT_CHILD + $STDERR_CHILD)
		c(GUICtrlRead($hInput2)&"done installing.")
;		$rzcmd = '@echo off' & @CRLF _
;				& 'call :isAdmin' & @CRLF _
;				& 'if %errorlevel% == 0 (' & @CRLF _
;				& 'goto :run' & @CRLF _
;				& ') else (' & @CRLF _
;				& 'echo Requesting administrative privileges...' & @CRLF _
;				& 'goto :UACPrompt' & @CRLF _
;				& ')' & @CRLF _
;				& 'exit /b' & @CRLF _
;				& ':isAdmin' & @CRLF _
;				& 'fsutil dirty query %systemdrive% >nul' & @CRLF _
;				& 'exit /b' & @CRLF _
;				& ':run' & @CRLF _
;				& 'cmd /c ' & FileGetShortName(@ScriptDir & "\Ressources\rzget.exe") & " install " & GUICtrlRead($hInput2) & @CRLF _
;				& 'echo done.' & @CRLF _
;				& 'exit /b' & @CRLF _
;				& ':UACPrompt' & @CRLF _
;				& 'echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"' & @CRLF _
;				& 'echo UAC.ShellExecute "cmd.exe", "/c %~s0 %~1", "", "runas", 1 >> "%temp%\getadmin.vbs"' & @CRLF _
;				& '"%temp%\getadmin.vbs"' & @CRLF _
;				& 'del "%temp%\getadmin.vbs"' & @CRLF _
;				& 'exit /B'
;;============================================================================
;				c("Creating Rzget process...")
;				FileWrite($ConfigDir & "\rzget.bat", $rzcmd)
;				c("waiting for Rzget to load...")
;				ShellExecute($ConfigDir & "\rzget.bat")
;				c("Done !")
	Else ; RZGET.EXE WAS NOT FOUND IN CONFIG FOLDER @@@@
		c("ruckzuck executable not found...")
	EndIf
	GUICtrlSetData($hInput2, "")
EndFunc   ;==>Rzget

Func _GetOnlineRZVersion()
	Local $Powershell = Run("powershell -Command (((curl 'https://ruckzuck.tools' -Usebasicparsing).Links | Select href)| Select-String -Pattern 'RZGet.exe')", @SystemDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
	$sORZVersion = ""
	While 1
		$sORZVersion &= StdoutRead($Powershell)
		If @error Then ExitLoop
	WEnd
	$sORZVersion = StringReplace(StringTrimRight(StringTrimLeft($sORZVersion, 63), 6), "/RZGet.exe}", "")
EndFunc   ;==>_GetOnlineRZVersion

Func _CheckRZGetVersion()
	$sRZVersion = FileGetVersion(@ScriptDir & "\Ressources\Rzget.exe")
	Call(_GetOnlineRZVersion)
	If $sORZVersion <> $sRZVersion Then
		c("Updating RZGet Version")
		InetGet("https://github.com/rzander/ruckzuck/releases/download/" & $sORZVersion & "/Rzget.exe", @ScriptDir & "\Ressources\Rzget.exe", 2, 0)
	Else
		c("RZGet is up to date.")
	EndIf
EndFunc   ;==>_CheckRZGetVersion

Func _RZCatalog()
	$sRZCatalog = 0
	c("Cleared RZ catalog data.")
	c("Retrieving updated catalog...")
	Local $Powershell = Run("powershell -Command ((& " & $sRZGet & " search | convertfrom-json) | Select ShortName | Out-String -Stream | foreach {$_.trimend()} | Where {$_ -ne ''} | Where {$_ -ne '---------'} | Where {$_ -ne 'ShortName'})", @SystemDir, @SW_HIDE, $STDOUT_CHILD + $STDERR_CHILD)
	ProcessWaitClose($Powershell)
	$sRZCatalog = StringSplit(StdoutRead($Powershell), @CRLF, 1)
	Local $ColCount = UBound($sRZCatalog, 1) - 1
	_ArrayDelete($sRZCatalog, $ColCount) ;Removes the blank line at the end
	If IsArray($sRZCatalog) = True Then
		c("Catalog updated.")
		;_ArraySort($sRZCatalog)
		c("There are " & $ColCount & " software available for install.")
		_Main()
	Else
		c("Error in updating catalog...")
	EndIf

EndFunc   ;==>_RZCatalog
#EndRegion RZGET <<<  Work flawless (RZGet is another solution like Ninite which offers a variety of softwares to download & install)

#Region SCREENSAVER
Func ScreenSaver($Alive) ;False = reset to default & True to prevent/disable sleep/power-savings modes (AND screensaver)

	If $Alive = True Then
		Local $ssKeepAlive = DllCall('kernel32.dll', 'long', 'SetThreadExecutionState', 'long', 0x80000003)
		If @error Then
			Return SetError(2, @error, 0x80000000)
			cw(@error)
		EndIf
		Return $ssKeepAlive[0] ; Previous state (typically 0x80000000 [-2147483648])
		c("Hibernate mode = " & $sSaverActive)
	ElseIf $Alive = False Then
		; Flag:	ES_CONTINUOUS (0x80000000) -> (default) -> used alone, it resets timers & allows regular sleep/power-savings mode
		Local $ssKeepAlive = DllCall('kernel32.dll', 'long', 'SetThreadExecutionState', 'long', 0x80000000)
		If @error Then
			Return SetError(2, @error, 0x80000000)
			cw(@error)
		EndIf
		Return $ssKeepAlive[0]    ; Previous state
	EndIf
EndFunc   ;==>ScreenSaver
#EndRegion SCREENSAVER

Func _FileCopy($sRemoteName, $d)
	Local $FOF_RESPOND_YES = 16
	Local $FOF_SIMPLEPROGRESS = 256
	$oShell = ObjCreate("shell.application")
	$oShell.namespace($d).CopyHere($sRemoteName, $FOF_SIMPLEPROGRESS)
EndFunc   ;==>_FileCopy

#Region NAS SCRIPT <<<<<<<<<
Func nas()

	c("Verifying if there is an existing connection to: >>> NAS-02 <<< ...")
	If _WinNet_GetConnection($sLocalName) = $sServerShare Then
		c("Connection aldready existing, skipping this step....")
	Else
		If $sUserName Or $sPassword = "CHANGEME" Then
			MsgBox(0, "", "Username or Password must be changed before attempting to connect NAS.")
			Exit
		EndIf
		c("Attributing letter: " & $letter)
		$sResult = _WinNet_AddConnection2($sLocalName, $sServerShare, $sUserName, $sPassword, 1, 2)  ;(1,2 = remember connection and user interact on error)
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

EndFunc   ;==>nas
#EndRegion NAS SCRIPT <<<<<<<<<

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

#Region DefaultBrowser >>>>>>>>>>>>
Func defaultBrowser()
	$chrome = "@echo off" & @CRLF _
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
			 & 'cmd /c ' & $ConfigDir & '\SetDefaultBrowser.exe chrome' & @CRLF _
			 & 'echo done.' & @CRLF _
			 & 'exit /b' & @CRLF _
			 & ':UACPrompt' & @CRLF _
			 & 'echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"' & @CRLF _
			 & 'echo UAC.ShellExecute "cmd.exe", "/c %~s0 %~1", "", "runas", 1 >> "%temp%\getadmin.vbs"' & @CRLF _
			 & '"%temp%\getadmin.vbs"' & @CRLF _
			 & 'del "%temp%\getadmin.vbs"' & @CRLF _
			 & 'exit /B'

	FileWrite($ConfigDir & "\defChrome.bat", $chrome)
	c("Wait...")
	ShellExecute($ConfigDir & "\defChrome.bat")
	c("Done !")
EndFunc   ;==>defaultBrowser
#EndRegion DefaultBrowser >>>>>>>>>>>>

#Region DefaultPDF
Func defaultPDF()
	c("Beginning to change default Application association to .pdf extension...")
	$vbs_def_pdf = @CRLF _
			 & "@echo off" & @CRLF _
			 & "cd " & $ConfigDir & "\Ressources\" & @CRLF _
			 & "SetUserFTA.exe .pdf AcroExch.Document.DC Everyone" & @CRLF _
			 & "SetUserFTA.exe .pdfxml AcroExch.Document.DC Everyone" & @CRLF _
			 & "exit /b"

	FileWrite(@TempDir & "\defaultPDF.bat", $vbs_def_pdf)
	ShellExecute("defaultPDF.bat", "", @TempDir, "open", @SW_SHOW) ; change to @SW_HIDE before release!
	c("Done, .pdf and .pdfxml extensions are now associated to: Adobe Reader DC !")
EndFunc   ;==>defaultPDF
#EndRegion DefaultPDF

#Region PIN TO TASKBAR >>>>>>>>>>>

;=======H   E   L    P   SYSPIN   H   E   L  P===========
;	syspin ["file"] #### or syspin ["file"] "commandstring"
;	5386 = Pin to taskbar
;	5387 = Unpin from taskbar
;	51201 = Pin to start
;	51394 = Unpin  from start
;Examples :
;	syspin "%PROGRAMFILES%\Internet Explorer\iexplore.exe" 5386
;	syspin "C:\Windows\notepad.exe" "Pin to taskbar"
;	syspin "%WINDIR%\System32\calc.exe" "Pin to start"
;	syspin "%WINDIR%\System32\calc.exe" "Unpin from taskbar"
;	syspin "C:\Windows\System32\calc.exe" 51201
;
;Note : You cannot pin any metro app or batch file.
;
;======== H   E   L    P   SYSPIN   H   E   L  P===========

Func syspin()
	; MAGIC NUMBERS "5386" SEEMS TO WORK BETTER THAN "PIN TO TASKBAR"

	$Excel = '"%PROGRAMFILES(X86)%\Microsoft Office\Office14\EXCEL.EXE" 5386'
	$Word = '"%PROGRAMFILES(X86)%\Microsoft Office\Office14\WORD.EXE" 5386'
	$Outlook = '"%PROGRAMFILES(X86)%\Microsoft Office\Office14\OUTLOOK.EXE" 5386'
	$Edge = '"%PROGRAMFILES(X86)%\Microsoft\Edge\Application\msedge.exe" 5387'
	$chrome = '"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe" 5386'

	c("Pinning Excel to taskbar...")
	ShellExecuteWait($syspin, $Excel, $ConfigDir, "open", @SW_HIDE)
	c("Pinning Outlook to taskbar...")
	ShellExecuteWait($syspin, $Outlook, $ConfigDir, "open", @SW_HIDE)
	c("Pinning Word to taskbar...")
	ShellExecuteWait($syspin, $Word, $ConfigDir, "open", @SW_HIDE)
	c("Pinning Chrome to taskbar...")
	ShellExecuteWait($syspin, $chrome, $ConfigDir, "open", @SW_HIDE)
	c("Unpinning crappy Edge from taskbar...")
	ShellExecuteWait($syspin, $Edge, $ConfigDir, "open", @SW_HIDE)
	c("==========>> Pin to taskbar completed ! <<=============")
EndFunc   ;==>syspin
#EndRegion PIN TO TASKBAR >>>>>>>>>>>


#Region 		 INSTALL OFFICE 2010
Func Office2010()  ;==> TO DO (Professional Plus version OR 365)
	MsgBox(16, @ScriptName, "If you see this message, it means this function is still under construction.")

	$officePath = "CHANGEME"

EndFunc   ;==>Office2010
#EndRegion 		 INSTALL OFFICE 2010

Func _IsChecked($control)
	Return BitAND(GUICtrlRead($control), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked


#Region WINDOWS UPDATE

Func _CreateMsUpdateSession($strhost = @ComputerName)
	c("Creating a Windows Update Session...")
	$objsession = ObjCreate("Microsoft.Update.Session", $strhost)
	If Not IsObj($objsession) Then Return 0
	Return $objsession
EndFunc   ;==>_CreateMsUpdateSession

Func _CreateSearcher($objsession)
	c("Creating Searcher Session..." & @CRLF & "Searching for updates available...")
	If Not IsObj($objsession) Then Return -1
	Return $objsession.createupdatesearcher
EndFunc   ;==>_CreateSearcher

Func _FetchNeededData($Host)
	$objsearcher = _CreateSearcher(_CreateMsUpdateSession($Host))
	$ColNeeded = _GetNeededUpdates($objsearcher)
	$objsearcher = 0
	Dim $arrneeded[1][2]
	For $i = 0 To $ColNeeded.updates.count - 1
		If $i < $ColNeeded.updates.count - 1 Then ReDim $arrneeded[$i + 2][2]
		$update = $ColNeeded.updates.item($i)
		$arrneeded[$i][0] = $update.title
		$arrneeded[$i][1] = $update.description
	Next
	If Not IsArray($arrneeded) Then
		cw("Windows Updates Service seems to have encounted a problem with: " & $Host)
		Return 0
	EndIf
	Return $arrneeded
EndFunc   ;==>_FetchNeededData

Func _GetNeededUpdates($objsearcher)
	If Not IsObj($objsearcher) Then Return -5
	$ColNeeded = $objsearcher.search("IsInstalled=0 and Type='Software'")
	Return $ColNeeded
EndFunc   ;==>_GetNeededUpdates

Func _GetTotalHistoryCount($objsearcher)
	If Not IsObj($objsearcher) Then Return -2
	Return $objsearcher.gettotalhistorycount
EndFunc   ;==>_GetTotalHistoryCount

Func _UserBannedList()
	$sBannedList = GUICtrlRead($label[8]) ;Read what is input into text box
	If $sBannedList = "" Then
		$BannedList = StringSplit("Silverlight", ", ", $STR_ENTIRESPLIT)     ;Just stick with Silverlight
	Else
		$BannedList = StringSplit("Silverlight, " & $sBannedList, ", ", $STR_ENTIRESPLIT)     ;Use Silverlight and whatever else is in the text box
	EndIf
EndFunc   ;==>_UserBannedList

Func _PopulateNeeded($Host)

	GUICtrlSetData($wul, "Searching for updates available...")
	GUICtrlSetData($label[7], "Searching for available updates...")
	c("Searching for updates available...")
	_GUICtrlListView_DeleteAllItems(ControlGetHandle($GUI, "", $wulv))
	$arrneeded = _FetchNeededData($Host)
	GUICtrlSetData($wul, "Retrieving results...")
	c("Retrieving results...")
	If IsArray($arrneeded) And $arrneeded[0][0] <> "" Then
		For $i = 0 To UBound($arrneeded) - 1
			_GUICtrlListView_AddItem($wulv, $arrneeded[$i][0])
			Call(_UserBannedList)
			$Dirty = False
			For $Check = 1 To $BannedList[0]
				If StringInStr($arrneeded[$i][0], $BannedList[$Check]) Then $Dirty = True
			Next
			If $Dirty == False Then _GUICtrlListView_SetItemSelected($wulv, $i, True)
		Next
		$objsearcher = 0
		$arrneeded = 0
		_UpdatesDownloadAndInstall()

	Else
		GUICtrlSetData($wup, "0")
		GUICtrlSetData($label[7], "Search not started yet.")
		GUICtrlSetData($wul, "Your windows is up to date.")
		c("Your Windows is up to date.")
		Return 0
	EndIf
EndFunc   ;==>_PopulateNeeded



Func _UpdatesDownloadAndInstall()

	$selected = _GUICtrlListView_GetSelectedIndices($wulv, True)
	If $selected[0] = 0 Then
		c("Results: Your Windows seems up to date.")
		GUICtrlSetData($wup, "0")
		GUICtrlSetData($label[7], "Search not started yet.")
		Return 0
	EndIf
	$objsearcher = _CreateMsUpdateSession($Host)
	For $x = 1 To $selected[0]
		$item = _GUICtrlListView_GetItemText($wulv, $selected[$x])
		For $i = 0 To $ColNeeded.updates.count - 1
			$update = $ColNeeded.updates.item($i)
			If $item = $update.title Then
				GUICtrlSetData($wul, "Downloaded : " & $x & " / Total : " & $selected[0] & " updates")
				c("Downloaded: " & $x & "  / Total: " & $selected[0] & " updates")
				Global $calculate = $x / $selected[0] * 100
				Global $rawpercent = Number($calculate, 1)
				Global $percent = $rawpercent & "%"
				GUICtrlSetData($wup, $rawpercent)
				GUICtrlSetData($label[7], $percent)
				_GUICtrlListView_SetItemText($wulv, $i, "Downloading...", 1)
				_GUICtrlListView_SetItemFocused($wulv, $i)
				_GUICtrlListView_EnsureVisible($wulv, $i)
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
	GUICtrlSetData($wup, "0")
	GUICtrlSetData($label[7], "0")
	c("All updates are downloaded, proceeding to install updates...")
	For $x = 1 To $selected[0]
		$item = _GUICtrlListView_GetItemText($wulv, $selected[$x])
		For $i = 0 To $ColNeeded.updates.count - 1
			$update = $ColNeeded.updates.item($i)
			If $item = $update.title And $update.isdownloaded Then
				GUICtrlSetData($wul, "Installed :" & $x & "  / Total :" & $selected[0] & " updates")
				c("Installed: " & $x & "  / Total: " & $selected[0] & " updates")
				$calculate = $x / $selected[0] * 100
				$rawpercent = Number($calculate, 1)
				$percent = $rawpercent & "%"
				GUICtrlSetData($wup, $rawpercent)
				GUICtrlSetData($label[7], $percent)
				_GUICtrlListView_SetItemText($wulv, $i, "Installing...", 1)
				_GUICtrlListView_SetItemFocused($wulv, $i)
				_GUICtrlListView_EnsureVisible($wulv, $i)
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
		c("A reboot is required to finish installing updates.")
		GUICtrlSetData($wup, "0")
		GUICtrlSetData($label[7], "0%")
;~ 		Shutdown(2 + 4 + 16)
	Else
		c("No reboot required, establishing a new Windows Updates Session...")
		_GUICtrlListView_DeleteAllItems($wulv)
		GUICtrlSetData($wup, "0")
		GUICtrlSetData($label[7], "Search not started yet.")
		_PopulateNeeded($Host)
	EndIf
	c("Your Windows seems up to date.")
	GUICtrlSetData($wup, "0")
	GUICtrlSetData($label[7], "Search not started yet.")
	$downloadsession = 0
	$updatestodownload = 0
	Return 0

EndFunc   ;==>_UpdatesDownloadAndInstall

#EndRegion WINDOWS UPDATE


#Region GATHER SYSTEM INFORMATION >>>>>>>>>>>>>>>
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
		$sInfos &= " Graphic card : " & $Obj_Item.Name & @CRLF & @CRLF
	Next

	; HDD
	Dim $Obj_Services = $Obj_WMIService.ExecQuery("Select * from Win32_DiskDrive")
	Local $Obj_Item
	For $Obj_Item In $Obj_Services
		If $Obj_Item.MediaType = "Fixed hard disk media" Then
			$sInfos &= " HDD " & $Obj_Item.Index & " : " & $Obj_Item.Model & " - " & Round($Obj_Item.Size / 1000000000, 0) & " Go - Status " & $Obj_Item.Status & @CRLF
			If $Obj_Item.Status <> "OK" Then
				$sReturn = "HDD SMART Results " & $Obj_Item.Index & "  : " & $Obj_Item.Status
			EndIf
		EndIf
	Next

	GUICtrlSetData($iEdit, $sInfos)

	Return $sReturn
EndFunc   ;==>_GetSystemInfo
#EndRegion GATHER SYSTEM INFORMATION >>>>>>>>>>>>>>>

#Region GET HDD SMART >>>>>>>>>
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

		Local $aSmartFound = [5, 9, 12, 194, 197, 198]
		Local $sMax
		Local $sPos
		Local $sMax

		If IsArray($Obj_Item.VendorSpecific) Then
			$bReturn = True
			$aHDDname = StringRegExp($Obj_Item.InstanceName, 'Ven_(.*)&Prod_(.*)\\', 3)
			If (IsArray($aHDDname) And UBound($aHDDname) = 2) Then
				$sHDDname = $aHDDname[0] & " " & $aHDDname[1]
			Else
				$aHDDname = StringRegExp($Obj_Item.InstanceName, 'Disk([A-Za-z0-9]*)', 3)
				If (IsArray($aHDDname)) Then
					$sHDDname = $aHDDname[0]
				Else
					$sHDDname = "Undefined"
				EndIf
			EndIf

			$aSmart = $Obj_Item.VendorSpecific

			$sMax = UBound($aSmart) - 1
			For $i = 2 To $sMax Step 12
				If _ArrayBinarySearch($aSmartFound, $aSmart[$i]) <> -1 Then
					_ArrayDelete($aSmartFound, "0;" & $aSmart[$i])
					If ($aSmart[$i] = 9 Or $aSmart[$i] = 12) Then
						; calcul POH
						$aSmartAttribute[$aSmart[$i]] = $aSmart[$i + 6] * 256 + $aSmart[$i + 5]
					Else
						$aSmartAttribute[$aSmart[$i]] = $aSmart[$i + 5]
					EndIf

				EndIf
			Next
		EndIf

		$aResults[$sHDDname] = $aSmartAttribute
	Next

	Return $bReturn

EndFunc   ;==>_GetSmart
#EndRegion GET HDD SMART >>>>>>>>>


Func _Exit()

	Exit

EndFunc   ;==>_Exit


