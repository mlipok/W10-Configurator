#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icon\gear.ico
#AutoIt3Wrapper_Outfile=W10-Configurator.exe
#AutoIt3Wrapper_Res_Fileversion=0.7.3.9
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=W10-Configurator
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_Field=Made By|Williamas Kumeliukas
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****



;#################################################
;#    W10-Configurator by WilliamasKumeliukas    #
;#################################################

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
#include <GuiEdit.au3>
#include <GuiTab.au3>
#include <GuiListView.au3>
#include <String.au3>
#include <StringConstants.au3>
#include <WinAPIShellEx.au3>
#include <Misc.au3>
#include ".\Ressources\UDF\ExtMsgBox.au3"
#include <APIDiagConstants.au3>
#include ".\Ressources\UDF\MS-Store.au3"
#include ".\Ressources\UDF\choco.au3"
Opt("GUIResizeMode", $GUI_DOCKAUTO)

#EndRegion Includes <<<<<<<<


#Region DECLARE VARIABLES FOR LATER USE
;EXPECT POSSIBILITIES OF {UNUSED} | {DUPLICATES} | {MISSING} VARIABLES

Global $ColCount
Global $ConfigDir = @ScriptDir & "\Ressources"
Global $hGUI	; INVISIBLE GUI FOR AutoComplete LISTBOX
Global $hList 	;AUTOCOMPLETE LISTBOX
Global $hInput 	;INPUT USED TO STORE SELECTED SOFTWARES IN AUTOCOMPLETE LISTBOX
Global $aSelected
Global $sChosen
Global $hUP 	;USED FOR NAVIGATING WITH ARROWS KEY
Global $hDOWN	;USED FOR NAVIGATING WITH ARROWS KEY
Global $hENTER  ;USED FOR SELECTING ITEMS IN AUTOCOMPLETE LISTBOX
Global $hESC	;USED FOR CLOSING AUTOCOMPLETE LISTBOX
Global $sCurrInput = ""
Global $aCurrSelected[2] = [-1, -1]
Global $iCurrIndex = -1
Global $hListGUI = -1
Global $nas ;UNUSED FOR NOW
Global $cw  ;WARNING message (ORANGE) in GUI Console
Global $cwe ;  ERROR message (RED) in GUI Console
Global $c 	;	INFO message (WHITE) in GUI Console
Global $e = False	; Error
Global $cmdfile
Global $sRemoteName
Global $regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
Global $iniFile = $ConfigDir & "\default.ini"
Global $iRead ;RESULT FROM INI READ
Global $isAdmin
Global $iniOpen =  $ConfigDir & "\default.ini"
Global $console ; GUI Console at bottom
Global $tree
Global $info
Global $sInfos
Global $task[11] 		;@ComputerName InputBox
Global $task[12]		;OEM
Global $task[14]		;OEM
Global $task[15]		;OEM
Global $task[12]		;OEM
Global $task[13]		;OEM
Global $task[16]		;OEM
Global $taskCount		;
Global $oemfile			;OEM LOGO FILE PATH
Global $cLogo			;OEM Logo Button
Global $wulv			;WINDOWS UPDATE LISTVIEW
Global $wul				;WINDOWS UPDATE LABEL
Global $wup				;WINDOWS UPDATE PROGRESSBAR
Global $GUI				;MAIN GUI
Global $Go		;RUN BUTTON
Global $sServer												;For $nas
Global $sShare												;For $nas
Global $letter												;For $nas
Global $sLocalName = $letter & ":" 							;For $nas
Global $sServerShare = $sServer & $sShare 					;For $nas
Global $sUserName = "CHANGEME", $sPassword = "CHANGEME" 	;For $nas
Global $sResult
Global $aResults
Global $sSaverActive 	;Retrieve state of ScreenSaver
Global $syspin = "syspin.exe"
Global $ColNeeded
Global $Host = @ComputerName
Global $sBannedList
Global $BannedList
Global $label[99]
Global $task[99]
Global $sMsg
Global $sRunTasks
Global $RZUpdated = False
Global $sRZVersion 		;Local RZGet Version
Global $sORZVersion 	;Online RZGet Version
Global $sRZCatalog 		;RZ Catalog
Global $sRZGet = FileGetShortName("'" & @ScriptDir & "\Ressources\RZget.exe'")

$task[0] = "99"		; WILL BE CHANGED TO REAL NUMBER OF TASKS BEFORE GUI SHOW
$label[0] = "99"
#EndRegion DECLARE VARIABLES  FOR LATER USE


GUI()

#Region  GUI <<<<<<<< @@@@@@@@@@@@@@@@@@@@

Func GUI()

	#Region GUI setup for Tasks

	Global $GUI = GUICreate(@ScriptName & " - Version: " & FileGetVersion(@ScriptFullPath), 1043, 820) ;740,632)
	GUISetIcon("Ressources\Icon\gear.ico", $GUI)
	$tab = GUICtrlCreateTab(0, 0, 697, 541, -1, -1)
	GUICtrlSetState($tab, $GUI_ONTOP) ;2048 = magic number

	GUICtrlCreateTabItem("System Info")
	GUICtrlCreateTabItem("Configuration")
	GUICtrlCreateTabItem("Windows Updater")
	GUICtrlCreateTabItem("")

	_GUICtrlTab_SetCurFocus($tab, -1)
	GUICtrlSetResizing(-1, 70)

	;Menu items
	local $FileMenu = GUICtrlCreateMenu("File")
	local $FileImp = GUICtrlCreateMenuItem("Import / Load", $FileMenu)
	GUICtrlSetState(-1, $GUI_DEFBUTTON)
	local $FileExp = GUICtrlCreateMenuItem("Export / Save",$FileMenu)
	GUICtrlCreateMenuItem("", $Filemenu, 2) ; create a separator line
	local $FileEx = GUICtrlCreateMenuItem("Exit",$FileMenu)
	local $HelpMenu = GUICtrlCreateMenu("Help")
	local $HelpInfo = GUICtrlCreateMenuItem("Configurator Help",$HelpMenu)
	local $HelpRep = GUICtrlCreateMenuItem("Report Issue",$HelpMenu)
	local $HelpGit = GUICtrlCreateMenuItem("Project website",$HelpMenu)
	local $HelpAbout = GUICtrlCreateMenuItem("About",$HelpMenu)


	GUICtrlCreateGraphic(705, 45, 300, 455, BitOR($GUI_SS_DEFAULT_GRAPHIC, $SS_WHITEFRAME, $WS_BORDER))
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetGraphic(-1, $GUI_GR_COLOR, 0x000000, 0xFFFFFF)
	$task[1] = GUICtrlCreateCheckbox("Windows Updater", 710, 48, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)
	$task[2] = GUICtrlCreateCheckBox("Install Softwares", 710, 70, 260, 24, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$hInput = GUICtrlCreateInput("", 710, 440, 280, 30)
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)
	GUICtrlSetState(-1, $GUI_DISABLE)

	$task[3] = GUICtrlCreateCheckbox("Google Chrome = Default Browser", 710, 95, 290, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[4] = GUICtrlCreateCheckbox(".pdf/pdfxml = ReaderDC", 710, 119, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[5] = GUICtrlCreateCheckbox("Apply OEM infos", 710, 142, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[6] = GUICtrlCreateCheckbox("Disable Hibernation", 710, 168, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[7] = GUICtrlCreateCheckbox("Install Office 365/2013", 710, 190, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[8] = GUICtrlCreateCheckbox("Change Computer Name", 710, 212, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[9] = GUICtrlCreateCheckbox("Add Office icons to taskbar", 710, 234, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$task[10] = GUICtrlCreateCheckbox("Select All", 710, 256, 260, 20, BitOR($TVS_DISABLEDRAGDROP, $TVS_CHECKBOXES))
	GUICtrlSetFont(-1, 13, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)

	$sRunTasks = GUICtrlCreateButton("Run", 910, 510, 80, 30)

	$console = _GUICtrlRichEdit_Create($GUI, "", 0, 549, 1040, 248, BitOR($ES_MULTILINE, $WS_VSCROLL, $ES_AUTOVSCROLL)) ;9, 364, 719, 258
	_GUICtrlRichEdit_SetEventMask($console, $ENM_LINK)
	_GUICtrlRichEdit_AutoDetectURL($console, True)
	_GUICtrlRichEdit_SetCharColor($console, 0xFFFFFF)
	_GUICtrlRichEdit_SetBkColor($console, 0x000000)
	_GUICtrlRichEdit_SetReadOnly($console, True)
	_GUICtrlRichEdit_HideSelection($console, True)
	_GUICtrlRichEdit_SetFont($console, 12, "Consolas")

	$label[1] = GUICtrlCreateLabel("Select tasks to do:", 780, 23, 160, 18, $SS_CENTER, -1)
	GUICtrlSetFont(-1, 12, 800, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	ConsoleWriteGUI($console, "=============== Win10 CONFIGURATOR CONSOLE ===============")

;~ 	Config Tab

	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 1) & GUICtrlRead($tab, 1))

GUICtrlCreateGroup("OEM",6,35,380,230)
	$label[2] = GUICtrlCreateLabel("Manufacturer:", 31, 52, 95)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[1], $GUI_DOCKAUTO)
	$task[12] = GUICtrlCreateInput("", 125, 50, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[3] = GUICtrlCreateLabel("Model:",71, 90, 95)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[2], $GUI_DOCKAUTO)
	$task[13] = GUICtrlCreateInput("", 125, 88, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[4] = GUICtrlCreateLabel("Support Hours:", 24, 128, 95)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$task[14] = GUICtrlCreateInput("", 125, 126, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[5] = GUICtrlCreateLabel("Support Website:", 10, 166, 110)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$task[15] = GUICtrlCreateInput("https://", 125, 164, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[7] = GUICtrlCreateLabel("Support Phone:", 21, 198, 110)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$task[16] = GUICtrlCreateInput("", 125, 196, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[6] = GUICtrlCreateLabel("Logo:", 77, 234, 101)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$task[17] = GUICtrlCreateInput("", 125, 232, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$oemfile = GUICtrlCreateButton("...", 350, 232, 30, 20)
	GUICtrlSetFont(-1, 10, 400, 0, "MS sans serif", 2)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	GUICtrlCreateGroup("Misc",6,265,380,200)
	$label[7] = GUICtrlCreateLabel("Computer Name:", 12, 282, 110, 15)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
	$task[11] = GUICtrlCreateInput($Host, 125, 280, 222, 20, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)


;~ 	System Info tab

	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 0) & GUICtrlRead($tab, 1))

	Global $iEdit = GUICtrlCreateEdit("", 0, 21, 694, 518, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $ES_MULTILINE, $ES_UPPERCASE, $WS_VSCROLL, $WS_HSCROLL), $WS_EX_TOOLWINDOW)

	_GetSystemInfo()


;~ 	Windows Updater tab

	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 2) & GUICtrlRead($tab, 1))

	$wulv = GUICtrlCreateListView("Description|Status", 17, 192, 647, 334, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	_GUICtrlListView_SetColumnWidth($wulv, 0, 500)
	GUICtrlSetBkColor(-1, $color_silver)
	GUICtrlSetColor(-1, 0x75EE3B)
	_GUICtrlListView_SetColumnWidth($wulv, 1, 140)
	GUICtrlSetBkColor(-1, $color_black)
	GUICtrlSetColor(-1, 0x75EE3B)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$label[7] = GUICtrlCreateLabel("Search not started yet", 16, 170, 250, 20, $SS_LEFTNOWORDWRAP)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing($label[7], $GUI_DOCKAUTO)
	$wup = GUICtrlCreateProgress(16, 142, 649, 20, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$task[18] = GUICtrlCreateInput("insert in this box KB or Keyword to be excluded from being applied: P.S :(large updates might fail to install)", 16, 90, 300, 40) ;User defined additions to banned list
	GUICtrlCreateLabel("Use the below box to exclude updates", 16, 55, 310, 40)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	$wus = GUICtrlCreateButton("Search", 500, 80, 100, 30, -1, -1)
	GUICtrlSetFont(-1, 10, 400, 0, "Lucida Bright", 5)
	GUICtrlSetResizing(-1, $GUI_DOCKAUTO)

	GUISwitch($GUI, _GUICtrlTab_SetCurFocus($tab, 1) & GUICtrlRead($tab, 1))
	_GUICtrlTab_SetCurFocus($tab, 0)

	For $i = UBound($task)-1 To 0 Step -1
		If $task[$i] = "" Then
			_ArrayDelete($task, $i)
		EndIf
	Next
	For $i = UBound($label)-1 To 0 Step -1
		If $label[$i] = "" Then
			_ArrayDelete($label, $i)
		EndIf
	Next
	$task[0] = UBound($task)
	$label[0] = UBound($label)

	GUISetState(@SW_SHOW, $GUI)
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
				;MsgBox(8256, "About",
				$sMsg =  "Windows-10 Configurator started as a project aiming to automate"
				$sMsg &= " some repetitive tasks or routine task on freshly installed windows being used by Network Administrators, Helpdesk personnel, etc."
				$sMsg &= @CRLF & @CRLF & "Founder: WilliamasKumeliukas" & @CRLF &  @CRLF & "***This project wouldn't have been gone this far without the GENEROUS help of these contributors*** :"
				$sMsg &= @CRLF & @CRLF &"lord_stewie" & @CRLF
				_ExtMsgBoxSet(4+ 32, 0, Default, Default, 10, "Lucida Bright", 700)
				$aRet = _StringSize($sMsg, Default, Default, 1)
				$retValue = _ExtMsgBox(32, "OK", "About", $sMsg)

			Case $HelpGit
				ShellExecute("https://github.com/Jivaross/W10-Configurator")

			Case $HelpInfo
				;MsgBox(8224,
				$sMsg  = "Feel Free to DM me any issues regarding W10-Configurator on Discord! " & @CRLF & @CRLF
				$sMsg &= "Discord Username: WilliamasKumeliukas#7342"
				_ExtMsgBoxSet(4 + 32, 0, Default, Default, 10, "Lucida Bright")
				$aRet = _StringSize($sMsg, Default, Default, 1)
				$retValue = _ExtMsgBox(32, "OK", "Configurator Help", $sMsg)

			Case $HelpRep
				ShellExecute("https://github.com/Jivaross/W10-Configurator/issues")

			Case $task[10]
				If GUICtrlRead($task[10]) == $GUI_CHECKED Then
					For $i = 10 To 1 Step -1
						GUICtrlSetState($task[$i], $GUI_CHECKED)
					Next


				ElseIf GUICtrlRead($task[10]) == $GUI_UNCHECKED Then
					For $i = 10 To 1 Step -1
						GUICtrlSetState($task[$i], $GUI_UNCHECKED)
					Next

				EndIf

			Case $task[2]

				If GUICtrlRead($task[2]) == $GUI_Checked Then

					If GUICtrlGetState($hInput) = "144" Then GUICtrlSetState($hInput, $GUI_ENABLE)

					If $sRZCatalog = BitAND("", 0 , -1) Then
						_RZCatalog()
					Else
						AdlibRegister("AutoComplete")
					EndIf

				Else
					If GUICtrlGetState($hInput) = "80" Then

						AdlibUnRegister("AutoComplete")
						GUICtrlSetData($hInput, "") ; input emptied
						GUICtrlSetState($hInput, $GUI_DISABLE)
					EndIf
				EndIf


			Case $sRunTasks
				Local $tasks=0
				For $i = 1 to 10 Step 1
					If GUICtrlRead($task[$i]) == $GUI_CHECKED Then $tasks +=1
					Next
					If $tasks < 1 Then
						cw("Choose tasks you want to execute before starting configuration!")
					Else

				c("Configuration started !")

				If GUICtrlRead($task[1]) == $GUI_Checked Then
					c("Running Windows Updater...")
					_PopulateNeeded($Host)
				EndIf

				If GUICtrlRead($task[2]) == $GUI_Checked Then

					Rzget()
					GUICtrlSetState($task[2], $GUI_UNCHECKED)
					GUICtrlSetData($hInput, "") ; emptied
				EndIf

				If GUICtrlRead($task[3]) == $GUI_Checked Then
					c("Setting Chrome as default browser..")
					defaultBrowser()
				EndIf

				If GUICtrlRead($task[4]) == $GUI_Checked Then
					c("Setting .PDF and .PDFXML to Adobe Reader...")
					defaultPDF()
				EndIf

				If GUICtrlRead($task[5]) == $GUI_Checked Then
					c("Installing OEM + Logo...")
					selfoem() ;WORK
				EndIf

				If GUICtrlRead($task[6]) == $GUI_Checked Then
					c("Disabling ScreenSaver...")
					ScreenSaver(False) ; >>>>>(UNTESTED)<<<< False to disable / True to reset by default
				EndIf


				If GUICtrlRead($task[7]) == $GUI_Checked Then
					c("Installing MS Office 365...")
					Office2013()
				EndIf

				If GUICtrlRead($task[8]) == $GUI_Checked Then
					c("Changing computer name...")
					 _SetComputerName($task[11])
				EndIf

				If GUICtrlRead($task[9]) == $GUI_Checked Then
					c("Adding Office icons to taskbar...")

				EndIf
				c("All selected tasks are completed!")

			EndIf

			Case $wus
				_PopulateNeeded($Host)

			Case $oemfile
				$cLogo = FileOpenDialog(@ScriptName, @DesktopDir, "BMP files (*.bmp)", $FD_FILEMUSTEXIST)
				If Not $cLogo = "" Then GUICtrlSetData($task[16], $cLogo)

		EndSwitch
	WEnd

EndFunc   ;==>GUI
#EndRegion  GUI SCRIPT <<<<<<<< @@@@@@@@@@@@@@@@@@@@

#Region SOFTWARE INSTALLER
Func AutoComplete() ; Main Loop of Software Installation

	debug("Input focused? : " & _IsFocused($GUI, $hInput))
		$hUP = GUICtrlCreateDummy()
		$hDOWN = GUICtrlCreateDummy()
		$hENTER = GUICtrlCreateDummy()
		$hESC = GUICtrlCreateDummy()
		Dim $AccelKeys[4][2] = [["{UP}", $hUP], ["{DOWN}", $hDOWN], ["{ENTER}", $hENTER], ["{ESC}", $hESC]]
		GUISetAccelerators($AccelKeys)

	While _IsFocused($GUI, $hInput) = True

		Do

			Switch GUIGetMsg()

				Case $GUI_EVENT_CLOSE
					Exit

				Case $hESC											;Action when User pressed: ESC Key
					If $hListGUI <> -1 Then ; List is visible.
						GUIDelete($hListGUI)
						$hListGUI = -1

					Else

						$mBox=MsgBox( 4, @ScriptName, "Do you want to quit application?")
						Select
							Case $mBox = 6
								Exit
							Case $mBox = 7
								ExitLoop
						EndSelect

					EndIf

				Case $hUP											;Action when User pressed: UP Arrow Key
					If $hListGUI <> -1 Then ; List is visible.
						$iCurrIndex -= 1
						If $iCurrIndex < 0 Then
							$iCurrIndex = 0
						EndIf
						_GUICtrlListBox_SetCurSel($hList, $iCurrIndex)
					EndIf

				Case $hDOWN											;Action when User pressed: DOWN Arrow Key
					If $hListGUI <> -1 Then ; List is visible.
						$iCurrIndex += 1
						If $iCurrIndex > _GUICtrlListBox_GetCount($hList) - 1 Then
							$iCurrIndex = _GUICtrlListBox_GetCount($hList) - 1
						EndIf
						_GUICtrlListBox_SetCurSel($hList, $iCurrIndex)
					EndIf

				Case $hENTER 										;Action when User pressed: ENTER Key

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
				$hList = _PopupSelector($GUI, $hListGUI, _CheckInputText($sCurrInput, $aCurrSelected)) ; ByRef $hListGUI, $aCurrSelected.
			EndIf
			If $sChosen <> "" Then
				GUICtrlSendMsg($hInput, 0x00B1, $aCurrSelected[0], $aCurrSelected[1]) ; $EM_SETSEL.
				_InsertText($hInput, "'" & $sChosen & "' ")
				$sCurrInput = GUICtrlRead($hInput)
				GUIDelete($hListGUI)
				$hListGUI = -1
				$sChosen = ""
			EndIf
			If _IsFocused($GUI,$sRunTasks) = True Then ExitLoop

		Until _IsFocused($GUI,$hInput) = False And _IsFocused($hListGUI, $hList) = False

		Select

			Case _IsFocused($GUI, $hInput) = False And _IsFocused($hListGUI, $hList) = False
				If $hListGUI <> -1 Then
					GUIDelete($hListGUI)
					$hListGUI = -1
					ExitLoop
				EndIf
		EndSelect


	WEnd

EndFunc



Func _IsFocused($h_Wnd, $i_ControlID) ; Check if control has focus.
    Return ControlGetHandle($h_Wnd, '', $i_ControlID) = ControlGetHandle($h_Wnd, '', ControlGetFocus($h_Wnd))
EndFunc

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

Func _PopupSelector($GUI, ByRef $hListGUI, $sCurr_List) ; FUNCTION THAT CREATE SUGGESTIONBOX WITH SOFTWARES AVAILABLE

	Local $hList = -1
	If $sCurr_List = "" Then
		Return $hList
	EndIf
	$hListGUI = GUICreate("", 710, 480, 710, 475, $WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_MDICHILD), $GUI) ; 280, 360
	$hList = GUICtrlCreateList("", 0, 0, 280, 350, BitOR($LBS_DISABLENOSCROLL, $LBS_SORT))
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
#EndRegion SOFTWARE INSTALLER

#Region Console
Func ConsoleSetCharColorNoSelection($hWnd, $iColor = Default)
If Not IsHWnd($hWnd) Then Return SetError(101, 0, False)

Local $tCharFormat = DllStructCreate($tagCHARFORMAT)
DllStructSetData($tCharFormat, 1, DllStructGetSize($tCharFormat))
If IsKeyword($iColor) Then
DllStructSetData($tCharFormat, 3, $CFE_AUTOCOLOR)
$iColor = 0
Else
If BitAND($iColor, 0xff000000) Then Return SetError(1022, 0, False)
EndIf

DllStructSetData($tCharFormat, 2, $CFM_COLOR)
DllStructSetData($tCharFormat, 6, $iColor)

Return _SendMessage($hWnd, $EM_SETCHARFORMAT, $SCF_SELECTION, DllStructGetPtr($tCharFormat)) <> 0

EndFunc

Func debug($cw)
ConsoleWrite(_NowTime() & " =-> " & $cw & @CRLF)
EndFunc


Func c($cw) ;INFORMATIVE MESSAGES IN CONSOLE
	$c = 0
	$e = False
	ConsoleWrite($cw & @CRLF)
	ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> " & $cw)

EndFunc   ;==>c

Func cw($cw, $e = False) ; @@@@ CONSOLE WARNINGS / ERRORS MESSAGE @@@@@
	$c = 1
	If $e = False Then
		ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> WARNING: " & $cw)
	ElseIf $e = True Then
		ConsoleWriteGUI($console, @CRLF & _NowTime() & " =-> ERROR: " & $cw)
EndIf

EndFunc   ;==>cw

Func ConsoleWriteGUI(Const ByRef $hConsole, Const $sTxt) ; READ C($cw) & cw($cw) ($cwe soon)
	;_GUICtrlRichEdit_GotoCharPos($console, -1)
	If $c = 1 or $e = True Then
		_GUICtrlRichEdit_SetFont($console, 12, "Consolas")
		ConsoleSetCharColorNoSelection($console, 0x0000FF)
		_GUICtrlEdit_AppendText($console, $sTxt)
		_GUICtrlRichEdit_ScrollLines($console, 1)

	Else
		_GUICtrlRichEdit_SetFont($console, 12, "Consolas")
		ConsoleSetCharColorNoSelection($console, 0xFFFFFF)
		_GUICtrlEdit_AppendText($console, $sTxt)
		_GUICtrlRichEdit_ScrollLines($console, 1)
	EndIf
	;_GUICtrlRichEdit_GotoCharPos($console, -1)
;	_GUICtrlRichEdit_HideSelection($console, True)

EndFunc   ;==>ConsoleWriteGUI
#EndRegion Console


#Region INI READ/WRITE
Func ini($iniFile, $section, $key, $value, $iniread = 1) ;Read or edit value in default.ini

	$iniOpen = FileOpen($iniFile, 2)
	c("read state = " & $iniread)

	Switch $iniread

		Case 0 ;0 = Write / Overwrite

			Return IniWrite($iniFile, $section, $key, $value)

		Case 1 ;1 = Read

			Return IniRead($iniFile, $section, $key, $value)

		Case 2 ;2 = Write Section

			Return IniWriteSection($iniFile, $section, $value)
	EndSwitch

EndFunc   ;==>ini

Func _ImportIni()
	;c("Importing settings...")
	$iniFile = FileOpenDialog("Choose File...", $ConfigDir, "Config files (*.ini)", 1, "default.ini")
	;Set values of check mark boxes, text boxes, etc. based on sections read within ini file array using inireadsection (creates 2d array) or iniread (reads single key value)

	For $i = 1 to $task[0]-1 Step 1
		$in = _GetCtrlClass($task[$i])
		Select
			Case $in="Checkbox"
				If GUICtrlRead($task[$i])==$GUI_CHECKED Then
					c("task(" & $i & ") = " & "1")
				ElseIf GUICtrlRead($task[$i])==$GUI_UNCHECKED then
						c("task(" & $i & ") = " & GUICtrlRead($task[$i],1))
				EndIf

			Case $in="Input"
				c("task(" & $i & ") = " & GUICtrlRead($task[$i]))
		EndSelect
	Next

EndFunc


Func _ExportIni()
	;c("Exporting current configuration settings...")
	$exportIni = FileSaveDialog("Select destination folder and desired File name...", @ScriptDir, "Configuration File (*.ini)", $FD_PROMPTOVERWRITE)
	If @error Then
		cw("Dialog was cancelled.")
	Else
		FileOpen($exportIni, "")

		For $i = 1 to $task[0]-1 Step 1
			$in = _GetCtrlClass($task[$i])
			If $in = "checkbox" Then

			Else
				msgbox(0,"","it's something else...")
			EndIf
		Next
	EndIf
EndFunc

#EndRegion INI READ/WRITE

#Region Get type of control

Func _GetCtrlClass($hHandle)
    Local Const $GWL_STYLE = -16
    Local $iLong, $sClass
    If IsHWnd($hHandle) = 0 Then
        $hHandle = GUICtrlGetHandle($hHandle)
        If IsHWnd($hHandle) = 0 Then
            Return SetError(1, 0, "Unknown")
        EndIf
    EndIf

    $sClass = _WinAPI_GetClassName($hHandle)
    If @error Then
        Return "Unknown"
    EndIf

    $iLong = _WinAPI_GetWindowLong($hHandle, $GWL_STYLE)
    If @error Then
        Return SetError(2, 0, 0)
    EndIf

    Switch $sClass
        Case "Button"
            Select
                Case BitAND($iLong, $BS_GROUPBOX) = $BS_GROUPBOX
                    Return "Group"
                Case BitAND($iLong, $BS_CHECKBOX) = $BS_CHECKBOX
                    Return "Checkbox"
                Case BitAND($iLong, $BS_AUTOCHECKBOX) = $BS_AUTOCHECKBOX
                    Return "Checkbox"
                Case BitAND($iLong, $BS_RADIOBUTTON) = $BS_RADIOBUTTON
                    Return "Radio"
                Case BitAND($iLong, $BS_AUTORADIOBUTTON) = $BS_AUTORADIOBUTTON
                    Return "Radio"
            EndSelect

        Case "Edit"
            Select
                Case BitAND($iLong, $ES_WANTRETURN) = $ES_WANTRETURN
                    Return "Edit"
                Case Else
                    Return "Input"
            EndSelect

        Case "Static"
            Select
                Case BitAND($iLong, $SS_BITMAP) = $SS_BITMAP
                    Return "Pic"
                Case BitAND($iLong, $SS_ICON) = $SS_ICON
                    Return "Icon"
                Case BitAND($iLong, $SS_LEFT) = $SS_LEFT
                    If BitAND($iLong, $SS_NOTIFY) = $SS_NOTIFY Then
                        Return "Label"
                    EndIf
                    Return "Graphic"
            EndSelect

        Case "ComboBox"
            Return "Combo"
        Case "ListBox"
            Return "List"
        Case "msctls_progress32"
            Return "Progress"
        Case "msctls_trackbar32"
            Return "Slider"
        Case "SysDateTimePick32"
            Return "Date"
        Case "SysListView32"
            Return "ListView"
        Case "SysMonthCal32"
            Return "MonthCal"
        Case "SysTabControl32"
            Return "Tab"
        Case "SysTreeView32"
            Return "TreeView"

    EndSwitch
    Return $sClass
EndFunc

#EndRegion Get type of control


Func selfoem() ;**** AutoIt version of oem() ****
	c("" & @CRLF & "OEM installation started !")
	c("(1/7)" & " | Installing OEM Logo...")
	_FileCopy($task[16], "C:\Windows\System32")
	RegWrite($regPath, "Manufacturer", "REG_SZ", $task[12])
	c("(2/7)" & " | Added Manufacturer registry key...")
	RegWrite($regPath, "Model", "REG_SZ", $task[13])
	c("(3/7)" & " | Added Model registry key...")
	RegWrite($regPath, "Logo", "REG_SZ", $task[17])
	c("(4/7)" & " | Added Logo (OEM Logo references) registry key...")
	RegWrite($regPath, "SupportHours", "REG_SZ", $task[14])
	c("(5/7)" & " | Added SupportHours registry key...")
	RegWrite($regPath, "SupportPhone", "REG_SZ", $task[16])
	c("(6/7)" & " | Added SupportPhone registry key...")
	RegWrite($regPath, "SupportURL", "REG_SZ",  $task[15])
	c("(7/7)" & " | Added SupportURL registry key...")
	c("OEM configuration done.")

EndFunc   ;==>selfoem


#Region RZGET <<<
Func Rzget()

	If FileExists(@ScriptDir & "\ressources\rzget.exe") Then ; IF RZGET IS IN CONFIG FOLDER @@@@
		;_CheckRZGetVersion()
		AdlibUnRegister("AutoComplete")

		$Powershell = Run('powershell -Command "& '& $sRZGet &' install '& GUICtrlRead($hInput), @SystemDir, @SW_HIDE, $STDERR_MERGED)
		While 1

        $sStreamOut = StdoutRead($Powershell)

        If @error Then
            ExitLoop
        Else
			$ReplaceStream = StringRegExpReplace($sStreamOut, '[\s]*$|([\n\r]){2,}', '$1')
			If StringLen($ReplaceStream) > 1 Then c($ReplaceStream)

        EndIf

        Sleep(100)

		WEnd

	c("Done!")

	Else ; RZGET.EXE WAS NOT FOUND IN CONFIG FOLDER @@@@
		c("W10-Configurator Ressources Folder is missing required dependencies")
	EndIf
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
		c("Updating catalog Version")
		InetGet("https://github.com/rzander/ruckzuck/releases/download/" & $sORZVersion & "/Rzget.exe", @ScriptDir & "\Ressources\Rzget.exe", 2, 0)
	Else
		c("Catalog is up to date.")
	EndIf
EndFunc   ;==>_CheckRZGetVersion

Func _RZCatalog()

If $RZUpdated = False Then

		c("Retrieving catalog, this may take few seconds...")
	Else

		c($sRZCatalog[0] & " softwares available.")
	EndIf

	Local $Powershell = Run("powershell -Command ((&" & $sRZGet & " search | convertfrom-json) | Select ShortName | Out-String -Stream | foreach {$_.trimend()} | Where {$_ -ne ''} | Where {$_ -ne '---------'} | Where {$_ -ne 'ShortName'})", @SystemDir, @SW_HIDE, $STDERR_MERGED)
	ProcessWaitClose($Powershell)
	$sRZCatalog = StringSplit(StdoutRead($Powershell), @CRLF, 1)

	If $sRZCatalog=Null Then c(StderrRead($Powershell))

	Local $ColCount = UBound($sRZCatalog, 1) - 1
	_ArrayDelete($sRZCatalog, $ColCount) ;Removes the blank line at the end

	If Not $sRZCatalog = Null Or IsArray($sRZCatalog) = True Then
		$RZUpdated = True
		c("Retrieve complete, " & $sRZCatalog[0] & " softwares has been found.")
		AdlibRegister("AutoComplete")
	Else

		c("Unable to retrieve catalog...")
	EndIf

EndFunc   ;==>_RZCatalog


#EndRegion RZGET <<<

#Region SCREENSAVER
Func ScreenSaver($Alive) ;False = reset to default & True to prevent/disable sleep/power-savings modes (AND screensaver)

	If $Alive = False Then
		Local $ssKeepAlive = DllCall('kernel32.dll', 'long', 'SetThreadExecutionState', 'long', 0x80000003)
		If @error Then
			Return SetError(2, @error, 0x80000000)
			cw(@error)
		EndIf
		Return $ssKeepAlive[0] ; Previous state (typically 0x80000000 [-2147483648])
		c("Hibernate mode = " & $sSaverActive)
	ElseIf $Alive = True Then
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

Func _FileCopy($sRemoteName, $d) ;UNUSED for now
	Local $FOF_RESPOND_YES = 16
	Local $FOF_SIMPLEPROGRESS = 256
	$oShell = ObjCreate("shell.application")
	$oShell.namespace($d).CopyHere($sRemoteName, $FOF_SIMPLEPROGRESS)
EndFunc   ;==>_FileCopy

#Region NAS SCRIPT <<<<<<<<<
Func nas()

	c("Verifying if there is an existing connection to: >>> " & $sServerShare & " <<< ...")
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
			cw("Connection Failed... Inform WilliamasKumeliukas of this problem please.")
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
	$defBrowser = Run('powershell -Command "& ' & $ConfigDir & '\SetDefaultBrowser.exe' & ' chrome',@SystemDir,@SW_HIDE, $STDERR_MERGED)

	While 1

        $sStreamOut = StdoutRead($defBrowser)

        If @error Then
            cw("Error occured at line: " & @ScriptLineNumber - 3, True)

        Else
			$ReplaceStream = StringRegExpReplace($sStreamOut, '[\s]*$|([\n\r]){2,}', '$1')
			If StringLen($ReplaceStream) > 1 Then c($ReplaceStream)

        EndIf

        Sleep(10)

		WEnd

	c("Done !")
	GUICtrlSetState($task[3], $GUI_UNCHECKED)
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
	ShellExecute("defaultPDF.bat", "", @TempDir, "open", @SW_HIDE)
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

	$Excel = '"%PROGRAMFILES(X86)%\Microsoft Office\Office15\EXCEL.EXE" 5386'
	$Word = '"%PROGRAMFILES(X86)%\Microsoft Office\Office15\WORD.EXE" 5386'
	$Outlook = '"%PROGRAMFILES(X86)%\Microsoft Office\Office15\OUTLOOK.EXE" 5386'
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


#Region 		 INSTALL OFFICE 2013
Func Office2013()  ;==> TO DO (Professional Plus version OR 365)
	MsgBox(16, @ScriptName, "If you see this message, it means this function is still under construction.")

	$officePath = "CHANGEME"

EndFunc   ;==>Office2013
#EndRegion 		 INSTALL OFFICE 2013

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
	c("Creating Update Searcher Session...")
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
	c("Searching for available updates...")
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

	GUICtrlSetData($wul, "Searching for available updates...")
	GUICtrlSetData($label[7], "Searching for available updates...")
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
		_GUICtrlListView_AddItem($wulv, "You are up to date.", 1)
		GUICtrlSetData($label[7], "You are up to date.")
		GUICtrlSetData($wul, "You are up to date.")
		c("You are up to date.")
		Return 0
	EndIf
EndFunc   ;==>_PopulateNeeded



Func _UpdatesDownloadAndInstall()

	$selected = _GUICtrlListView_GetSelectedIndices($wulv, True)
	If $selected[0] = 0 Then
		c("Results: You are up to date.")
		GUICtrlSetData($wup, "0")
		GUICtrlSetData($label[7], "You are up to date ")
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
				Global $calculate = $x-1 / $selected[0] * 100
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
				$calculate = $x-1 / $selected[0] * 100
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
	c("You are up to date.")
	_GUICtrlListView_AddItem($wulv, "You are up to date.", 1)
	GUICtrlSetData($wup, "0")
	GUICtrlSetData($label[7], "You are up to date.")
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
		$sInfos &= @CRLF & " OS : " & $Obj_Item.Caption & " " & @OSArch & " version " & $Obj_Item.Version & @CRLF
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