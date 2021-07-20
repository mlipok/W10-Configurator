#requireAdmin
#include "UIAWrappers.au3"
#include <Array.au3>
#include <File.au3>
#include <GUIConstants.au3>


Func _mstore()


Local $mStoreAppsPath = "C:\Program Files\WindowsApps\"
Local $msAppName = ""
Local $appPath
;~ _ArrayDisplay( _FileListToArray($mStoreAppsPath,"*" & $msAppName & "*" & "x64" & "*",2, True) )
Local $msTitle
Local $msControlType
Local $msTypeId
Local $msClass
Global $countdown = 10, $cCount



Local $msgTxt = "The following procedure could fail" & _
" to install application if you were to move your mouse or keyboard, "& _
"please do not touch anything until task is completed." & _
@LF & @LF & 'By pressing "Accept" button,' & " you're giving your consent to follow these conditions."


Global $mStoreGUI = GUICreate("", 615, 210, 592, 374, $WS_POPUP, $WS_EX_APPWINDOW)
Global $ButtonAgree = GUICtrlCreateButton("Accept", 192, 136, 75, 25)
GUICtrlSetState($ButtonAgree, $GUI_DISABLE)
Global $ButtonCancel = GUICtrlCreateButton("Cancel", 336, 136, 75, 25)
Global $iMsg = GUICtrlCreateLabel( $msgTxt, 25, 28, 548, 94, $SS_CENTER);BitOR($WS_BORDER, $SS_CENTER))
GUICtrlSetFont($iMsg, 11.9, 400, 0, "Lucida Console")

GUISetState(@SW_SHOW)

AdlibRegister("countdown", 1000)
While 1


	Local $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $ButtonAgree
			GUIDelete($mStoreGUI)
			ShellExecute("ms-windows-store:")

			Local $Window = WinWaitActive ("Microsoft Store")

			If WinExists ($Window) Then
				;
				_UIA_setVar("AppBarButton","Title:=Rechercher;controltype:=UIA_ButtonControlTypeId;class:=AppBarButton")
				_UIA_action("AppBarButton","highlight")
				_UIA_action("AppBarButton","leftclick")

				_UIA_setVar("TextBox","Title:=Rechercher;controltype:=UIA_EditControlTypeId;class:=TextBox")
				_UIA_action("TextBox","settextvalue" , "Lenovo Vantage")

				_UIA_setVar("ListViewItem3","ListItem:=Lenovo Vantage Application;controltype:=UIA_ListItemControlTypeId;class:=ListViewItem")
				_UIA_action("ListViewItem3", "leftclick")


				_UIA_setVar("GetButton", "Button:=Obtenir;controltype:=UIA_ButtonControlTypeId;automationid:=buttonPanel_AppIdentityBuyButton")
				_UIA_action("GetButton", "leftclick")

				_UIA_setVar("PlayButton", "Button:=Lancer;controltype:=UIA_ButtonControlTypeId;automationid:=PlayBar_AppIdentityOpenButton")
				_UIA_action("PlayButton", "leftclick")
				ExitLoop
			EndIf

		Case $ButtonCancel
			GUIDelete($mStoreGUI)
			ExitLoop
	EndSwitch
WEnd

EndFunc

Func countdown()

$cCount = $countdown-1
$countdown = $cCount
	If $cCount = 0 Then
		GUICtrlSetData($ButtonAgree, "Accept")
		GUICtrlSetState($ButtonAgree, $GUI_ENABLE)
		AdlibUnRegister("countdown")
	Else
		ConsoleWrite($cCount & @CRLF)
		GUICtrlSetData($ButtonAgree, $cCount)
	EndIf
EndFunc
