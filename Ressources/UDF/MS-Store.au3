#requireAdmin
#include "UIAWrappers.au3"
#include <Array.au3>
#include <File.au3>

;~ Local $lenovoPart1_example = "E046963F"
;~ Local $lenovoPart2_example = "_10.2105.16.0_x64__k1h2ywk1493x8"

Func _mstore()


Local $mStoreAppsPath = "C:\Program Files\WindowsApps\"
Local $msAppName = ".LenovoCompanion"
Local $appPath = _FileListToArray($mStoreAppsPath,"*" & $msAppName & "*" & "x64" & "*",2, True)
Local $msTitle
Local $msControlType
Local $msTypeId
Local $msClass
_ArrayDisplay($appPath)


Local $Msgtxt = "The following procedure could fail" & _
" to install application if you were to move your mouse or keyboard, "& _
"please do not touch anything until task is completed." & _
@CRLF & 'By pressing "Accept" button,' & " you're accepting to follow these conditions."

Local $mStoreGUI = GUICreate("", 615, 210, 592, 374, $WS_POPUP, $WS_EX_APPWINDOW)
Local $ButtonAgree = GUICtrlCreateButton("Accept", 192, 136, 75, 25)
Local $ButtonCancel = GUICtrlCreateButton("Cancel", 336, 136, 75, 25)
Local $iMsg = GUICtrlCreateLabel($Msgtxt, 65, 48, 448, 84)
GUICtrlSetFont(-1, 12, 400, 0, "Lucida Console")
GUISetState(@SW_SHOW)


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

