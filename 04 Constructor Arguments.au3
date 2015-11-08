; OOP UDF Example 4: Constructor Arguments

#AutoIt3Wrapper_Run_AU3Check=n
#include <Misc.au3>
#include 'OOP.au3'

$hGUI = GUICreate("OOP Example")
Local $hPlayers = [GUICtrlCreateLabel("", 25, 175, 50, 50), GUICtrlCreateLabel("", 325, 175, 50, 50)]
GUICtrlSetBkColor($hPlayers[0], 0xFF0000)
GUICtrlSetBkColor($hPlayers[1], 0xFFCC00)
GUISetState()

; The constructor is called with arguments:
Local $aParameters = [25, 175, $hPlayers[0]]
# <classdef:Player $oPlayer1 from $aParameters>
Local $aParameters = [325, 175, $hPlayers[1]]
# <classdef:Player $oPlayer2 from $aParameters>

While GUIGetMsg() <> -3
	$DiffX = _IsPressed("27") - _IsPressed("25")
	$DiffY = _IsPressed("28") - _IsPressed("26")
	$oPlayer1.Move($DiffX, $DiffY)
	$oPlayer2.Move($DiffX, $DiffY)
WEnd

#Region Class Player
	Local $iX, $iY, $iHandle

	Func _Player($aArgs)
		$This.iX      = $aArgs[0]
		$This.iY      = $aArgs[1]
		$This.iHandle = $aArgs[2]
	EndFunc

	Func Move($iStepX, $iStepY)
		$This.iX += $iStepX
		$This.iY += $iStepY
		GUICtrlSetPos($This.iHandle, $This.iX, $This.iY)
	EndFunc
#EndRegion



