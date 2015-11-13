#AutoIt3Wrapper_Run_AU3Check=n
#include <Misc.au3>
#include '../OOPE/OOPE.au3'

; Define a simple loop macro:
#define doevents($body) While GUIGetMsg() <> -3 \ For $n = 0 To UBound($aPlayers)-1 \ $body \ Next \ WEnd

Local $aPlayers[10]

$hGUI = GUICreate("OOP Example")
#classdef <Player> $aPlayers
GUISetState()

doevents($aPlayers[$n].Move(_IsPressed("27") - _IsPressed("25"), _IsPressed("28") - _IsPressed("26")))

#Region Class Player
	Local $iX, $iY, $iHandle

	Func _Player()
		$This.iX      = Random(0, 390, 1)
		$This.iY      = Random(0, 390, 1)
		$This.iHandle = GUICtrlCreateLabel("", $This.iX, $This.iY, 10, 10)
		GUICtrlSetBkColor($This.iHandle, Random(0xAAAAAA, 0xFFFFFF, 1))
	EndFunc

	Func Move($iStepX, $iStepY)
		$This.iX += $iStepX
		$This.iY += $iStepY
		GUICtrlSetPos($This.iHandle, $This.iX, $This.iY)
	EndFunc
#EndRegion