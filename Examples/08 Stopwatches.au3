; Example inspired by @Chimp's issues with https://www.autoitscript.com/forum/topic/161337-simple-stopwatch-udf/

#AutoIt3Wrapper_Run_AU3Check=n
#include <WindowsConstants.au3>
#include <GUIConstants.au3>
#include '../OOPE/OOPE.au3'

; How many watches to create (try to adjust this number):
Global Const $iWatches = 6


; Keep track of how many SW's are created:
Global $iStopWatchCounter = 0
Global $aWatches[$iWatches]
Global $iGuiWidth = $iWatches * 150 + ($iWatches-1) * 50 + 100

$hGUI = GUICreate("OOP Stopwatch Example", $iGuiWidth, 250, -1, -1, Default, 34078728)
GUISetBkColor(0x303030)
GUISetState()

#classdef <Stopwatch> $aWatches

GUICtrlCreateLabel($iStopWatchCounter & " stopwatches up and running!", 50, 180, $iGuiWidth - 100, 30, 1)
GUICtrlSetFont(-1, 12, 0, 0, "Segoe UI Semilight")
GUICtrlSetColor(-1, 0xF0F0F0)

While 1
	$nMsg = GUIGetMsg()
	For $n = 0 To $iWatches-1
		$aWatches[$n].HandleMessage($nMsg)
		$aWatches[$n].Update
	Next
WEnd

#Region Class Stopwatch
	Local $hLabel, $hStartStop, $hReset, $hTimer
	Local $iHalt, $fTotal
	Local $fSec, $iMin, $iHour

	Func _Stopwatch()
		; Autoposition this watch
		$iLeft = 50 + $iStopWatchCounter * 150 + $iStopWatchCounter * 50
		$This.hLabel  = GUICtrlCreateLabel("00:00:00", $iLeft, 50, 150, 40, 1)
		GUICtrlSetColor(-1, 0xF0F0F0)
		GUICtrlSetFont(-1, 25, 0, 0, "Segoe UI Light")
		$This.hStartStop = GUICtrlCreateLabel(ChrW(0x23F8), $iLeft, 95, 70, 30, 1)
		GUICtrlSetColor(-1, 0xF0F0F0)
		GUICtrlSetBkColor(-1, 0x808080)
		GUICtrlSetFont(-1, 14, 0, 0, "Segoe UI Symbol")
		GUICtrlSetCursor(-1, 0)
		$This.hReset = GUICtrlCreateLabel(ChrW(0x23EE), $iLeft + 75, 95, 70, 30, 1)
		GUICtrlSetColor(-1, 0xF0F0F0)
		GUICtrlSetBkColor(-1, 0x808080)
		GUICtrlSetFont(-1, 14, 0, 0, "Segoe UI Symbol")
		GUICtrlSetCursor(-1, 0)
		$This.hTimer = TimerInit()
		$iStopWatchCounter += 1
	EndFunc

	Func HandleMessage($iMsg)
		Switch $iMsg
			Case -3
				Exit
			Case $This.hStartStop
				GUICtrlSetData($This.hStartStop, $This.iHalt ? ChrW(0x23F8) : ChrW(0x23F5))
				$This.iHalt = $This.iHalt ? 0 : 1
			Case $This.hReset
				$This.fTotal = 0
		EndSwitch
	EndFunc

	Func Update()
		If Not $This.iHalt Then $This.fTotal += TimerDiff($This.hTimer)
		$fElapsed = $This.fTotal
		$This.iHour = Int($fElapsed / 1000 / 60 / 60)
		$fElapsed  -= $This.iHour * 60 * 60 * 1000
		$This.iMin  = Int($fElapsed / 1000 / 60)
		$fElapsed  -= $This.iMin * 60 * 1000
		$iNSec      = Round($fElapsed / 1000, 1)
		$This.fSec  = $iNSec
		GUICtrlSetData($This.hLabel, StringFormat("%02i:%02i:%04.1f", $This.iHour, $This.iMin, $This.fSec))
		$This.hTimer = TimerInit()
	EndFunc
#EndRegion



