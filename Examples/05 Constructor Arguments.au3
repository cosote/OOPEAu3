#AutoIt3Wrapper_Run_AU3Check=n
#include '../OOPE/OOPE.au3'

#classdef <Mouse> $oTest from <MouseGetPos()>

$oTest.SetPos(50, 50) ; Move slowly
$oTest.SetSpeed(0)
$oTest.SetPos(200, 200) ; Move instantly

#Region Class Mouse
	Local $iX, $iY, $iSpeed

	Func _Mouse($aArgs)
		$This.iX     = $aArgs[0]
		$This.iY     = $aArgs[1]
		$This.iSpeed = 50
	EndFunc

	Func SetSpeed($iNewSpeed)
		$This.iSpeed = $iNewSpeed
	EndFunc

	Func SetPos($iNewX, $iNewY)
		$This.iX = $iNewX
		$This.iY = $iNewY
		MouseMove($This.iX, $This.iY, $This.iSpeed)
	EndFunc
#EndRegion