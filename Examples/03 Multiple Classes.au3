#AutoIt3Wrapper_Run_AU3Check=n
#include '../OOPE/OOPE.au3'

#classdef <Test> $oTest, <Second> $oSecond

$oTest.SetValue(2)
MsgBox(0, 0, $oTest.GetValue)

$oSecond.SetBase(2)
MsgBox(0, 0, $oSecond.GetSquare)

#Region Class Test
	Local $iValue

	Func SetValue($iNewValue)
		$This.iValue = $iNewValue
	EndFunc

	Func GetValue()
		Return $This.iValue
	EndFunc
#EndRegion

#Region Class Second
	Local $iValue

	Func SetBase($iNewValue)
		$This.iValue = $iNewValue
	EndFunc

	Func GetSquare()
		Return $This.iValue ^ 2
	EndFunc
#EndRegion