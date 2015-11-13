#AutoIt3Wrapper_Run_AU3Check=n
#include '../OOPE/OOPE.au3'

; Create two instances (same as <Test> $oTest, <Test> $oSecond)
#classdef <Test> $oTest, $oSecond

$oTest.SetValue(42)
MsgBox(0, 0, "The answer to everything is " & $oTest.GetValue & ".")

$oSecond.SetValue(99)
MsgBox(0, 0, "I've got " & $oSecond.GetValue & " problems.")

#Region Class Test
	Local $iValue

	Func SetValue($iNewValue)
		$This.iValue = $iNewValue
	EndFunc

	Func GetValue()
		Return $This.iValue
	EndFunc
#EndRegion