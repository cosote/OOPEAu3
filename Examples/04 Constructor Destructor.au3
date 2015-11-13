#AutoIt3Wrapper_Run_AU3Check=n
#include '../OOPE/OOPE.au3'

#classdef <Test> $oTest

MsgBox(0, 0, "Default value: " & $oTest.GetValue)
$oTest.SetValue(42)
MsgBox(0, 0, "New value: " & $oTest.GetValue)

#Region Class Test
	Local $iValue

	; Set a default value when instantiated
	Func _Test()
		$This.iValue = 2
	EndFunc

	Func __Test()
		MsgBox(0, 0, "I've been destroyed!")
	EndFunc

	Func SetValue($iNewValue)
		$This.iValue = $iNewValue
	EndFunc

	Func GetValue()
		Return $This.iValue
	EndFunc
#EndRegion