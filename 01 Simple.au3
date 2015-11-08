; OOP UDF Example 1: A Simple Class

; AU3Check will NOT actually be disabled!
#AutoIt3Wrapper_Run_AU3Check=n
#include 'OOP.au3'

; Create a new instance of the class "Example" in $oTest
# <classdef:Test $oTest>

$oTest.SetFoo(42)
MsgBox(0, "OOP Example", "Value of iFoobar: " & $oTest.iGetValue())

#Region Class Test
	; Declare an integer(64) field
	Local $iFoobar

	; Declare methods
	Func SetFoo($iNewValue)
		$This.iFoobar = $iNewValue
	EndFunc

	Func iGetValue()
		Return $This.iFoobar
	EndFunc
#EndRegion