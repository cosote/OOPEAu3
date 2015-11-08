; OOP UDF Example 2: Multiple Instances

#AutoIt3Wrapper_Run_AU3Check=n
#include 'OOP.au3'

# <classdef:Test $oTest>
# <classdef:Test $oSecond>

$oTest.SetFoo(21)
$oSecond.SetFoo(42)

; Empty parentheses are optional on objects:
If $oTest.iGetValue * 2 = $oSecond.iGetValue Then MsgBox(0, "OOP Example", "Works.")

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