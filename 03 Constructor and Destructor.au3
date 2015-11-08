; OOP UDF Example 3: Constructor and Destructor

#AutoIt3Wrapper_Run_AU3Check=n
#include 'OOP.au3'

; The constructor is called when this statement is executed:
# <classdef:Test $oTest>

MsgBox(0, "OOP Example", "Some code here...")

#Region Class Test
	Local $iFoobar

	; Constructor
	Func _Test()
		MsgBox(0, "OOP Example", "Yay, I've been instantiated!")
	EndFunc

	; Desctructor
	Func __Test()
		MsgBox(0, "OOP Example", "I've been destroyed :/")
	EndFunc

	Func SetFoo($iNewValue)
		$This.iFoobar = $iNewValue
	EndFunc

	Func iGetValue()
		Return $This.iFoobar
	EndFunc
#EndRegion