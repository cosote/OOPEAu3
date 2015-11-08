#include <Array.au3>
#include <AutoItConstants.au3>
#include <WinAPICom.au3>
#include <Memory.au3>
#include-once

; +---------------------------------------------------------------------------------------------------+
; |    _         _               _             _                                                      |
; |   \ \      / /_ _ _ __ _ __ (_)_ __   __ _| |                                                     |
; |    \ \ /\ / / _` | '__| '_ \| | '_ \ / _` | |    DO NOT call any function from this include or    |
; |     \ '  ' / (_| | |  | | | | | | | | (_| |_|    violate it's namespace!                          |
; |      \_/\_/ \__,_|_|  |_| |_|_|_| |_|\__, (_)                                                     |
; |                                      |___/    		                                              |
; +---------------------------------------------------------------------------------------------------+
;
; [Object Oriented Programming UDF]
;
; This UDF implements classes in AutoIt code. It is not just another Preprocessor, this actually
; creates real objects using native AutoIt functions available since AutoIt v3.14.X.X (Stable).
;
; If you want to redistribute compiled scripts, take the generated *_stable.au3 file and compile it.
;
; [TODO]
;
; - Parse includes recursive
; - Implement string fields in classes
; - Make object arrays ReDim-able while keeping the type
;
; [CHANGELOG]
;
; [UTC 2015-11-08 9:37:00]: (minx): Create release version v0.2.

Global $__CR, $__CQI, $__CAR
Global $__ClsTbl[0][2]
Global Const $__s_ObjVars       = 'ptr Vtbl; int RefCnt; ptr paiCallback; ptr psVarsString; ptr psMethodsString; ptr psDestruktor; '
Global Const $__i_ObjVarsOffset = DllStructGetSize(DllStructCreate($__s_ObjVars))
Global Const $__s_ObjMethods    = 'ptr __QueryInterface; ptr __AddRef; ptr __Release; '
Global Const $__ThisFile        = 'OOP.au3'
Global Const $__Timer           = TimerInit()

; This  is a function  for the sake of keeping the global
; namespace clean, DO NOT CALL THIS FUNCTION AT ANY TIME!
_OOPExtendFromClassTemplates()

; by minx
Func _OOPExtendFromClassTemplates()
	Local $RegEnScript     = StringReplace(@ScriptFullPath, ".au3", "_stable.au3")
	Local $ScriptText      = FileRead(@ScriptFullPath)
	Local $VarHeader       = " = _OOPCreateClassInstance('"
	Local $sSourceCode     = $ScriptText
	Local $aIncludes       [Null]
	Local $stdRefVar       = '$___selfObjRef'

	If StringInStr($ScriptText, '# ' & '<classdef:') = 0 Or StringInStr(@ScriptFullPath, '_stable') <> 0 Then Return

	; Get all Class templates from code and all Includes
	$aIncList = StringRegExp($ScriptText, "(?s)\Q#include '\E(.*?)(?=\Q'\E)", 3)
	If UBound($aIncList) Then _ArrayAdd($aIncludes, $aIncList)
	$aIncList = StringRegExp($ScriptText, '(?s)\Q#include "\E(.*?)(?=\Q"\E)', 3)
	If UBound($aIncList) Then _ArrayAdd($aIncludes, $aIncList)
	If UBound($aIncludes) Then
		For $sIFile In $aIncludes
			$sSourceCode = $sIFile <> $__ThisFile ? FileRead(@ScriptDir & "\" & $sIFile) & @CRLF & $sSourceCode : $sSourceCode
		Next
	EndIf
	$aClassRegions = StringRegExp($sSourceCode, "(?s)\Q#Region Class \E(.*?)(?=\Q#EndRegion\E)", 3)

	; Do nothing if there are no classes in the code
	If Not UBound($aClassRegions) Then Return

	; Go through classes and build executable template commands
	For $sCurTemplate In $aClassRegions
		$sClassName = StringRegExp($sCurTemplate, "([a-zA-Z0-9_]+)", 3)[0]
		If $sClassName = "" Then
			Exit ConsoleWrite("!> Class name " & ($sClassName ? "contains invalid characters" : "is empty") & "." & @LF)
		EndIf
		$sCurTemplate = StringTrimLeft($sCurTemplate, $sClassName)
		If Not StringInStr($sCurTemplate, "Func ") Then Exit ConsoleWrite("!> Class " & $sClassName & " has no methods. Use Maps for structured data." & @LF)
		$sVarSpace = StringLeft($sCurTemplate, StringInStr($sCurTemplate, "Func") - 1)
		$aVariables = StringRegExp($sVarSpace, "(\$[a-zA-Z0-9_]+)", 3)
		If Not UBound($aVariables) Then Exit ConsoleWrite("!> Class " & $sClassName & " has no public variables. Use UDFs as collections of functions." & @LF)
		For $sEach In $aVariables
			__NamedArray($sClassName, __VarGetType($sEach) & " " & StringMid($sEach, 2) & ";")
		Next
		__NamedArray($sClassName, StringTrimRight(__NamedArray($sClassName), 1), True)
		__NamedArray($sClassName, "%")
		$sCurTemplate = StringTrimLeft($sCurTemplate, StringLen($sVarSpace))
		$aFunctions = StringRegExp($sCurTemplate, "(?s)\QFunc \E(.*?)(?=\QEndFunc\E)", 3)
		$sConstructor = 'Default'
		$sDestructor = 'Default'
		For $sEach In $aFunctions
			$sFuncName = StringLeft($sEach, StringInStr($sEach, '(') - 1)
			If $sFuncName == "_" & $sClassName Then
				$sConstructor = $sEach
			ElseIf $sFuncName == "__" & $sClassName Then
				$sDestructor = $sEach
			Else
				$sEach = StringTrimLeft($sEach, StringLen($sFuncName) + 1)
				$sArgHeader = StringLeft($sEach, StringInStr($sEach, ')') - 1)
				__NamedArray($sClassName, "?" & $sFuncName & "(" & $sArgHeader & ")?{" & StringMid($sEach, StringInStr($sEach, ')') + 1) & "} " & __VarGetType($sFuncName, True) & "(")
				If $sArgHeader <> "" Then
					If StringInStr($sArgHeader, ' = ') Then
						Exit ConsoleWrite("!> Method " & $sFuncName & " in class " & $sClassName & " has default values. This is not (yet) supported." & @LF)
					EndIf
					$aArgs = StringRegExp($sArgHeader, "(\$[a-zA-Z0-9_]+)", 3)
					For $sParameter In $aArgs
						__NamedArray($sClassName, __VarGetType($sParameter) & ",")
					Next
					__NamedArray($sClassName, StringTrimRight(__NamedArray($sClassName), 1), True)
				EndIf
				__NamedArray($sClassName, "); ")
			EndIf
		Next
		__NamedArray($sClassName, StringTrimRight(__NamedArray($sClassName), 2), True)
		__NamedArray($sClassName, "%" & $sConstructor & "%" & $sDestructor)
	Next
	For $sReg In $aClassRegions
		$sSourceCode = StringReplace($sSourceCode, '#Region Class ' & $sReg & '#EndRegion', @CRLF)
	Next
	$aClassDefs =  StringRegExp($sSourceCode, "(?s)\Q# <classdef:\E(.*?)(?=\Q>\E)", 3)
	$sOPClean = ''
	If UBound($aClassDefs) Then
		For $sEach In $aClassDefs
			$sORG = $sEach
			Local $sInsertInstance = ["", ""]
			$sGetTName = StringLeft($sEach, StringInStr($sEach, " ") - 1)
			If StringLen(__NamedArray($sGetTName)) < 1 Then Exit ConsoleWrite("!> Class " & $sGetTName & " does not exist :(" & @LF)
			$sEach = StringTrimLeft($sEach, StringLen($sGetTName) + 1)
			$aVar = StringRegExp($sEach, "(\$[a-zA-Z0-9_]+\[?[0-9]?+\]?)", 3)
			If Not UBound($aVar) Then Exit ConsoleWrite("!> classdef without any variables." & @LF)
			$sAssignTo = $aVar[0]
			Local $bArray = IsArray(Execute($sAssignTo))
			If $bArray Then
				$sOPClean &= "For $n = 0 To UBound(" & $sAssignTo & ")-1" & @CRLF & @TAB & $sAssignTo & "[$n] = 0" & @CRLF & "Next" & @CRLF
				$sInsertInstance[0] &= "For $n = 0 To UBound(" & $sAssignTo & ")-1" & @CRLF & @TAB & $sAssignTo & "[$n]" & _
				" = _OOPCreateClassInstance('" & StringSplit(__NamedArray($sGetTName),'%',3)[0] & "','"
			Else
				$sOPClean &= $sAssignTo & " = 0" & @CRLF
				$sInsertInstance[0] &= $sAssignTo & " = _OOPCreateClassInstance('" & StringSplit(__NamedArray($sGetTName),'%',3)[0] & "','"
			EndIf
			$sFuncDef = StringSplit(__NamedArray($sGetTName),'%',3)[1];([a-zA-Z0-9_]+\{)
			$aFuncNms = StringRegExp($sFuncDef, "(\Q?\E.*\Q?{\E)", 3)
;~ 			ClipPut($sFuncDef)
;~ 			_ArrayDisplay($aFuncNms)
			$prfx = "___o_" & StringRegExpReplace($sAssignTo, "([^A-Za-z0-9])", "") & "_"
			If Not UBound($aFuncNms) Then Exit ConsoleWrite("!> Class " & $sGetTName & " has no methods. Use Maps for structured data." & @LF)
			For $sFCT In $aFuncNms
				$sFCT = StringTrimRight(StringMid($sFCT, 2), 2)
;~ 				If $sFCT = 0 Then ContinueLoop
				$sInsertInstance[1] &= @CRLF & "Func " & $prfx & StringReplace(StringTrimRight($sFCT, 1), "(", "(" & $stdRefVar & (StringInStr($sFCT, "$") ? "," : "")) & ")"
				$sInsertInstance[1] &= @CRLF & @TAB & "Local $This = DllStructCreate(__PointerToString(DllStructCreate($__s_ObjVars, " & $stdRefVar & ").psVarsString), " & $stdRefVar & " + $__i_ObjVarsOffset)"
;~ 				MsgBox(0,0,$sFCT)
				$sFCB = StringRegExp($sFuncDef, "(?s)\Q?" & $sFCT & "?{\E(.*?)(?=\Q}\E)", 3)[0]
				$sInsertInstance[1] &= __UndoIndent(StringTrimRight($sFCB,1)) & "EndFunc" & @CRLF
				$sInsertInstance[0] &= $prfx & StringLeft($sFCT, StringInStr($sFCT, "(") - 1)
				$sInsertInstance[0] &= StringRegExp($sFuncDef, "(?s)\Q" & $sFCB & "}\E(.*?)(?=\Q)\E)", 3)[0] & ");"
				$sMethdName = StringLeft($sFCT, StringInStr($sFCT, "(") - 1)
				; Find all references and point to instantiated class:
				If $bArray Then
					$aRefs = StringRegExp($sSourceCode, "(\$\Q" & StringMid($sAssignTo, 2) & "\E\[.*?\]\Q." & $sMethdName & "(\E)", 3)
					For $sReference In $aRefs
						$sSourceCode = StringReplace($sSourceCode, $sReference, StringReplace($sReference, "." & $sMethdName, "." & $prfx & $sMethdName, 0, 1))
					Next
				Else
					$sSourceCode = StringReplace($sSourceCode, $sAssignTo & "." & $sMethdName, $sAssignTo & "." & $prfx & $sMethdName, 0, 1)
				EndIf
			Next
			$sInsertInstance[0] = StringTrimRight($sInsertInstance[0], 1) & "',"
			$sInitArg = StringInStr($sORG, ' from ') ? StringMid($sORG, StringInStr($sORG, ' from ') + 6) : ''
			$sCallC = StringSplit(__NamedArray($sGetTName),'%',3)[2]
			If $sCallC <> "Default" Then
				$sInsertInstance[1] &= @CRLF & _
				StringReplace(__UndoIndent(StringTrimRight(StringReplace($sCallC, "_" & $sGetTName & "(", "Func " & _
				$prfx & "constructor(" & $stdRefVar & ($sInitArg ? "," : "")),1)), _
				')', ')' & @CRLF & @TAB & _
				"Local $This = DllStructCreate(__PointerToString(DllStructCreate($__s_ObjVars, " & _
				$stdRefVar & ").psVarsString), " & $stdRefVar & " + $__i_ObjVarsOffset)" & @CRLF, 1) & _
				"EndFunc"
				$sInsertInstance[0] &= $prfx & "constructor,"
			Else
				$sInsertInstance[0] &= $sCallC & ","
			EndIf
			$sCallC = StringSplit(__NamedArray($sGetTName),'%',3)[3]
			If $sCallC <> "Default" Then
				$sInsertInstance[1] &= @CRLF & __UndoIndent(StringTrimRight(StringReplace($sCallC, "__" & $sGetTName & "(", "Func " & $prfx & "destructor(" & $stdRefVar),1)) & "EndFunc"
				$sInsertInstance[0] &= $prfx & "destructor"
			Else
				$sInsertInstance[0] &= $sCallC
			EndIf
			$sInsertInstance[0] &= ($sInitArg ? "," & $sInitArg : "") & ")" & @CRLF
			If $bArray Then $sInsertInstance[0] &= "Next" & @CRLF
			$sSourceCode = StringReplace($sSourceCode, '# <classdef:' & $sORG & '>', $sInsertInstance[0] & $sInsertInstance[1] & ';')
		Next
	EndIf
	$sSourceCode = StringReplace($sSourceCode, '#AutoIt3Wrapper_Run_AU3Check=n', '#AutoIt3Wrapper_Change2CUI=y' & @CRLF & 'OnAutoItExitRegister("__cleanup")')
	$sSourceCode &= @CRLF & 'Func __cleanup()' & @CRLF & $sOPClean & 'EndFunc' & @CRLF
	FileDelete($RegEnScript)
	FileWrite($RegEnScript, $sSourceCode)
	$iPID = Run(@AutoItExe & ' /ErrorStdOut "' & $RegEnScript & '"', @ScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	If @error Then Exit ConsoleWrite("!> Error running script." & @LF)
	ConsoleWrite("> Running script (lag: " & Round(TimerDiff($__Timer), 2) & "ms)..." & @LF & @LF)
	Local $aLatestMem = [-1, -1]
	While ProcessExists($iPID)
		ConsoleWrite(StdoutRead($iPID))
		$aLatestMem = ProcessGetStats($iPID)
	WEnd
	If Not UBound($aLatestMem) Then Dim $aLatestMem = [-1, -1]
	Exit ConsoleWrite(@LF & "> Exit script (Mem: " & Round($aLatestMem[0] / 1024^2) & "M - Peak: " & Round($aLatestMem[1] / 1024^2) & "M - Time: " & Round(TimerDiff($__Timer)/1e3, 2) & "s)." & @LF) * 0
EndFunc

; by Mars, minx
Func _OOPCreateClassInstance($sObjVars, $sObjMethods, $xFuncConst = Default, $xFuncDest = Default, $aParamConst = Null)
	If $xFuncConst = Default Then $xFuncConst = __DefaultConst
	If $xFuncDest  = Default Then $xFuncDest  = __DefaultDest

	If Not $__CR Then
		If Not $__CQI Then $__CQI = DllCallbackRegister('__QueryInterface', 'none', 'ptr;ptr;ptr*')
		If Not $__CAR Then $__CAR = DllCallbackRegister('__AddRef', 'uint', 'ptr')
		If Not $__CR Then $__CR   = DllCallbackRegister('__Release', 'uint', 'ptr')
	EndIf

	Local $aSplit = StringSplit($sObjMethods, ';', 2), $sObjTypes = ''
	Local $nMethods = UBound($aSplit), $aRetTypes[$nMethods], $aParams[$nMethods]
	Local $sSingleMethod = '', $aSingleMethod, $aMethods[$nMethods], $aiCallback[$nMethods], $sMethods

	For $i = 0 To $nMethods - 1 Step 1
		$sSingleMethod = StringStripWS($aSplit[$i], 7)
		$aSingleMethod = StringSplit($sSingleMethod, ' ', 2)
		$sObjTypes &= 'ptr ' & $aSingleMethod[0] & ';'
		$aMethods[$i] = $aSingleMethod[0]
		$sMethods &= $aSingleMethod[0] & ';'
		$aSingleMethod = StringSplit($aSingleMethod[1], '(', 2)
		$aRetTypes[$i] = $aSingleMethod[0]
		$aParams[$i] = 'ptr;' & StringReplace(StringTrimRight($aSingleMethod[1], 1), ',', ';', 0, 1)
	Next

	$sObjMethods = StringReplace($sObjMethods, ',', ';', 0, 1)

	Local $tObj = __DllStructCreateProtected($__s_ObjVars & $sObjVars & '; ' & $__s_ObjMethods & StringTrimRight($sObjTypes, 1))
	$tObj.Vtbl = DllStructGetPtr($tObj, '__QueryInterface')
	$tObj.psVarsString = DllStructGetPtr(__StringToStructProtected($sObjVars))
	$tObj.psMethodsString = DllStructGetPtr(__StringToStructProtected(StringTrimRight($sMethods, 1)))
	$tObj.psDestruktor = DllStructGetPtr(__StringToStructProtected(IsFunc($xFuncDest) ? FuncName($xFuncDest) : $xFuncDest))
	$tObj.RefCnt = 1

	For $i = 0 To $nMethods - 1 Step 1
		$aiCallback[$i] = DllCallbackRegister($aMethods[$i], $aRetTypes[$i], $aParams[$i])
		DllStructSetData($tObj, $aMethods[$i], DllCallbackGetPtr($aiCallback[$i]))
	Next

	$tObj.__QueryInterface = DllCallbackGetPtr($__CQI)
	$tObj.__AddRef = DllCallbackGetPtr($__CAR)
	$tObj.__Release = DllCallbackGetPtr($__CR)

	Local $viaCallback = __DllStructCreateProtected('int iCallbackCnt; int aiCallback[' & $nMethods & ']')
	$viaCallback.iCallbackCnt = $nMethods

	For $i = 1 To $nMethods Step 1
		$viaCallback.aiCallback(($i)) = $aiCallback[$i - 1]
	Next

	$tObj.paiCallback = DllStructGetPtr($viaCallback)
	$oObj = ObjCreateInterface(DllStructGetPtr($tObj), _WinAPI_CreateGUID(), $sObjMethods)
	If Not IsFunc($xFuncConst) Then $xFuncConst = Execute($xFuncConst)
	Execute('$xFuncConst(DllStructGetPtr($tObj)' & ($aParamConst = Null ? ')' : ', $aParamConst)'))

	Return $oObj
EndFunc

Func __AddRef($pObj)
	Local $tObj = DllStructCreate('ptr Vtbl; int RefCnt', $pObj)
	$tObj.RefCnt += 1
	Return $tObj.RefCnt
EndFunc

Func __DefaultConst($pObj)
EndFunc

Func __DefaultDest($pObj)
EndFunc

Func __DllStructCreateProtected($sStruct)
	Return DllStructCreate($sStruct, _MemGlobalAlloc(DllStructGetSize(DllStructCreate($sStruct))))
EndFunc

Func __PointerToString($iPointer)
	Local $vStruct = DllStructCreate('uint iLen; char sString[' & DllStructCreate('uint iLen', $iPointer).iLen & ']', $iPointer)
	Return $vStruct.sString
EndFunc

Func __QueryInterface($pObj, $IID, $pPTR)
	Local $tObj = DllStructCreate('ptr Vtbl', $pPTR)
	$tObj.Vtbl = $pObj
	__AddRef($pObj)
	Return 0
EndFunc

Func __Release($pObj)
	Local $tObj = DllStructCreate(StringTrimRight($__s_ObjVars, 2), $pOBJ)
	$tObj.RefCnt -= 1
	If $tObj.RefCnt = 0 Then
		Local $xFuncDest = Execute(__PointerToString($tObj.psDestruktor)), $vStruct
		$xFuncDest($pObj)
		$vStruct = DllStructCreate('int iCallbackCnt', $tObj.paiCallback)
		Local $vaiCallback = DllStructCreate('int iCallbackCnt; int aiCallback[' & $vStruct.iCallbackCnt & ']', $tObj.paiCallback)
		For $i = 1 To $vStruct.iCallbackCnt
			DllCallbackFree($vaiCallback.aiCallback(($i)))
		Next
		_MemGlobalFree($tObj.paiCallback)
		_MemGlobalFree($tObj.psVarsString)
		_MemGlobalFree($tObj.psMethodsString)
		_MemGlobalFree($tObj.psDestruktor)
		_MemGlobalFree($pObj)
	EndIf
	Return $tObj.RefCnt
EndFunc

Func __StringToStructProtected($sString)
	Local $iLen = StringLen($sString)
	Local $vStruct = __DllStructCreateProtected('uint iLen; char sString[' & $iLen & ']')
	$vStruct.iLen = $iLen
	$vStruct.sString = $sString
	Return $vStruct
EndFunc

Func __NamedArray($sKey, $nVal = Null, $bOverwrite = False)
	$iKey = _ArraySearch($__ClsTbl, $sKey, 0, 0, 1, 2, 1, 0)
	If $iKey = -1 Then
		If $nVal = Null Then Return ''
		ReDim $__ClsTbl[UBound($__ClsTbl)+1][2]
		$__ClsTbl[UBound($__ClsTbl)-1][0] = $sKey
		$__ClsTbl[UBound($__ClsTbl)-1][1] = $nVal
	Else
		If $nVal = Null Then Return $__ClsTbl[$iKey][1]
		$__ClsTbl[$iKey][1] = $bOverwrite ? $nVal : $__ClsTbl[$iKey][1] & $nVal
	EndIf
EndFunc

Func __UndoIndent($s)
	Return StringReplace($s, @TAB & @TAB, @TAB)
EndFunc

Func __VarGetType($sVar, $bFunc = False)
	Local Static $iTrim = $bFunc ? 1 : 2
	Switch StringMid($sVar, $iTrim, 1)
		Case "f"
			Return "double"
		Case Else
			Return "int64"
	EndSwitch
EndFunc