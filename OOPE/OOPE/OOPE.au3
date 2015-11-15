; OOP Extender

#include <Array.au3>
#include <AutoItConstants.au3>
#include <WinAPICom.au3>
#include <Memory.au3>
#include <StringConstants.au3>
#include <Crypt.au3>
#include-once

Global Const $__OOPE_MacroStatement             = 'define'
Global Const $__OOPE_MacroPattern               = '((?m)^#' & $__OOPE_MacroStatement & ' (?:.*\\\r?\n)*.*$)'
Global Const $__OOPE_MacroNamePattern           = '([\w]+)'
Global Const $__OOPE_MacroTilCParenPattern      = '(.*[\Q\)\E])'
Global Const $__OOPE_MacroArgpPattern           = '(\$[\w]+)'
Global Const $__OOPE_MacroEscBackslash          = '{bcksl}'
Global Const $__OOPE_ClassTagStatement          = 'classdef'
Global Const $__OOPE_ClassTagPattern            = '((?m)^#' & $__OOPE_ClassTagStatement & ' (?:.*\\\r?\n)*.*$)'
Global Const $__OOPE_ClassTagMembersPattern     = '((<.+?>?\h?)?\$[\w\[\]]+(?i)( +from +<[^>]+>)?)'
Global Const $__OOPE_ClassRegionStatement       = ' Class '
Global Const $__OOPE_ClassRegionPattern         = '(?s)\Q#Region' & $__OOPE_ClassRegionStatement & '\E(.*?)(?=\Q#EndRegion\E)'
Global Const $__OOPE_ClassFieldsPattern         = '\$([\w]+)'
Global Const $__OOPE_AutoIt3UniqStringsPattern  = '("[^"\r\n]*?"|''[^''\r\n]*?'')'
Global Const $__OOPE_AutoIt3UniqCommentsPattern = '(;.*|#cs[^.]*#ce)'
Global Const $__OOPE_ClassUserMethodsPattern    = '(?si)Func ([^_].*?)(?=\QEndFunc\E)'
Global Const $__OOPE_ClassConDesMethodsPattern  = '(?si)Func (_+%CLASSNAME%.*?)(?=\QEndFunc\E)'
Global Const $__OOPE_Debug                      = False
Global Const $__OOPE_ObjectInstanceVariables    = 'ptr Vtbl; int RefCnt; ptr paiCallback; ptr psVarsString; ptr psMethodsString; ptr psDestruktor; '
Global Const $__OOPE_ObjectVariableOffset       = DllStructGetSize(DllStructCreate($__OOPE_ObjectInstanceVariables))
Global Const $__OOPE_ObjectStdMethods           = 'ptr __QueryInterface; ptr __AddRef; ptr __Release; '
Global Const $__OOPE_OOPEStdIncludeName         = 'OOPE.au3'
Global Const $__OOPE_GlobalLagTimer             = TimerInit()
Global $__OOPE_GarbageCollectionBuildString     = ''
Global $__CR, $__CQI, $__CAR


_OOPE_ParseAndRun()

Func _OOPE_ParseAndRun()
	Local $sScriptPath      = @ScriptFullPath
	Local $sScriptFName     = @ScriptName
	Local $bParserFinishOK  = False

	If StringInStr($sScriptPath, '_stable') Then Return
	Local $sGeneratedScript = StringReplace($sScriptPath, '.au3', '_stable.au3')
	If Not @extended Then _OOPE_Panic("This script (" & $sScriptFName & ") has an invalid extension! Please rename to .au3.")

	Local $sSourceCode = FileRead($sScriptPath)
	If $sSourceCode = "" Or @error Then _OOPE_Panic("Error reading script file!")
	If _OOPE_WriteTest() Then _OOPE_Panic("Error writing in script directory. Permissions?")

	; Purge source file of nasty comments and strings that might contain confusing code.
	Local $aStringsInCode = StringRegExp($sSourceCode, $__OOPE_AutoIt3UniqStringsPattern, 3)
	For $sEach In $aStringsInCode
		$sSourceCode = StringReplace($sSourceCode, $sEach, _Crypt_HashData($sEach, $CALG_MD5))
	Next
	$sSourceCode = StringRegExpReplace($sSourceCode, $__OOPE_AutoIt3UniqCommentsPattern, @CRLF)

	; (1) Run the macro extender
	If _OOPE_ContainsMacros($sSourceCode) Then $bParserFinishOK = _OOPE_ExtendMacros($sSourceCode)

	; (2) Run the object extender
	If _OOPE_ContainsClasses($sSourceCode) Then $bParserFinishOK = _OOPE_ExtendObjects($sSourceCode)

	; Re-insert all strings
	For $sEach In $aStringsInCode
		$sSourceCode = StringReplace($sSourceCode, _Crypt_HashData($sEach, $CALG_MD5), $sEach)
	Next

	_OOPE_WriteAndRun($sSourceCode, $sGeneratedScript)
EndFunc

Func _OOPE_WriteAndRun($sCode, $sFile)
	$bIsConsole = StringInStr($sCode, '#AutoIt3Wrapper_Change2CUI=y') Or StringInStr($sCode, 'ConsoleWrite')
	$sCode = StringReplace($sCode, '#AutoIt3Wrapper_Run_AU3Check=n', ($bIsConsole?'#AutoIt3Wrapper_Change2CUI=y':'') & @CRLF & 'OnAutoItExitRegister("__cleanup")')
	$sCode &= @CRLF & 'Func __cleanup()' & @CRLF & $__OOPE_GarbageCollectionBuildString & 'EndFunc' & @CRLF
	FileDelete($sFile)
	FileWrite($sFile, $sCode)
	If Not $bIsConsole Then Exit Run(@AutoItExe & ' "' & $sFile & '"', @ScriptDir) = 0
	$iPID = Run(@AutoItExe & ' /ErrorStdOut "' & $sFile & '"', @ScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	If @error Then Exit ConsoleWrite("!> Error running script." & @LF)
	ConsoleWrite("> Running script (lag: " & Round(TimerDiff($__OOPE_GlobalLagTimer), 2) & "ms)..." & @LF & @LF)
	Local $aLatestMem = [-1, -1]
	While ProcessExists($iPID)
		ConsoleWrite(StdoutRead($iPID))
		$aLatestMem = ProcessGetStats($iPID)
	WEnd
	If Not UBound($aLatestMem) Then Dim $aLatestMem = [-1, -1]
	Exit ConsoleWrite(@LF & "> Exit script (Mem: " & Round($aLatestMem[0] / 1024^2) & _
	"M - Peak: " & Round($aLatestMem[1] / 1024^2) & _
	"M - Time: " & Round(TimerDiff($__OOPE_GlobalLagTimer)/1e3, 2) & "s)." & @LF) * 0
EndFunc

Func _OOPE_ExtendObjects(ByRef $sCode)
	; Get all class templates
	Local $aClsRegions = StringRegExp($sCode, $__OOPE_ClassRegionPattern, 3)
	Local $aClasses[0][8] ; Name, Fields, Methods, Constructor, Destructor, DestructorHasArguments, Parsed Functions, IsClassUsed

	For $sClsRegion In $aClsRegions
		$sClassName  = StringLeft($sClsRegion, StringInStr($sClsRegion, @CRLF) - 1)
		If $__OOPE_Debug Then __OOPE_Debug("Entering definition of class <" & $sClassName & ">.")
		$sClassBody = StringTrimLeft($sClsRegion, StringLen($sClassName) + 1)
		$sFieldSpace = StringLeft($sClassBody, StringInStr($sClassBody, 'Func') - 1)
		$aFields = StringRegExp($sFieldSpace, $__OOPE_ClassFieldsPattern, 3)
		If Not UBound($aFields) Then _OOPE_Panic("Error: Class " & $sClassName & " doesn't contain any fields.")
		For $nEach = 0 To UBound($aFields)-1
			$aFields[$nEach] = (StringLeft($aFields[$nEach], 1) = 'f' ? 'double ' : 'int64 ') & $aFields[$nEach]
		Next
		; build explicitely typed field declaration
		$sClassFields = _ArrayToString($aFields, ";")

		$sClassBody = StringTrimLeft($sClassBody, StringLen($sFieldSpace))
		$aClassMethodsUser = StringRegExp($sClassBody, $__OOPE_ClassUserMethodsPattern, 3)
		If Not UBound($aClassMethodsUser) Then _OOPE_Panic("Class " & $sClassName & " doesn't contain any methods (not counting destructors / constructors. Use Maps / Arrays for structured data!")
		; Consructors and Destructors are AutoIt functions called from AutoIt, they don't have to be typed!
		; Everything else is called from the interface and needs types.
		Local $aMTypeDecls[0]
		Local $sCompMethods = ''
		For $sUM In $aClassMethodsUser
			$sMName = StringLeft($sUM, StringInStr($sUM, '(')-1)
			$sMTypeDecl = $sMName & " " & (StringLeft($sUM, 1) = 'f' ? 'double' : 'int64') & "("
			$sUM = StringTrimLeft($sUM, StringInStr($sUM, '('))
			$sArgHeader = StringLeft($sUM, StringInStr($sUM, ')')-1)
			$aMArgs = StringRegExp($sArgHeader, $__OOPE_ClassFieldsPattern, 3)
			$sCompMethods &= _
			"Func " & $sMName & "($___selfObjRef" & (UBound($aMArgs)?',$'&_ArrayToString($aMArgs, ",$"):'') & ")" & _
				@CRLF & '$This = DllStructCreate(__PointerToString(DllStructCreate($__OOPE_ObjectInstanceVariables, $___selfObjRef).psVarsString), $___selfObjRef + $__OOPE_ObjectVariableOffset)' & _
				@CRLF & StringMid($sUM, StringInStr($sUM, ')')+1) & @CRLF & _
			"EndFunc" & @CRLF
			For $kArg = 0 To UBound($aMArgs)-1
				$aMArgs[$kArg] = StringLeft($kArg, 1) = 'f' ? 'double' : 'int64'
			Next
			$sMTypeDecl &= (UBound($aMArgs)?_ArrayToString($aMArgs, ","):'') & ")"
			_ArrayAdd($aMTypeDecls, $sMTypeDecl)
		Next
		$sClassMethodsUser = _ArrayToString($aMTypeDecls, ";")

		$aClassConDes = StringRegExp($sClassBody, StringReplace($__OOPE_ClassConDesMethodsPattern, '%CLASSNAME%', $sClassName), 3)
		$sConstructor = 'Default'
		$sDestructor  = 'Default'
		$sConHasArgs  = False
		If UBound($aClassConDes) Then
			For $sConDes In $aClassConDes
				$bIsDestructor = StringLeft($sConDes, 2) == '__'
				$sFName        = StringLeft($sConDes, StringInStr($sConDes, '(')-1)
				$sConstructor  = $bIsDestructor ? $sConstructor : $sFName
				$sDestructor   = $bIsDestructor ? $sFName       : $sDestructor
				$sConDes       = StringTrimLeft($sConDes, StringLen($sFName)+1)
				$sArgs         = StringRegExp(StringLeft($sConDes, StringInStr($sConDes, ')')-1), $__OOPE_ClassFieldsPattern, 3)
				If UBound($sArgs) And $bIsDestructor Then _OOPE_Panic("Whoops! Destructor has arguments. That makes no sense.")
				If Not $bIsDestructor And UBound($sArgs) Then $sConHasArgs = True
				$sCompMethods &= _
				@CRLF & 'Func ' & $sFName & '($___selfObjRef' & (UBound($sArgs)?',$'&$sArgs[0]:'') & ")" & _
					@CRLF & '$This = DllStructCreate(__PointerToString(DllStructCreate($__OOPE_ObjectInstanceVariables, $___selfObjRef).psVarsString), $___selfObjRef + $__OOPE_ObjectVariableOffset)' & _
					@CRLF & StringMid($sConDes, StringInStr($sConDes, ')')+1) & @CRLF & _
				'EndFunc' & @CRLF
			Next
		EndIf
		ReDim $aClasses[UBound($aClasses)+1][8]
		$aClasses[UBound($aClasses)-1][0] = $sClassName
		$aClasses[UBound($aClasses)-1][1] = $sClassFields
		$aClasses[UBound($aClasses)-1][2] = $sClassMethodsUser
		$aClasses[UBound($aClasses)-1][3] = $sConstructor
		$aClasses[UBound($aClasses)-1][4] = $sDestructor
		$aClasses[UBound($aClasses)-1][5] = $sConHasArgs
		$aClasses[UBound($aClasses)-1][6] = $sCompMethods

		$sCode = StringReplace($sCode, '#Region Class ' & $sClsRegion & '#EndRegion', '')
	Next

	Local $aInstances[0][3] ; Type, Name
	; Get all declarations from classdef tags and fill array
	$aClassDefTags = StringRegExp($sCode, $__OOPE_ClassTagPattern, 3)
	For $sCDT In $aClassDefTags
		$sLastType = ''
		$aDecls = StringRegExp($sCDT, $__OOPE_ClassTagMembersPattern, 3)
		If Not UBound($aDecls) Then _OOPE_Panic("Classdef contains invalid symbols.")
		For $sEach In $aDecls
			If StringLeft(StringStripWS($sEach, $STR_STRIPLEADING), 4) = "from" Then ContinueLoop
			If StringInStr($sEach, '$') = 0 Then ContinueLoop
			If StringInStr($sEach, '<') Then
				$sLastType = StringTrimLeft($sEach, 1)
				$sLastType = StringLeft($sLastType, StringInStr($sLastType, '>') - 1)
				$sEach = StringMid($sEach, StringInStr($sEach, '>') + 1)
			EndIf
			If StringInStr($sEach, '<') And StringInStr($sEach, 'from') Then
				$aCArg = StringTrimRight(StringStripWS($sEach, $STR_STRIPTRAILING), 1)
				$aCArg = StringMid($aCArg, StringInStr($aCArg, '<') + 1)
				$sEach = StringStripWS(StringLeft($sEach, StringInStr($sEach, 'from') - 1), $STR_STRIPTRAILING)
			Else
				$aCArg = 'noop'
			EndIf
			ReDim $aInstances[UBound($aInstances)+1][3]
			$aInstances[UBound($aInstances)-1][0] = StringStripWS($sLastType, 3)
			$aInstances[UBound($aInstances)-1][1] = StringStripWS($sEach, 3)
			$aInstances[UBound($aInstances)-1][2] = $aCArg
		Next

		$sClsDimStr = ''
		For $k = 0 To UBound($aInstances)-1
			$nClass = -1
			For $l = 0 To UBound($aClasses)-1
				If $aClasses[$l][0] = $aInstances[$k][0] Then $nClass = $l
			Next
			If $nClass = -1 Then _OOPE_Panic($aInstances[$k][1] & " wants to be a " & $aInstances[$k][0] & " object, but there is no class with that name!")
			$sOOPEArgument = $aClasses[$nClass][1] & '", "' & $aClasses[$nClass][2] & '", ' & $aClasses[$nClass][3] & ', ' & $aClasses[$nClass][4]
			If $aInstances[$k][2] <> 'noop' And Not $aClasses[$nClass][5] Then _OOPE_Panic("Error: classdef tried to pass arguments to " & $aInstances[$k][0] & " (parameterless).")
			If $aInstances[$k][2] == 'noop' And $aClasses[$nClass][5] Then _OOPE_Panic("Error: classdef didn't pass any arguments to " & $aInstances[$k][0] & " (requires arguments).")
			$sOOPEArgument &= $aInstances[$k][2] <> 'noop' ? ', ' & $aInstances[$k][2] : ''
			$sClsDimStr &= _
			'#forcedef ' & $aInstances[$k][1] & @CRLF & _
			'If UBound(Execute("' & $aInstances[$k][1] & '")) Then' & @CRLF & _
			'	For $n = 0 To UBound(Execute("' & $aInstances[$k][1] & '"))-1' & @CRLF & _
			'		' & $aInstances[$k][1] & '[$n] = __OOPE_InstantiateClass("' & $sOOPEArgument & ')' & _
			@CRLF & '	Next' & @CRLF & _
			'Else' & @CRLF & _
			'	' & $aInstances[$k][1] & ' = __OOPE_InstantiateClass("' & $sOOPEArgument & ')' & @CRLF & _
			'EndIf' & @CRLF
			$aClasses[$nClass][7] = True
			$__OOPE_GarbageCollectionBuildString &= _
			@CRLF & '#forcedef ' & $aInstances[$k][1] & @CRLF & _
			@CRLF & 'If UBound(Execute("' & $aInstances[$k][1] & '")) Then' & @CRLF & _
			'	For $n = 0 To UBound(Execute("' & $aInstances[$k][1] & '"))-1' & @CRLF & _
			'	' & $aInstances[$k][1] & '[$n] = 0' & @CRLF & _
			'	Next' & @CRLF & _
			'Else' & @CRLF & _
			'	' & $aInstances[$k][1] & ' = 0' & @CRLF & _
			'EndIf' & @CRLF
		Next
		$sCode = StringReplace($sCode, $sCDT, $sClsDimStr)
	Next

	; Now compile every used class into the code, leave out unused classes:
	$sCode &= @CRLF
	For $k = 0 To UBound($aClasses)-1
		If $aClasses[$k][7] Then $sCode &= $aClasses[$k][6] & @CRLF
	Next
EndFunc

Func _OOPE_ContainsClasses($sCode)
	Return UBound(StringRegExp($sCode, $__OOPE_ClassTagPattern, 3))
EndFunc

Func _OOPE_ExtendMacros(ByRef $sCode)
	Local $aRawMacros = StringRegExp($sCode, $__OOPE_MacroPattern, 3)
	If $__OOPE_Debug Then __OOPE_Debug("Source file containts " & UBound($aRawMacros) & " macros.")

	; Cycle through all macros and parse them (they are NOT regular)
	For $sRawMacro In $aRawMacros
		$sTemplate  = StringTrimLeft($sRawMacro, StringLen($__OOPE_MacroStatement) + 1)
		$sTemplate  = StringStripWS($sTemplate, $STR_STRIPLEADING)
		$sMacroName = StringRegExp($sTemplate, $__OOPE_MacroNamePattern, 3)

		If Not UBound($sMacroName) Or $sMacroName = "" Then _OOPE_Panic("Invalid macro name in >" & $sRawMacro & "<.")
		$sMacroName = $sMacroName[0]
		If $__OOPE_Debug Then __OOPE_Debug("Declare macro: " & $sMacroName)

		; Macro arguments are static+regular and do not pollute the global namespace, so RegEx is it then
		$sTemplate  = StringTrimLeft($sTemplate, StringLen($sMacroName) + 1)
		; ^ TODO: Make arg-less macros.
		$sArgumentSpace = StringLeft($sTemplate, StringInStr($sTemplate, ')') - 1)
		$aArguments = StringRegExp($sArgumentSpace, $__OOPE_MacroArgpPattern, 3)
		If Not UBound($aArguments) Then _OOPE_Panic("Macro has weird arguments >" & $sRawMacro & "<.")

		; Explode macro code
		$sTemplate  = StringMid($sTemplate, StringInStr($sTemplate, ')') + 1)
		$sTemplate  = StringTrimLeft($sTemplate, 1)
		$sTemplate  = StringReplace($sTemplate, "\\", $__OOPE_MacroEscBackslash)
		$sTemplate  = StringReplace($sTemplate, "\", @CRLF)
		$sTemplate  = StringReplace($sTemplate, $__OOPE_MacroEscBackslash, "\\")
		$sMacroCode = $sTemplate

		; Find references and parse them (cause RegEx can't to all the escapement AutoIt does)
		$aReferences = __OOPE_StrFindAll($sCode, $sMacroName & '(')
		If Not UBound($aReferences) And $__OOPE_Debug Then __OOPE_Debug("Macro " & $sMacroName & " has no references.")
		If UBound($aReferences) Then
			For $iPosRef In $aReferences
				; Check if this is the definition and skip it
				ContinueLoop StringMid($sCode, $iPosRef - (StringLen($__OOPE_MacroStatement) + 2), StringLen($__OOPE_MacroStatement) + 2) = '#' & $__OOPE_MacroStatement & " "
				$sReference = ''
				$iPosRef += StringLen($sMacroName) + 1
				$nParenthesesLevel = 1
				Local $aRefArgs[0]
				Do
					$sScanChar = StringMid($sCode, $iPosRef, 1)
					$nParenthesesLevel += $sScanChar == '(' ? 1 : $sScanChar == ')' ? -1 : 0
					If ($sScanChar == ',' And $nParenthesesLevel < 2) Or Not $nParenthesesLevel Then
						_ArrayAdd($aRefArgs, $sReference)
						$sReference = ''
					Else
						$sReference &= $sScanChar
					EndIf
					$iPosRef += 1
				Until Not $nParenthesesLevel

				If UBound($aArguments) <> UBound($aRefArgs) Then _OOPE_Panic("Macro error: Argument count mismatch! " & $sMacroName & " must have " & UBound($aArguments) & " args.")

				$sLocalMacroCode = $sMacroCode
				For $n = 0 To UBound($aArguments)-1
					$sLocalMacroCode = StringReplace($sLocalMacroCode, $aArguments[$n], $aRefArgs[$n])
					If Not @extended Then _OOPE_Panic("Internal error parsing macros. Please submit your script as an issue.")
				Next

				$sCode = StringReplace($sCode, $sMacroName & "(" & _ArrayToString($aRefArgs, ",") & ")", $sLocalMacroCode)
			Next
		EndIf
	Next

	$sCode = StringRegExpReplace($sCode, $__OOPE_MacroPattern, '')
EndFunc

Func _OOPE_ContainsMacros($sCode)
	Return UBound(StringRegExp($sCode, $__OOPE_MacroPattern, 3))
EndFunc

Func __OOPE_StrFindAll($sString, $sSubString)
	Local $aPositions[0]
	While 1
		$iNewPos = StringInStr($sString, $sSubString, 0, UBound($aPositions)+1)
		If $iNewPos = 0 Then ExitLoop
		_ArrayAdd($aPositions, $iNewPos)
	WEnd
	Return $aPositions
EndFunc

Func __OOPE_Debug($sMessage)
	Return ConsoleWrite("-> " & $sMessage & @LF)
EndFunc

Func _OOPE_WriteTest()
	$hDummy = FileOpen(@ScriptDir & "\.tmp", 2)
	If @error Then Return 1
	FileDelete(@ScriptDir & "\.tmp")
	Return @error
EndFunc

Func _OOPE_Panic($sMessage = "Unknown Error")
	Exit ConsoleWrite("!> " & $sMessage & @LF)
EndFunc



;- - - - - - - - - - - - - - - - -- - - - - - - -- - - - - - - -- - - - - - - -- - - - -


; by Mars, minx
Func __OOPE_InstantiateClass($sObjVars, $sObjMethods, $xFuncConst = Default, $xFuncDest = Default, $aParamConst = Null)
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

	Local $tObj = __DllStructCreateProtected($__OOPE_ObjectInstanceVariables & $sObjVars & '; ' & $__OOPE_ObjectStdMethods & StringTrimRight($sObjTypes, 1))
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
	Local $tObj = DllStructCreate(StringTrimRight($__OOPE_ObjectInstanceVariables, 2), $pOBJ)
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







