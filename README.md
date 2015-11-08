# AutoIt OOP Extender

Since the introduction of ObjCreateInterface, AutoIt is able to natively create real objects. However, this is quite difficult for non-experienced users. This UDF here enables a "classic" syntax for declaring classes and instantiating objects from them. Before you ask: **NO**, it is *not* just another preprocessor that's faking OOP syntax - this is the real deal.

Care has been put into this UDF and it is in the authors interest to fix all bugs remaining and implement features as long as the result works in the **Stable** release of AutoIt.

### Features

- Define an unlimited number of classes.
- Classes can be defined in other includes of the script.
- Create unlimited instances of objects.
- Create arrays of objects.
- Mix and match different data types in arrays (one or more elements can be objects).
- Define custom constructors and destructors.
- Pass an unlimited number of arguments to the constructor (even to all objects in one array at the same time).
- Automatic garbage collection.
- Compatible with Object-enabled AutoIt keywords (`With` etc.), optional parentheses on parameterless functions.
- Fully AU3Check enabled.
- IntelliSense catches class-names for auto-completion.
- Automatically generates a compilable version of the script.
- Non-instantated classes get optimzed away.

## Tutorials

### Example 1: Creating a simple class

If you're coming from a VB/C#/JScript etc. background, you may be quite familiar with this syntax. The first thing we have to do is disable AU3Check. *But minxomat, you said above that this is AU3Check enable?!* Yes, it is. After the extender did it's magic AU3Check is run on all code that is not a OOP Extender command, so no worries. So the script starts off with

```
#AutoIt3Wrapper_Run_AU3Check=n
```

Now we can call the extender. This is implicitely done by this include statement. Now, be aware that the placement of this statement is very important if you want to create an array of objects. Read all tutorials here to understand why that is. For this example, we don't have anything else to do before calling the extender, so let's just include it:

```au3
#include 'OOP.au3'
```

Time to define a class. This is typically done at the very end of the code. To define a class, the `#Region` code is used together with the OOP Extender keyword "class". A class *must* have both variables (fields) and functions (methods). Inside of the class, everything has real data types (int64, double, ...). Those are denoted using the [wrong](http://www.joelonsoftware.com/articles/Wrong.html) [hungarian notation](https://www.autoitscript.com/wiki/Best_coding_practices#Names_of_Variables) common in AutoIt. A method or field that starts with `i` will be `int64`. `f` is for `double`. Everything else will also be `int64`.

The variable `$This` inside of a method is a reference to the fields of the currently running instance. Technical info: Because methods are the same for all instances, they are only declared once. However, fields are values bound to a single instance, so the self-reference is neccesary.

Here's a class with one field and a function for setting and getting the value of this field:

```au3
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
```

Great, now we have a class! If we'd run the script now, the class would be optimized away and nothing is run. So we need some code. First, let's create a new instance of the class. A class is instantiated using the `<classdef` tag. The position of the classdef tag in the code is the same position the class will be instantiated (makes sense, doesn't it?). The tag has the following grammar:

```au3
# <classdef:ClassName $nVariable[ from %Arguments%]>
```

We don't care about the from... part yet, we just want a new instance. After a new instance is created, we can call the methods of the instance. Let's set and get the value.

So the whole example code would be:

```au3
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
```

### Example 2: Multiple class instances

A single class can be instantiated multiple times. All fields remain bound to the respective instance, so here we create two instances of the same class. Note that methods that do not have arguments can be called without empty parentheses:

```au3
; OOP UDF Example 2: Multiple Instances

; AU3Check will NOT actually be disabled!
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
```

### Example 3: Constructors and Destructors

A class can have a destructor, constructor, none or both. For every class, a method that is declared as `_ClassName` is the constructor and `__ClassName` is the destructor. The constructor is called when the class is instantiated (every single time). The destructor is called either when the user deletes the object himself (by `$oObject = 0`) or when it is cleaned by the garbage collection (which occurs when the script exits). The constructor can also have arguments, but we'll get to that.

Here's the modified class from above with a constructor and destructor:

```au3
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
```

Easy.

### Example 4: Constructor Arguments

It makes sense to pass arguments to the constructor. When instantiating a class you can pass an array (or a plain variable) to the constructor. Attention: Contrary to other methods, the arguments to the constructor **do not have types**. Why? Because there are no types in AutoIt. This makes things easier. Arguments are passed using the "from" part in the classdef tag. By now you should be familiar with the class syntax, so here's a more advanded example (which we'll optimize in the next section):

```au3
; OOP UDF Example 4: Constructor Arguments

#AutoIt3Wrapper_Run_AU3Check=n
#include <Misc.au3>
#include 'OOP.au3'

$hGUI = GUICreate("OOP Example")
Local $hPlayers = [GUICtrlCreateLabel("", 25, 175, 50, 50), GUICtrlCreateLabel("", 325, 175, 50, 50)]
GUICtrlSetBkColor($hPlayers[0], 0xFF0000)
GUICtrlSetBkColor($hPlayers[1], 0xFFCC00)
GUISetState()

; The constructor is called with arguments:
Local $aParameters = [25, 175, $hPlayers[0]]
# <classdef:Player $oPlayer1 from $aParameters>
Local $aParameters = [325, 175, $hPlayers[1]]
# <classdef:Player $oPlayer2 from $aParameters>

While GUIGetMsg() <> -3
	$DiffX = _IsPressed("27") - _IsPressed("25")
	$DiffY = _IsPressed("28") - _IsPressed("26")
	$oPlayer1.Move($DiffX, $DiffY)
	$oPlayer2.Move($DiffX, $DiffY)
WEnd

#Region Class Player
	Local $iX, $iY, $iHandle

	Func _Player($aArgs)
		$This.iX      = $aArgs[0]
		$This.iY      = $aArgs[1]
		$This.iHandle = $aArgs[2]
	EndFunc

	Func Move($iStepX, $iStepY)
		$This.iX += $iStepX
		$This.iY += $iStepY
		GUICtrlSetPos($This.iHandle, $This.iX, $This.iY)
	EndFunc
#EndRegion
```

Run that and press the direction keys. Both "players" move synchronously.

### Example 5: Arrays of Objects

Now, creating objects in an array element couldn't be simple, just use it in the tag:

```au3
# <classdef:SomeClass $aSomeArray[42]>
```

But what if the *whole* array should be objects? Easy, just *declare* and *dimension* your array before you call the extender:

```au3
Local $aPlayers[10]

#include 'OOP.au3'
```

So we can create an optimized version of the "game" from the last section:

```au3
; OOP UDF Example 5: Array of Objects

#AutoIt3Wrapper_Run_AU3Check=n
#include <Misc.au3>

Local $aPlayers[10]

#include 'OOP.au3'

$hGUI = GUICreate("OOP Example")
# <classdef:Player $aPlayers>
GUISetState()

While GUIGetMsg() <> -3
	For $n = 0 To UBound($aPlayers)-1
		$aPlayers[$n].Move(_IsPressed("27") - _IsPressed("25"), _IsPressed("28") - _IsPressed("26"))
	Next
WEnd

#Region Class Player
	Local $iX, $iY, $iHandle

	Func _Player()
		$This.iX      = Random(0, 390, 1)
		$This.iY      = Random(0, 390, 1)
		$This.iHandle = GUICtrlCreateLabel("", $This.iX, $This.iY, 10, 10)
		GUICtrlSetBkColor($This.iHandle, Random(0xAAAAAA, 0xFFFFFF, 1))
	EndFunc

	Func Move($iStepX, $iStepY)
		$This.iX += $iStepX
		$This.iY += $iStepY
		GUICtrlSetPos($This.iHandle, $This.iX, $This.iY)
	EndFunc
#EndRegion
```

### Creating compilable scripts

You don't have to :-). The extender creates a compilable scrip everytime the current script is executed and no error occured. The generated script is placed in the same directory as `ScriptName_stable.au3`.
