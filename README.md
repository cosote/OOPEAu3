# AutoIt OOP Extender

Since the introduction of ObjCreateInterface, AutoIt is able to natively create real objects. However, this is quite difficult for non-experienced users. This UDF here enables a "classic" syntax for declaring classes and instantiating objects from them. Before you ask: **NO**, it is *not* just another preprocessor that's faking OOP syntax - this is the real deal.

Care has been put into this UDF and it is in the authors interest to fix all bugs remaining and implement features as long as the result works in the **Stable** release of AutoIt.

### Features

- Define an unlimited number of classes.
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
- Use `C`-style macros

## [&raquo; Visit the Wiki!](https://github.com/minxomat/AutoIt-OOP-Extender/wiki)
