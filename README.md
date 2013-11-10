# MacroMapper: Visual Studio 2012 Extensibility project written in CSharp

Visual Studio 2012 no longer has the Macro functionality that was available in previous versions.

This project creates a loadable Add-in that maps Macro functionality so that it can be easily bound to the TOOLS toolbar or Keyboard shortcuts.

As an example, the Macros from Visual Studio 2010 for turning line numbers on and off were mapped in this Add-in as Commands and added to the TOOLS toolbar:

Visual Studio 2010 Macro                       ->      Visual Studio 2012 Add-in Command
--------------------------------------------           ---------------------------------
Macros.Samples.Utilities.TurnOnLineNumbers     ->      MacroMapper.Connect.TurnOnLineNumbers
Macros.Samples.Utilities.TurnOffLineNumbers    ->      MacroMapper.Connect.TurnOffLineNumbers

To get this functionality without building the project, close all like Step 3 below then copy the 2 files from the INSTALL directory to the location in Step 4 below.  Then Steps 5 and 6 are in play.


To add more Macros as Add-in Commands:

1. Get the definition via Visual Studio 2010 - Tools - Macros - Macro Explorer, browse and right-click - Edit to bring up the definition (in Visual Basic).

2. Create command by adding 3 things in Connect.cs file: Create new command and optionally add to toolbar in OnConnection(), set new command to supported and enabled in QueryStatus(), and finally port the definition from Step 1 to DTE2 and add for the new command in Exec().

3. Build and then close all Visual Studio 2012 instances.

4. Copy 2 files into Addins where have Visual Studio 2012 installed, which is probably like C:\Users\YOUR LOGIN\Documents\Visual Studio 2012\Addins: MacroMapper\MacroMapper.AddIn and MacroMapper\bin\MacroMapper.dll

For this project:

5  Now when you open Visual Studio 2012, you will see 2 (happy!) commands under TOOLS toolbar, TurnOnLineNumbers and TurnOffLineNumbers.

6  To get really interactive, map these commands to Keyboard shortcuts via TOOLS - Options - Environment - Keyboard - Show commands containing: MacroMapper.Connect.  Go ahead and map those 2 commands to appropriate keyboard shortcuts and you can now toggle line numbers in files on and off quickly.






