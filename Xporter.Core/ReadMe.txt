============================
# This is Xporter's ReadMe #
============================
__  __ __     ___   ___  _____  ___  ___   
\ \/ /| _ \  / _ \ | _ \|_   _|| __|| _ \  
 >  < |  _/ | (_) ||   /  | |  | _| |   /  
/_/\_\|_|    \___/ |_|_\  |_|  |___||_|_\

"Xporter" allows users to export Spreadsheet files easily from either any object type or 
list of properties as a source, using pre-existing .xlsx files as Templates.

================================================================================================
 __    __   ______   __       __        ________   ______         __    __   ______   ________ 
|  \  |  \ /      \ |  \  _  |  \      |        \ /      \       |  \  |  \ /      \ |        \
| $$  | $$|  $$$$$$\| $$ / \ | $$       \$$$$$$$$|  $$$$$$\      | $$  | $$|  $$$$$$\| $$$$$$$$
| $$__| $$| $$  | $$| $$/  $\| $$         | $$   | $$  | $$      | $$  | $$| $$___\$$| $$__    
| $$    $$| $$  | $$| $$  $$$\ $$         | $$   | $$  | $$      | $$  | $$ \$$    \ | $$  \   
| $$$$$$$$| $$  | $$| $$ $$\$$\$$         | $$   | $$  | $$      | $$  | $$ _\$$$$$$\| $$$$$   
| $$  | $$| $$__/ $$| $$$$  \$$$$         | $$   | $$__/ $$      | $$__/ $$|  \__| $$| $$_____ 
| $$  | $$ \$$    $$| $$$    \$$$         | $$    \$$    $$       \$$    $$ \$$    $$| $$     \
 \$$   \$$  \$$$$$$  \$$      \$$          \$$     \$$$$$$         \$$$$$$   \$$$$$$  \$$$$$$$$
================================================================================================

To use this library, add a using statement for Xporter { using Xporter; }

Now you can call the base static class named " Xport. "

and on that class call one of three methods:

	* Xport.LoadFromFileInfo()
	* Xport.LoadFromFileStream()
	* Xport.CreateNewPackage()

Then you can use the library extensions methods listed below as you like:

	* .LoadTempl()						--> to load a template from another xlsx file to the current one
	* .InsertData()		+3 overloads	--> to insert any kind of data from a model or other source
	* .WriteToCells()					--> to write something in specified cells
	* .Clear()			+2 overloads	--> to clear the data of a worksheet or clear all sheets of a file
