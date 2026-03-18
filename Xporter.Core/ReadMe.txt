━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# This is Xporter's ReadMe #
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 /$$   /$$ /$$$$$$$   /$$$$$$  /$$$$$$$  /$$$$$$$$ /$$$$$$$$ /$$$$$$$ 
| $$  / $$| $$__  $$ /$$__  $$| $$__  $$|__  $$__/| $$_____/| $$__  $$
|  $$/ $$/| $$  \ $$| $$  \ $$| $$  \ $$   | $$   | $$      | $$  \ $$
 \  $$$$/ | $$$$$$$ | $$  | $$| $$$$$$$    | $$   | $$$$$   | $$$$$$$
   $$  $$ | $$____/ | $$  | $$| $$__  $$   | $$   | $$__/   | $$__  $$
 /$$/\  $$| $$      | $$  | $$| $$  \ $$   | $$   | $$      | $$  \ $$
| $$  \ $$| $$      |  $$$$$$/| $$  | $$   | $$   | $$$$$$$$| $$  | $$
|__/  |__/|__/       \______/ |__/  |__/   |__/   |________/|__/  |__/     
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        

"Xporter" allows users to export Spreadsheet files easily from either any object type or 
list of properties as a source, using pre-existing .xlsx files as Templates.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
╷ ╷╭─╮╷ ╷   ╶┬╴╭─╮   ╷ ╷╭─╮╭─╴
├─┤│ ││╷│    │ │ │   │ │╰─╮├╴ 
╵ ╵╰─╯╰┴╯    ╵ ╰─╯   ╰─╯╰─╯╰─╴
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

To use this library, add a using statement for Xporter { using Xporter; }

Now you can call the base static class named " Xport. "

and on that class call one of three methods:

	* Xport.LoadFromFileInfo()
	* Xport.LoadFromFileStream()
	* Xport.CreateNewPackage()

Then you can use the library extensions methods listed below as you like:

	* .Clear()			+1 overload	--> to clear the data of a worksheet or clear all sheets of a file
	* .LoadTempl()		+1 overload	--> to load a template from another xlsx file to the current one
	* .InsertData()		+2 overloads	--> to insert any kind of data from a model all other source
	* .WriteToCells()	+1 overload	--> to write something in specified cells
	* .InsertToCells() 	+1 overload	--> to replace all cells containing a specific string with another string


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
╭─╮╷╭┬╮╭─╮╷  ╭─╴   ╭─╴╭─╮╶┬╮╭─╴   ╭─╴╷ ╷╭─╮╭┬╮╭─╮╷  ╭─╴
╰─╮││││├─╯│  ├╴    │  │ │ ││├╴    ├╴ ╭┼╯├─┤│││├─╯│  ├╴ 
╰─╯╵╵ ╵╵  ╰─╴╰─╴   ╰─╴╰─╯╶┴╯╰─╴   ╰─╴╵ ╵╵ ╵╵ ╵╵  ╰─╴╰─╴
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  C# code:

	using Xporter;

		var cellProps = new CellProperties();
			cellProps.Add("E2", "Stats");
			cellProps.Add("E3", "TypeOfProduct");
			cellProps.Add("E4", "Images");
			cellProps.Add("I2", DateTime.Now.ToString("yyyy-MM-dd"));

		Xport.LoadFromFileInfo(new FileInfo("C:\\Users\\{YourPcName}\\Desktop\\MyFile.xlsx"))
			.Clear()
			.LoadTempl(new FileInfo("C:\\Users\\{YourPcName}\\Desktop\\TemplateFile.xlsx"))
			.InsertData(yourListOfAnyType, startingRow, startingCol)
			.WriteToCells(cellProps)
			.InsertToCells("NAME", "John")    //Replace all cells containing "NAME" with "John"
			.Save();

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
It's that simple! Awesome, right? ♥︎
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━