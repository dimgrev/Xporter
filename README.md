# Xporter
![CI](https://raw.githubusercontent.com/dimgrev/Xporter/refs/heads/main/Xporter.Core/XporterIcon164.png) 

"Xporter" allows users to export Spreadsheet files easily from either any object type or list of properties as a source, using also if they want pre-existing .xlsx files as Templates.

## Contents
[The problem](#The-problem)

[Installation](#Instalation)

[How to use](#Usage)

[ToDo](#ToDo)

[License](#License)

## The problem
Imagine having different types of data and you want to export them easily in an xlsx file.
You may need this at your work, to present any kind of statistics.. Its necessary to have 
a service that exports any kind of data that you will provide to it.

Like the method below:

```C#
public static void InsertData(List<object> objects)
{
}
```

OR

```C#
public static void InsertData(List<AnyTypeHere> listOfAnyType)
{
}
```

## Instalation (4 ways)
##### (1) [Using NuGet]
Search into the NuGet (prerelease) packages the library or run the following command:

PM> Install-Package Xporter -Version $(AssemblyVersion)

##### (2) [Manual]
* Download this repository: <a href="https://github.com/dimgrev/Xporter/archive/main.zip" target="_blank">here</a>
* Unzip downloaded file
* Copy the resulting folder to `app/Plugin`
* Rename the folder you copied to utilityXporter

##### (3) [GIT Submodule]
In your app directory type:
```bash
  git submodule add -b master git://github.com/dimgrev/Xporter.git 
Plugin/utilityXporter
  git submodule init
  git submodule update
```

##### (4) [GIT Clone]
In your `Plugin` directory type:
```bash
  git clone -b master git://github.com/dimgrev/Xporter.git 
UtilityXporter
```

## Usage
To use this library, add a using statement for Xporter {using Xporter;}

Now you can call the base static class named "Xport."

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

## Simple Example

	```C#
		using Xporter;

		var cellProps = new CellProperties();
			cellProps.Add("E2", "Stats");
			cellProps.Add("E3", "TypeOfProduct");
			cellProps.Add("E4", "Images");
			cellProps.Add("I2", DateTime.Now.ToString("yyyy-MM-dd"));

		Xport.LoadFromFileInfo(new FileInfo("C:\\Users\\YourName\\Desktop\\MyFile.xlsx"))
			.Clear()
			.LoadTempl(new FileInfo("C:\\Users\\YourName\\Desktop\\TemplateFile.xlsx"))
			.InsertData(yourListOfAnyType)
			.WriteToCells(cellProps)
			.InsertToCells("NAME", "John")    //Replace all cells containing "NAME" with "John"
			.Save();
	```

It's that simple! Awesome, right?

## ToDo
- Maybe the ability to modify xlsx file's style

## License

This project is licensed under the MIT License - see the LICENSE file for details.