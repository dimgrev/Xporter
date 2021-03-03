# Xporter
![CI](https://github.com/dimgrev/Xporter/actions/workflows/ci.yml/badge.svg)

"Xporter" allows users to export Spreadsheet files easily from either any object type or list of properties as a source, using pre-existing .xlsx files as Templates.

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

Like the method above:

```C#
public static void DataToXport(List<object> objects)
{
}
```

## Instalation
##### [Using NuGet]
-- not yet

-- be patient..

-- it will be something like this:

Search into the NuGet packages the library or run the following command:

PM> Install-Package Xporter.Core -Version 0.1.0-alpha

##### [Manual]
* Download this repository: <a href="https://github.com/dimgrev/Xporter/archive/main.zip" target="_blank">here</a>
* Unzip downloaded file
* Copy the resulting folder to `app/Plugin`
* Rename the folder you copied to utilityXporter

##### [GIT Submodule]
In your app directory type:
```bash
  git submodule add -b master git://github.com/dimgrev/Xporter.git 
Plugin/utilityXporter
  git submodule init
  git submodule update
```

##### [GIT Clone]
In your `Plugin` directory type:
```bash
  git clone -b master git://github.com/dimgrev/Xporter.git 
UtilityXporter
```

## Usage
-- not yet

-- be patient..

## ToDo
-- not yet

-- be patient..

## License

The MIT License (MIT)

Copyright ©2021 Dimitris Grevenos

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.