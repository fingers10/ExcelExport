[![NuGet Badge](https://buildstats.info/nuget/fingers10.excelexport)](https://www.nuget.org/packages/fingers10.excelexport/)
[![Open Source Love svg1](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)](https://github.com/fingers10/open-source-badges/)

[![GitHub license](https://img.shields.io/github/license/fingers10/ExcelExport.svg)](https://github.com/fingers10/ExcelExport/blob/master/LICENSE)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://GitHub.com/fingers10/ExcelExport/graphs/commit-activity)
[![Ask Me Anything !](https://img.shields.io/badge/Ask%20me-anything-1abc9c.svg)](https://GitHub.com/fingers10/ExcelExport)
[![HitCount](http://hits.dwyl.io/fingers10/badges.svg)](http://hits.dwyl.io/fingers10/badges)

[![GitHub forks](https://img.shields.io/github/forks/fingers10/ExcelExport.svg?style=social&label=Fork)](https://GitHub.com/fingers10/ExcelExport/network/)
[![GitHub stars](https://img.shields.io/github/stars/fingers10/ExcelExport.svg?style=social&label=Star)](https://GitHub.com/fingers10/ExcelExport/stargazers/)
[![GitHub followers](https://img.shields.io/github/followers/fingers10.svg?style=social&label=Follow)](https://github.com/fingers10?tab=followers)

[![GitHub contributors](https://img.shields.io/github/contributors/fingers10/ExcelExport.svg)](https://GitHub.com/fingers10/ExcelExport/graphs/contributors/)
[![GitHub issues](https://img.shields.io/github/issues/fingers10/ExcelExport.svg)](https://GitHub.com/fingers10/ExcelExport/issues/)
[![GitHub issues-closed](https://img.shields.io/github/issues-closed/fingers10/ExcelExport.svg)](https://GitHub.com/fingers10/ExcelExport/issues?q=is%3Aissue+is%3Aclosed)

# Excel Export
Simple classes to generate Excel/CSV Report in ASP.NET Core.

To export/download the `IEnumerable<T>` data as an excel file, add action method in your controller as shown below. Return the data as `ExcelResult<T>`/`CSVResult<T>` by passing filtered/ordered data, sheet name and file name. **ExcelResult**/**CSVResult** Action Result that I have added in the Nuget package. This will take care of converting your data as excel/csv file and return it back to browser.

>**Note: This tutorial contains example for downloading/exporting excel/csv from Asp.Net Core Backend.**

# Give a Star ⭐️
If you liked `ExcelExport` project or if it helped you, please give a star ⭐️ for this repository. That will not only help strengthen our .NET community but also improve development skills for .NET developers in around the world. Thank you very much 👍

## Columns 
### Name
Column names in Excel Export can be configured using the below attributes
* `[Display(Name = "")]`
* `[Display(Name = "", ResourceType = typeof(MyResourceFile))]`
* `[DisplayName(“”)]`

### Report Setup
To customize the Excel Column display in your report, use the following attribute
* `[IncludeInReport]`
* `[NestedIncludeInReport]`

And here are the options,

|Option |Type    |Example                                  |Description|
|-------|--------|-----------------------------------------|-----------|
|Order  |`int`   |`[IncludeInReport(Order = N)]`           |To control the order of columns in Excel Report|

**Please note:** From **v.2.0.0**, simple properties in your models **with `[IncludeInReport]` attribute** will be displayed in excel report. You can add any level of nesting to your models using **`[NestedIncludeInReport]` attribute**.

# NuGet:
* [Fingers10.ExcelExport](https://www.nuget.org/packages/Fingers10.ExcelExport/) **v3.0.0**

## Package Manager:
```c#
PM> Install-Package Fingers10.ExcelExport
```

## .NET CLI:
```c#
> dotnet add package Fingers10.ExcelExport
```

## Root Model

```c#
public class DemoExcel
{
    public int Id { get; set; }

    [IncludeInReport(Order = 1)]
    public string Name { get; set; }

    [IncludeInReport(Order = 2)]
    public string Position { get; set; }

    [Display(Name = "Office")]
    [IncludeInReport(Order = 3)]
    public string Offices { get; set; }

    [NestedIncludeInReport]
    public DemoNestedLevelOne DemoNestedLevelOne { get; set; }
}
```

## Nested Level One Model:

```c#
public class DemoNestedLevelOne
{
    [IncludeInReport(Order = 4)]
    public short? Experience { get; set; }

    [DisplayName("Extn")]
    [IncludeInReport(Order = 5)]
    public int? Extension { get; set; }

    [NestedIncludeInReport]
    public DemoNestedLevelTwo DemoNestedLevelTwos { get; set; }
}
```

## Nested Level Two Model:

```c#
public class DemoNestedLevelTwo
{
    [DisplayName("Start Date")]
    [IncludeInReport(Order = 6)]
    public DateTime? StartDates { get; set; }

    [IncludeInReport(Order = 7)]
    public long? Salary { get; set; }
}
```

## Action Method
 
```c#
public async Task<IActionResult> GetExcel()
{
    // Get you IEnumerable<T> data
    var results = await _demoService.GetDataAsync();
    return new ExcelResult<DemoExcel>(results, "Demo Sheet Name", "Fingers10");
}

public async Task<IActionResult> GetCSV()
{
    // Get you IEnumerable<T> data
    var results = await _demoService.GetDataAsync();
    return new CSVResult<DemoExcel>(results, "Fingers10");
}
```
 
## Page Handler
 
```c#
public async Task<IActionResult> OnGetExcelAsync()
{
     // Get you IEnumerable<T> data
    var results = await _demoService.GetDataAsync();
    return new ExcelResult<DemoExcel>(results, "Demo Sheet Name", "Fingers10");
}

public async Task<IActionResult> OnGetCSVAsync()
{
     // Get you IEnumerable<T> data
    var results = await _demoService.GetDataAsync();
    return new CSVResult<DemoExcel>(results, "Fingers10");
}
```

# Target Platform
* .Net Standard 2.0
 
# Tools Used
* Visual Studio Community 2019
 
# Other Nuget Packages Used
* ClosedXML (0.95.0) - For Generating Excel Bytes
* Microsoft.AspNetCore.Mvc.Abstractions (2.2.0) - For using IActionResult
* System.ComponentModel.Annotations (4.7.0) - For Reading Column Names from Annotations
 
# Author
* **Abdul Rahman** - Software Developer - from India. Software Consultant, Architect, Freelance Lecturer/Developer and Web Geek.  
 
# Contributions
Feel free to submit a pull request if you can add additional functionality or find any bugs (to see a list of active issues), visit the  Issues section. Please make sure all commits are properly documented.
  
# License
ExcelExport is release under the MIT license. You are free to use, modify and distribute this software, as long as the copyright header is left intact.

Enjoy!

# Sponsors/Backers
I'm happy to help you with my Nuget Package. Support this project by becoming a sponsor/backer. Your logo will show up here with a link to your website. Sponsor/Back via  [![Sponsor via PayPal](https://www.paypalobjects.com/webstatic/mktg/Logo/pp-logo-100px.png)](https://paypal.me/arsmtb)
