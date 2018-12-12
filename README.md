

Branch | Status
------ | -------
Master | [![Build Status](https://dev.azure.com/xBIMTeam/xBIMToolkit/_apis/build/status/xBimTeam.XbimCobieExpress?branchName=master)](https://dev.azure.com/xBIMTeam/xBIMToolkit/_build/latest?definitionId=2&branchName=master)
Develop | [![Build Status](https://dev.azure.com/xBIMTeam/xBIMToolkit/_apis/build/status/xBimTeam.XbimCobieExpress?branchName=develop)](https://dev.azure.com/xBIMTeam/xBIMToolkit/_build/latest?definitionId=2&branchName=develop)


# Xbim COBie Express
Part of Xbim; the eXtensible [Building Information Modelling (BIM) Toolkit](https://xbimteam.github.io/)

This code was originally part of xBIM Essentials but was moved out to make Essentials smaller and more focussed.

COBie Express is our attempt to support COBie in a way that is a lot easier to maintain compared to spreadsheets. 

This library enables you to both read and write spreadsheets adhering to the COBie Schema (MVD), 
while also providing the power of the XBIM toolkit to query, interrogate and build data transactionally.

COBie Express is modeled using EXPRESS modelling language (the same as IFC) and the implementation is 
generated using the same tooling as we use for IFC. 
As a result you can use all advanced data processing features of xBIM to work with the data. 

## Code Examples

### 1 Reading from an Excel COBie spreadsheet

If you're familiar with *LINQ* or have ever queried IFC models with XBIM, you'll be right at home querying
COBie data sources. 

```csharp

    using (IModel model = CobieModel.ImportFromTable("MyCobieSpreadsheet.xlsx", out string report))
    {
        // Get all the contacts
        var contacts = model.Instances.OfType<CobieContact>().ToList();
        
        // Query the spaces but filter by a Uniclass2015 category. And then my the room area
        var largeCommercialSpaces = model.Instances.OfType<CobieSpace>()
            .Where(space => space.Categories.Any(cat => cat.Classification.Name.StartsWith("SL_20_50")))
            .Where(space => space.GrossArea > 1000);
        
        // We can then drill across to other parts of the model
        var occupancyRatingOfFirstSpaces = largeCommercialSpaces.FirstOrDefault().Attributes
            .Where(attr => attr.Name.StartsWith("Occupancy"));
    }
```

### 2. Converting IFCs to COBIe spreadheets

This is a more sophisticated example where we convert an IFC to COBie. Here we're using some 
built-in mappings, but these can all be over-ridden in the `IfcToCoBieExpressExchanger`
constructor. See `OutPutFilters` and [CobieAttributes.config](Xbim.CobieExpress.Exchanger/IfcToCOBieExpress/CobieAttributes.config)

```csharp
    const string input = @"SampleHouse4.ifc";

    var ifc = MemoryModel.OpenReadStep21(input);

    var cobie = new CobieModel();
    using (var txn = cobie.BeginTransaction("Sample house conversion"))
    {
        var exchanger = new IfcToCoBieExpressExchanger(ifc, cobie
            /*,     // More advanced configuration options available
            reportProgress: reportProgressDelegate,
            filter: outputFilters,
            configFile: pathToAttributeMappingConfigFile,
            extId: EntityIdentifierMode.GloballyUniqueIds,
            sysMode: SystemExtractionMode.System,
            classify: true*/
            );
        exchanger.Convert();
        txn.Commit();
    }

    // We can persists our model to disk for faster access in future
    var output = Path.ChangeExtension(input, ".cobie");
    cobie.SaveAsEsent(output);

    // We can fix up the data to make it valid - e.g. Deduplicate some names
    using (var txn = cobie.BeginTransaction("Make some changes"))
    {
        MakeUniqueNames<CobieSpace>(cobie);
        MakeUniqueNames<CobieType>(cobie);
        txn.Commit();
    }

    // Finally export as a COBie spreadsheet
    output = Path.ChangeExtension(input, ".xlsx");
    cobie.ExportToTable(output, out string report);
```

### Using the library

To get started, the simplest approach is to add the `Xbim.COBieExpress.Exchanger` and `Xbim.COBieExpress.IO` 
nuget packages to your Visual Studio Project from Nuget or get the latest versions from our [MyGet feeds](nuget.config)

Alternatively you can add the packages using Nuget's Package Manager Console and issuing the following command:

```
PM> Install-Package Xbim.COBieExpress.Exchanger
PM> Install-Package Xbim.COBieExpress.IO
```



## Building yourself

You will need Visual Studio 2017 or newer to compile the Solution. 
Prior versions of Visual Studio may work, but we'd recommend 2017 where possible.
The [free VS 2017 Community Edition](https://visualstudio.microsoft.com/downloads/) should work fine. 
All projects target .NET Framework *net47*, as well as *netstandard2.0*, which should 
permit limited trials of XBIM with .NET Core / Mono etc.


## Licence

The XBIM library is made available under the CDDL Open Source licence.  See the licences folder for a full text.

All licences should support the commercial usage of the XBIM system within a 'Larger Work', as long as you honour 
the licence agreements.
