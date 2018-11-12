# Xbim COBie Express
Part of Xbim; the eXtensible [Building Information Modelling (BIM) Toolkit](https://xbimteam.github.io/)

This code was originally part of xBIM Essentials but was moved out to make Essentials smaller and more clean.

COBie Express is our attempt to support COBie in a way ahich is a lot easier to maintain compared to spreadsheets. It is possible to export data into spreadsheets
and it is also possible to load the data from them but you should be aware that spreadsheets are not reliable source of information and that design of 
COBie spreadsheets makes it very hard to maintain integrity of the data using several bad practises of database design in its core.

COBie Express is modeled using EXPRESS modelling language (the same as IFC) and the implementation is generated using the same tooling as we use for IFC.
As a result you can use all advanced data processing features of xBIM to work with the data. 

## What is xBIM Toolkit?

The xBIM Tookit (eXtensible Building Information Modelling) is an open-source, software development BIM toolkit that 
supports the BuildingSmart Data Model (aka the [Industry Foundation Classes IFC](http://en.wikipedia.org/wiki/Industry_Foundation_Classes)).

xBIM allows developers to read, create and view [Building Information (BIM)](http://en.wikipedia.org/wiki/Building_information_modeling) Models in the IFC format. 
There is full support for geometric, topological operations and visualisation. In addition xBIM supports 
bi-directional translation between IFC and COBie formats

## Getting Started

You will need Visual Studio 2017 or newer to compile the Solution. All solutions target .NET 4.7 and .NET Standard 2.0 where possible


## Licence

The XBIM library is made available under the CDDL Open Source licence.  See the licences folder for a full text.

All licences should support the commercial usage of the XBIM system within a 'Larger Work', as long as you honour 
the licence agreements.
