using FluentAssertions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Xbim.CobieExpress.Exchanger;
using Xbim.Common;
using Xbim.Ifc4.Interfaces;
using Xbim.IO.CobieExpress;
using Xbim.IO.Memory;
using Xunit;
using Xunit.Abstractions;


namespace Xbim.CobieExpress.Tests
{
    public class CobieExpressTests
    {
        private readonly ITestOutputHelper console;

        public CobieExpressTests(ITestOutputHelper output)
        {
            this.console = output;
        }

        [Theory]
        [InlineData(@"TestFiles\SampleHouse4.ifc")]
        public void ConvertIfcToCoBieExpress(string input)
        {
            var inputInfo = new FileInfo(input);
            
            var ifc = IO.Memory.MemoryModel.OpenReadStep21(input);
            var inputCount = ifc.Instances.Count;

            var w = new Stopwatch();
            var cobie = new CobieModel();
            using (var txn = cobie.BeginTransaction("Sample house conversion"))
            {
                var exchanger = new IfcToCoBieExpressExchanger(ifc, cobie);
                w.Start();
                exchanger.Convert();
                w.Stop();
                txn.Commit();
            }
            var output = Path.ChangeExtension(input, ".cobie");
            cobie.SaveAsStep21(output);

            var outputInfo = new FileInfo(output);
            console.WriteLine ("Time to convert {0:N}MB file ({2} entities): {1}ms", inputInfo.Length/1e6f, w.ElapsedMilliseconds, inputCount);
            console.WriteLine("Resulting size: {0:N}MB ({1} entities)", outputInfo.Length / 1e6f, cobie.Instances.Count);

            using (var txn = cobie.BeginTransaction("Renaming"))
            {
                MakeUniqueNames<CobieFacility>(cobie);
                MakeUniqueNames<CobieFloor>(cobie);
                MakeUniqueNames<CobieSpace>(cobie);
                MakeUniqueNames<CobieZone>(cobie);
                MakeUniqueNames<CobieComponent>(cobie);
                MakeUniqueNames<CobieSystem>(cobie);
                MakeUniqueNames<CobieType>(cobie);
                txn.Commit();
            }

            //save as XLSX
            output = Path.ChangeExtension(input, ".xlsx");
            cobie.ExportToTable(output, out string report);
        }

        [Theory]
        [InlineData(@"TestFiles\SampleHouse4.ifc")]
        public void ConvertIfcToCoBieExpressWithConfig(string input)
        {
            var inputInfo = new FileInfo(input);

            var ifc = IO.Memory.MemoryModel.OpenReadStep21(input);
            var inputCount = ifc.Instances.Count;

            var w = new Stopwatch();
            var cobie = new CobieModel();
            using (var txn = cobie.BeginTransaction("Sample house conversion"))
            {
                var exchanger = new IfcToCoBieExpressExchanger(default);
                var configuration = new IfcToCOBieExchangeConfiguration
                {
                     
                };
                exchanger.Initialise(configuration, ifc, cobie);
                w.Start();
                exchanger.Convert();
                w.Stop();
                txn.Commit();
            }
            var output = Path.ChangeExtension(input, ".cobie");
            cobie.SaveAsStep21(output);

            var outputInfo = new FileInfo(output);
            console.WriteLine("Time to convert {0:N}MB file ({2} entities): {1}ms", inputInfo.Length / 1e6f, w.ElapsedMilliseconds, inputCount);
            console.WriteLine("Resulting size: {0:N}MB ({1} entities)", outputInfo.Length / 1e6f, cobie.Instances.Count);

           
            //save as XLSX
            output = Path.ChangeExtension(input +"-output", ".xlsx");
            cobie.ExportToTable(output, out string report);
        }


        [Fact]
        public void CanDedupeTypes()
        {
            MemoryModel ifcModel = LoadIfc(@"TestFiles\Primary_School.ifc");

            var dict = new Dictionary<int, IIfcTypeObject>();
            var types = ifcModel.Instances.OfType<Ifc4.Interfaces.IIfcTypeObject>();

            int duplicates = 0;

            foreach (var typeObject in types)
            {
                var hashCode = typeObject.CalculateHash();

                if(!dict.ContainsKey(hashCode))
                {
                    dict.Add(hashCode, typeObject);
                }
                else
                {
                    console.WriteLine("Object {0} is duplicate of {1}", typeObject, dict[hashCode]);
                    duplicates++;
                }
                
            }
            dict.Count.Should().Be(215);
            duplicates.Should().Be(21);
            types.Count().Should().Be(236);
        }


        [Fact]
        public void BaselineConfig()
        {
            MemoryModel ifcModel = LoadIfc(@"TestFiles\Primary_School.ifc");
            using var cobieModel = new CobieModel();
            using var txn = cobieModel.BeginTransaction("Sample house conversion");
            var exchanger = new IfcToCoBieExpressExchanger(default);
            // Act
            var configuration = new IfcToCOBieExchangeConfiguration
            {
                // Take defaults
            };

            exchanger.Initialise(configuration, ifcModel, cobieModel);
            exchanger.Convert();
            txn.Commit();

            var facility = cobieModel.Instances.OfType<CobieFacility>().FirstOrDefault();

            facility.ExternalId.Should().HaveLength(22, "ifcGuid default");

            cobieModel.Instances.OfType<CobieFloor>().Should().HaveCount(3);
            cobieModel.Instances.OfType<CobieFloor>().Should().AllSatisfy(f => 
            {
                f.Name.Should().StartWith("Level");
                f.CreatedBy.Name.Should().Be("info@xbim.net");
                f.CreatedOn.Value.Value.Should().Be("2017-07-02T06:24:42");
                f.Categories.Should().HaveCount(1).And.Satisfy(s => s.Value == "Floor" || s.Value == "Roof");
                f.ExternalId.Should().HaveLength(22);
                f.Description.Should().NotBeNullOrEmpty();
                f.Elevation.Should().BeInRange(0, 7300);
                f.Height.Should().BeInRange(1400, 4000);
            });

            cobieModel.Instances.OfType<CobieSpace>().Should().HaveCount(52);
            cobieModel.Instances.OfType<CobieSpace>().Where(s => s.ExternalObject.Name != "IfcBuildingStorey").Should().AllSatisfy(f =>
            {
                f.Name.Should().NotBeNullOrEmpty();
                f.CreatedBy.Name.Should().Be("info@xbim.net");
                f.CreatedOn.Value.Value.Should().Be("2017-07-02T06:24:42");
                f.Categories.Count.Should().BeInRange(1, 2);
                f.Categories.Should().AllSatisfy(s =>
                {
                    if (s.Classification.Name.Contains("Uniclass"))
                        s.Value.Should().StartWith("SL_");
                    else
                        s.Value.Should().HaveLength(5); // ADS
                });
                f.ExternalObject.Name.Should().Be("IfcSpace");
                f.RoomTag.Should().NotBeNullOrEmpty();
                f.UsableHeight.Should().BeGreaterThan(0);
                f.GrossArea.Should().BeGreaterThan(0);
                f.NetArea.Should().BeGreaterThan(0);

            });


            cobieModel.Instances.OfType<CobieType>().Where(t => t.ExternalObject.Name == "IfcWallType").Should().BeEmpty("Not maintainable filtered out");
            cobieModel.Instances.OfType<CobieType>().Where(t => t.ExternalObject.Name == "IfcFurnitureType").Should().NotBeEmpty("Furnishings included");
            cobieModel.Instances.OfType<CobieType>().Should().HaveCount(135);
            cobieModel.Instances.OfType<CobieType>().Where(s => s.ExternalObject.Name != "IfcBuildingElementProxyType" && s.ExternalObject.Name != "IfcCoveringType").Should().AllSatisfy(t =>
            {
                t.Description.Should().NotBeNullOrEmpty();
            });

            cobieModel.Instances.OfType<CobieComponent>().Where(t => t.ExternalObject.Name == "IfcFlowTerminal").Should().NotBeEmpty("Electricals included");

            cobieModel.Instances.OfType<CobieSystem>().Should().HaveCount(0);
            cobieModel.Instances.OfType<CobieZone>().Should().HaveCount(6);

            //var output = Path.ChangeExtension("Primary_School-output2", ".xlsx");
            //cobieModel.ExportToTable(output, out string report);
        }

        [Fact]
        public void Can_Configure_ExternalIds()
        {
            MemoryModel ifcModel = LoadIfc(@"TestFiles\Primary_School.ifc");
            using var cobieModel = new CobieModel();
            using var txn = cobieModel.BeginTransaction("Sample house conversion");
            var exchanger = new IfcToCoBieExpressExchanger(default);
            // Act
            var configuration = new IfcToCOBieExchangeConfiguration
            {
                ExternalIdentifierSource = EntityIdentifierMode.IfcEntityLabels
            };

            exchanger.Initialise(configuration, ifcModel, cobieModel);
            exchanger.Convert();
            txn.Commit();

            var facility = cobieModel.Instances.OfType<CobieFacility>().FirstOrDefault();

            facility.ExternalId.Length.Should().BeInRange(1, 8);
        }


        private static MemoryModel LoadIfc(string input)
        {
            var inputInfo = new FileInfo(input);
            var ifc = IO.Memory.MemoryModel.OpenReadStep21(input);
            return ifc;
        }

        private static void MakeUniqueNames<T>(IModel model) where T : CobieAsset
        {
            var groups = model.Instances.OfType<T>().GroupBy(a => a.Name);
            foreach (var @group in groups)
            {
                if (group.Count() == 1)
                {
                    var item = group.First();
                    if (string.IsNullOrEmpty(item.Name))
                        item.Name = item.ExternalObject.Name;
                    continue;
                }

                var counter = 1;
                foreach (var item in group)
                {
                    if (string.IsNullOrEmpty(item.Name))
                        item.Name = item.ExternalObject.Name;
                    item.Name = string.Format("{0} ({1})", item.Name, counter++);
                }
            }
        }
    }
}
