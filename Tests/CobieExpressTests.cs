using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Xbim.CobieExpress.Exchanger;
using Xbim.Common;
using Xbim.IO.CobieExpress;
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

        [Fact]
        public void ConvertIfcToCoBieExpress()
        {
            const string input = @"TestFiles\SampleHouse4.ifc";
            var inputInfo = new FileInfo(input);
            
#pragma warning disable CS0618 // Type or member is obsolete  TODO: Needs correct non-obsolete signature in Essentials - defaults to ILogger
            var ifc = IO.Memory.MemoryModel.OpenReadStep21(input);
#pragma warning restore CS0618 // Type or member is obsolete
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
