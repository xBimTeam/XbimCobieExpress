using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xbim.CobieExpress;
using Xbim.CobieExpress.Exchanger;
using Xbim.Common;
using Xbim.Ifc;
using Xbim.IO.CobieExpress;
using Xbim.IO.Memory;


namespace Tests
{
    [TestClass]
    public class CobieExpressTests
    {
        [TestMethod]
        [DeploymentItem("TestFiles")]
        public void ConvertIfcToCoBieExpress()
        {
            const string input = @"SampleHouse4.ifc";
            var inputInfo = new FileInfo(input);
            var ifc = MemoryModel.OpenReadStep21(input);
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
            Console.WriteLine("Time to convert {0:N}MB file ({2} entities): {1}ms", inputInfo.Length/1e6f, w.ElapsedMilliseconds, inputCount);
            Console.WriteLine("Resulting size: {0:N}MB ({1} entities)", outputInfo.Length / 1e6f, cobie.Instances.Count);

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
