using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Xbim.IO.Esent;
using System;
using Xbim.IO.CobieExpress;
using Xbim.Ifc;

namespace Xbim.CobieExpress.Tests
{
    [TestClass]
    public class EmptyModelTests
    {

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            IfcStore.ModelProviderFactory.UseHeuristicModelProvider();
        }


        [TestMethod]
        public void EsentInCobieModelTest()
        {

            string file = Guid.NewGuid() + ".xbim";
            try
            {
                
                using (var model = new CobieModel(file))
                {
                    using (var txn = model.BeginTransaction("Creation"))
                    {
                        var wall = model.Instances.New<CobieComponent>(w => w.Name = "Wall A");
                        txn.Commit();
                    }
                }
            }
            finally
            {
                if (File.Exists(file))
                    File.Delete(file);
            }
        }

        [TestMethod]
        public void EsentInOuterScope()
        {
            //creating esent model in outer scope for better control
            using (var esent = EsentModel.CreateModel(new EntityFactoryCobieExpress(), "test2.xbim"))
            {
                using (var model = new CobieModel(esent))
                {
                    using (var txn = model.BeginTransaction("Creation"))
                    {
                        var wall = model.Instances.New<CobieComponent>(w => w.Name = "Wall A");
                        txn.Commit();
                    }
                }

                //we can close Esent model or do anything we need with it
                esent.Close();
            }

            using (var model = CobieModel.OpenEsent("test2.xbim"))
            {
                var wall = model.Instances.FirstOrDefault<CobieComponent>();
                Assert.IsNotNull(wall);
                Assert.IsTrue(wall.Name == "Wall A");
            }
        }

        [TestMethod]
        public void MemoryToEsentCobieModelTest()
        {

            //creating as in-memory model
            using (var model = new CobieModel())
            {
                CreateSimpleModel(model);

                //saving to Esent (will change extention to *.xbim even if you define something else)
                model.SaveAsEsent(esentName);

                //save as step21
                model.SaveAsStep21(stpName);

                //save as step21zip
                model.SaveAsStep21Zip(stpZipName);
            }


            AssertAllModelTypes();

            //delete these files to make sure it starts from empty
            File.Delete(stpName);
            File.Delete(stpZipName);
            File.Delete(esentName);

            string file = Guid.NewGuid() + ".xbim";
            try
            {
                using (var model = new CobieModel(file))
                {
                    CreateSimpleModel(model);

                    //saving to Esent (will change extention to *.xbim even if you define something else)
                    model.SaveAsEsent(esentName);

                    //save as step21
                    model.SaveAsStep21(stpName);

                    //save as step21zip
                    model.SaveAsStep21Zip(stpZipName);
                }

                AssertAllModelTypes();
            }
            finally
            {
                if (File.Exists(file))
                    File.Delete(file);
            }
            
        }

        private const string esentName = "test.xbim";
        private const string stpName = "test.stp";
        private const string stpZipName = "test.stpzip";
        private const string wallName = "Wall A";

        private void CreateSimpleModel(CobieModel model)
        {
            using (var txn = model.BeginTransaction("Creation"))
            {
                var wall = model.Instances.New<CobieComponent>(w => w.Name = wallName);
                txn.Commit();
            }
        }

        private void AssertSimpleModel(CobieModel model)
        {
            var wall = model.Instances.FirstOrDefault<CobieComponent>();
            Assert.IsNotNull(wall);
            Assert.IsTrue(wall.Name == wallName);
        }

        private void AssertAllModelTypes()
        {
            using (var model = CobieModel.OpenEsent(esentName))
            {
                AssertSimpleModel(model);
            }

            //open into memory
            using (var model = CobieModel.OpenStep21Zip(stpZipName))
            {
                AssertSimpleModel(model);
            }

            //open into esent
            using (var model = CobieModel.OpenStep21Zip(stpZipName, true))
            {
                AssertSimpleModel(model);
            }

            //open into memory
            using (var model = CobieModel.OpenStep21Zip(stpZipName))
            {
                AssertSimpleModel(model);
            }

            //open into ESENT
            using (var model = CobieModel.OpenStep21Zip(stpZipName, true))
            {
                AssertSimpleModel(model);
            }
        }
    }
}
