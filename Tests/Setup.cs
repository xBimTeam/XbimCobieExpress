using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xbim.IO.CobieExpress;
using Xunit;

namespace Xbim.CobieExpress.Tests
{

    // Initialise the XbimServices early in unit tests - regardless of whether MsTest or xUnit tests run first.
    // We could leave the initialisation to occur implicitly in COBieModel static ctor, but this would
    // introduce subtle timing issues in the singleton XbimServices if for example, IfcStore is accessed before COBieModel in the initial test.

    [TestClass]
    public class MsTestBootstrap
    {
        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            // Trigger initialisation
            _ = new CobieModel();
        }
    }


    [CollectionDefinition(nameof(xUnitBootstrap))]
    public class xUnitBootstrap : ICollectionFixture<xUnitSetup>
    {
        // Does nothing but trigger xUnitSetup construction at beginning of test run
    }

    public class xUnitSetup 
    {
        public xUnitSetup()
        {
            // Trigger initialisation
            _ = new CobieModel();
        }
    }
}
