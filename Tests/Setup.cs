using Xbim.IO.CobieExpress;
using Xunit;

namespace Xbim.CobieExpress.Tests
{

    // Initialise the XbimServices early in unit tests
    // We could leave the initialisation to occur implicitly in COBieModel static ctor, but this would
    // introduce subtle timing issues in the singleton XbimServices if for example, IfcStore is accessed before COBieModel in the initial test.



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
