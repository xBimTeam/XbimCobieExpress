using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.CobieExpress;
using Xbim.Ifc;

namespace Xbim.IO.CobieExpress
{
    [Obsolete("ModelProviders are now created via the XbimServices.Current.ServiceProvider")]
    public class COBieModelProviderFactory : IModelProviderFactory
    {
        public IModelProvider CreateProvider()
        {
            throw new NotImplementedException();
        }

        public void Use(Func<IModelProvider> providerFn)
        {
            throw new NotImplementedException();
        }
    }
}
