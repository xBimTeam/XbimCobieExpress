using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.CobieExpress;
using Xbim.Ifc;

namespace Xbim.IO.CobieExpress
{
    public class COBieModelProviderFactory : IModelProviderFactory
    {
        IModelProviderFactory inner;

        public COBieModelProviderFactory()
        {
            inner = new DefaultModelProviderFactory();
        }

        public IModelProvider CreateProvider()
        {
            var modelProvider = inner.CreateProvider();

            // override the modelProvider to use our COBie EntityFactory
            modelProvider.EntityFactoryResolver = (version) =>
            {
                if (version == Common.Step21.XbimSchemaVersion.Cobie2X4)
                {
                    return new EntityFactoryCobieExpress();
                }
                return null;
            };
            return modelProvider;
        }

        public void Use(Func<IModelProvider> providerFn)
        {
            inner.Use(providerFn);
        }
    }
}
