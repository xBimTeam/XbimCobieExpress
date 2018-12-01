using Xbim.CobieExpress;
using Xbim.Common;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger
{
    internal class MappingIfcSiteToSite : XbimMappings<IModel, IModel, int, IIfcSite, CobieSite>
    {
        protected override CobieSite Mapping(IIfcSite ifcSite, CobieSite site)
        {
            var helper = ((IfcToCoBieExpressExchanger)Exchanger).Helper;
            site.ExternalObject = helper.GetExternalObject(ifcSite);
            site.ExternalId = helper.ExternalEntityIdentity(ifcSite);
            site.AltExternalId = ifcSite.GlobalId;
            site.Name = ifcSite.LongName;
            site.Description = ifcSite.Description;
            return site;
        }

        public override CobieSite CreateTargetObject()
        {
            return Exchanger.TargetRepository.Instances.New<CobieSite>();
        }
    }
}
