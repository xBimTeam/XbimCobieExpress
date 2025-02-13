using Xbim.CobieExpress;
using Xbim.Common;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger
{
    internal class MappingIfcProjectToProject : XbimMappings<IModel, ICOBieModel, int, IIfcProject, CobieProject>
    {
        protected override CobieProject Mapping(IIfcProject source, CobieProject target)
        {
            var helper = ((IfcToCoBieExpressExchanger)Exchanger).Helper;
            target.ExternalObject = helper.GetExternalObject(source);
            target.ExternalId = helper.ExternalEntityIdentity(source);
            target.AltExternalId = source.GlobalId;
            target.Name = source.Name;
            target.Description = source.LongName;
            if(string.IsNullOrWhiteSpace(target.Description))
            {
                target.Description = source.Description;
            }
            return target;
        }

        public override CobieProject CreateTargetObject()
        {
            return Exchanger.TargetRepository.Instances.New<CobieProject>();
        }
    }
}
