using Xbim.CobieExpress;
using Xbim.Common;
using Xbim.Ifc;

namespace Xbim.CobieExpress.Exchanger
{
    internal class MappingStringToExternalObject : XbimMappings<IModel, ICOBieModel, string, string, CobieExternalObject>
    {
        public override CobieExternalObject CreateTargetObject()
        {
            return Exchanger.TargetRepository.Instances.New<CobieExternalObject>();
        }

        protected override CobieExternalObject Mapping(string source, CobieExternalObject target)
        {
            target.Name = source;
            return target;
        }
    }
}
