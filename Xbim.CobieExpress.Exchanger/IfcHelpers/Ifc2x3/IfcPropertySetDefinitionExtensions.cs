using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger.IfcHelpers.Ifc2x3
{
    public static class IfcPropertySetDefinitionExtensions
    {
        public static bool Add(this IIfcPropertySetDefinition pSetDefinition, IIfcSimpleProperty prop)
        {
            var propSet = pSetDefinition as IIfcPropertySet;
            if(propSet!=null) propSet.HasProperties.Add(prop);
            return propSet != null;
        }

        public static bool Add(this IIfcPropertySetDefinition pSetDefinition, IIfcPhysicalQuantity quantity)
        {
            var quantSet = pSetDefinition as IIfcElementQuantity;
            if (quantSet != null) quantSet.Quantities.Add(quantity);
            return quantSet != null;
        }
    }
}
