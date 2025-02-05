using System.Collections.Generic;
using System.IO;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;

namespace Xbim.CobieExpress.Abstractions
{
    public interface IOutputFilters
    {
        RoleFilter AppliedRoles { get; set; }
        IPropertyFilter CommonFilter { get; set; }
        IPropertyFilter ComponentFilter { get; set; }
        bool DefaultsNotSet { get; }
        IPropertyFilter FacilityFilter { get; set; }
        bool FlipResult { get; set; }
        IPropertyFilter FloorFilter { get; set; }
        IObjectFilter IfcAssemblyFilter { get; set; }
        IObjectFilter IfcProductFilter { get; set; }
        IObjectFilter IfcTypeObjectFilter { get; set; }
        IPropertyFilter SpaceFilter { get; set; }
        IPropertyFilter SpareFilter { get; set; }
        IPropertyFilter TypeFilter { get; set; }
        IPropertyFilter ZoneFilter { get; set; }

        void ApplyRoleFilters(RoleFilter roles, bool append = false, IDictionary<RoleFilter, IOutputFilters> rolesFilter = null);
        void Clear();
        void Copy(IOutputFilters copyFilter);
        IOutputFilters GetRoleFilter(RoleFilter role);
        bool IsEmpty();
        void Merge(IOutputFilters mergeFilter);
        bool ObjFilter(IPersistEntity obj, bool checkType = true);
        void SerializeXml(FileInfo filename);
    }
}