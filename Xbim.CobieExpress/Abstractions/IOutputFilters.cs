using System.Collections.Generic;
using System.IO;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;

namespace Xbim.CobieExpress.Abstractions
{
    /// <summary>
    /// <para>
    /// Interface representing the contract by which IFC models are filtered (or selected) to generate the desired
    /// results in a COBieExpress model. 
    /// </para>
    /// <para>
    /// Filters can be Role based so that COBie extraction can be tuned per Discipline. 
    /// </para>
    /// </summary>
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
        /// <summary>
        /// Loads a new filter for a Role from either the supplied filename, or the built-in default when no fileName supplied
        /// </summary>
        /// <remarks>Overwrites any prior Filter</remarks>
        /// <param name="role"></param>
        /// <param name="filterFile">Optional filePath</param>
        IOutputFilters LoadFilter(RoleFilter role, string filterFile = null);
        /// <summary>
        /// Loads a new filter for a Role from the supplied Stream
        /// </summary>
        /// <remarks>Overwrites any prior Filter</remarks>
        /// <param name="filterStream"></param>
        /// <param name="role"></param>
        IOutputFilters LoadFilter(Stream filterStream, RoleFilter role);
        IOutputFilters GetRoleFilter(RoleFilter role);

        bool IsEmpty();
        void Merge(IOutputFilters mergeFilter);
        bool ObjFilter(IPersistEntity obj, bool checkType = true);
        void SerializeXml(FileInfo filename);
    }
}