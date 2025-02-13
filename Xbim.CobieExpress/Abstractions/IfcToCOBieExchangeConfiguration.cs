using System;
using Xbim.CobieExpress.Abstractions;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;

namespace Xbim.CobieExpress.Exchanger
{
    /// <summary>
    /// Class defining the IFC to COBie Exchange configuration
    /// </summary>
    public class IfcToCOBieExchangeConfiguration
    {
        /// <summary>
        /// Gets and sets the default filters used to select / filter the relevant IFC entities for the given role
        /// </summary>
        public IOutputFilters SelectionFilters { get; set; }

        /// <summary>
        /// Allow any filters to be inspected or manipulated before conversion
        /// </summary>
        public Func<IOutputFilters, IOutputFilters> SelectionBuilder { get; set; } = (f) => { f.LoadFilter(RoleFilter.Default); return f; };

        /// <summary>
        /// Path to an Attribute Configuration allowing the mapping of IFC Properties to designed COBie Attributes
        /// </summary>
        /// <remarks>
        /// If null, defaults to <c>COBieAttributes.config</c> embedded in the Exchanger</remarks>
        public string AttributeMappingFile { get; set; }

        /// <summary>
        /// A delegate to receive progress updates
        /// </summary>
        public ReportProgressDelegate ReportProgressDelegate { get; set; } = null;

        /// <summary>
        /// Indicates the strategy for applying ExternalIdentifiers
        /// </summary>
        public EntityIdentifierMode ExternalIdentifierSource { get; set; } = EntityIdentifierMode.GloballyUniqueIds;

        /// <summary>
        /// Indicates the strategy for identifying Systems
        /// </summary>
        public SystemExtractionMode SystemExtractionMode { get; set; } = SystemExtractionMode.System | SystemExtractionMode.Types;

        /// <summary>
        /// When <c>true</c> infers Categories from well known IFC Classifications
        /// </summary>
        public bool Classify { get; set; } = false;

        /// <summary>
        /// Indicates strategy for locating a Type's properties from its associated Components
        /// </summary>
        /// <remarks>Some models flatten properties so COBieType properties are on the IfcObjects rather than the IfcTypeObjects that
        /// define the instances.</remarks>
        public InferFromComponentMode InferTypePropertiesFromComponents { get; set; } = InferFromComponentMode.Disabled;

        /// <summary>
        /// Controls which External Systems data is exported
        /// </summary>
        public ExternalReferenceMode ExternalReferenceMode { get; set; } = ExternalReferenceMode.OutputAll;
    }

    /// <summary>
    /// Defines the strategies to be used to locate Type Properties when not found on the type
    /// </summary>
    public enum InferFromComponentMode
    {
        /// <summary>
        /// The Type is the only place where Type properties are checked
        /// </summary>
        Disabled = 0,

        /// <summary>
        /// Fall back to the first matching property on a Type's Components
        /// </summary>
        FirstComponent,

        /// <summary>
        /// Fall back to the matching properties on a Type's Component, but only where the value in unambiguous
        /// </summary>
        UnambiguousComponents


    }

    /// <summary>
    /// 
    /// </summary>
    public enum EntityIdentifierMode
    {
        /// <summary>
        /// Use the Entity Label in the Ifc file (e.g. #23)
        /// </summary>
        IfcEntityLabels = 0,
        /// <summary>
        /// Use the GlobalId of the Entity (e.g. "10mjSDZJj9gPS2PrQaxa3z")
        /// </summary>
        GloballyUniqueIds = 1,
        /// <summary>
        /// Does not write any External Identifier for Entities
        /// </summary>
        None = 2
    }

    /// <summary>
    /// Control what we extract from IFC as systems
    /// </summary>
    [Flags]
    public enum SystemExtractionMode
    {
        /// <summary>
        /// Identify System from IfcSystems
        /// </summary>
        System = 0x1, //default and should always be set
        /// <summary>
        /// Identify Systems with properties defined in 'SystemMaps' config
        /// </summary>
        PropertyMaps = 0x2, //include properties as set by GetPropMap("SystemMaps")
        /// <summary>
        /// Include Types in System listings
        /// </summary>
        Types = 0x4, //include types as system listing all defined objects in componentnsnames
    }

    /// <summary>
    /// 
    /// </summary>
    public enum ExternalReferenceMode
    {
        // Output always
        OutputAll = 0,

        /// <summary>
        /// Does not write out the External Entity Type Name or the External System Name
        /// </summary>
        IgnoreSystemAndEntityName = 3,
        /// <summary>
        /// Does not write out the External System Name but does write out the External Entity Type Name
        /// </summary>
        IgnoreSystem = 1,
        /// <summary>
        /// Does not write out the External Entity Type Name but does write the External System Name
        /// </summary>
        IgnoreEntityName = 2

    }
}
