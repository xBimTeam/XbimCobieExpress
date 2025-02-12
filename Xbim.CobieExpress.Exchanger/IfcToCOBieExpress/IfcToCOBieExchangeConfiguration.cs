using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;
using Xbim.Common.Configuration;

namespace Xbim.CobieExpress.Exchanger
{
    /// <summary>
    /// Class defining the IFC to COBie Exchange configuration
    /// </summary>
    public class IfcToCOBieExchangeConfiguration
    {
        /// <summary>
        /// Gets and sets the filters used to select / filter the relevant IFC entities for the given role
        /// </summary>
        public OutputFilters SelectionFilters { get; set; } = OutputFilters.GetDefaults(RoleFilter.Unknown, XbimServices.Current.CreateLogger<OutputFilters>());

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
}
