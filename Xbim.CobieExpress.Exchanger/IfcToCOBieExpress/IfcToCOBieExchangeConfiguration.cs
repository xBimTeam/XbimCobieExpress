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
    }
}
