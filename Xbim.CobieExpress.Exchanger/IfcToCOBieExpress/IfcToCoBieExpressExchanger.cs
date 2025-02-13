using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Linq;
using Xbim.CobieExpress.Exchanger.Classifications;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger
{

    /// <summary>
    /// Provide conversion between models in IFC schema to a well defined COBie schema
    /// </summary>
    public class IfcToCoBieExpressExchanger : XbimExchanger<IModel, ICOBieModel>, IIfcToCOBieExpressExchanger
    {
        private readonly ILogger<IfcToCoBieExpressExchanger> logger;

        protected IfcToCOBieExchangeConfiguration Configuration { get; private set; }


        /// <summary>
        /// Constructs a new <see cref="IfcToCoBieExpressExchanger"/>
        /// </summary>
        /// <param name="logger"></param>
        public IfcToCoBieExpressExchanger(ILogger<IfcToCoBieExpressExchanger> logger) : base()
        {
            this.logger = logger;
        }

        /// <summary>
        /// Instantiates a new IfcToCOBieLiteUkExchanger class.
        /// </summary>
        /// <param name="source">The source IFC model</param>
        /// <param name="target">The target <see cref="ICOBieModel"/></param>
        /// <param name="reportProgress">A progress delegate</param>
        /// <param name="filter">The IFC <see cref="OutputFilters"/> to apply</param>
        /// <param name="configFile">The IFC Property mapping configuration file to apply (defaults to embedded CobieAttributes.config)</param>
        /// <param name="extId">The External Identifier strategy (defaults to Ifc EntityLabels)</param>
        /// <param name="sysMode">The strategy for extracting Systems</param>
        /// <param name="classify">Determines whether category should be inferred from IFC Classifications (defaults false)</param>
        public IfcToCoBieExpressExchanger(IModel source, ICOBieModel target, ReportProgressDelegate reportProgress = null, 
            OutputFilters filter = null, 
            string configFile = null, 
            EntityIdentifierMode extId = EntityIdentifierMode.IfcEntityLabels, 
            SystemExtractionMode sysMode = SystemExtractionMode.System | SystemExtractionMode.Types, 
            bool classify = false) 
            : base()
        {
            IfcToCOBieExchangeConfiguration config = new IfcToCOBieExchangeConfiguration()
            {
                ReportProgressDelegate = reportProgress,
                SelectionFilters = filter,
                AttributeMappingFile = configFile,
                ExternalIdentifierSource = extId,
                SystemExtractionMode = sysMode,
                Classify = classify

            };
            Initialise(config, source, target);
        }

        /// <summary>
        /// Initialisers the Exchanger with runtime parameters before a conversion.
        /// </summary>
        /// <param name="configuration"></param>
        /// <param name="source"></param>
        /// <param name="cobieModel"></param>
        public void Initialise(IfcToCOBieExchangeConfiguration configuration, IModel source, ICOBieModel cobieModel)
        {
            base.Initialise(source, cobieModel);
            Configuration = configuration;
            ReportProgress.Progress = configuration.ReportProgressDelegate; //set reporter
            Helper = new COBieExpressHelper(this, ReportProgress, logger, configuration);
            Helper.Init();
        }

        /// <summary>
        /// 
        /// </summary>
        internal COBieExpressHelper Helper { get; private set; }

        /// <summary>
        /// Converts a source (IFC) model to COBie schema using the defined mapping configuration
        /// </summary>
        /// <returns></returns>
        public override ICOBieModel Convert()
        {
            if (Configuration == null || Helper == null)
                throw new System.InvalidOperationException("IfcToCoBieExpressExchanger not initialised");

            var mapping = GetOrCreateMappings<MappingIfcBuildingToFacility>();
            var classifier = new Classifier(this);
            var buildings = SourceRepository.Instances.OfType<IIfcBuilding>().ToList();
            var facilities = new List<CobieFacility>(buildings.Count);
            foreach (var building in buildings)
            {
                var facility = TargetRepository.Instances.New<CobieFacility>();
                facility = mapping.AddMapping(building, facility);
                if(Configuration.Classify)
                    classifier.Classify(facility);
                facilities.Add(facility);
            }
            return TargetRepository;
        }
    }
}
