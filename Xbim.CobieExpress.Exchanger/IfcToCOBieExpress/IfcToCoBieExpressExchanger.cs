using Microsoft.Extensions.Logging.Abstractions;
using System.Collections.Generic;
using System.Linq;
using Xbim.CobieExpress.Exchanger.Classifications;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.CobieExpress.Exchanger.IfcHelpers;
using Xbim.Common;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger
{
    public class IfcToCoBieExpressExchanger : XbimExchanger<IModel, IModel>
    {
        private readonly bool _classify;
        internal COBieExpressHelper Helper ;
        /// <summary>
        /// Instantiates a new IIfcToCOBieLiteUkExchanger class.
        /// </summary>
        public IfcToCoBieExpressExchanger(IModel source, IModel target, ReportProgressDelegate reportProgress = null, OutputFilters filter = null, string configFile = null, EntityIdentifierMode extId = EntityIdentifierMode.IfcEntityLabels, SystemExtractionMode sysMode = SystemExtractionMode.System | SystemExtractionMode.Types, bool classify = false) 
            : base(source, target)
        {

            ReportProgress.Progress = reportProgress; //set reporter
            Helper = new COBieExpressHelper(this, ReportProgress, default, filter, configFile, extId, sysMode);
            Helper.Init();

            _classify = classify;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override IModel Convert()
        {
            var mapping = GetOrCreateMappings<MappingIfcBuildingToFacility>();
            var classifier = new Classifier(this);
            var buildings = SourceRepository.Instances.OfType<IIfcBuilding>().ToList();
            var facilities = new List<CobieFacility>(buildings.Count);
            foreach (var building in buildings)
            {
                var facility = TargetRepository.Instances.New<CobieFacility>();
                facility = mapping.AddMapping(building, facility);
                if(_classify)
                    classifier.Classify(facility);
                facilities.Add(facility);
            }
            return TargetRepository;
        }
    }
}
