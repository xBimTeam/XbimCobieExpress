using System;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;

namespace XbimExchanger.IfcHelpers
{
    /// <summary>
    /// Params Class, holds parameters for worker to access
    /// </summary>
    public class CobieConversionParams
    {
        public IModel Source { get; set; }
        public Func<IModel> NewCobieModel { get; set; }
        public EntityIdentifierMode ExtId { get; set; }
        public SystemExtractionMode SysMode { get; set; }
        public OutPutFilters Filter { get; set; }
        public string ConfigFile { get; set; }
        /// <summary>
        /// Produce COBie validation log file
        /// </summary>
        public bool Log { get; set; } 
        public ReportProgressDelegate ReportProgress { get; set; }
    }
}
