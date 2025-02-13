using System;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common;

namespace Xbim.CobieExpress.Exchanger
{
    /// <summary>
    /// Params Class, holds parameters for worker to access
    /// </summary>
    public class CobieConversionParams
    {
        public IModel Source { get; set; }
        public Func<ICOBieModel> NewCobieModel { get; set; }
        public EntityIdentifierMode ExtId { get; set; }
        public SystemExtractionMode SysMode { get; set; }
        public OutputFilters Filter { get; set; }
        public string ConfigFile { get; set; }
        public ReportProgressDelegate ReportProgress { get; set; }
    }
}
