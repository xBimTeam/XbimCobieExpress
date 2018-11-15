using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.Common;
using Xbim.Common.Federation;
using XbimExchanger.IfcHelpers;


namespace XbimExchanger.IfcToCOBieExpress.Conversion
{

    public class CobieExpressConverter : ICobieConverter
    {
        private readonly ILogger _logger;

        public CobieExpressConverter(ILogger logger)
        {
            _logger = logger;
        }


        /// <summary>
        /// Run the worker
        /// </summary>
        /// <param name="args"></param>
        public async Task<IModel> Run(CobieConversionParams args)
        {
            if (args.ReportProgress == null)
                args.ReportProgress = (int percentProgress, object userState) => { };
            return await Task.Run<IModel>(() => {
                return GetCobieModel(args);
            });
        }

        private IModel GetCobieModel(CobieConversionParams parameters)
        {
            var timer = new Stopwatch();
            timer.Start();

            var model = parameters.Source;
            if (model is IFederatedModel fm && fm.ReferencedModels.Count() > 1)
            {
                throw new NotImplementedException("Work to do on COBie Federated");
                //see COBieLitConverter for Lite code
            }
            var cobie = parameters.NewCobieModel();
            using (var txn = cobie.BeginTransaction("begin conversion"))
            {
                var exchanger = new IfcToCoBieExpressExchanger
                    (model,
                    cobie,
                    _logger,
                    parameters.ReportProgress,
                    parameters.Filter,
                    parameters.ConfigFile,
                    parameters.ExtId,
                    parameters.SysMode
                    );
                exchanger.Convert();
                txn.Commit();
            }

            timer.Stop();
            parameters.ReportProgress(0, string.Format("Time to generate COBieLite data: {0} seconds", timer.Elapsed.TotalSeconds.ToString("F3")));
            return cobie;
        }
    }
}
