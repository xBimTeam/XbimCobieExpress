using System.ComponentModel;
using System.Threading.Tasks;
using Xbim.Common;

namespace XbimExchanger.IfcHelpers
{
    public interface ICobieConverter
    {
        Task<IModel> Run(CobieConversionParams args);
    }
}