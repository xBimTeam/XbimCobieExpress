using System.ComponentModel;
using System.Threading.Tasks;
using Xbim.Common;

namespace Xbim.CobieExpress.Exchanger
{
    public interface ICobieConverter
    {
        Task<IModel> Run(CobieConversionParams args);
    }
}