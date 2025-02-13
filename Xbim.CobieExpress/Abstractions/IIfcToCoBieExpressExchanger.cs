using Xbim.Common;

namespace Xbim.CobieExpress.Exchanger
{
    public interface IIfcToCOBieExpressExchanger
    {
        void Initialise(IfcToCOBieExchangeConfiguration configuration, IModel source, ICOBieModel cobieModel);
        ICOBieModel Convert();
    }
}
