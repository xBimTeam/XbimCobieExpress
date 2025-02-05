using System.Collections.Generic;
using System.Xml.Serialization;

namespace Xbim.CobieExpress.Abstractions
{
    public interface ISerializableXmlDictionary<TKey, TValue> : IDictionary<TKey, TValue>, IXmlSerializable
    {

    }
}