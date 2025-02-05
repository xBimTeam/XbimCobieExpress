using Xbim.Common;

namespace Xbim.CobieExpress.Abstractions
{
    public interface IObjectFilter
    {
        ISerializableXmlDictionary<string, bool> Items { get; set; }
        ISerializableXmlDictionary<string, string[]> PreDefinedType { get; set; }

        void Clear();
        void Copy(IObjectFilter copyFilter);
        bool IsEmpty();
        bool ItemsFilter(IPersistEntity obj);
        bool ItemsFilter(string testStr, string preDefinedType = null);
        void MergeInc(IObjectFilter mergeFilter);
        bool SetPreDefinedType(string ifcElement, string[] definedTypes);
    }
}