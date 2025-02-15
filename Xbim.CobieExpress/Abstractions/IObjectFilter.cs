using Xbim.Common;

namespace Xbim.CobieExpress.Abstractions
{
    public interface IObjectFilter
    {
        /// <summary>
        /// Dictionary of type names and whether they are permitted or denied
        /// </summary>
        ISerializableXmlDictionary<string, bool> Items { get; set; }
        /// <summary>
        /// Dictionary of type names to their applicable PreDefined Types
        /// </summary>
        ISerializableXmlDictionary<string, string[]> PreDefinedType { get; set; }

        /// <summary>
        /// Clears all Items and PreDefinedTypes
        /// </summary>
        void Clear();
        /// <summary>
        /// Rebuilds the internal lists. Can be used to apply updates retrospectively
        /// </summary>
        void Rebuild();
        void Copy(IObjectFilter copyFilter);
        bool IsEmpty();
        bool ItemsFilter(IPersistEntity obj);
        bool ItemsFilter(string testStr, string preDefinedType = null);
        void MergeInc(IObjectFilter mergeFilter);
        bool SetPreDefinedType(string ifcElement, string[] definedTypes);
    }
}