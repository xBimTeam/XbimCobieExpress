using System.Collections.Generic;

namespace Xbim.CobieExpress.Abstractions
{
    public interface IPropertyFilter
    {
        List<string> Contain { get; set; }
        List<string> EqualTo { get; set; }
        List<string> PropertySetsEqualTo { get; set; }
        List<string> StartWith { get; set; }

        void Clear();
        void Copy(IPropertyFilter copyFilter);
        bool IsEmpty();
        void Merge(IPropertyFilter mergeFilter);
        bool NameFilter(string testStr);
        bool PSetNameFilter(string testStr);
    }
}