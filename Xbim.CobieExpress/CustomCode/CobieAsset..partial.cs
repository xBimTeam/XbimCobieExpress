using System;
using System.Collections.Generic;
using System.Linq;
using Xbim.CobieExpress.Interfaces;
using Xbim.Common;

namespace Xbim.CobieExpress
{

    internal static class ItemSet
    {
        public static ItemSet<T> Empty<T>(IPersistEntity entity)
        {
            return new ItemSet<T>(entity, 0, 0);
        }
    }

    public partial class CobieContact : ICobieAsset
    {
        public string Name { get => Email; set => Email = value; }
        public string Description { get => $"{GivenName} {FamilyName}" ; set => throw new NotImplementedException(); }

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }

   

    public partial class CobieSpare : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }
    public partial class CobieResource : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }
    public partial class CobieJob : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }

    public partial class CobieDocument : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);     // TODO: DocumentType
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }

    public partial class CobieConnection : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }

    public partial class CobieIssue : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }

    public partial class CobieImpact : ICobieAsset
    {

        IItemSet<ICobieCategory> ICobieAsset.Categories => ItemSet.Empty<ICobieCategory>(this);
        IItemSet<ICobieImpact> ICobieAsset.Impacts => ItemSet.Empty<ICobieImpact>(this);
        IItemSet<ICobieDocument> ICobieAsset.Documents => ItemSet.Empty<ICobieDocument>(this);
        IItemSet<ICobieAttribute> ICobieAsset.Attributes => ItemSet.Empty<ICobieAttribute>(this);
        IItemSet<ICobieCoordinate> ICobieAsset.Representations => ItemSet.Empty<ICobieCoordinate>(this);
        IEnumerable<ICobieIssue> ICobieAsset.CausingIssues => Enumerable.Empty<ICobieIssue>();
        IEnumerable<ICobieIssue> ICobieAsset.AffectedBy => Enumerable.Empty<ICobieIssue>();
    }
}
