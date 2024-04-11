using DocumentFormat.OpenXml.Spreadsheet;
using System;
using Xbim.Common.Metadata;

namespace Xbim.IO.Table.Resolvers
{
    /// <summary>
    /// Implementatios of ITypeResolver can be used to resolve abstract types when data is being read into object model.
    /// You can add as many resolvers as necessary to TableStore.
    /// </summary>
    public interface ITypeResolver
    {
        /// <summary>
        /// Checks if this resolver can resolve the type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        bool CanResolve(Type type);
        bool CanResolve(ExpressType type);
        Type Resolve(Type type, Cell cell, ClassMapping cMapping, PropertyMapping pMapping, SharedStringTable sharedStringTable);
        ExpressType Resolve(ExpressType abstractType, ReferenceContext context, ExpressMetaData metaData);
    }
}
