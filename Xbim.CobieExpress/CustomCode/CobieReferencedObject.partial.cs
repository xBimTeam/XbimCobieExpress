using System;

namespace Xbim.CobieExpress.Interfaces
{
    public partial interface ICobieReferencedObject
    {
        /// <summary>
        /// The original Excel row number. one-based not zero-based in order to align with user's expectation.
        /// </summary>
        /// <remarks>Used for reference back to the origin COBie row. Can be used to ensure round-trip ordering is maintained.</remarks>
        int RowNumber { get; set; }
    }
}
