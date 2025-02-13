using System;

namespace Xbim.CobieExpress.Exchanger.FilterHelper
{
    /// <summary>
    /// Merge Flags for roles in deciding if an object is allowed or discarded depending on the role of the model
    /// </summary>
    [Flags] //allows use to | and & values for multiple boolean tests
    public enum RoleFilter
    {
        /// <summary>
        /// Any role - a combination of all roles
        /// </summary>
        Unknown = 0x1,

        /// <summary>
        /// Architectural disciplines
        /// </summary>
        Architectural = 0x2,
        /// <summary>
        /// Mechanical Disciplines
        /// </summary>
        Mechanical = 0x4,
        /// <summary>
        /// Electrical Disciplines
        /// </summary>
        Electrical = 0x8,
        /// <summary>
        /// Plumbing Disciplines
        /// </summary>
        Plumbing = 0x10,
        /// <summary>
        /// Fire Protection Disciplines
        /// </summary>
        FireProtection = 0x20,

        /// <summary>
        /// The default - all roles. Synonymous with <see cref="RoleFilter.Unknown"/>
        /// </summary>
        Default = Unknown
    }
}