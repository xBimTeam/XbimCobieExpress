using System;
using System.Collections.Generic;
using System.Linq;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger
{
    public static class IfcTypeObjectHashExtensions
    {
        /// <summary>
        /// Calculates a hash of the <see cref="IIfcTypeObject"/> to determine if it is a 
        /// duplicate, using the values of any Properties, Quantities and Classifications
        /// </summary>
        /// <param name="typeObject"></param>
        /// <param name="seedHash">An initial hash allowing hashes to ba 'chained'</param>
        /// <returns>A hashcode identifing the signature of the type</returns>
        public static int CalculateHash(this IIfcTypeObject typeObject, int seedHash = 0)
        {
            if (typeObject == null) return (seedHash);

            var hashCode = HashCode.Combine(seedHash, typeObject.Name.Value, typeObject.GetType().Name);

            if (typeObject.HasPropertySets != null)
            {
                // Account for Properties and Quantities
                hashCode = typeObject.HasPropertySets
                            .OrderBy(e => e.Name?.Value)
                            .Aggregate(hashCode, (current, next) => CalculateHash(next, current));
            }
            if (typeObject.HasAssociations != null)
            {
                hashCode = typeObject.HasAssociations.OfType<IIfcRelAssociatesClassification>()
                    .OrderBy(r => r.RelatingClassification.EntityLabel)
                    .Aggregate(hashCode, (current, next) => CalculateHash(next, current));
            }

            return hashCode;
        }

        private static int CalculateHash(IIfcRelAssociatesClassification relClass, int hashCode)
        {
            return relClass.RelatingClassification switch
            {
                // TODO: we could traverse up the hierarchy
                IIfcClassificationReference classRef => CalculateHash(hashCode, classRef.Identification),
                IIfcClassification system => CalculateHash(hashCode, system.Name),
                _ => hashCode
            };
        }

        private static int CalculateHash(IIfcPropertySetDefinition set, int hashCode)
        {
            return set switch
            {
                IIfcPropertySet pset => CalculateHash(pset, hashCode),
                IIfcElementQuantity pset => CalculateHash(pset, hashCode),
                IIfcPreDefinedPropertySet pdt => CalculateHash(pdt, hashCode),
                _ => hashCode
            };
        }

        private static int CalculateHash(IIfcPropertySet pset, int current)
        {
            int hashCode = HashCode.Combine(current, pset.Name?.Value);
            return pset.HasProperties
                    .OrderBy(e => e.Name.Value)
                    .Aggregate(hashCode, (curr, p) => CalculateHash(p, curr));
        }

        private static int CalculateHash(IIfcElementQuantity qset, int current)
        {
            int hashCode = HashCode.Combine(current, qset.Name?.Value);
            return qset.Quantities
                    .OrderBy(e => e.Name.Value)
                    .Aggregate(hashCode, (curr, q) => CalculateHash(q, curr));
        }

        private static int CalculateHash(IIfcProperty property, int current)
        {
            return property switch
            {
                IIfcPropertySingleValue single => HashCode.Combine(current, single.Name.Value, single.NominalValue?.Value),
                // IIfcPropertyEnumeratedValue enumerated => HashCode.Combine(current, enumerated.Name.Value, enumerated.E?.Value),
                IIfcPropertyListValue list => HashCode.Combine(current, list.Name.Value, CalculateHash(0, list.ListValues.ToArray())),
                IIfcPropertyBoundedValue bounded => HashCode.Combine(current, bounded.Name.Value, bounded.LowerBoundValue?.Value, bounded.UpperBoundValue?.Value),
                _ => current
            };
        }

        private static int CalculateHash(IIfcPhysicalQuantity quant, int current)
        {

            return quant switch
            {
                IIfcQuantityCount cnt => HashCode.Combine(current, cnt.Name.Value, cnt.CountValue.Value),
                IIfcQuantityArea area => HashCode.Combine(current, area.Name.Value, area.AreaValue.Value),
                IIfcQuantityLength len => HashCode.Combine(current, len.Name.Value, len.LengthValue.Value),
                IIfcQuantityVolume v => HashCode.Combine(current, v.Name.Value, v.VolumeValue.Value),
                IIfcQuantityWeight w => HashCode.Combine(current, w.Name.Value, w.WeightValue.Value),
                IIfcQuantityTime t => HashCode.Combine(current, t.Name.Value, t.TimeValue.Value),

                _ => HashCode.Combine(current, quant.Name.Value),
            };
        }

        private static int CalculateHash(IIfcPreDefinedPropertySet property, int current)
        {
            return property switch
            {
                IIfcDoorLiningProperties dl => CalculateHash(current,
                    dl.CasingDepth,
                    dl.CasingThickness,
                    dl.LiningDepth,
                    dl.LiningOffset,
                    dl.LiningThickness,
                    dl.LiningToPanelOffsetX,
                    dl.LiningToPanelOffsetY,
                    dl.TransomOffset,
                    dl.TransomThickness,
                    dl.ThresholdDepth,
                    dl.ThresholdOffset,
                    dl.ThresholdThickness,
                    dl.ShapeAspectStyle?.Name
                    ),
                IIfcDoorPanelProperties dp => HashCode.Combine(CalculateHash(current,
                    dp.PanelDepth,
                    dp.PanelWidth,
                    dp.ShapeAspectStyle?.Name
                    ), dp.PanelOperation, dp.PanelPosition),

                IIfcWindowLiningProperties wl => CalculateHash(current, 
                    wl.LiningDepth,
                    wl.LiningThickness,
                    wl.LiningOffset,
                    wl.LiningToPanelOffsetX,
                    wl.LiningToPanelOffsetY,
                    wl.TransomThickness,
                    wl.FirstTransomOffset,
                    wl.SecondTransomOffset,
                    wl.MullionThickness,
                    wl.FirstMullionOffset,
                    wl.SecondMullionOffset,
                    wl.ShapeAspectStyle?.Name
                ),
                IIfcWindowPanelProperties wp => HashCode.Combine(CalculateHash(current,
                    wp.FrameDepth,
                    wp.FrameThickness,
                    wp.ShapeAspectStyle?.Name
                ), wp.PanelPosition, wp.OperationType),
                
                _ => current
            };

        }

        private static int CalculateHash(int current, params IEnumerable<IIfcValue> values)
        {
            return values.Select(v => v?.Value).Aggregate(current, (curr, v) => HashCode.Combine(curr, v));
        }
    }
}
