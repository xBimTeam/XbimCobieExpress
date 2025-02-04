using FluentAssertions;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Concurrent;
using System.Linq;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common.Configuration;
using Xbim.Ifc2x3.BuildingcontrolsDomain;
using Xbim.Ifc2x3.HVACDomain;
using Xbim.Ifc2x3.Kernel;
using Xbim.Ifc2x3.SharedBldgElements;
using Xbim.Ifc2x3.SharedBldgServiceElements;
using Xbim.Ifc4.Interfaces;
using Xbim.IO.Memory;
using Xunit;

namespace Xbim.CobieExpress.Tests
{
    public class OutputFiltersTests
    {

        private readonly ILogger logger = XbimServices.Current.CreateLogger<OutputFilters>();

        private readonly string productKey = nameof(Xbim.Ifc4.SharedBldgServiceElements.IfcEnergyConversionDevice).ToUpper();
        private readonly string predefinedTypeKey = nameof(Xbim.Ifc4.BuildingControlsDomain.IfcSensorType).ToUpper();

        [Fact]
        public void Can_Merge_Filters()
        {

            var architectFilters = new OutputFilters(logger, RoleFilter.Architectural);
            var mechanicalFilters = new OutputFilters(logger, RoleFilter.Mechanical);

            // Sanity Checks
            architectFilters.IfcProductFilter.Items[productKey].Should().BeFalse();
            mechanicalFilters.IfcProductFilter.Items[productKey].Should().BeTrue();
            architectFilters.IfcTypeObjectFilter.PreDefinedType.Should().NotContainKey(predefinedTypeKey);
            mechanicalFilters.IfcTypeObjectFilter.PreDefinedType.Should().ContainKey(predefinedTypeKey);

            // Act
            architectFilters.Merge(mechanicalFilters);

            // Assert
            architectFilters.IfcProductFilter.Items[productKey].Should().BeTrue();
            architectFilters.IfcTypeObjectFilter.PreDefinedType.Should().ContainKey(predefinedTypeKey);

        }


        [InlineData(false, RoleFilter.Architectural, typeof(IfcWall))]
        [InlineData(false, RoleFilter.Unknown, typeof(IfcWall))]
        [InlineData(true, RoleFilter.Architectural, typeof(IfcDoor))]
        [InlineData(true, RoleFilter.Unknown, typeof(IfcDoor))]
        [InlineData(false, RoleFilter.Architectural, typeof(IfcEnergyConversionDevice))]
        [InlineData(true, RoleFilter.Unknown, typeof(IfcEnergyConversionDevice))]

        [InlineData(true, RoleFilter.Architectural, typeof(IfcDistributionControlElement))]
        [InlineData(false, RoleFilter.Architectural, typeof(IfcDistributionControlElement), typeof(IfcSensorType))]
        [InlineData(true, RoleFilter.Unknown, typeof(IfcDistributionControlElement), typeof(IfcSensorType))]

        [InlineData(true, RoleFilter.Electrical, typeof(IfcFlowMeterType), null, "ENERGYMETER")]
        [InlineData(false, RoleFilter.Electrical, typeof(IfcFlowMeterType), null, "GASMETER")]
        [Theory]
        public void DoesFilterElements(bool expectedToInclude, RoleFilter role, Type productType, Type familyType = default, string productPdt = null)
        {

            using var model = new MemoryModel(new Ifc2x3.EntityFactoryIfc2x3());
            using var txn = model.BeginTransaction("");
            var filters = new OutputFilters(logger, role);

            IfcObjectDefinition entity = (IfcObjectDefinition)model.Instances.New(productType);
            if (familyType != default && entity is IfcObject instance)
            {
                IfcTypeObject parent = (IfcTypeObject)model.Instances.New(familyType);
                instance.AddDefiningType(parent);
            }

            if(productPdt != null)
            {
                entity.SetPredefinedTypeValue(productPdt);
            }

            if(expectedToInclude)
            {
                filters.ObjFilter(entity).Should().Be(false, "Role {0} should output {1}", role, entity);
            }
            else
            {
                filters.ObjFilter(entity).Should().Be(true, "Role {0} should filter out {1}", role, entity);
            }

        }

        [Fact]
        public void Merging_Is_Additive()
        {
            // Repeat the standard test in reverse. You can't turn off a filter that's already true

            var architectFilters = new OutputFilters(logger, RoleFilter.Architectural);
            var mechanicalFilters = new OutputFilters(logger, RoleFilter.Mechanical);

            // Act
            mechanicalFilters.Merge(architectFilters);

            // Assert
            mechanicalFilters.IfcProductFilter.Items[productKey].Should().BeTrue();
            mechanicalFilters.IfcTypeObjectFilter.PreDefinedType.Should().ContainKey(predefinedTypeKey);
        }

        [Fact]
        public void Loading_Multiple_Roles_Throws()
        {
            // Act

            var ex = Record.Exception(() =>
                new OutputFilters(logger, RoleFilter.Architectural | RoleFilter.Mechanical)
            );

            // Assert
            ex.Should().BeOfType<InvalidOperationException>();
        }

        [Fact]
        public void RolesFilters_Can_Have_Multiple_Values()
        {
            RoleFilter.Architectural.HasMultipleFlags().Should().BeFalse();
            (RoleFilter.Architectural | RoleFilter.Electrical).HasMultipleFlags().Should().BeTrue();

            RoleFilter undefined = 0;

            undefined.HasMultipleFlags().Should().BeFalse();
        }
    }

    public static class ObjectDefinitionExtensions
    {
        private const string PredefinedType = nameof(IIfcPile.PredefinedType);

        // A thread-safe cache between the type and a Setter for its PredefinedType property. By caching setters we elimiminate all Reflection
        static ConcurrentDictionary<Type, (Type, Action<object, object>)> _predefinedSetterDict = new ConcurrentDictionary<Type, (Type, Action<object, object>)>();

        /// <summary>
        /// Gets the string value of any PredefinedType property on the <see cref="IIfcObjectDefinition"/> instance, if the type has one; else returns null
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        public static bool SetPredefinedTypeValue(this IIfcObjectDefinition instance, string value)
        {
            if (instance is null)
            {
                return false;
            }

            // Locate the setter for the PredefinedType property on this type from the cache, lazily creating one if not present
            var (type, setter) = _predefinedSetterDict.GetOrAdd(instance.GetType(), (_) => BuildPredefinedTypeSetter(instance));
            if (setter != null)
            {
                if(Enum.TryParse(type, value, true, out var pdt))
                {
                    setter(instance, pdt);
                    return true;
                }
            }
            // this type has no PredefinedType property, or the enum was not applicable
            return false;
        }

        /// <summary>
        /// Returns the Type and its setter method for the PredefinedType property on the concrete type of the object
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private static (Type, Action<object, object>) BuildPredefinedTypeSetter(IIfcObjectDefinition obj)
        {
            var predefinedMetadata = obj.ExpressType.Properties.FirstOrDefault(p => p.Value.Name == PredefinedType).Value;
            if (predefinedMetadata != null)
            {
                var type = predefinedMetadata.PropertyInfo.PropertyType;
                // return the Type and the setter function - Input: 1) the instance as param 2) enum object. Output=void
                return (type, predefinedMetadata.PropertyInfo.SetValue);
            }
            return (null, null);
        }

    }
}
