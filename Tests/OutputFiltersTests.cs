using FluentAssertions;
using Microsoft.Extensions.Logging;
using System;
using Xbim.CobieExpress.Exchanger.FilterHelper;
using Xbim.Common.Configuration;
using Xunit;

namespace Tests
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
}
