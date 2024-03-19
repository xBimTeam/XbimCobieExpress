using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Xbim.CobieExpress;
using Xbim.Common.Metadata;
using Xbim.IO.Table;
using Xbim.IO.Table.Resolvers;

namespace Xbim.IO.CobieExpress.Resolvers
{
    public class AttributeTypeResolver : ITypeResolver
    {
        public bool CanResolve(Type type)
        {
            return type == typeof(AttributeValue);
        }

        public bool CanResolve(ExpressType type)
        {
            return CanResolve(type.Type);
        }

        private static readonly Regex DateTimeRegex = new Regex("[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}",
                            RegexOptions.Compiled);
        private static readonly Regex FirstLetterRegex = new Regex("^[0-9].*",
                            RegexOptions.Compiled);

        public Type Resolve(Type type, Cell cell, ClassMapping cMapping, PropertyMapping pMapping,SharedStringTable sharedStringTable)
        {

            if (cell.DataType == null)
            {
                if (int.TryParse(cell.InnerText, out int intValue))
                {
                    return typeof(IntegerValue);
                }
                else if(double.TryParse(cell.InnerText, out double doubleValue))
                {
                    return typeof(FloatValue);
                }
                else
                {
                    return typeof(StringValue);
                }
            }
            if (cell.DataType == CellValues.Number)
            {
                //it might be integer or float
                if (double.TryParse(cell.InnerText, out double numericValue))
                {
                    return Math.Abs(numericValue % 1) < 1e-9 ? typeof(IntegerValue) : typeof(FloatValue);
                }
                return typeof(StringValue);
            }
            else if (cell.DataType == CellValues.String)
            {
                //it might be string or datetime
                var str = cell.CellValue.Text;
                if (str.Length >= 19 && FirstLetterRegex.IsMatch(str[0].ToString())) //2009-06-15T13:45:30
                {
                    var dStr = str.Substring(0, 19);
                    if (DateTimeRegex.IsMatch(dStr))
                        return typeof(DateTimeValue);
                }
                return typeof(StringValue);
            }
            else if (cell.DataType == CellValues.SharedString)
            {
                //it might be string or datetime
                var str = sharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                if (str.Length >= 19 && FirstLetterRegex.IsMatch(str[0].ToString())) //2009-06-15T13:45:30
                {
                    var dStr = str.Substring(0, 19);
                    if (DateTimeRegex.IsMatch(dStr))
                        return typeof(DateTimeValue);
                }
                return typeof(StringValue);
            }
            else if (cell.DataType == CellValues.Boolean)
                return typeof(BooleanValue);

            return typeof(StringValue);
        }

        public ExpressType Resolve(ExpressType abstractType, ReferenceContext context, ExpressMetaData metaData)
        {
            return metaData.ExpressType(typeof(StringValue));
        }
    }
}
