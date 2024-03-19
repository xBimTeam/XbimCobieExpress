using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Xbim.Common;
using Xbim.Common.Exceptions;
using Xbim.Common.Metadata;
using Xbim.IO.Table.Resolvers;


namespace Xbim.IO.Table
{
    public partial class TableStore
    {
        public IModel Model { get; private set; }
        public ModelMapping Mapping { get; private set; }
        public ExpressMetaData MetaData { get { return Model.Metadata; } }
        public List<ITypeResolver> Resolvers { get; private set; }

        #region Writing data out to a spreadsheet
        /// <summary>
        /// Limit of the length of the text in a cell before the row gets repeated if MultiRow == MultiRow.IfNecessary
        /// </summary>
        private const int CellTextLimit = 1024;

        //dictionary of all styles for different data statuses
        private Dictionary<DataStatus, uint> _styles;
        private SharedStringTable _sharedStringTable;

        //cache of latest row number in different sheets
        private Dictionary<string, uint> _rowNumCache = new Dictionary<string, uint>();

        //cache of class mappings and respective express types
        private readonly Dictionary<ExpressType, ClassMapping> _typeClassMappingsCache = new Dictionary<ExpressType, ClassMapping>();

        //cache of meta properties so it doesn't have to look them up in metadata all the time
        private readonly Dictionary<ExpressType, Dictionary<string, ExpressMetaProperty>> _typePropertyCache = new Dictionary<ExpressType, Dictionary<string, ExpressMetaProperty>>();

        // cache of index column indices for every table in use
        private Dictionary<string, int[]> _multiRowIndicesCache;

        //preprocessed enum aliases to speed things up
        private readonly Dictionary<string, string> _enumAliasesCache;
        private readonly Dictionary<string, string> _aliasesEnumCache;

        //list of forward references to be resolved
        private readonly Queue<ForwardReference> _forwardReferences = new Queue<ForwardReference>();

        // cache used by forwwrdReferences as optimisation Resolving parents across references
        internal readonly List<IPersistEntity> _forwardReferenceParentCache = new List<IPersistEntity>();

        //cached check if the mapping contains any potentially multi-value columns
        private Dictionary<ClassMapping, bool> _isMultiRowMappingCache;

        //cache of global types so that it is not necessary to search and validate in configuration
        private List<ExpressType> _globalTypes;
        private readonly Dictionary<ExpressType, Dictionary<string, IPersistEntity>> _globalEntities = new Dictionary<ExpressType, Dictionary<string, IPersistEntity>>();

        //cache of all reference contexts which are only built once (string parsing, search for express properties and types)
        private Dictionary<ClassMapping, ReferenceContext> _referenceContexts;

        public Dictionary<string, Dictionary<uint, int>> RowNoToEntityLabelLookup = new Dictionary<string, Dictionary<uint, int>>();

        public TableStore(IModel model, ModelMapping mapping)
        {
            Model = model;
            Mapping = mapping;
            Resolvers = new List<ITypeResolver>();

            if (mapping.EnumerationMappings != null && mapping.EnumerationMappings.Any())
            {
                _enumAliasesCache = new Dictionary<string, string>();
                _aliasesEnumCache = new Dictionary<string, string>();
                foreach (var enumMapping in mapping.EnumerationMappings)
                {
                    if (enumMapping.Aliases == null || !enumMapping.Aliases.Any())
                        continue;
                    foreach (var alias in enumMapping.Aliases)
                    {
                        _enumAliasesCache.Add(enumMapping.Enumeration + "." + alias.EnumMember, alias.Alias);
                        _aliasesEnumCache.Add(enumMapping.Enumeration + "." + alias.Alias, alias.EnumMember);
                    }
                }
            }
            else
            {
                _enumAliasesCache = null;
                _aliasesEnumCache = null;
            }

            Mapping.Init(MetaData);
        }

        public void Store(string path, Stream template = null)
        {
            if (path == null)
                throw new ArgumentNullException("path");
            var ext = Path.GetExtension(path).ToLower().Trim('.');
            if (ext != "xls" && ext != "xlsx")
            {
                //XLSX is Spreadsheet XML representation which is capable of storing more data
                path += ".xlsx";
                ext = "xlsx";
            }
            using (var file = File.Create(path))
            {
                var type = ext == "xlsx" ? ExcelTypeEnum.XLSX : ExcelTypeEnum.XLS;
                Store(file, type, template: template);
                file.Close();
            }


        }

        public Stream Store(Stream stream, ExcelTypeEnum type, Stream template = null, bool recalculate = false)
        {
            Log = new StringWriter();
            SpreadsheetDocument spreadsheetDocument;
            if (template != null)
            {
                spreadsheetDocument = SpreadsheetDocument.Open(template, true);
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart wbPart = spreadsheetDocument.AddWorkbookPart();
                wbPart.Workbook = new Workbook();
            }

            var workbookPart = spreadsheetDocument.WorkbookPart;

            //Add a WorkbookStylesPart to the WorkbookPart
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();
            CellFormats cellFormats = new CellFormats();
            Fonts fonts = new Fonts();
            Borders borders = new Borders();
            Fills fills = new Fills();
            stylesPart.Stylesheet.CellFormats = cellFormats;
            stylesPart.Stylesheet.Fonts = fonts;
            stylesPart.Stylesheet.Borders = borders;
            stylesPart.Stylesheet.Fills = fills;
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            //create spreadsheet representaion 
            Store(workbookPart);

            if (template != null)
            {
                var newSpreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                // Add a WorkbookPart to the document
                WorkbookPart newworkbookPart = newSpreadsheetDocument.AddWorkbookPart();
                newworkbookPart.Workbook = new Workbook();

                // Copy the content of the template document to the output document
                WorkbookPart templateWorkbookPart = spreadsheetDocument.WorkbookPart;

                foreach (var part in templateWorkbookPart.Parts)
                {
                    newworkbookPart.AddPart(part.OpenXmlPart, part.RelationshipId);
                }
                workbookPart = newworkbookPart;
            }
            if (!recalculate || template == null)
            {
                spreadsheetDocument.Save();
                spreadsheetDocument.Dispose();
                return stream;
            }


            //write to output stream
            spreadsheetDocument.Save();
            spreadsheetDocument.Dispose();
            return stream;
        }

        public void Store(WorkbookPart workbook)
        {
            //if there are no mappings do nothing
            if (Mapping.ClassMappings == null || !Mapping.ClassMappings.Any()) return;

            _rowNumCache = new Dictionary<string, uint>();
            _styles = new Dictionary<DataStatus, uint>();

            //creates tables in defined order if they are not there yet
            SetUpTables(workbook, Mapping);

            //start from root definitions
            var rootClasses = Mapping.ClassMappings.Where(m => m.IsRoot);
            foreach (var classMapping in rootClasses)
            {
                if (classMapping.PropertyMappings == null)
                    continue;

                var eType = classMapping.Type;
                if (eType == null)
                {
                    Debug.WriteLine("Type not found: " + classMapping.Class);
                    continue;
                }

                //root definitions will always have parent == null
                Store(workbook, classMapping, eType, null);
            }
        }

        private void Store(WorkbookPart workbook, ClassMapping mapping, ExpressType expType, IPersistEntity parent)
        {
            if (mapping.PropertyMappings == null)
                return;

            var context = parent == null ?
                new EntityContext(Model.Instances.OfType(expType.Name.ToUpper(), false).ToList()) { LeavesDepth = 1 } :
                mapping.GetContext(parent);

            if (!context.Leaves.Any()) return;

            var tableName = mapping.TableName ?? "Default";
            Sheet sheet = workbook.Workbook.Sheets.FirstOrDefault(x => (x as Sheet).Name == tableName) as Sheet;
            var workSheetPart = (WorksheetPart)workbook.GetPartById(sheet.Id);
            foreach (var leaveContext in context.Leaves)
            {
                Store(workSheetPart, leaveContext.Entity, mapping, expType, leaveContext, tableName);

                foreach (var childrenMapping in mapping.ChildrenMappings)
                {
                    Store(workbook, childrenMapping, childrenMapping.Type, leaveContext.Entity);
                }
            }
        }

        private void Store(WorksheetPart sheet, IPersistEntity entity, ClassMapping mapping, ExpressType expType, EntityContext context, string sheetName)
        {
            Row multiRow = new Row() { RowIndex = 0 };
            List<string> multiValues = null;
            PropertyMapping multiMapping = null;
            var row = GetRow(sheet, sheetName);

            //fix on "Special Case" Assembly Row to Entity mapping
            if ((context?.RootEntity != null) && (expType?.ExpressNameUpper == "TYPEORCOMPONENT")) //without CobieExpress reference and not using reflection this is as good as it gets to ID Assembly
            {
                RowNoToEntityLabelLookup[sheetName].Add(row.RowIndex, context.RootEntity.EntityLabel);
            }
            else
            {
                RowNoToEntityLabelLookup[sheetName].Add(row.RowIndex, entity.EntityLabel);
            }

            foreach (var propertyMapping in mapping.PropertyMappings)
            {
                object value = null;
                foreach (var path in propertyMapping.Paths)
                {
                    value = GetValue(entity, expType, path, context);
                    if (value != null) break;
                }
                if (value == null && propertyMapping.Status == DataStatus.Required)
                    value = propertyMapping.DefaultValue ?? "n/a";

                var isMultiRow = IsMultiRow(value, propertyMapping);
                if (isMultiRow)
                {
                    multiRow = row;
                    var values = new List<string>();
                    var enumerable = value as IEnumerable<string>;
                    if (enumerable != null)
                        values.AddRange(enumerable);

                    //get only first value and store it
                    var first = values.First();
                    Store(row, first, propertyMapping, sheet);

                    //set the rest for the processing as multiValue
                    values.Remove(first);
                    multiValues = values;
                    multiMapping = propertyMapping;
                }
                else
                {
                    Store(row, value, propertyMapping, sheet);
                }

            }

            //adjust width of the columns after the first and the eight row 
            //adjusting fully populated workbook takes ages. This should be almost all right
            if (row.RowIndex == 1 || row.RowIndex == 8)
                AdjustAllColumns(sheet, mapping, row);

            //it is not a multi row so return
            if ((multiRow != null && multiRow.RowIndex <= 1) || multiValues == null || !multiValues.Any())
                return;

            //add repeated rows if necessary
            foreach (var value in multiValues)
            {
                var rowNum = GetNextRowNum(sheetName);
                var copy = CopyRow(multiRow, rowNum);
                if (copy != null)
                {
                    SheetData sheetData = sheet.Worksheet.Elements<SheetData>().First();
                    sheetData.Append(copy);
                }
                Store(copy, value, multiMapping, sheet);
                RowNoToEntityLabelLookup[sheetName].Add(rowNum, entity.EntityLabel);
            }
        }

        private Row CopyRow(Row sourceRow, uint destinationIndex)
        {
            if (sourceRow != null)
            {
                // Clone the source row
                Row destinationRow = (Row)sourceRow.CloneNode(true);

                // Set the row index of the destination row
                destinationRow.RowIndex = destinationIndex;

                // Update the cell references in the cloned row
                foreach (Cell cell in destinationRow.Elements<Cell>())
                {
                    string cellReference = cell.CellReference;
                    cell.CellReference = new StringValue(cellReference.Substring(0, 1) + destinationIndex);
                }
                return destinationRow;
            }
            return null;
        }

        private bool IsMultiRow(object value, PropertyMapping mapping)
        {
            if (value == null)
                return false;
            if (mapping.MultiRow == MultiRow.None)
                return false;

            var values = value as IEnumerable<string>;
            if (values == null)
                return false;

            var strings = values.ToList();
            var count = strings.Count;
            if (count > 1 && mapping.MultiRow == MultiRow.Always)
                return true;

            var single = string.Join(Mapping.ListSeparator, strings);
            return single.Length > CellTextLimit && mapping.MultiRow == MultiRow.IfNecessary;
        }

        private void Store(Row row, object value, PropertyMapping mapping, WorksheetPart worksheetPart)
        {
            if (value == null)
                return;

            Cell cell = row.Elements<Cell>()?.FirstOrDefault(x => x.CellReference == mapping.Column + row.RowIndex.Value);
            if (cell is null)
            {
                cell = new Cell() { CellReference = mapping.Column + row.RowIndex.Value };
                row.Append(cell);
            }
            //set column style to cell

            Columns columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
            Column column = columns.Elements<Column>().FirstOrDefault(c =>
            c.Min == GetColumnIndexFromString(mapping.Column) &&
            c.Max == GetColumnIndexFromString(mapping.Column)
            );

            cell.StyleIndex = column.Style;

            //simplify any eventual enumeration into a single string
            var enumVal = value as IEnumerable;
            if (enumVal != null && !(value is string))
            {
                var strValue = string.Join(Mapping.ListSeparator, enumVal.Cast<object>());
                if (string.IsNullOrEmpty(strValue))
                    return;
                cell.DataType = CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(strValue);
                return;
            }

            //string value
            var str = value as string;
            if (str != null)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(str);
                return;
            }

            //numeric point types
            if (value is double || value is float || value is int || value is long || value is short || value is byte || value is uint || value is ulong ||
                value is ushort)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(Convert.ToDouble(value));
                return;
            }

            //boolean value
            if (value is bool)
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue((bool)value);
                return;
            }

            //enumeration
            if (value is Enum)
            {
                var eType = value.GetType();
                var eValue = Enum.GetName(eType, value);
                var alias = GetEnumAlias(eType, eValue);

                cell.DataType = CellValues.String;
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(alias ?? eValue);
                return;
            }

            throw new NotSupportedException("Only base types are supported");
        }

        private string GetEnumAlias(Type type, string value)
        {
            string result;
            return _enumAliasesCache.TryGetValue(type.Name + "." + value, out result) ? result : null;
        }

        private ClassMapping GetTable(ExpressType type)
        {
            ClassMapping mapping;
            if (_typeClassMappingsCache.TryGetValue(type, out mapping))
                return mapping;

            var mappings = Mapping.ClassMappings.Where(m => m.Type == type || m.Type.AllSubTypes.Contains(type)).ToList();
            if (!mappings.Any())
                throw new XbimException("No sheet mapping defined for " + type.Name);

            var root = mappings.FirstOrDefault(m => m.IsRoot);
            if (root != null)
            {
                _typeClassMappingsCache.Add(type, root);
                return root;
            }

            mapping = mappings.FirstOrDefault();
            _typeClassMappingsCache.Add(type, mapping);
            return mapping;
        }

        private readonly Dictionary<string, ExpressType> _tableTypeCache = new Dictionary<string, ExpressType>();

        internal ExpressType GetType(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
                return null;

            ExpressType type;
            if (_tableTypeCache.TryGetValue(tableName, out type))
                return type;

            var mapping = Mapping.ClassMappings.FirstOrDefault(
                m => string.Equals(m.TableName, tableName, StringComparison.OrdinalIgnoreCase));
            if (mapping != null)
                type = MetaData.ExpressType(mapping.Class.ToUpper());
            _tableTypeCache.Add(tableName, type);
            return type;
        }

        private object GetValue(IPersistEntity entity, ExpressType type, string path, EntityContext context)
        {
            while (true)
            {
                if (string.IsNullOrWhiteSpace(path))
                    return null;

                //if it is parent, skip to the root of the context
                //optimization: check first letter before StartsWith() function. 
                if (path[0] == 'p' && path.StartsWith("parent."))
                {
                    if (context == null)
                        return null;

                    path = path.Substring(7); //trim "parent." from the beginning
                    entity = context.RootEntity;
                    type = entity.ExpressType;
                    context = null;
                    continue;
                }

                //one level up in the context hierarchy
                //optimization: check first letter before StartsWith() function. 
                if (path[0] == '(' && path.StartsWith("()."))
                {
                    if (context == null)
                        return null;

                    path = path.Substring(3); //trim "()." from the beginning
                    entity = context.Parent.Entity;
                    type = entity.ExpressType;
                    context = context.Parent;
                    continue;
                }

                if (string.Equals(path, "[table]", StringComparison.Ordinal))
                {
                    var mapping = GetTable(type);
                    return mapping.TableName;
                }

                if (string.Equals(path, "[type]", StringComparison.Ordinal))
                {
                    return entity.ExpressType.ExpressName;
                }

                var parts = path.Split('.');
                var multiResult = new List<string>();
                for (var i = 0; i < parts.Length; i++)
                {
                    var value = GetPropertyValue(parts[i], entity, type);

                    if (value == null)
                        return null;

                    var ent = value as IPersistEntity;
                    if (ent != null)
                    {
                        entity = ent;
                        type = ent.ExpressType;
                        continue;
                    }

                    var expVal = value as IExpressValueType;
                    if (expVal != null)
                    {
                        //if the type of the value is what we want
                        if (i < parts.Length - 1 && parts[parts.Length - 1] == "[type]")
                            return expVal.GetType().Name;
                        //return actual value as an underlying system type
                        return expVal.Value;
                    }

                    var expValEnum = value as IEnumerable<IExpressValueType>;
                    if (expValEnum != null)
                        return expValEnum.Select(v => v.Value);

                    var entEnum = value as IEnumerable<IPersistEntity>;
                    //it must be a simple value
                    if (entEnum == null) return value;

                    //it is a multivalue result
                    var subParts = parts.ToList().GetRange(i + 1, parts.Length - i - 1);
                    var subPath = string.Join(".", subParts);
                    foreach (var persistEntity in entEnum)
                    {
                        var subValue = GetValue(persistEntity, persistEntity.ExpressType, subPath, null);
                        if (subValue == null) continue;
                        var subString = subValue as string;
                        if (subString != null)
                        {
                            multiResult.Add(subString);
                            continue;
                        }
                        multiResult.Add(subValue.ToString());
                    }
                    return multiResult;

                }

                //if there is only entity itself to return, try to get 'Name' or 'Value' property as a fallback
                return GetFallbackValue(entity, type);
            }
        }

        internal ExpressMetaProperty GetProperty(ExpressType type, string name)
        {
            ExpressMetaProperty property;
            Dictionary<string, ExpressMetaProperty> properties;
            if (!_typePropertyCache.TryGetValue(type, out properties))
            {
                properties = new Dictionary<string, ExpressMetaProperty>();
                _typePropertyCache.Add(type, properties);
            }
            if (properties.TryGetValue(name, out property))
                return property;

            property = type.Properties.Values.FirstOrDefault(p => p.Name == name) ??
                    type.Inverses.FirstOrDefault(p => p.Name == name) ??
                    type.Derives.FirstOrDefault(p => p.Name == name);
            if (property == null)
                return null;

            properties.Add(name, property);
            return property;
        }

        private object GetPropertyValue(string pathPart, IPersistEntity entity, ExpressType type)
        {
            var propName = pathPart;
            var ofType = GetPropertyTypeOf(ref propName);
            var propIndex = GetPropertyIndex(ref propName);
            var pInfo = GetPropertyInfo(propName, type, propIndex);
            var value = pInfo.GetValue(entity, propIndex == null ? null : new[] { propIndex });

            if (ofType == null || value == null) return value;

            var vType = value.GetType();
            if (!typeof(IEnumerable).IsAssignableFrom(vType))
                return ofType.IsAssignableFrom(vType) ? value : null;

            var ofTypeMethod = vType.GetMethod("OfType");
            ofTypeMethod = ofTypeMethod.MakeGenericMethod(ofType);
            return ofTypeMethod.Invoke(value, null);
        }

        internal PropertyInfo GetPropertyInfo(string name, ExpressType type, object index)
        {
            var isIndexed = index != null;
            PropertyInfo pInfo;
            if (isIndexed && string.IsNullOrWhiteSpace(name))
            {
                pInfo = type.Type.GetProperty("Item"); //anonymous index accessors are automatically named 'Item' by compiler
                if (pInfo == null)
                    throw new XbimException(string.Format("{0} doesn't have an index access", type.Name));

                var iParams = pInfo.GetIndexParameters();
                if (iParams.All(p => p.ParameterType != index.GetType()))
                    throw new XbimException(string.Format("{0} doesn't have an index access for type {1}", type.Name, index.GetType().Name));
            }
            else
            {
                var expProp = GetProperty(type, name);
                if (expProp == null)
                    throw new XbimException(string.Format("It wasn't possible to find property {0} in the object of type {1}", name, type.Name));
                pInfo = expProp.PropertyInfo;
                if (isIndexed)
                {
                    var iParams = pInfo.GetIndexParameters();
                    if (iParams.All(p => p.ParameterType != index.GetType()))
                        throw new XbimException(string.Format("Property {0} in the object of type {1} doesn't have an index access for type {2}", name, type.Name, index.GetType().Name));
                }
            }
            return pInfo;
        }

        public Type GetPropertyTypeOf(ref string pathPart)
        {
            var isTyped = pathPart.Contains("\\");
            if (!isTyped) return null;

            var match = TypeOfRegex.Match(pathPart);
            pathPart = match.Groups["name"].Value;
            var typeName = match.Groups["type"].Value;
            var eType = MetaData.ExpressType(typeName.ToUpper());
            return eType != null ? eType.Type : null;
        }

        //static precompiled regular expressions
        private static readonly Regex TypeOfRegex = new Regex("((?<name>).+)?\\\\(?<type>.+)", RegexOptions.Compiled);
        private static readonly Regex PropertyIndexRegex = new Regex("((?<name>).+)?\\[(?<index>.+)\\]", RegexOptions.Compiled);

        public static object GetPropertyIndex(ref string pathPart)
        {
            var isIndexed = pathPart.Contains("[") && pathPart.Contains("]");
            if (!isIndexed) return null;

            object propIndex;
            var match = PropertyIndexRegex.Match(pathPart);

            var indexString = match.Groups["index"].Value;
            pathPart = match.Groups["name"].Value;

            if (indexString.Contains("'") || indexString.Contains("\""))
            {
                propIndex = indexString.Trim('\'', '"');
            }
            else
            {
                //try to convert it to integer access
                int indexInt;
                if (int.TryParse(indexString, out indexInt))
                    propIndex = indexInt;
                else
                    propIndex = indexString;
            }

            return propIndex;
        }

        private static object GetFallbackValue(IPersistEntity entity, ExpressType type)
        {
            var nameProp = type.Properties.Values.FirstOrDefault(p => p.Name == "Name");
            var valProp = type.Properties.Values.FirstOrDefault(p => p.Name == "Value");
            if (nameProp == null && valProp == null)
                return entity.ToString();

            if (nameProp != null && valProp != null)
            {
                var nValue = nameProp.PropertyInfo.GetValue(entity, null);
                var vValue = valProp.PropertyInfo.GetValue(entity, null);
                if (nValue != null && vValue != null)
                    return string.Join(":", vValue, nValue);
            }

            if (nameProp != null)
            {
                var nameValue = nameProp.PropertyInfo.GetValue(entity, null);
                if (nameValue != null)
                    return nameValue.ToString();
            }

            if (valProp != null)
            {
                var valValue = valProp.PropertyInfo.GetValue(entity, null);
                if (valValue != null)
                    return valValue.ToString();
            }
            return entity.ToString();
        }

        private Row GetRow(WorksheetPart sheet, string sheetName)
        {
            SheetData sheetData = sheet.Worksheet.Elements<SheetData>().First();
            //get the next row in rowNumber is less than 1 or use the argument to get or create new row
            uint lastIndex;
            if (!_rowNumCache.TryGetValue(sheetName, out lastIndex))
            {
                lastIndex = 0;
                _rowNumCache.Add(sheetName, 0);
            }
            var row = lastIndex < 1
                ? GetNextEmptyRow(sheet)
                : sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == lastIndex + 1);

            if (row is null)
            {
                row = new Row() { RowIndex = (uint)lastIndex + 1 };
                sheetData.Append(row);
            }
            if (row.RowIndex == 1)
            {
                row = new Row() { RowIndex = 2 };
                sheetData.Append(new Row() { RowIndex = 2 });
            }
            //cache the latest row index
            _rowNumCache[sheetName] = row.RowIndex;
            return row;
        }

        private uint GetNextRowNum(string sheetName)
        {
            uint lastIndex;
            //no raws were created in this sheet so far
            if (!_rowNumCache.TryGetValue(sheetName, out lastIndex))
                return 0;

            lastIndex++;
            _rowNumCache[sheetName] = lastIndex;
            return lastIndex;
        }

        private static Row GetNextEmptyRow(WorksheetPart sheet)
        {
            SheetData sheetData = sheet.Worksheet.Elements<SheetData>().First();
            int lastIndex = 0;
            foreach (Row row in sheetData.Elements<Row>())
            {
                lastIndex++;
                var isEmpty = true;
                foreach (Cell cell in row)
                {
                    if (string.IsNullOrEmpty(cell.CellValue.InnerText)) continue;

                    isEmpty = false;
                    break;
                }
                if (isEmpty) return row;
            }
            var newRow = new Row() { RowIndex = (uint)lastIndex + 1 };
            sheetData.Append(newRow);
            return newRow;
        }



        private void SetUpHeader(WorksheetPart sheetPart, WorkbookPart workbook, ClassMapping classMapping)
        {
            SheetData sheetData = sheetPart.Worksheet.Elements<SheetData>().First();
            //var workbook = sheetPart.Workbook;
            var row = sheetData?.Elements<Row>()?.FirstOrDefault();
            if (row is null)
            {
                row = new Row() { RowIndex = 1 };
                sheetData.Append(row);
            }
            InitMappingColumns(classMapping);
            CacheColumnIndices(classMapping);

            //freeze header row
            //SheetViews sheetViews = new SheetViews();
            //SheetView sheetView = new SheetView();
            //Pane pane = new Pane() { VerticalSplit = 0, HorizontalSplit = 1, TopLeftCell = "A1", ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen };

            //sheetView.Append(pane);
            //sheetViews.Append(sheetView);
            //sheetPart.Worksheet.Append(sheetViews);

            //create header and column style for every mapped column
            foreach (var mapping in classMapping.PropertyMappings)
            {
                Cell cell = row.Elements<Cell>()?.FirstOrDefault(x => x.CellReference == mapping.Column + row.RowIndex.Value);
                if (cell is null)
                {
                    cell = new Cell() { CellReference = mapping.Column + (int)row.RowIndex.Value, DataType = CellValues.String };
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(mapping.Header);
                    cell.StyleIndex = GetStyleIndex(DataStatus.Header, workbook.WorkbookStylesPart.Stylesheet);
                    // Add the cell to the row
                    row.Append(cell);
                }

                //set default column style if not defined but available
                if (mapping.Status == DataStatus.None) continue;

                Columns columns = sheetPart.Worksheet.GetFirstChild<Columns>();

                if (columns == null)
                {
                    columns = new Columns();
                    sheetPart.Worksheet.InsertAt(columns, 0);
                }

                Column column = columns.Elements<Column>().FirstOrDefault(c => c.Min == GetColumnIndexFromString(mapping.Column) && c.Max == mapping.ColumnIndex);

                if (column == null)
                {
                    column = new Column { Min =(uint)mapping.ColumnIndex, Max = (uint)mapping.ColumnIndex, CustomWidth = true, Width = 15 };
                    columns.Append(column);
                }
                if (mapping.Hidden)
                    column.Hidden = true;

                var existStyle = column.Style;
                if (
                    existStyle != null
                    ) continue;
                column.Style = GetStyleIndex(mapping.Status, workbook.WorkbookStylesPart.Stylesheet);
            }

            //set up filter
            var lastPropMap = classMapping.PropertyMappings.OrderBy(p => p.ColumnIndex).LastOrDefault();

            AutoFilter autoFilter = new AutoFilter { Reference = new StringValue($"A1:{lastPropMap.Column}1") };

            // Check if AutoFilter exists, if not, add it
            if (sheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault() == null)
            {
                sheetPart.Worksheet.Append(autoFilter);
            }
            else // If AutoFilter exists, update its reference
            {
                sheetPart.Worksheet.Elements<AutoFilter>().First().Reference = autoFilter.Reference;
            }
        }

        private static void InitMappingColumns(ClassMapping mapping)
        {
            if (mapping.PropertyMappings == null ||
                !mapping.PropertyMappings.Any() ||
                mapping.PropertyMappings.All(m => !string.IsNullOrWhiteSpace(m.Column)))
                return;

            var letter = 'A';
            var number = (int)letter;
            foreach (var pMapping in mapping.PropertyMappings)
            {
                pMapping.Column = ((char)number++).ToString();
            }

        }

        private uint GetStyleIndex(DataStatus status, Stylesheet stylesheet)
        {

            if (_styles == null)
                _styles = new Dictionary<DataStatus, uint>();

            uint styleIndex;
            if (_styles.TryGetValue(status, out styleIndex))
                return styleIndex;

            var representation = Mapping.StatusRepresentations.FirstOrDefault(r => r.Status == status);


            if (representation == null)
            {
                _styles.Add(status, styleIndex);
                return styleIndex;
            }


            Fill fill = new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Indexed = (uint)GetClosestColour(representation.Colour) } });

            if (representation.Border)
            {

                var borders = stylesheet.Elements<Borders>().FirstOrDefault();

                Border border = new Border(
                    new LeftBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new RightBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } });

                if (stylesheet.Borders.Count == null)
                {
                    stylesheet.Borders.Count = 0;
                }
                stylesheet.Borders.Count++;
                stylesheet.Borders.Append(border);
            }

            if (stylesheet.Fills.Count == null)
            {
                stylesheet.Fills.Count = 0;
            }
            stylesheet.Fills.Count++;
            stylesheet.Fills.Append(fill);

            Font font = new Font();

            switch (representation.FontWeight)
            {
                case FontWeight.Normal:
                    break;
                case FontWeight.Bold:
                    font.Bold = new Bold();
                    break;
                case FontWeight.Italics:
                    font.Italic = new DocumentFormat.OpenXml.Spreadsheet.Italic();
                    break;
                case FontWeight.BoldItalics:
                    font.Bold = new Bold();
                    font.Italic = new DocumentFormat.OpenXml.Spreadsheet.Italic();
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
            if (stylesheet.Fonts.Count == null)
            {
                stylesheet.Fonts.Count = 0;
            }
            stylesheet.Fonts.Count++;
            stylesheet.Fonts.Append(font);



            stylesheet.CellFormats.AppendChild(new CellFormat()
            {
                FormatId = stylesheet.Borders.Count - 1,
                FontId = stylesheet.Fonts.Count - 1,
                BorderId = stylesheet.Borders.Count - 1,
                FillId = stylesheet.Fills.Count - 1,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyFont = true
            }
            );
            if (stylesheet.CellFormats.Count == null)
            {
                stylesheet.CellFormats.Count = 1;
            }
            stylesheet.CellFormats.Count++;

            styleIndex = (uint)stylesheet.CellFormats.Count - 1;
            _styles.Add(status, styleIndex);

            return styleIndex;
        }

        //This operation takes very long time if applied at the end when spreadsheet is fully populated
        private static void AdjustAllColumns(WorksheetPart sheet, ClassMapping mapping, Row row)
        {
            var columns = sheet.Worksheet.GetFirstChild<Columns>();
            var cells = row.Elements<Cell>();
            foreach (var col in mapping.PropertyMappings)
            {
                var cellWidth = GetCellWidth(cells.FirstOrDefault(x => x.CellReference == col.Column + row.RowIndex.Value));
                columns.Elements<Column>().FirstOrDefault(x => x.Min == col.ColumnIndex && x.Max == col.ColumnIndex).Width = cellWidth < 15 ? 15 : cellWidth;

            }
        }
        static double GetCellWidth(Cell cell)
        {
            double cellWidth = 0;

            if (cell != null && cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                // Calculate width based on length of inline string
                if (cell.InlineString != null && cell.InlineString.HasChildren)
                {
                    foreach (OpenXmlElement element in cell.InlineString.ChildElements)
                    {
                        if (element is Text)
                        {

                            string text = ((Text)element).Text;
                            cellWidth += text.Length * 1.4; // Adjust this value based on your font and size
                        }
                    }
                }
            }
            else if (cell != null && cell.CellValue != null)
            {
                // Calculate width based on length of cell value
                string text = cell.CellValue.Text;
                cellWidth = text.Length * 1.4; // Adjust this value based on your font and size
            }
            return cellWidth;
        }
        private void SetUpTables(WorkbookPart workbook, ModelMapping mapping)
        {
            if (mapping == null || mapping.ClassMappings == null || !mapping.ClassMappings.Any())
                return;

            var i = 0;
            foreach (var classMapping in Mapping.ClassMappings.Where(classMapping => string.IsNullOrWhiteSpace(classMapping.TableName)))
            {
                classMapping.TableName = string.Format("NoName({0})", i++);
            }

            var names = Mapping.ClassMappings.OrderBy(m => m.TableOrder).Select(m => m.TableName).Distinct();
            uint count = 1;
            Sheets sheets = workbook.Workbook.AppendChild(new Sheets());

            foreach (var name in names)
            {
                WorksheetPart worksheetPart = workbook.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add a new sheet to the workbook
                Sheet sheet = new Sheet() { Id = workbook.GetIdOfPart(worksheetPart), SheetId = count, Name = name };
                sheets.Append(sheet);

                RowNoToEntityLabelLookup.Add(sheet.Name, new Dictionary<uint, int>());
                var classMapping = Mapping.ClassMappings.First(m => m.TableName == name);
                SetUpHeader(worksheetPart, workbook, classMapping);
                count++;
                ////set colour of the tab: Not implemented exception in NPOI
                //if (classMapping.TableStatus == DataStatus.None) continue;
                //var style = GetStyle(classMapping.TableStatus, workbook);
                //sheet.TabColorIndex = style.FillForegroundColor;
            }
        }

        private static readonly Dictionary<string, short> ColourCodeCache = new Dictionary<string, short>();
        private static readonly List<IndexedColor> IndexedColoursList = new List<IndexedColor>();

        private static short GetClosestColour(string rgb)
        {
            if (!IndexedColoursList.Any())
            {
                var props = typeof(IndexedColor).GetFields(BindingFlags.Static | BindingFlags.Public).Where(p => p.FieldType == typeof(IndexedColor));
                foreach (var info in props)
                {
                    IndexedColoursList.Add((IndexedColor)info.GetValue(null));
                }
            }

            if (string.IsNullOrWhiteSpace(rgb))
                return IndexedColor.Automatic.Index;
            rgb = rgb.Trim('#').Trim();
            short result;
            if (ColourCodeCache.TryGetValue(rgb, out result))
                return result;

            var triplet = rgb.Length == 3;
            var hR = triplet ? rgb.Substring(0, 1) + rgb.Substring(0, 1) : rgb.Substring(0, 2);
            var hG = triplet ? rgb.Substring(1, 1) + rgb.Substring(1, 1) : rgb.Substring(2, 2);
            var hB = triplet ? rgb.Substring(2, 1) + rgb.Substring(1, 1) : rgb.Substring(4, 2);

            var r = Convert.ToByte(hR, 16);
            var g = Convert.ToByte(hG, 16);
            var b = Convert.ToByte(hB, 16);

            var rgbBytes = new[] { r, g, b };
            var distance = double.NaN;
            var colour = IndexedColor.Automatic;
            foreach (var col in IndexedColoursList)
            {
                var dist = ColourDistance(rgbBytes, col.RGB);
                if (double.IsNaN(distance)) distance = dist;

                if (!(distance > dist)) continue;
                distance = dist;
                colour = col;
            }
            ColourCodeCache.Add(rgb, colour.Index);
            return colour.Index;
        }

        private static double ColourDistance(byte[] a, byte[] b)
        {
            return Math.Sqrt(Math.Pow(a[0] - b[0], 2) + Math.Pow(a[1] - b[1], 2) + Math.Pow(a[2] - b[2], 2));
        }
        #endregion
    }

    public enum ExcelTypeEnum
    {
        XLS,
        XLSX
    }
}
