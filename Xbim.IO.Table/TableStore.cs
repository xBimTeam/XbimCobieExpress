using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
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

        /// <summary>
        /// Stores the <see cref="Model"/> in an Excel file according to the <see cref="Mapping"/> table.
        /// </summary>
        /// <param name="path">The filename of the Excel file to create.</param>
        /// <param name="template">A readonly stream containing an Excel template document which will be cloned to used as the base content</param>
        /// <param name="recalculate">Indicates if excel formulas should be refreshed on completion</param>
        /// <remarks>Note: any existing sheets in any template with the same name as target sheets will be replaced</remarks>
        /// <exception cref="ArgumentNullException"></exception>
        public void Store(string path, Stream template = null, bool recalculate = true)
        {
            if (path == null)
                throw new ArgumentNullException("path");
            var ext = Path.GetExtension(path).ToLower().Trim('.');
            if (ext != "xls" && ext != "xlsx")
            {
                // XLSX is Spreadsheet XML representation which is capable of storing more data
                path += ".xlsx";
                ext = "xlsx";
            }
            using (var file = File.Create(path))
            {
                var type = ext == "xlsx" ? ExcelTypeEnum.XLSX : ExcelTypeEnum.XLS;
                Store(file, type, template, recalculate);
                file.Close();
            }


        }

        /// <summary>
        /// Stores the <see cref="Model"/> in Excel format to the provided stream based on the <see cref="Mapping"/> table.
        /// </summary>
        /// <param name="stream">A writeable stream to output the Excel file to</param>
        /// <param name="type">The Excel file format version</param>
        /// <param name="template">An options stream containing a template to clone as the base document</param>
        /// <param name="recalculate">Indicates if excel formulas should be refreshed on completion</param>
        /// <returns>A <paramref name="stream"/></returns>
        /// <remarks>Note: any existing sheets in any template with the same name as target sheets will be replaced</remarks>
        public Stream Store(Stream stream, ExcelTypeEnum type, Stream template = null, bool recalculate = true)
        {
            Log = new StringWriter();
            SpreadsheetDocument spreadsheetDocument;
            if (template != null)
            {
                var templateFile = SpreadsheetDocument.Open(template, false);
                spreadsheetDocument = templateFile.Clone(stream, true);
                templateFile.Dispose();
            }
            else
            {
                spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            }

            WorkbookPart workbookPart = spreadsheetDocument.GetOrCreatePart(d => d.AddWorkbookPart(), d => d.WorkbookPart);
            
            if (workbookPart.Workbook == null)
            {
                workbookPart.Workbook = new Workbook();
            }

            //Add a WorkbookStylesPart to the WorkbookPart
            WorkbookStylesPart stylesPart = workbookPart.GetOrCreatePart(w => w.AddNewPart<WorkbookStylesPart>());

            var styleSheet = stylesPart.Stylesheet ??= new Stylesheet();

            styleSheet.Fonts ??= new Fonts();
            styleSheet.Fills ??= new Fills();
            styleSheet.Borders ??= new Borders();
            styleSheet.CellFormats ??= new CellFormats();

            styleSheet.CellFormats.GetOrCreate(w => w.AppendChild(new CellFormat(){ FillId = 0, BorderId=0, FontId = 0, FormatId = 0 }));

            SetStyles(styleSheet);
            //create spreadsheet representation 
            SerialiseModel(workbookPart);

            if(recalculate && workbookPart.Workbook.CalculationProperties != null)
            {
                workbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                workbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            }

            // Validate
            var validator = new OpenXmlValidator();
            var err = validator.Validate(spreadsheetDocument);


            //write to output stream
            spreadsheetDocument.Save();
            spreadsheetDocument.Dispose();
            return stream;
        }


        /// <summary>
        /// Serialise the model content in the supplied <see cref="WorkbookPart"/>
        /// </summary>
        /// <param name="workbookPart"></param>
        private void SerialiseModel(WorkbookPart workbookPart)
        {
            //if there are no mappings do nothing
            if (Mapping.ClassMappings == null || !Mapping.ClassMappings.Any()) return;

            _rowNumCache = new Dictionary<string, uint>();
            if(_styles == null) 
                _styles = new Dictionary<DataStatus, uint>();

            //creates tables in defined order if they are not there yet
            SetUpTables(workbookPart, Mapping);

            //start from root definitions - those that don't have a parent.
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
                SerialiseSheet(workbookPart, classMapping, eType, null);
            }
        }


        private void SerialiseSheet(WorkbookPart workbookPart, ClassMapping mapping, ExpressType expType, IPersistEntity parent)
        {
            if (mapping.PropertyMappings == null)
                return;

            var context = parent == null ?
                new EntityContext(Model.Instances.OfType(expType.Name.ToUpper(), false).ToList()) { LeavesDepth = 1 } :
                mapping.GetContext(parent);

            var tableName = mapping.TableName ?? "Default";
            Sheet sheet = workbookPart.Workbook.Sheets.FirstOrDefault(x => (x as Sheet).Name == tableName) as Sheet;
            var workSheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            if (!context.Leaves.Any())
            {
                if(parent is null)
                { 
                    // On root tabs, set Tab colour to 'empty' when no rows
                    SetTabColour(workSheetPart, "#AAAAAA");
                }
                return;
            }

            foreach (var leafContext in context.Leaves)
            {
                SerialiseEntity(workSheetPart, leafContext.Entity, mapping, expType, leafContext, tableName);

                foreach (var childMapping in mapping.ChildrenMappings)
                {
                    // E.g. tables with a parentPath such as COBie Attributes/Documents/Impacts, which are not 'Roots'
                    SerialiseSheet(workbookPart, childMapping, childMapping.Type, leafContext.Entity);
                }
                //workSheetPart.Worksheet.Save();
            }
        }

        private void SerialiseEntity(WorksheetPart worksheetPart, IPersistEntity entity, ClassMapping mapping, ExpressType expType, EntityContext context, string sheetName)
        {
            Row multiRow = new Row() { RowIndex = 0 };
            List<string> multiValues = null;
            PropertyMapping multiMapping = null;
            var row = GetRow(worksheetPart, sheetName);

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
                    SerialiseCell(row, first, propertyMapping, worksheetPart);

                    //set the rest for the processing as multiValue
                    values.Remove(first);
                    multiValues = values;
                    multiMapping = propertyMapping;
                }
                else
                {
                    SerialiseCell(row, value, propertyMapping, worksheetPart);
                }

            }

            // adjust width of the columns for the initial rows and then sample every 100
            if (row.RowIndex <= 8 || row.RowIndex % 100 == 0)
                AdjustAllColumnsWidths(worksheetPart, mapping, row, row.RowIndex <= 2);

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
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    sheetData.Append(copy);
                }
                SerialiseCell(copy, value, multiMapping, worksheetPart);
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

        private void SerialiseCell(Row row, object value, PropertyMapping mapping, WorksheetPart worksheetPart)
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

            if (column != null)
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

        private Row GetRow(WorksheetPart worksheetPart, string sheetName)
        {
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            //get the next row in rowNumber is less than 1 or use the argument to get or create new row
            uint lastIndex;
            if (!_rowNumCache.TryGetValue(sheetName, out lastIndex))
            {
                lastIndex = 0;
                _rowNumCache.Add(sheetName, 0);
            }
            var row = lastIndex < 1
                ? GetNextEmptyRow(worksheetPart)
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

        private static Row GetNextEmptyRow(WorksheetPart worksheetPart)
        {
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            int lastIndex = 0;
            foreach (Row row in sheetData.Elements<Row>())
            {
                lastIndex++;
                var isEmpty = true;
                foreach (Cell cell in row)
                {
                    if (string.IsNullOrEmpty(cell.CellValue?.InnerText)) continue;

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

                    cell.StyleIndex = GetOrSetStyleIndex(DataStatus.Header, workbook.WorkbookStylesPart.Stylesheet);

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
                    column = new Column { Min = (uint)mapping.ColumnIndex, Max = (uint)mapping.ColumnIndex, CustomWidth = true, Width = 15 };
                    columns.Append(column);
                }
                if (mapping.Hidden)
                    column.Hidden = true;

                column.Style = GetOrSetStyleIndex(mapping.Status, workbook.WorkbookStylesPart.Stylesheet);
                if (mapping.IsKey == true && mapping.Status == DataStatus.Required)
                {
                    // Create named ranges for Keys columns
                    DefineNamedKeys(workbook, classMapping, mapping);
                }
                if (!string.IsNullOrEmpty(mapping.LookUp))
                {
                    // Create Data validations back to keyed columns
                    AddDataValidation(sheetPart, mapping);
                }
            }

            AddAutofilters(sheetPart, classMapping, sheetData);
        }

        private static void AddAutofilters(WorksheetPart sheetPart, ClassMapping classMapping, SheetData sheetData)
        {
            //set up filter
            var lastPropMap = classMapping.PropertyMappings.OrderBy(p => p.ColumnIndex).LastOrDefault();

            AutoFilter autoFilter = new AutoFilter { Reference = new StringValue($"A1:{lastPropMap.Column}1") };

            // Check if AutoFilter exists, if not, add it
            if (sheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault() == null)
            {
                sheetPart.Worksheet.InsertAfter(autoFilter, sheetData);
            }
            else // If AutoFilter exists, update its reference
            {
                sheetPart.Worksheet.Elements<AutoFilter>().First().Reference = autoFilter.Reference;
            }
        }

        private static void DefineNamedKeys(WorkbookPart workbook, ClassMapping classMapping, PropertyMapping mapping)
        {
            DefinedNames definedNames = workbook.Workbook.GetOrCreate(c => c.AppendChild(new DefinedNames()));
            var keyName = $"{classMapping.TableName}.{mapping.Header}";
            var appliesToRange = $"${mapping.Column}:${mapping.Column}";
            var targetRange = $"{classMapping.TableName}!{appliesToRange}"; // Range to cover
            var definition = definedNames.ChildElements.OfType<DefinedName>().FirstOrDefault(c => c.Name == keyName);
            if (definition != null)
            {
                definition.Text = targetRange;
            }
            else
            {
                definition = new DefinedName()
                {
                    Name = keyName,
                    Text = targetRange
                };
                definedNames.Append(definition);
            }
        }

        private void AddDataValidation(WorksheetPart sheetPart, PropertyMapping mapping)
        {
            var lookup = mapping.LookUp;
            var lookupParts = lookup.Split('.');
            if (lookupParts.Length > 1)
            {
                var tableName = lookupParts[0];
                var columnName = lookupParts[1];

                var workSheet = sheetPart.Worksheet;
                DataValidations dataValidations = workSheet.GetOrCreate(
                        c => c.InsertBefore(new DataValidations(), workSheet.Descendants<PageMargins>().FirstOrDefault()),
                        getter => getter.GetFirstChild<DataValidations>());
                var appliesToRange = $"{mapping.Column}:{mapping.Column}";

                DataValidation dataValidation = new DataValidation()
                {
                    Type = DataValidationValues.List,
                    AllowBlank = true,
                    SequenceOfReferences = new ListValue<StringValue>() { InnerText = appliesToRange }
                };
                Formula1 formula = new Formula1();

                if (string.Compare(tableName, Mapping.PickTableName, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    // Picklists are usually just the 'Picklists.<namedRange>'
                    // TODO: Check for missing namedRanges and fall back to column header?
                    formula.Text = $"={columnName}";
                }
                else
                {
                    var targetSheet = Mapping.ClassMappings.FirstOrDefault(c => c.TableName == tableName);
                    if (targetSheet == null)
                    {
                        // Typically dynamic Sheet names e.g. [SheetName].Name
                        return;
                    }
                    formula.Text = $"={mapping.LookUp}";
                    
                }
                dataValidation.Append(formula);
                dataValidations.Append(dataValidation);
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

        private uint GetOrSetStyleIndex(DataStatus status, Stylesheet stylesheet)
        {

            if (_styles == null)
                _styles = new Dictionary<DataStatus, uint>();

            uint styleIndex;
            if (_styles.TryGetValue(status, out styleIndex))
                return _styles[status];

            var representation = Mapping.StatusRepresentations.FirstOrDefault(r => r.Status == status);


            if (representation == null)
            {
                //_styles.Add(status, 0);
                //return styleIndex;

                representation = new StatusRepresentation() { Colour = "#FF0000", Status = status };
            }
            if (stylesheet.Fills.Count == null)
            {
                stylesheet.Fills.Count = 0;
            }
            if (stylesheet.Fonts.Count == null)
            {
                stylesheet.Fonts.Count = 0;
            }
            if (stylesheet.CellFormats.Count == null)
            {
                stylesheet.CellFormats.Count = 1;
            }
            if (stylesheet.Borders.Count == null)
            {
                stylesheet.Borders.Count = 0;
            }
            
            // Fonts
            // Fills
            // Borders
            // Cell Formats
            Fill fill = new Fill(new PatternFill() 
            { 
                PatternType = PatternValues.Solid, 
                ForegroundColor = new ForegroundColor { Indexed = (uint)GetClosestColour(representation.Colour) } 
            });

            if (representation.Border)
            {

                Border border = new Border(
                    new LeftBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new RightBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new TopBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } },
                    new BottomBorder() { Style = BorderStyleValues.Thin, Color = new Color() { Indexed = (uint)IndexedColor.Black.Index } });

               
                stylesheet.Borders.Count++;
                stylesheet.Borders.Append(border);
            }
            else
            {
                Border border = new Border();
                stylesheet.Borders.Count++;
                stylesheet.Borders.Append(border);
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
           
            
            stylesheet.Fonts.Count++;
            stylesheet.Fonts.Append(font);


            stylesheet.CellFormats.Count++;
            var cellFormat = new CellFormat()
            {
               // FormatId = (uint)(stylesheet.CellFormats.Count - 1),
                FontId = (uint)(stylesheet.Fonts.Count -1),
                BorderId = (uint)(stylesheet.Borders.Count - 1),
                FillId = (uint)(stylesheet.Fills.Count - 1),
                ApplyBorder = true,
                ApplyFont = true,
                
            };

            if (status == DataStatus.Header)
            {
                // Rotate headers 90% and horizonally centre
                var alignment = new Alignment()
                {
                    TextRotation = 90,
                    Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center)
                };
                cellFormat.Alignment = alignment;
            }
            stylesheet.CellFormats.AppendChild(cellFormat);

            styleIndex = (uint)(stylesheet.CellFormats.Count - 1);
            _styles.Add(status, styleIndex);

            return styleIndex;
        }

        /// <summary>
        /// Register global styles defined by the <see cref="StatusRepresentation"/>s
        /// </summary>
        /// <param name="stylesheet"></param>
        private void SetStyles( Stylesheet stylesheet)
        {

            if (_styles == null)
                _styles = new Dictionary<DataStatus, uint>();

            // Add two dummy styles which we never use. OpenXml styles are a dumpster fire of magic indexex
            // and this just bumps the indexes on a bit so we can start 
            GetOrSetStyleIndex(DataStatus.None, stylesheet);
            GetOrSetStyleIndex(DataStatus.UserDefined, stylesheet);
            foreach (var representation in Mapping.StatusRepresentations)
            {
                GetOrSetStyleIndex(representation.Status, stylesheet);
            }

         
        }

        const double MinColumnWidth = 8;
        const double MaxColumnWidth = 50;
        const double ColumnPadding = 1.5;

        //This operation takes very long time if applied at the end when spreadsheet is fully populated
        private static void AdjustAllColumnsWidths(WorksheetPart sheet, ClassMapping mapping, Row row, bool initialResize = false)
        {
            var columns = sheet.Worksheet.GetFirstChild<Columns>();
            var cells = row.Elements<Cell>();
            foreach (var col in mapping.PropertyMappings)
            {
                var column = columns.Elements<Column>().FirstOrDefault(x => x.Min == col.ColumnIndex && x.Max == col.ColumnIndex);
                if(column != null)
                {
                    var cellWidth = GetCellWidth(cells.FirstOrDefault(x => x.CellReference == col.Column + row.RowIndex.Value));
                    var optimalWidth = Math.Min(Math.Max(cellWidth + ColumnPadding, MinColumnWidth), MaxColumnWidth);
                    // Set the width on first row, and grow it on subsequent.
                    if(initialResize || optimalWidth > column.Width)
                    {
                        column.BestFit = true;
                        column.Width = optimalWidth;
                    }
                }

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
                            cellWidth += text.Length * 1.0; // Adjust this value based on your font and size
                        }
                    }
                }
            }
            else if (cell != null && cell.CellValue != null)
            {
                // Calculate width based on length of cell value
                string text = cell.CellValue.Text;
                if(cell.DataType == CellValues.Number)
                {
                    // Account for Floating point precision. E.g. 2500d => "2500.000000000000000002"
                    try
                    {
                        var value = Convert.ToDouble(text);
                        text = Math.Round(value, 8).ToString();
                    }
                    catch(FormatException) { }
                }
                cellWidth = text.Length * 1.0; // Adjust this value based on your font and size
            }
            return cellWidth;
        }


        /// <summary>
        /// Set up sheets for the 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="mapping"></param>
        private void SetUpTables(WorkbookPart workbook, ModelMapping mapping)
        {
            if (mapping == null || mapping.ClassMappings == null || !mapping.ClassMappings.Any())
                return;

            Sheets sheets = workbook.Workbook.GetOrCreate(w => w.AppendChild(new Sheets()));

            uint count = (uint)sheets.Count() + 1;

            // Provide default unique sheet name for edgecase when tableName not supplied
            var i = 0;
            foreach (var classMapping in Mapping.ClassMappings.Where(classMapping => string.IsNullOrWhiteSpace(classMapping.TableName)))
            {
                classMapping.TableName = string.Format("{0}({1})", classMapping.Class ?? "NoName", i++);
            }

            var tableNames = Mapping.ClassMappings.OrderBy(m => m.TableOrder).Select(m => m.TableName).Distinct();

            foreach (var name in tableNames)
            {
                OpenXmlElement insertBefore = null;
                var classMapping = Mapping.ClassMappings.First(m => m.TableName == name);

                // Check for existing sheets (e.g. in Template)
                Sheet sheet = sheets.ChildElements.OfType<Sheet>().FirstOrDefault(s => s.Name == name);
                if (sheet != null)
                {
                    // Remove any matching sheet so we have a 'clean start'. 
                    // Saves us from having to map mis-placed columns, clear existing data etc.
                    insertBefore = sheet.NextSibling();
                    var sheetPart = (WorksheetPart)(workbook.GetPartById(sheet.Id));
                    workbook.DeletePart(sheetPart);
                    sheet.Remove();
                }

                WorksheetPart worksheetPart = workbook.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add a new sheet to the workbook
                sheet = new Sheet() { Id = workbook.GetIdOfPart(worksheetPart), SheetId = count, Name = name };

                // Appends - or insert where insertBefore is not null (when updating a Template)
                sheets.InsertBefore(sheet, insertBefore);

                RowNoToEntityLabelLookup.Add(sheet.Name, new Dictionary<uint, int>());

                SetUpHeader(worksheetPart, workbook, classMapping);
                count++;

                // Initialise Tab colour for status/ May get updated later
                InitialiseTabColour(mapping, classMapping, worksheetPart);

            }
        }

        private static void InitialiseTabColour(ModelMapping mapping, ClassMapping classMapping, WorksheetPart worksheetPart)
        {
            var sheetProps = worksheetPart.Worksheet.SheetProperties;
            if (worksheetPart.Worksheet.SheetProperties == null)
                worksheetPart.Worksheet.SheetProperties = new SheetProperties();
            if (worksheetPart.Worksheet.SheetProperties.TabColor == null)
                worksheetPart.Worksheet.SheetProperties.TabColor = new TabColor();

            if (classMapping.TableStatus == DataStatus.None) return;

            var representation = mapping.StatusRepresentations.FirstOrDefault(r => r.Status == classMapping.TableStatus);
            if (representation != null)
            {
                SetTabColour(worksheetPart, representation.Colour);
            }
        }

        private static void SetTabColour(WorksheetPart worksheetPart, string colour)
        {
            var colorArgb = colour.Replace("#", "").ToUpperInvariant();

            if (colorArgb.Length == 6)
                colorArgb = $"FF{colorArgb}";   // Add Alpha
            else if (colorArgb.Length == 3)
                colorArgb = $"F{colorArgb}";   // Add Alpha
            worksheetPart.Worksheet.SheetProperties.TabColor.Rgb = HexBinaryValue.FromString(colorArgb);
        }

        private static readonly Dictionary<string, short> ColourCodeCache = new Dictionary<string, short>();
        private static List<IndexedColor> IndexedColoursList
        {
            get => LazyColoursList.Value;
        }

        private static Lazy<List<IndexedColor>> LazyColoursList = new Lazy<List<IndexedColor>>(() => 
        {
            var props = typeof(IndexedColor).GetFields(BindingFlags.Static | BindingFlags.Public).Where(p => p.FieldType == typeof(IndexedColor));
            return props.Select(p => (IndexedColor)p.GetValue(null)).ToList();
        });

        private static short GetClosestColour(string rgb)
        {
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
