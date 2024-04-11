using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Xbim.Common;
using Xbim.Common.Metadata;

namespace Xbim.IO.Table
{
    public partial class TableStore
    {
        public TextWriter Log { get; private set; }

        #region Reading in from a spreadsheet
        public void LoadFrom(string path)
        {
            if (path == null)
                throw new ArgumentNullException("path");

            Log = new StringWriter();

            var ext = Path.GetExtension(path).ToLower().Trim('.');
            if (ext != "xls" && ext != "xlsx")
            {
                //XLSX is Spreadsheet XML representation which is capable of storing more data
                path += ".xlsx";
                ext = "xlsx";
            }
            var type = ext == "xlsx" ? ExcelTypeEnum.XLSX : ExcelTypeEnum.XLS;
            LoadFrom(path, type);

        }

        public void LoadFrom(string filePath, ExcelTypeEnum type)
        {

            WorkbookPart workbook;

            switch (type)
            {
                case ExcelTypeEnum.XLS:

                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                    {
                        workbook = spreadsheetDocument.WorkbookPart;
                        _multiRowIndicesCache = new Dictionary<string, int[]>();
                        _isMultiRowMappingCache = new Dictionary<ClassMapping, bool>();
                        _referenceContexts = new Dictionary<ClassMapping, ReferenceContext>();
                        _forwardReferences.Clear();
                        _forwardReferenceParentCache.Clear();
                        _globalEntities.Clear();
                        LoadFromWorkbook(workbook);
                    }
                    break;
                case ExcelTypeEnum.XLSX: //this is as it should be according to a standard
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                    {
                        workbook = spreadsheetDocument.WorkbookPart;
                        _multiRowIndicesCache = new Dictionary<string, int[]>();
                        _isMultiRowMappingCache = new Dictionary<ClassMapping, bool>();
                        _referenceContexts = new Dictionary<ClassMapping, ReferenceContext>();
                        _forwardReferences.Clear();
                        _forwardReferenceParentCache.Clear();
                        _globalEntities.Clear();
                        LoadFromWorkbook(workbook);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException("type");
            }



        }

        private void LoadFromWorkbook(WorkbookPart workbook)
        {
            //get all data tables
            if (Mapping.ClassMappings == null || !Mapping.ClassMappings.Any())
                return;

            var partialSheets = new List<WorksheetPart>();
            _sharedStringTable = workbook?.SharedStringTablePart?.SharedStringTable;
            foreach (Sheet worksheet in workbook.Workbook.Sheets)
            {
                var sheetName = worksheet.Name;
                var mapping =
                    Mapping.ClassMappings.FirstOrDefault(
                        m => string.Equals(sheetName, m.TableName, StringComparison.OrdinalIgnoreCase));
                if (mapping == null)
                    continue;
                var workSheetPart = (WorksheetPart)workbook.GetPartById(worksheet.Id);

                if (mapping.IsPartial)
                {
                    ProcessPartialSheets(workSheetPart, sheetName);
                    continue;
                }

                LoadFromSheet(workSheetPart, mapping);
            }




            //be happy
        }
        public void ResolveReferences()
        {

            //resolve references (don't use foreach as new references might be added to the queue during the processing)
            while (_forwardReferences.Count != 0)
            {
                var reference = _forwardReferences.Dequeue();
                reference.Resolve();
            }
        }
        private void ProcessPartialSheets(WorksheetPart sheet, string sheetName)
        {
            var mapping =
                Mapping.ClassMappings.FirstOrDefault(
                    m => string.Equals(sheetName, m.TableName, StringComparison.OrdinalIgnoreCase));
            if (mapping == null)
                return;
            SheetData sheetData = sheet.Worksheet.Elements<SheetData>().First();
            AdjustMapping(sheetData, mapping);
            CacheColumnIndices(mapping);
            var context = GetReferenceContext(mapping);
            var emptyRows = 0;
            foreach (Row row in sheetData.Elements<Row>())
            {
                //skip header row
                if (row == null || row.RowIndex == 1)
                    continue;
                var cells = row.Elements<Cell>().ToList();
                if (!cells.Any() ||
                 cells.All(c => c.DataType == null) ||
                     cells.All(c => c.DataType == CellValues.SharedString && string.IsNullOrWhiteSpace(c.InnerText)))
                {
                    emptyRows++;
                    if (emptyRows == 3)
                        //break processing if this is third empty row
                        break;
                    //skip empty row
                    continue;
                }
                emptyRows = 0;

                context.LoadData(cells, false);
                var entities = GetReferencedEntities(context);
                var parentContext = context.Children.FirstOrDefault(c => c.ContextType == ReferenceContextType.Parent);
                if (parentContext == null)
                {
                    Log.WriteLine("Table {0} is marked as a partial table but it doesn't have any parent mapping defined");
                    continue;
                }
                foreach (var entity in entities)
                {
                    _forwardReferences.Enqueue(new ForwardReference(entity, parentContext, this));
                }
            }

        }
        private void LoadFromSheet(WorksheetPart sheetPart, ClassMapping mapping)
        {


            SheetData sheetData = sheetPart.Worksheet.Elements<SheetData>().First();
            //if there is only header in a sheet, don't waste resources
            if (sheetData.Elements<Row>().LastOrDefault().RowIndex < 2)
                return;
            //adjust mapping to sheet in case columns are in a different order
            AdjustMapping(sheetData, mapping);
            CacheColumnIndices(mapping);

            //cache key columns
            CacheMultiRowIndices(mapping);

            ////cache contexts
            var context = GetReferenceContext(mapping);

            ////iterate over rows (be careful about MultiRow != None, merge values if necessary)
            foreach (Row row in sheetData.Elements<Row>())
            {
                IPersistEntity lastEntity = null;
                var emptyCells = 0;
                Row lastRow = null;

                //skip header row
                if (row == null || row.RowIndex.Value == 1)
                    continue;
                var cells = row.Elements<Cell>().ToList();

                if (!cells.Any()
                    || cells.All(c => !CheckIfCellHasValue(c, "", out string value)))
                {
                    emptyCells++;
                    if (emptyCells == 3)
                        //break processing if this is third empty row
                        break;
                    //skip empty row
                    continue;
                }
                emptyCells = 0;

                //load data into the context
                context.LoadData(cells, true);

                // check if there are any data to create entity
                if (!context.HasData)
                {
                    continue;
                }

                // if mapping defines identity data, check if there is any
                if (context.HasKeyRequirements && !context.HasKeyData)
                {
                    continue;
                }

                //last row might be used in case this is a multirow
                lastEntity = LoadFromRow(row, context, lastRow, lastEntity);
                AddRowNumber(lastEntity, context, (int)row.RowIndex.Value);
                lastRow = row;
            }

        }

        private ReferenceContext GetReferenceContext(ClassMapping mapping)
        {
            ReferenceContext context;
            if (_referenceContexts.TryGetValue(mapping, out context))
                return context;
            context = new ReferenceContext(this, mapping);
            _referenceContexts.Add(mapping, context);
            return context;
        }

        /// <summary>
        /// All indices should be cached already
        /// </summary>
        /// <param name="mapping"></param>
        /// <returns></returns>
        private IEnumerable<int> GetIdentityIndices(ClassMapping mapping)
        {
            return _multiRowIndicesCache[mapping.TableName];
        }

        private void CacheMultiRowIndices(ClassMapping mapping)
        {
            int[] existing;
            //one table might be defined for multiple classes but it has to have the same structure and constrains
            _multiRowIndicesCache.TryGetValue(mapping.TableName, out existing);

            var indices = new int[0];
            if (mapping.PropertyMappings != null && mapping.PropertyMappings.Any())
                indices = mapping.PropertyMappings
                    .Where(p => p.IsMultiRowIdentity)
                    .Select(m => m.ColumnIndex)
                    .ToArray();

            if (existing != null)
            {
                //update and check if it is consistent. Report inconsistency.
                _multiRowIndicesCache[mapping.TableName] = indices;
                if (existing.Length != indices.Length || !existing.SequenceEqual(indices))
                    Log.WriteLine("Table {0} is defined in multiple class mappings with different key columns for a multi-value records", mapping.TableName);
            }
            else
                _multiRowIndicesCache.Add(mapping.TableName, indices);

        }


        private IPersistEntity LoadFromRow(Row row, ReferenceContext context, Row lastRow, IPersistEntity lastEntity)
        {
            var multirow = IsMultiRow(row, context.CMapping, lastRow);
            if (multirow)
            {
                //only add multivalue to the multivalue properties of last entity
                var subContexts = context.AllScalarChildren
                    .Where(c => c.Mapping.MultiRow != MultiRow.None)
                    .Select(c =>
                    {
                        //get to the first list level up or on the base level if it is a scalar list
                        if (c.ContextType == ReferenceContextType.ScalarList)
                            return c;
                        while (c != null && c.ContextType != ReferenceContextType.EntityList)
                            c = c.ParentContext;
                        return c;
                    })
                    .Where(c => c != null)
                    .Distinct();
                foreach (var ctx in subContexts)
                    ResolveMultiContext(ctx, lastEntity);
                return lastEntity;
            }


            //get type of the coresponding object from ClassMapping or from a type hint, create instance
            var entity = ResolveContext(context, -1, false);
            return entity;
        }
        private void AddRowNumber(IPersistEntity entity, ReferenceContext context, int rowNum)
        {
            if (string.IsNullOrEmpty(Mapping.RowNumber))
                return;

            // TODO: Consider caching the delegate to avoid reflection lookup.
            var field = context.SegmentType.Derives.FirstOrDefault(d => d.Name == Mapping.RowNumber);
            if (field == null)
                return;
            if (rowNum == 0)
            {

            }
            field.PropertyInfo.SetValue(entity, rowNum);  // 
        }

        /// <summary>
        /// Search the model for the entities satisfying the conditions in context
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        internal IEnumerable<IPersistEntity> GetReferencedEntities(ReferenceContext context)
        {
            var type = context.SegmentType;

            //return empty enumeration in case there are identifiers but no data
            if (context.TypeHintMapping == null && context.TableHintMapping == null && context.ScalarChildren.Any() && !context.HasData)
                return Enumerable.Empty<IPersistEntity>();

            //we don't have any data so use just a type for the search
            return !context.ScalarChildren.Any() ?
                Model.Instances.OfType(type.Name, true) :
                Model.Instances.OfType(type.Name, true).Where(e => IsValidEntity(context, e));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="context">Reference context of the data</param>
        /// <param name="scalarIndex">Index of value to be used in a value list in case of multi values</param>
        /// <param name="onlyScalar"></param>
        /// <returns></returns>
        internal IPersistEntity ResolveContext(ReferenceContext context, int scalarIndex, bool onlyScalar)
        {
            IPersistEntity entity = null;
            var eType = GetConcreteType(context);
            if (IsGlobalType(eType.Type))
            {
                //it is a global type but there are no values to fill in
                if (!context.AllScalarChildren.Any(c => c.Values != null && c.Values.Length > 0))
                    return null;

                //it is a global entity and it was filled in with the data before
                if (GetOrCreateGlobalEntity(context, out entity, eType, scalarIndex))
                    return entity;
            }

            //create new entity if new global one was not created
            if (entity == null)
                entity = Model.Instances.New(eType.Type);

            //scalar values to be set to the entity
            foreach (var scalar in context.ScalarChildren)
            {
                var values = scalar.Values;
                if (values == null || values.Length == 0)
                    continue;
                if (scalar.ContextType == ReferenceContextType.ScalarList)
                {
                    //is should be ItemSet which is always initialized and inherits from IList
                    var list = scalar.PropertyInfo.GetValue(entity, null) as IList;
                    if (list == null)
                        continue;
                    foreach (var value in values)
                        list.Add(value);
                    continue;
                }

                //it is a single value
                var val = scalarIndex < 0 ? values[0] : (values.Length >= scalarIndex + 1 ? values[scalarIndex] : null);
                if (val != null)
                    scalar.PropertyInfo.SetValue(entity, val, scalar.Index != null ? new[] { scalar.Index } : null);
            }

            if (onlyScalar)
                return entity;

            //nested entities (global, local, referenced)
            foreach (var childContext in context.EntityChildren)
            {
                if (childContext.IsReference)
                {
                    _forwardReferences.Enqueue(new ForwardReference(entity, childContext, this));
                    continue;
                }

                if (childContext.ContextType == ReferenceContextType.EntityList)
                {
                    var depth =
                        childContext.ScalarChildren.Where(c => c.Values != null)
                            .Select(c => c.Values.Length)
                            .OrderByDescending(v => v)
                            .FirstOrDefault();
                    for (var i = 0; i < depth; i++)
                    {
                        var child = depth == 1 ? ResolveContext(childContext, -1, false) : ResolveContext(childContext, i, false);
                        AssignEntity(entity, child, childContext);
                    }
                    continue;
                }

                //it is a single entity
                var cEntity = ResolveContext(childContext, -1, false);
                AssignEntity(entity, cEntity, childContext);
            }

            var parentContext = context.Children.FirstOrDefault(c => c.ContextType == ReferenceContextType.Parent);
            if (parentContext != null)
                _forwardReferences.Enqueue(new ForwardReference(entity, parentContext, this));

            return entity;
        }

        internal void AssignEntity(IPersistEntity parent, IPersistEntity entity, ReferenceContext context)
        {
            if (context.MetaProperty != null && context.MetaProperty.IsDerived)
            {
                Log.WriteLine("It wasn't possible to add entity {0} as a {1} to parent {2} because it is a derived value",
                    entity.ExpressType.ExpressName, context.Segment, parent.ExpressType.ExpressName);
                return;
            }

            var index = context.Index == null ? null : new[] { context.Index };
            //inverse property
            if (context.MetaProperty != null && context.MetaProperty.IsInverse)
            {
                var remotePropName = context.MetaProperty.InverseAttributeProperty.RemoteProperty;
                var entityType = entity.ExpressType;
                var remoteProp = GetProperty(entityType, remotePropName);
                //it is enumerable inverse
                if (remoteProp.EnumerableType != null)
                {
                    var list = remoteProp.PropertyInfo.GetValue(entity, index) as IList;
                    if (list != null)
                    {
                        list.Add(parent);
                        return;
                    }
                }
                //it is a single inverse entity
                else
                {
                    remoteProp.PropertyInfo.SetValue(entity, parent, index);
                    return;
                }
                Log.WriteLine("It wasn't possible to add entity {0} as a {1} to parent {2}",
                    entity.ExpressType.ExpressName, context.Segment, entityType.ExpressName);
                return;
            }
            //explicit property
            var info = context.PropertyInfo;
            if (context.ContextType == ReferenceContextType.EntityList)
            {
                var list = info.GetValue(parent, index) as IList;
                if (list != null)
                {
                    list.Add(entity);
                    return;
                }
            }
            else
            {
                if ((context.MetaProperty != null && context.MetaProperty.IsExplicit) || info.GetSetMethod() != null)
                {
                    info.SetValue(parent, entity, index);
                    return;
                }
            }
            Log.WriteLine("It wasn't possible to add entity {0} as a {1} to parent {2}",
                entity.ExpressType.ExpressName, context.Segment, parent.ExpressType.ExpressName);
        }

        /// <summary>
        /// This is used for a multi-row instances where only partial context needs to be processed
        /// </summary>
        /// <param name="subContext"></param>
        /// <param name="rootEntity"></param>
        private void ResolveMultiContext(ReferenceContext subContext, IPersistEntity rootEntity)
        {
            //get context path from root entity
            var ctxStack = new Stack<ReferenceContext>();
            var context = subContext;
            while (context != null)
            {
                ctxStack.Push(context);
                context = context.ParentContext;
            }

            //use path to get to the bottom of the stact and add the value to it
            context = ctxStack.Pop();
            var entity = rootEntity;
            //stop one level above the original subcontext
            while (ctxStack.Peek() != subContext)
            {
                //browse to the level of the bottom context and call ResolveContext there
                var index = context.Index != null ? new[] { context.Index } : null;
                var value = context.PropertyInfo.GetValue(rootEntity, index);
                if (value == null)
                {
                    Log.WriteLine("It wasn't possible to browse to the data entry point.");
                    return;
                }

                if (context.ContextType == ReferenceContextType.Entity)
                {
                    entity = value as IPersistEntity;
                    continue;
                }

                var entities = value as IEnumerable;
                if (entities == null)
                {
                    Log.WriteLine("It wasn't possible to browse to the data entry point.");
                    return;
                }
                foreach (var e in entities)
                {
                    if (!IsValidEntity(context, e))
                        continue;
                    entity = e as IPersistEntity;
                    break;
                }
            }
            if (subContext.IsReference)
            {
                var reference = new ForwardReference(entity, subContext, this);
                _forwardReferences.Enqueue(reference);
                return;
            }

            if (subContext.ContextType == ReferenceContextType.EntityList)
            {
                var child = ResolveContext(subContext, -1, false);
                AssignEntity(entity, child, subContext);
                return;
            }

            if (subContext.ContextType == ReferenceContextType.ScalarList)
            {
                var list = subContext.PropertyInfo.GetValue(entity, null) as IList;
                if (list != null && subContext.Values != null && subContext.Values.Length > 0)
                    list.Add(subContext.Values[0]);
            }

        }

        internal static bool IsValidEntity(ReferenceContext context, object entity)
        {
            if (context.ScalarChildren.Count == 0)
                return true;

            //if it might have identifiers but doesn't have a one it can't find any
            if (!context.HasData)
                return false;

            return context.ScalarChildren
                .Where(s => s.Values != null && s.Values.Length > 0)
                .All(scalar => IsValidInContext(scalar, entity));
        }

        private static bool IsValidInContext(ReferenceContext scalar, object entity)
        {
            var prop = scalar.PropertyInfo;
            var vals = scalar.Values;
            var eVal = prop.GetValue(entity, null);
            if (scalar.ContextType != ReferenceContextType.ScalarList)
                return eVal != null && vals.Any(v => v != null && v.Equals(eVal));
            var list = eVal as IEnumerable;
            return list != null &&
                   //it might be a multivalue
                   list.Cast<object>().All(item => vals.Any(v => v.Equals(item)));
        }

        private bool IsGlobalType(Type type)
        {
            var gt = _globalTypes ??
                     (_globalTypes =
                         Mapping.Scopes.Where(s => s.Scope == ClassScopeEnum.Model)
                             .Select(s => MetaData.ExpressType(s.Class.ToUpper()))
                             .ToList());
            return gt.Any(t => t.Type == type || t.SubTypes.Any(st => st.Type == type));
        }

        private bool IsMultiRow(Row row, ClassMapping mapping, Row lastRow)
        {
            if (lastRow == null) return false;

            bool isMultiMapping;
            if (_isMultiRowMappingCache.TryGetValue(mapping, out isMultiMapping))
            {
                if (!isMultiMapping) return false;
            }
            else
            {
                if (mapping.PropertyMappings == null || !mapping.PropertyMappings.Any())
                {
                    _isMultiRowMappingCache.Add(mapping, false);
                    return false;
                }

                var multiRowProperty = mapping.PropertyMappings.FirstOrDefault(m => m.MultiRow != MultiRow.None);
                if (multiRowProperty == null)
                {
                    _isMultiRowMappingCache.Add(mapping, false);
                    return false;
                }

                _isMultiRowMappingCache.Add(mapping, true);
            }


            var keyIndices = GetIdentityIndices(mapping);
            foreach (var index in keyIndices)
            {
                var cellA = row.Elements<Cell>().FirstOrDefault(x => GetColumnIndexFromCell(x) == index);
                var cellB = lastRow.Elements<Cell>().FirstOrDefault(x => GetColumnIndexFromCell(x) == index);

                if (cellA == null || cellB == null)
                    return false;

                if (cellA.DataType == null || string.IsNullOrEmpty(cellA.InnerText) || cellB.DataType == null || string.IsNullOrEmpty(cellA.InnerText))
                    return false;

                if (cellA.DataType.Value != cellB.DataType.Value)
                    return false;
                if (cellA.DataType == CellValues.Number)
                {
                    if (Math.Abs(double.Parse(cellA.InnerText) - double.Parse(cellB.InnerText)) > 1e-9)
                        return false;
                }
                else if (cellA.DataType == CellValues.String)
                {
                    if (cellA.InnerText != cellB.InnerText)
                        return false;
                }
                else if (cellA.DataType == CellValues.Boolean)
                {
                    if (bool.Parse(cellA.InnerText) != bool.Parse(cellB.InnerText))
                        return false;
                }
            }

            return true;
        }
        /// <summary>
        /// Returns true if it exists, FALSE if new entity fas created and needs to be filled in with data
        /// </summary>
        /// <param name="context"></param>
        /// <param name="entity"></param>
        /// <param name="type"></param>
        /// <param name="scalarIndex">Index to field of values to be used to create the key. If -1 no index is used and all values are used.</param>
        /// <returns></returns>
        private bool GetOrCreateGlobalEntity(ReferenceContext context, out IPersistEntity entity, ExpressType type, int scalarIndex)
        {
            type = type ?? GetConcreteType(context);
            Dictionary<string, IPersistEntity> entities;
            if (!_globalEntities.TryGetValue(type, out entities))
            {
                entities = new Dictionary<string, IPersistEntity>();
                _globalEntities.Add(type, entities);
            }

            var keys = scalarIndex > -1 ?
                   context.AllScalarChildren.OrderBy(c => c.Segment)
                        .Where(c => c.Values != null)
                        .Select(c =>
                        {
                            if (c.Values.Length == 1) return c.Values[0];
                            return c.Values.Length >= scalarIndex + 1 ? c.Values[scalarIndex] : null;
                        }).Where(v => v != null) :

                    context.AllScalarChildren.OrderBy(c => c.Segment)
                        .Where(c => c.Values != null)
                        .SelectMany(c => c.Values.Where(cv => cv != null).Select(v => v.ToString()));
            var key = string.Join(", ", keys);
            if (entities.TryGetValue(key, out entity))
                return true;

            entity = Model.Instances.New(type.Type);
            entities.Add(key, entity);
            return false;
        }


        internal Type GetConcreteType(ReferenceContext context, Cell cell)
        {
            var cType = context.SegmentType;
            if (cType != null && !cType.Type.IsAbstract)
                return cType.Type;

            //use custom type resolver if there is a one which can resolve this type
            if (cType != null && Resolvers != null && Resolvers.Any())
            {
                var resolver = Resolvers.FirstOrDefault(r => r.CanResolve(cType));
                if (resolver != null)
                    return resolver.Resolve(cType.Type, cell, context.CMapping, context.Mapping, _sharedStringTable);
            }

            if (context.PropertyInfo != null)
            {
                var pType = context.PropertyInfo.PropertyType;
                pType = GetNonNullableType(pType);
                if (pType.IsValueType || pType == typeof(string))
                    return pType;

                if (typeof(IEnumerable).IsAssignableFrom(pType))
                {
                    pType = pType.GetGenericArguments()[0];
                    if (pType.IsValueType || pType == typeof(string))
                        return pType;
                }

                if (Resolvers != null && Resolvers.Any())
                {
                    var resolver = Resolvers.FirstOrDefault(r => r.CanResolve(pType));
                    if (resolver != null)
                        return resolver.Resolve(pType, cell, context.CMapping, context.Mapping,_sharedStringTable);
                }
            }

            Log.WriteLine("It wasn't possible to find a non-abstract type for table {0}, class {1}",
                context.CMapping.TableName, context.CMapping.Class);
            return null;
        }
        private ExpressType GetConcreteType(ReferenceContext context)
        {
            var cType = context.SegmentType;
            if (cType != null && !cType.Type.IsAbstract)
                return cType;


            //use fallback to retrieve a non-abstract type (defined in a configuration file?)
            var fbTypeName = context.CMapping.FallBackConcreteType;
            if (context.IsRoot && !string.IsNullOrWhiteSpace(fbTypeName))
            {
                var eType = MetaData.ExpressType(fbTypeName.ToUpper());
                if (eType != null && !eType.Type.IsAbstract)
                    return eType;
            }


            //use custom type resolver if there is a one which can resolve this type
            if (cType != null && Resolvers != null && Resolvers.Any())
            {
                var resolver = Resolvers.FirstOrDefault(r => r.CanResolve(cType));
                if (resolver != null)
                    return resolver.Resolve(cType, context, MetaData);
            }

            Log.WriteLine("It wasn't possible to find a non-abstract type for table {0}, class {1}",
                context.CMapping.TableName, context.CMapping.Class);
            return null;
        }

        private static void CacheColumnIndices(ClassMapping mapping)
        {
            foreach (var pMap in mapping.PropertyMappings)
                pMap.ColumnIndex = GetColumnIndexFromString(pMap.Column);
        }
      
        private void AdjustMapping(SheetData sheetData, ClassMapping mapping)
        {
            if (sheetData.Elements<Row>().LastOrDefault().RowIndex < 2)
                return;


            Row headerRow = sheetData.Elements<Row>().FirstOrDefault();
            //get the header row and analyze it
            if (headerRow == null)
                return;

            var headings = headerRow.Elements<Cell>().Where(c => c.DataType.Value == CellValues.String || c.DataType.Value == CellValues.SharedString || !string.IsNullOrWhiteSpace(c.InnerText)).ToList();
            if (!headings.Any())
                return;
            var mappings = mapping.PropertyMappings;
            if (mappings == null || !mappings.Any())
                return;

            foreach (var heading in headings)
            {
                var index = GetColumnIndexFromCell(heading);
                var column = ColumnIndexToName(index).ToUpper();
                string header;
                if (heading.DataType.Value == CellValues.SharedString)
                    header = _sharedStringTable.ElementAt(int.Parse(heading.InnerText)).InnerText;
                else
                    header = heading.CellValue.Text;

                var pMapping = mappings.FirstOrDefault(m => string.Equals(m.Header, header, StringComparison.OrdinalIgnoreCase));
                //if no mapping is found things might go wrong or it is just renamed
                if (pMapping == null || string.Equals(pMapping.Column, column, StringComparison.OrdinalIgnoreCase))
                    continue;

                //if the letter is not assigned at all, assign this letter
                if (string.IsNullOrWhiteSpace(pMapping.Column))
                {
                    pMapping.Column = column;
                    continue;
                }

                //move the column mapping to the new position
                var current = mappings.FirstOrDefault(m => string.Equals(m.Column, column, StringComparison.OrdinalIgnoreCase));
                if (current != null)
                    current.Column = null;
                pMapping.Column = column;
            }

            var unassigned = mappings.Where(m => string.IsNullOrWhiteSpace(m.Column)).ToList();
            if (!unassigned.Any())
                return;

            //try to assign letters to the unassigned columns
            foreach (var heading in headings)
            {
                var index = GetColumnIndexFromCell(heading);
                var column = ColumnIndexToName(index).ToUpper();
                var pMapping = mappings.FirstOrDefault(m => string.Equals(m.Column, column, StringComparison.OrdinalIgnoreCase));
                if (pMapping != null)
                    continue;

                var first = unassigned.FirstOrDefault();
                if (first == null) break;

                first.Column = column;
                unassigned.Remove(first);
            }

            //remove unassigned mappings
            if (unassigned.Any())
                return;
            foreach (var propertyMapping in unassigned)
                mapping.PropertyMappings.Remove(propertyMapping);
        }
        static string ColumnIndexToName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                columnName = (char)(65 + remainder) + columnName;
                columnIndex = (columnIndex - 1) / 26;
            }
            return columnName;
        }
        public static int GetColumnIndexFromCell(Cell cell)
        {
            // Extract the column portion from the cell reference
            string columnPart = "";
            foreach (char c in cell.CellReference.Value)
            {
                if (char.IsLetter(c))
                    columnPart += c;
                else
                    break;
            }

            // Subtract 1 to make it zero-based index
            return GetColumnIndexFromString(columnPart);
        }
        public static int GetColumnIndexFromString(string columnPart)
        {

            // Convert column letters to column index
            int columnIndex = 0;
            foreach (char c in columnPart)
            {
                columnIndex *= 26;
                columnIndex += char.ToUpper(c) - 'A' + 1;
            }

            // Subtract 1 to make it zero-based index
            return columnIndex;
        }
        private static Type GetNonNullableType(Type type)
        {
            //only value types can be nullable
            if (!type.IsValueType) return type;

            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(type) : type;
        }

        internal object CreateSimpleValue(Type type, string value)
        {
            var underlying = GetNonNullableType(type);

            var propType = underlying;
            var isExpress = false;

            //dig deeper if it is an express value type
            if (underlying.IsValueType && typeof(IExpressValueType).IsAssignableFrom(underlying))
            {
                var eType = MetaData.ExpressType(underlying);
                if (eType != null)
                {
                    isExpress = true;
                    underlying = GetNonNullableType(eType.UnderlyingType);
                }
            }

            //chack base types
            if (typeof(string) == underlying)
            {
                return isExpress ? Activator.CreateInstance(propType, value) : value;
            }
            if (underlying == typeof(double) || underlying == typeof(float))
            {
                double d;
                if (double.TryParse(value, out d))
                    return isExpress
                    ? Activator.CreateInstance(propType, d)
                    : d;
                return null;
            }
            if (underlying == typeof(int) || underlying == typeof(long))
            {
                var l = type == typeof(int) ? Convert.ToInt32(value) : Convert.ToInt64(value);
                return isExpress ? Activator.CreateInstance(propType, l) : l;
            }
            if (underlying == typeof(DateTime))
            {
                DateTime date;
                return !DateTime.TryParse(value, null, DateTimeStyles.RoundtripKind, out date) ?
                    DateTime.Parse("1900-12-31T23:59:59", null, DateTimeStyles.RoundtripKind) :
                    date;
            }
            if (underlying == typeof(bool))
            {
                bool i;
                if (bool.TryParse(value, out i))
                    return isExpress ? Activator.CreateInstance(propType, i) : i;
                return null;
            }
            if (underlying.IsEnum)
            {
                try
                {
                    var eMember = GetAliasEnumName(underlying, value);
                    //if there was no alias try to parse the value
                    var val = Enum.Parse(underlying, eMember ?? value, true);
                    return val;
                }
                catch (Exception)
                {
                    Log.WriteLine("Enumeration {0} doesn't have a member {1}", underlying.Name, value);
                }
            }
            return null;
        }

        internal bool CheckIfCellHasValue(Cell cell, string defaultValue, out string value)
        {
            try
            {
                var cellState = cell != null;
                value = cellState ? cell.DataType != null ? (cell.DataType == CellValues.SharedString ? _sharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText : cell.CellValue?.Text) : cell.CellValue?.Text : "";
                return cellState && (value != null &&
                         !string.Equals(value, defaultValue, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(value) && !string.IsNullOrWhiteSpace(value));
            }
            catch
            {
                value = "";
                return false;
            }

        }
        internal object CreateSimpleValue(Type type, Cell cell, string strValue)
        {

            type = GetNonNullableType(type);

            var propType = type;
            var isExpress = false;

            // Dig deeper if it is an express value type
            if (type.IsValueType && typeof(IExpressValueType).IsAssignableFrom(type))
            {
                var eType = MetaData.ExpressType(type);
                if (eType != null)
                {
                    isExpress = true;
                    type = GetNonNullableType(eType.UnderlyingType);
                }
            }

            if (typeof(string) == type)
            {
                string value = null;
                if (cell.DataType == null)
                {
                    value = strValue;
                }
                else if (cell.DataType.Value == CellValues.Number)
                {
                    value = strValue.ToString(CultureInfo.InvariantCulture);
                }
                else if (cell.CellFormula != null || cell.DataType.Value == CellValues.String)
                {
                    value = strValue;
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    value = strValue;
                }
                else if (cell.DataType.Value == CellValues.Boolean)
                {
                    value = bool.Parse(strValue).ToString();
                }
                else
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                return isExpress ? Activator.CreateInstance(propType, value) : value;
            }

            if (type == typeof(DateTime))
            {
                var date = default(DateTime);
                if (cell.DataType == null)
                {
                    date = DateTime.FromOADate(double.Parse(cell.InnerText));
                }
                else if (cell.DataType.Value == CellValues.Number)
                {
                    date = DateTime.FromOADate(double.Parse(cell.InnerText));
                }
                else if (cell.DataType.Value == CellValues.String)
                {
                    if (!DateTime.TryParse(strValue, null, DateTimeStyles.RoundtripKind, out date))
                    {
                        Log.WriteLine("There is no suitable value for {0} in cell {1} Unable to parse '{2}'", propType.Name, cell.CellReference.Value, cell.InnerText);
                        // Set to default value according to specification
                        date = DateTime.Parse("1900-12-31T23:59:59", null, DateTimeStyles.RoundtripKind);
                    }
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    if (_sharedStringTable != null)
                    {
                        if (!DateTime.TryParse(strValue, null, DateTimeStyles.RoundtripKind, out date))
                        {
                            Log.WriteLine("There is no suitable value for {0} in cell {1} Unable to parse '{2}'", propType.Name, cell.CellReference.Value, cell.InnerText);
                            // Set to default value according to specification
                            date = DateTime.Parse("1900-12-31T23:59:59", null, DateTimeStyles.RoundtripKind);
                        }
                    }

                }
                else
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                return date;
            }

            if (type == typeof(double) || type == typeof(float))
            {
                if (cell.DataType == null)
                {
                    double doubleValue;

                    if (double.TryParse(strValue, out doubleValue))
                    {
                        return isExpress ? Activator.CreateInstance(propType, doubleValue) : doubleValue;
                    }
                    else
                    {
                        Log.WriteLine("There is no suitable value for {0} in cell {1} Unable to parse '{2}'", propType.Name, cell.CellReference.Value, strValue);
                    }
                }
                if (cell.CellFormula != null || cell.DataType.Value == CellValues.Number)
                {
                    if(double.TryParse(strValue,out double value))
                    {
                        return isExpress ? Activator.CreateInstance(propType, value) : value;

                    }
                }
                else if (cell.DataType.Value == CellValues.String)
                {
                    double d;
                    if (double.TryParse(strValue, out d))
                    {
                        return isExpress ? Activator.CreateInstance(propType, d) : d;
                    }
                    else
                    {
                        Log.WriteLine("There is no suitable value for {0} in cell {1} Unable to parse '{2}'", propType.Name, cell.CellReference.Value, strValue);
                    }
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    double doubleValue;

                    if (double.TryParse(strValue, out doubleValue))
                    {
                        return isExpress ? Activator.CreateInstance(propType, doubleValue) : doubleValue;
                    }
                    else
                    {
                        Log.WriteLine("There is no suitable value for {0} in cell {1} Unable to parse '{2}'", propType.Name, cell.CellReference.Value, strValue);
                    }

                }
                else
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                return null;
            }

            if (type == typeof(int) || type == typeof(long))
            {
                if (cell.DataType == null)
                {
                    var l = type == typeof(int) ? Convert.ToInt32(strValue) : Convert.ToInt64(strValue);
                    return isExpress ? Activator.CreateInstance(propType, l) : l;
                }
                else if (cell.DataType.Value == CellValues.Number || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.InlineString)
                {
                    var l = type == typeof(int) ? Convert.ToInt32(strValue) : Convert.ToInt64(strValue);
                    return isExpress ? Activator.CreateInstance(propType, l) : l;
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    var l = type == typeof(int) ? Convert.ToInt32(strValue) : Convert.ToInt64(strValue);
                    return isExpress ? Activator.CreateInstance(propType, l) : l;
                }
                else
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                return null;
            }

            if (type == typeof(bool))
            {
                if (cell.DataType == null)
                {
                    if( int.TryParse(strValue,out int intBool))
                    return isExpress ? Activator.CreateInstance(propType, intBool != 0) : intBool != 0;
                }
               else if (cell.DataType.Value == CellValues.Number)
                {
                    var b = int.Parse(strValue) != 0;
                    return isExpress ? Activator.CreateInstance(propType, b) : b;
                }
                else if (cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.InlineString)
                {
                    bool i;
                    if (bool.TryParse(strValue, out i))
                        return isExpress ? Activator.CreateInstance(propType, i) : i;
                    Log.WriteLine("Wrong boolean format of {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                else if (cell.DataType.Value == CellValues.SharedString)
                {
                    bool i;
                    if (bool.TryParse(strValue, out i))
                        return isExpress ? Activator.CreateInstance(propType, i) : i;
                    Log.WriteLine("Wrong boolean format of {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                else if (cell.DataType.Value == CellValues.Boolean)
                {
                    bool i;
                    if (bool.TryParse(strValue, out i))
                        return isExpress ? Activator.CreateInstance(propType, (object)i) : i;
                    i = strValue == "1";
                    return isExpress ? Activator.CreateInstance(propType, (object)i) : i;
                }
                else
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
                return null;
            }

            // Enumeration
            if (type.IsEnum)
            {
                if (cell.DataType != CellValues.String && cell.DataType != CellValues.SharedString)
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                    return null;
                }
                try
                {
                    var eValue = strValue.Replace("-", "_");  // Hyphens aren't valid in C# enums, but have been seen in live data
                    var eMember = GetAliasEnumName(type, eValue);
                    // If there was no alias, try to parse the value
                    var val = Enum.Parse(type, eMember ?? eValue, true);
                    return val;
                }
                catch (Exception)
                {
                    Log.WriteLine("There is no suitable value for {0} in cell {1}", propType.Name, cell.CellReference.Value);
                }
            }

            // If no suitable type was found, report it 
            throw new Exception("Unsupported type " + type.Name + " for value '" + cell + "'");
        }
        private string GetAliasEnumName(Type type, string alias)
        {
            string result;
            return _aliasesEnumCache.TryGetValue(type.Name + "." + alias, out result) ? result : null;
        }

        #endregion
    }
}
