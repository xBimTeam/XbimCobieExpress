using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Xbim.CobieExpress;
using Xbim.Common;
using Xbim.Common.Configuration;
using Xbim.Common.Geometry;
using Xbim.Common.Metadata;
using Xbim.Common.Step21;
using Xbim.Ifc;
using Xbim.IO.CobieExpress.Resolvers;
using Xbim.IO.Memory;
using Xbim.IO.Table;

namespace Xbim.IO.CobieExpress
{
    public class CobieModel : IModel, IDisposable
    {
        private readonly IModel _model;

        private static readonly IEntityFactory factory = new EntityFactoryCobieExpress();

        private static Lazy<IModelProvider> lazyModelProvider = new Lazy<IModelProvider>(() => BuildModelProvider());

        /// <summary>
        /// Provides access to model persistance capabilities
        /// </summary>
        protected static IModelProvider ModelProvider
        {
            get => lazyModelProvider.Value;
        }

        /// <summary>
        /// Factory to create ModelProvider instances. 
        /// </summary>
        /// <remarks>Consumers can use this instance of <see cref="IModelProviderFactory"/> to control the 
        /// implementations of IModel it uses.
        /// In particular you can tell the factory to always use MemoryModel, or Esent model, or a blend (Heuristic)
        /// </remarks>
        [Obsolete("ModelProviders are now created via the XbimServices.Current.ServiceProvider")]
        public static IModelProviderFactory ModelProviderFactory
        {
            get;
            set;
        }

        public CobieModel(IModel model)
        {

            _model = model;

            InitEvents();
        }

        static CobieModel()
        {
            if(!XbimServices.Current.IsBuilt)
            {
                XbimServices.Current.ConfigureServices(cfg => cfg.AddXbimToolkit(builder => builder.AddHeuristicModel()));
            }
        }

        /// <summary>
        /// Creates memory model inside
        /// </summary>
        public CobieModel() : this(CreateModel())
        {
        }

        /// <summary>
        /// Creates EsentModel inside
        /// </summary>
        /// <param name="esentDbFile"></param>
        public CobieModel(string esentDbFile): this(CreateModel(esentDbFile))
        {
        }


        private static IModelProvider BuildModelProvider()
        {
            
            var provider = XbimServices.Current.ServiceProvider.GetService<IModelProvider>() ?? new HeuristicModelProvider(XbimServices.Current.GetLoggerFactory());

            // Here we hook into the defined ModelProvider implementation and provide our CobieExpress EntityFactory when a Cobie2x4 is opened
            provider.EntityFactoryResolver = (version) =>
            {
                if (version == Common.Step21.XbimSchemaVersion.Cobie2X4)
                {
                    return new EntityFactoryCobieExpress();
                }
                return null;
            };

            return provider;
        }

        private static IModel CreateModel(string file = "")
        {
            var provider = ModelProvider;
            if (string.IsNullOrEmpty(file))
            {
                return provider.Create(XbimSchemaVersion.Cobie2X4, XbimStoreType.InMemoryModel);
            }
            else
            { 
                return provider.Create(XbimSchemaVersion.Cobie2X4, file);
            }
        }

        public object Tag { get; set; }

        /// <summary>
        /// This factory only opens an in memory model
        /// </summary>
        /// <param name="input"></param>
        /// <param name="streamSize"></param>
        /// <param name="labelFrom"></param>
        /// <returns></returns>
        public static CobieModel OpenStep21(Stream input, long streamSize, int labelFrom)
        {
            var model = new MemoryModel(factory, default(ILoggerFactory), labelFrom);
            model.LoadStep21(input, streamSize);
            return new CobieModel(model);
        }

        public static CobieModel OpenStep21(string input, bool esentDB = false)
        {

            var provider = ModelProvider;
            XbimSchemaVersion ifcVersion = GetSchemaVersion(input, provider);

            var model = provider.Open(input, ifcVersion);

            return new CobieModel(model);
        }

        public static CobieModel OpenStep21(Stream input, long streamSize, bool esentDB = false)
        {
            var provider = ModelProvider;
            var modelType = esentDB ? XbimModelType.EsentModel : XbimModelType.MemoryModel;
            // TODO: Should determine the schema. not hardwire
            var model = provider.Open(input, StorageType.Stp, XbimSchemaVersion.Ifc4, modelType);

            return new CobieModel(model);
        }

        public void SaveAsStep21(string file)
        {
            using (var fileStream = File.Create(file))
            {
                this.SaveAsIfc(fileStream);
            }
        }

        public static CobieModel OpenEsent(string esentDB)
        {
            var provider = ModelProvider;

            var model = provider.Open(esentDB, XbimSchemaVersion.Cobie2X4, accessMode: XbimDBAccess.ReadWrite);

            return new CobieModel(model);
        }

        public void SaveAsEsent(string dbName)
        {
            ModelProvider.Persist(this, dbName);
        }

        public void SaveAsStep21Zip(string file)
        {
            using (var fileStream = File.Create(file))
            {
                this.SaveAsIfcZip(fileStream, "Cobie.stp", StorageType.Stp);
            }
        }

        public static CobieModel OpenStep21Zip(string input, bool esentDB = false)
        {
            var provider = ModelProvider;

            var model = provider.Open(input, XbimSchemaVersion.Cobie2X4);

            return new CobieModel(model);
        }

        private static XbimSchemaVersion GetSchemaVersion(string esentDB, IModelProvider provider)
        {
            var ifcVersion = provider.GetXbimSchemaVersion(esentDB);
            if (ifcVersion == XbimSchemaVersion.Unsupported)
            {
                throw new FileLoadException(esentDB + " is not a valid IFC file format, ifc, ifcxml, ifczip and xBIM are supported.");
            }

            return ifcVersion;
        }



        public static ModelMapping GetMapping()
        {
            return ModelMapping.Load(GetCobieConfigurationXml());
        }

        private static string GetCobieConfigurationXml()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = assembly.GetManifestResourceNames().Single(str => str.EndsWith("COBieUK2012.xml"));

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void ExportToTable(string file, out string report, ModelMapping mapping = null, Stream template = null)
        {
            var ext = (Path.GetExtension(file) ?? "").ToLower();
            if (ext != ".xls" && ext != ".xlsx")
                file = Path.ChangeExtension(file, ".xlsx");

            mapping = mapping ?? GetMapping();
            var storage = GetTableStore(this, mapping);
            storage.Store(file, template);
            report = storage.Log.ToString();
        }

        public void ExportToTable(Stream file, ExcelTypeEnum typeEnum, out string report, ModelMapping mapping = null, Stream template = null)
        {
            mapping = mapping ?? GetMapping();
            var storage = GetTableStore(this, mapping);
            storage.Store(file, typeEnum, template);
            report = storage.Log.ToString();
        }


        public static CobieModel ImportFromTable(string file, out string report, ModelMapping mapping = null)
        {
            var loaded = new CobieModel();
            mapping = mapping ?? GetMapping();
            var storage = GetTableStore(loaded, mapping);
            using (var txn = loaded.BeginTransaction("Loading XLSX"))
            {
                storage.LoadFrom(file);

                //assign all levels to facility because COBie XLS standard contains this implicitly
                var facility = loaded.Instances.FirstOrDefault<CobieFacility>();
                var floors = loaded.Instances.OfType<CobieFloor>().ToList();
                if (facility != null && floors.Any())
                    floors.ForEach(f => f.Facility = facility);

                txn.Commit();
            }

            report = storage.Log.ToString();
            return loaded;
        }

        public static CobieModel ImportFromTable(Stream file, ExcelTypeEnum typeEnum, out string report, ModelMapping mapping = null)
        {
            var loaded = new CobieModel();
            mapping = mapping ?? GetMapping();
            var storage = GetTableStore(loaded, mapping);
            using (var txn = loaded.BeginTransaction("Loading XLSX"))
            {
                storage.LoadFrom(file, typeEnum);
                txn.Commit();
            }

            report = storage.Log.ToString();
            return loaded;
        }

        private static TableStore GetTableStore(IModel model, ModelMapping mapping)
        {
            var storage = new TableStore(model, mapping);
            storage.Resolvers.Add(new AttributeTypeResolver());
            return storage;
        }

        
        public void InsertCopy(IEnumerable<CobieComponent> components, bool keepLabels, XbimInstanceHandleMap mappings)
        {
            foreach (var component in components)
            {
                InsertCopy(component, mappings, InsertCopyComponentFilter, true, keepLabels);
            }
        }

        private object InsertCopyComponentFilter(ExpressMetaProperty property, object parentObject)
        {
            if (!property.IsInverse) 
                return property.PropertyInfo.GetValue(parentObject, null);
            
            if (property.Name == "InSystems")
                return property.PropertyInfo.GetValue(parentObject, null);

            return null;
        }

        #region IModel implementation using inner model
        public int UserDefinedId 
        { 
            get { return _model.UserDefinedId; }
            set { _model.UserDefinedId = value; } 
        }

        public IGeometryStore GeometryStore
        {
            get { return _model.GeometryStore; }
        }
        public IStepFileHeader Header { get { return _model.Header; } }
        public bool IsTransactional { get { return _model.IsTransactional; } }
        public IList<XbimInstanceHandle> InstanceHandles { get { return _model.InstanceHandles; } }
        public IEntityCollection Instances {
            get { return _model.Instances; }
        }
        public bool Activate(IPersistEntity owningEntity)
        {
            return _model.Activate(owningEntity);
        }

        public void Delete(IPersistEntity entity)
        {
            _model.Delete(entity);
        }

        public ITransaction BeginTransaction(string name)
        {
            return _model.BeginTransaction(name);
        }

        public ITransaction CurrentTransaction {
            get { return _model.CurrentTransaction; }
        }
        public ExpressMetaData Metadata {
            get { return _model.Metadata; }
        }
        public IModelFactors ModelFactors {
            get { return _model.ModelFactors; }
        }

        public T InsertCopy<T>(T toCopy, XbimInstanceHandleMap mappings, PropertyTranformDelegate propTransform, bool includeInverses,
            bool keepLabels) where T : IPersistEntity
        {
            return _model.InsertCopy(toCopy, mappings, propTransform, includeInverses, keepLabels);
        }

        public void ForEach<TSource>(IEnumerable<TSource> source, Action<TSource> body) where TSource : IPersistEntity
        {
            _model.ForEach(source, body);
        }

        public event NewEntityHandler EntityNew;
        public event ModifiedEntityHandler EntityModified;
        public event DeletedEntityHandler EntityDeleted;
        public IInverseCache BeginInverseCaching()
        {
            return _model.BeginInverseCaching();
        }

        public IInverseCache InverseCache
        {
            get { return _model.InverseCache; }
        }

        public XbimSchemaVersion SchemaVersion
        {
            get { return _model.SchemaVersion; }
        }

        [Obsolete("Best practice is to get your own logger via XbimServices.Current.CreateLogger<T>()")]
        public ILogger Logger { get; private set; } = XbimServices.Current.CreateLogger<CobieModel>();

        public IEntityCache EntityCache => _model.EntityCache;

        private void InitEvents()
        {
            _model.EntityNew += OnEntityNew;
            _model.EntityDeleted += OnEntityDeleted;
            _model.EntityModified += OnEntityModified;
        }

        protected virtual void OnEntityNew(IPersistEntity entity)
        {
            var handler = EntityNew;
            if (handler != null) handler(entity);
        }

        protected virtual void OnEntityModified(IPersistEntity entity, int property)
        {
            var handler = EntityModified;
            if (handler != null) handler(entity, property);
        }

        protected virtual void OnEntityDeleted(IPersistEntity entity)
        {
            var handler = EntityDeleted;
            if (handler != null) handler(entity);
        }
        #endregion

        #region Entity created and modified default CreatedInfo assignment
        private CobieCreatedInfo _entityInfo;
        private bool _ownChange;
        private IPersistEntity _lastEntity;

        protected virtual void SetEntityCreatedInfo(IPersistEntity entity, int property)
        {
            SetEntityCreatedInfo(entity);
        }

        protected virtual void SetEntityCreatedInfo(IPersistEntity entity)
        {
            if (_ownChange)
            {
                if (ReferenceEquals(_lastEntity, entity))
                    return;
                _ownChange = false;
            }
            var refObj = entity as CobieReferencedObject;
            if (refObj == null) return;

            _ownChange = true;
            _lastEntity = entity;
            refObj.Created = _entityInfo;
        }

        public CobieCreatedInfo SetDefaultEntityInfo(DateTime date, string email, string givenName, string familyName)
        {
            _entityInfo = Instances.New<CobieCreatedInfo>(ci =>
            {
                ci.CreatedOn = date;
                ci.CreatedBy =
                    Instances.FirstOrDefault<CobieContact>(
                        c => c.Email == email && c.GivenName == givenName && c.FamilyName == familyName) ??
                    Instances.New<CobieContact>(
                            c =>
                            {
                                c.Created = ci;
                                c.Email = email;
                                c.GivenName = givenName;
                                c.FamilyName = familyName;
                            });
            });
            EntityNew += SetEntityCreatedInfo;
            EntityModified += SetEntityCreatedInfo;
            return _entityInfo;
        }
        #endregion

        #region IDispose implementation
        public void Dispose()
        {
            //detach event handlers
            _model.EntityNew -= OnEntityNew;
            _model.EntityDeleted -= OnEntityDeleted;
            _model.EntityModified -= OnEntityModified;

            if (_entityInfo != null)
                EntityNew -= SetEntityCreatedInfo;

            //dispose model if it is disposable
            var dispModel = _model as IDisposable;
            if(dispModel != null) dispModel.Dispose();
        }

        public IEntityCache BeginEntityCaching()
        {
            return _model.BeginEntityCaching();
        }
        #endregion
    }
}
