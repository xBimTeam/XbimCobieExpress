using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Xbim.Ifc4.Interfaces;

namespace Xbim.CobieExpress.Exchanger.FilterHelper
{

    public class OutputFilters
    {
        private readonly ILogger _log;

        #region Properties

        /// <summary>
        /// Flip results of true/false 
        /// </summary>
        [XmlIgnore][JsonIgnore]
        public bool FlipResult { get; set; }

        /// <summary>
        /// Roles set on this filter
        /// </summary>
        public RoleFilter AppliedRoles  { get; set; }
        /// <summary>
        /// IfcProduct Exclude filters
        /// </summary>
        public ObjectFilter IfcProductFilter { get;  set; }
        /// <summary>
        /// IfcTypeObject Exclude filters
        /// </summary>
        public ObjectFilter IfcTypeObjectFilter { get;  set; }
        /// <summary>
        /// IfcAssembly Exclude filters
        /// </summary>
        public ObjectFilter IfcAssemblyFilter { get; set; }

        /// <summary>
        /// Zone attribute filters
        /// </summary>
        public PropertyFilter ZoneFilter { get;  set; }
        /// <summary>
        /// Type attribute filters
        /// </summary>
        public PropertyFilter TypeFilter { get;  set; }
        /// <summary>
        /// Space attribute filters
        /// </summary>
        public PropertyFilter SpaceFilter { get;  set; }
        /// <summary>
        /// Floor attribute filters
        /// </summary>
        public PropertyFilter FloorFilter { get;  set; }
        /// <summary>
        /// Facility attribute filters
        /// </summary>
        public PropertyFilter FacilityFilter { get;  set; }
        /// <summary>
        /// Spare attribute filters
        /// </summary>
        public PropertyFilter SpareFilter { get;  set; }
        /// <summary>
        /// Component attribute filters
        /// </summary>
        public PropertyFilter ComponentFilter { get;  set; }
        /// <summary>
        /// Common attribute filters
        /// </summary>
        public PropertyFilter CommonFilter { get;  set; }

        /// <summary>
        /// Temp storage for role OutPutFilters
        /// </summary>
        [XmlIgnore][JsonIgnore]
        private Dictionary<RoleFilter, OutputFilters> RolesFilterHolder { get; set; }

        /// <summary>
        /// Nothing set in RolesFilterHolder
        /// </summary>
        public bool DefaultsNotSet 
        {
            get
            {
                return RolesFilterHolder.Count == 0;
            }
        }
        #endregion

        #region Constructor methods

        /// <summary>
        /// Empty constructor for Serialize
        /// </summary>
        public OutputFilters(ILogger logger)
        {
            _log = logger;
            //will flip filter result from true to false
            FlipResult = false;

            //object filters
            IfcProductFilter = new ObjectFilter();
            IfcTypeObjectFilter = new ObjectFilter();
            IfcAssemblyFilter = new ObjectFilter();

            //Property name filters
            ZoneFilter = new PropertyFilter();
            TypeFilter = new PropertyFilter();
            SpaceFilter = new PropertyFilter();
            FloorFilter = new PropertyFilter();
            FacilityFilter = new PropertyFilter();
            SpareFilter = new PropertyFilter();
            ComponentFilter = new PropertyFilter();
            CommonFilter = new PropertyFilter();

            //role storage
            RolesFilterHolder = new Dictionary<RoleFilter, OutputFilters>();
        }

        /// <summary>
        /// Constructor for default set configFileName = null, or passed in configuration file path
        /// </summary>
        /// <param name="configFileName"></param>
        /// <param name="logger"></param>
        /// <param name="roleFlags"></param>
        /// <param name="setsToImport"></param>
        public OutputFilters(string configFileName, ILogger logger, RoleFilter roleFlags, ImportSet setsToImport = ImportSet.All) : this(logger)
        {
            AppliedRoles = roleFlags;
            FiltersHelperInit(configFileName, setsToImport);
        }

        public OutputFilters(ILogger logger, RoleFilter roleFlags) : this(logger)
        {
            AppliedRoles = roleFlags;
            if(roleFlags.HasMultipleFlags())
            {
                throw new InvalidOperationException("Cannot construct with multiple roles. Use OutputFilters.Merge to add additional roles");
            }
            FiltersHelperInit(roleFlags.ToResourceName());
        }

        /// <summary>
        /// Constructor to apply roles, and pass custom role OutPutFilters
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="roles">RoleFilter flags on roles to filter on</param>
        /// <param name="rolesFilter">Dictionary of role to OutPutFilters</param>
        public OutputFilters(ILogger logger, RoleFilter roles, Dictionary<RoleFilter, OutputFilters> rolesFilter = null) : this(logger)
        {
            ApplyRoleFilters(roles, false, rolesFilter);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Test for empty object
        /// </summary>
        /// <returns>bool</returns>
        public bool IsEmpty()
        {
            return IfcProductFilter.IsEmpty() && IfcTypeObjectFilter.IsEmpty() && IfcAssemblyFilter.IsEmpty() && 
            ZoneFilter.IsEmpty() && TypeFilter.IsEmpty() && SpaceFilter.IsEmpty() && 
            FloorFilter.IsEmpty() && FacilityFilter.IsEmpty() && SpareFilter.IsEmpty() && 
            ComponentFilter.IsEmpty() && CommonFilter.IsEmpty();
        }

        /// <summary>
        /// Will read Configuration file if passed, or default COBieAttributesFilters.config
        /// </summary>
        /// <param name="configFileName">Full path/name for config file</param>
        /// <param name="import"></param>
        private void FiltersHelperInit(string configFileName = null, ImportSet import = ImportSet.All)
        {
            //set default
            var sourceFile = configFileName ?? RoleFilter.Unknown.ToResourceName();
            var config = GetConfig(sourceFile);

            //IfcProduct and IfcTypeObject filters
            if (import == ImportSet.All || import == ImportSet.IfcFilters)
            {
                IfcProductFilter = new ObjectFilter(config.GetSection("IfcElementInclusion"));
                IfcTypeObjectFilter = new ObjectFilter(config.GetSection("IfcTypeInclusion"));
                IfcTypeObjectFilter.FillPreDefinedTypes(config.GetSection("IfcPreDefinedTypeFilter"));
                IfcAssemblyFilter = new ObjectFilter(config.GetSection("IfcAssemblyInclusion"));
            }
            
            //Property name filters
            if (import == ImportSet.All || import == ImportSet.PropertyFilters)
            {
                ZoneFilter = new PropertyFilter(config.GetSection("ZoneFilter"));
                TypeFilter = new PropertyFilter(config.GetSection("TypeFilter"));
                SpaceFilter = new PropertyFilter(config.GetSection("SpaceFilter"));
                FloorFilter = new PropertyFilter(config.GetSection("FloorFilter"));
                FacilityFilter = new PropertyFilter(config.GetSection("FacilityFilter"));
                SpareFilter = new PropertyFilter(config.GetSection("SpareFilter"));
                ComponentFilter = new PropertyFilter(config.GetSection("ComponentFilter"));
                CommonFilter = new PropertyFilter(config.GetSection("CommonFilter"));
            }
            // API restructure:
            // a call to File.Delete(config.FilePath); has been removed
            // it is strange for the configuration reading routine to delete a configuration file
        }

        /// <summary>
        /// Get Configuration object from the passed file path or embedded resource file
        /// </summary>
        /// <param name="fileOrResourceName">file path or resource name; an existing file gets the priortiy over an omonymous resource name</param>
        /// <returns></returns>
        private Configuration GetConfig(string fileOrResourceName)
        {
            
            if (!File.Exists(fileOrResourceName))
            {
                // try to save resource to temporary file
                
                var asss = global::System.Reflection.Assembly.GetExecutingAssembly();
                using (var input = asss.GetManifestResourceStream(fileOrResourceName))
                {
                    if (input == null)
                    {
                        _log.LogError("Could not load configuration file: {0}.", fileOrResourceName);
                        return null;
                    }
                    var tmpFile = Path.GetTempPath() + Guid.NewGuid() + ".tmp";
                    using (var output = File.Create(tmpFile))
                    {
                        input.CopyTo(output);
                    }
                    fileOrResourceName = tmpFile;
                }               
            }

            Configuration config;
            try
            {
                var configMap = new ExeConfigurationFileMap { ExeConfigFilename = fileOrResourceName };
                config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
            }
            catch (Exception ex)
            {
                var message = string.Format(@"Error loading configuration file '{0}'.", fileOrResourceName);
                _log.LogError(message, ex);
                throw;
            }
            return config;
        }

        /// <summary>
        /// Copy the OutPutFilters
        /// </summary>
        /// <param name="copyFilter">OutPutFilters to copy </param>
        public void Copy(OutputFilters copyFilter)
        {
            AppliedRoles = copyFilter.AppliedRoles;

            IfcProductFilter.Copy(copyFilter.IfcProductFilter);
            IfcTypeObjectFilter.Copy(copyFilter.IfcTypeObjectFilter);
            IfcAssemblyFilter.Copy(copyFilter.IfcAssemblyFilter);

            ZoneFilter.Copy(copyFilter.ZoneFilter);
            TypeFilter.Copy(copyFilter.TypeFilter);
            SpaceFilter.Copy(copyFilter.SpaceFilter);
            FloorFilter.Copy(copyFilter.FloorFilter);
            FacilityFilter.Copy(copyFilter.FacilityFilter);
            SpareFilter.Copy(copyFilter.SpareFilter);
            ComponentFilter.Copy(copyFilter.ComponentFilter);
            CommonFilter.Copy(copyFilter.CommonFilter);
        }


        /// <summary>
        /// Clear OutPutFilters
        /// </summary>
        public void Clear()
        {
            AppliedRoles = 0;

            IfcProductFilter.Clear();
            IfcTypeObjectFilter.Clear();
            IfcAssemblyFilter.Clear();

            ZoneFilter.Clear();
            TypeFilter.Clear();
            SpaceFilter.Clear();
            FloorFilter.Clear();
            FacilityFilter.Clear();
            SpareFilter.Clear();
            ComponentFilter.Clear();
            CommonFilter.Clear();
        }
        #endregion


        #region Filter Methods

        //TODO: Check function below, see if it works!
        /// <summary>
        /// filter on IfcObjectDefinition objects
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="checkType">Flag indicating whether any <see cref="IIfcTypeObject"/> should be checked</param>
        /// <returns>bool true = exclude</returns>
        public bool ObjFilter(IIfcObjectDefinition obj, bool checkType = true)
        {
            bool exclude = false;
            if (obj is IIfcProduct product)
            {
                exclude = IfcProductFilter.ItemsFilter(obj);
                //check the element is not defined by a type which is excluded, by default if no type, then no element included
                if (!exclude && checkType)
                {
                    IIfcTypeObject objType = product.IsTypedBy.FirstOrDefault()?.RelatingType;
                    if (objType != null) //if no type defined lets include it for now
                    {
                        exclude = IfcTypeObjectFilter.ItemsFilter(objType); 
                    }
                }
                
            }
            else if (obj is IIfcTypeProduct)
            {
                exclude = IfcTypeObjectFilter.ItemsFilter(obj);
            }
            return FlipResult ? !exclude : exclude;
        }
        #endregion

        #region Merge Roles
        /// <summary>
        /// Merge OutPutFilters
        /// </summary>
        /// <param name="mergeFilter">OutPutFilters</param>
        public void Merge(OutputFilters mergeFilter)
        {
            IfcProductFilter.MergeInc(mergeFilter.IfcProductFilter);
            IfcTypeObjectFilter.MergeInc(mergeFilter.IfcTypeObjectFilter);
            IfcAssemblyFilter.MergeInc(mergeFilter.IfcAssemblyFilter);

            ZoneFilter.Merge(mergeFilter.ZoneFilter);
            TypeFilter.Merge(mergeFilter.TypeFilter);
            SpaceFilter.Merge(mergeFilter.SpaceFilter);
            FloorFilter.Merge(mergeFilter.FloorFilter);
            FacilityFilter.Merge(mergeFilter.FacilityFilter);
            SpareFilter.Merge(mergeFilter.SpareFilter);
            ComponentFilter.Merge(mergeFilter.ComponentFilter);
            CommonFilter.Merge(mergeFilter.CommonFilter);

        }

        /// <summary>
        /// Extension method to use default role configuration resource files
        /// </summary>
        /// <param name="roles">MergeRoles, Flag enum with one or more roles</param>
        /// <param name="append">true = add, false = overwrite existing </param>
        /// <param name="rolesFilter">Dictionary of roles to OutPutFilters to use for merge, overwrites current assigned dictionary</param>
        public void ApplyRoleFilters(RoleFilter roles, bool append = false, Dictionary<RoleFilter, OutputFilters> rolesFilter = null)
        {
            if (rolesFilter != null)
            {
                RolesFilterHolder = rolesFilter;
            }

            var init = append && !IsEmpty();
            
            OutputFilters mergeFilter = null;
            foreach (RoleFilter role in Enum.GetValues(typeof(RoleFilter)))
            {
                if (!roles.HasFlag(role)) 
                    continue;
                if (RolesFilterHolder.ContainsKey(role))
                {
                    mergeFilter = RolesFilterHolder[role];
                }
                else
                {
                    // load defaults
                    mergeFilter = GetDefaults(role, _log);
                    RolesFilterHolder[role] = mergeFilter;
                }
                if (mergeFilter != null)
                {
                    if (!init)
                    {
                        Copy(mergeFilter); //start a fresh
                        init = true;//want to merge on next loop iteration
                    }
                    else
                    {
                        Merge(mergeFilter);
                    }
                }
                mergeFilter = null;
            }
            //add the default property filters
            OutputFilters defaultPropFilters = new OutputFilters(null, _log, RoleFilter.Unknown, ImportSet.PropertyFilters);
            Merge(defaultPropFilters);

            //save the applied roles at end as this.Copy(mergeFilter) would set to first role in RoleFilter
            AppliedRoles = roles; 
        }

        

        /// <summary>
        /// Set filters for Federated Model, referenced models
        /// </summary>
        /// <param name="modelRoleMap"></param>
        public Dictionary<T, OutputFilters> SetFedModelFilter<T>(Dictionary<T, RoleFilter> modelRoleMap)
        {
            Dictionary<T, OutputFilters> modelFilterMap = new Dictionary<T, OutputFilters>();

            //save this filter before working out all fed models
            OutputFilters saveFilter = new OutputFilters(_log);
            saveFilter.Copy(this);

            foreach (var item in modelRoleMap)
            {
                ApplyRoleFilters(item.Value);
                OutputFilters roleFilter = new OutputFilters(_log);
                roleFilter.Copy(this);
                modelFilterMap.Add(item.Key, roleFilter);
            }

            Copy(saveFilter); //reset this filter back to state at top of function
            return modelFilterMap;
        }

        /// <summary>
        /// Fill RolesFilterHolder with default values
        /// </summary>
        public void FillDefaultRolesFilterHolder()
        {
            foreach (RoleFilter role in Enum.GetValues(typeof(RoleFilter)))
            {
                string roleFile = role.ToResourceName();
                if (!string.IsNullOrEmpty(roleFile))
                {
                    RolesFilterHolder[role] = new OutputFilters(roleFile, _log, role);
                }
            }
        }


        /// <summary>
        /// Fill FilterHolder From Directory, if no file use defaults config files in assembly
        /// </summary>
        /// <param name="dir">DirectoryInfo</param>
        public void FillRolesFilterHolderFromDir(DirectoryInfo dir)
        {
            if (!dir.Exists) 
                return;
            foreach (RoleFilter role in Enum.GetValues(typeof(RoleFilter)))
            {
                var fileName = Path.Combine(dir.FullName, role + "Filters.xml");
                var fileInfo = new FileInfo(fileName);
                if (fileInfo.Exists)
                {
                    RolesFilterHolder[role] = DeserializeXml(fileInfo);
                }
                else
                {
                    var roleFile = role.ToResourceName();
                    if (!string.IsNullOrEmpty(roleFile))
                    {
                        RolesFilterHolder[role] = new OutputFilters(roleFile, _log, role);
                    }
                }
            }
        }

        /// <summary>
        /// Write to xml roleFilter files on passed directory
        /// </summary>
        /// <param name="dir">DirectoryInfo</param>
        public void WriteXmlRolesFilterHolderToDir(DirectoryInfo dir)
        {
            if (!dir.Exists) 
                return;
            foreach (var item in RolesFilterHolder)
            {
                var fileName = Path.Combine(dir.FullName, item.Key + "Filters.xml");
                var fileInfo = new FileInfo(fileName);
                item.Value.SerializeXml(fileInfo);
            }
        }

        /// <summary>
        /// Get stored role filter
        /// </summary>
        /// <param name="role">RoleFilter with single flag(role) set</param>
        /// <returns>OutPutFilters</returns>
        public OutputFilters GetRoleFilter(RoleFilter role)
        {
            if ((role & (role - 1)) != 0)
            {
                throw new ArgumentException("More than one flag set on role");
            }

            if (RolesFilterHolder.ContainsKey(role))
            {
                return RolesFilterHolder[role];
            }
            //load defaults
            var objFilter = GetDefaults(role, _log);
            if (objFilter != null)
            {
                RolesFilterHolder[role] = objFilter;
            }
            return RolesFilterHolder[role];
        }

        /// <summary>
        /// Get the default filters for a single role
        /// </summary>
        /// <param name="role">RoleFilter with single flag(role) set</param>
        /// <param name="log"></param>
        /// <returns>OutPutFilters</returns>
        public static OutputFilters GetDefaults(RoleFilter role, ILogger log)
        {
            if ((role & (role - 1)) != 0)
            {
                throw new ArgumentException("More than one flag set on role");
            }
            var filterFile = role.ToResourceName();
            return !string.IsNullOrEmpty(filterFile) 
                ? new OutputFilters(filterFile, log, role) 
                : null;
        }



        /// <summary>
        /// Add filter for a role, used by ApplyRoleFilters for none default filters
        /// </summary>
        /// <param name="role">RoleFilter, single flag RoleFilter</param>
        /// <param name="filter">OutPutFilters to assign to role</param>
        /// <remarks>Does not apply filter to this object, used ApplyRoleFilters after setting the RolesFilterHolder items </remarks>
        public void AddRoleFilterHolderItem(RoleFilter role, OutputFilters filter)
        {
            if ((role & (role - 1)) != 0)
            {
                throw new ArgumentException("More than one flag set on role");
            }

            RolesFilterHolder[role] = filter;
        }

        #endregion

        #region Serialize

        /// <summary>
        /// Save object as xml file
        /// </summary>
        /// <param name="filename">FileInfo</param>
        public void SerializeXml(FileInfo filename)
        {
            var writer = new XmlSerializer(typeof(OutputFilters));
            using (var file = new StreamWriter(filename.FullName))
            {
                writer.Serialize(file, this);
            }
        }

        /// <summary>
        /// Create a OutPutFilters object from a XML file
        /// </summary>
        /// <param name="filename">FileInfo</param>
        /// <returns>OutPutFilters</returns>
        public static OutputFilters DeserializeXml(FileInfo filename)
        {
            OutputFilters result;
            var writer = new XmlSerializer(typeof(OutputFilters));
            using (var file = new StreamReader(filename.FullName))
            {
                result =  (OutputFilters)writer.Deserialize(file);
            }
            return result;
        }

        /// <summary>
        /// Save object as JSON 
        /// </summary>
        /// <param name="filename">FileInfo</param>
        public void SerializeJson (FileInfo filename)
        {
            var writer = new JsonSerializer();
            using (var file = new StreamWriter(filename.FullName))
            {
                writer.Serialize(file, this);
            }
        }

        /// <summary>
        /// Create a OutPutFilters object from a JSON file
        /// </summary>
        /// <param name="filename">FileInfo</param>
        /// <returns>OutPutFilters</returns>
        public static OutputFilters DeserializeJson(FileInfo filename)
        {
            OutputFilters result;
            var writer = new JsonSerializer();
            using (var file = new StreamReader(filename.FullName))
            {
                result = (OutputFilters)writer.Deserialize(file, typeof(OutputFilters));
            }
            return result;
        }

        #endregion
    }

    [Obsolete("Use OutputFilters instead")]
    public class OutPutFilters : OutputFilters
    {
        public OutPutFilters(ILogger logger) : base(logger)
        {
        }

        public OutPutFilters(ILogger logger, RoleFilter roleFlags) : base(logger, roleFlags)
        {
        }

        public OutPutFilters(ILogger logger, RoleFilter roles, Dictionary<RoleFilter, OutputFilters> rolesFilter = null) : base(logger, roles, rolesFilter)
        {
        }

        public OutPutFilters(string configFileName, ILogger logger, RoleFilter roleFlags, ImportSet setsToImport = ImportSet.All) : base(configFileName, logger, roleFlags, setsToImport)
        {
        }
    }
}
