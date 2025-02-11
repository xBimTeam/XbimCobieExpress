using System;
using System.Collections.Concurrent;
using System.Globalization;

namespace Xbim.CobieExpress.Exchanger
{
    public abstract class XbimExchanger<TSourceRepository, TTargetRepository> 
    {
        private readonly ConcurrentDictionary<Type, ConcurrentDictionary<Type, IXbimMappings<TSourceRepository, TTargetRepository>>> _mappings = new ConcurrentDictionary<Type, ConcurrentDictionary<Type, IXbimMappings<TSourceRepository, TTargetRepository>>>();

        /// <summary>
        /// This property can be used by Exchanger to set up a context for all mappings (like a specific stage of project for example).
        /// </summary>
        public object Context { get; protected set; }

        protected XbimExchanger()
        {
            ReportProgress = new ProgressReporter(); //no ReportProgressDelegate set so will not report
        }

        public TTargetRepository TargetRepository { get; private set; }
        public TSourceRepository SourceRepository { get; private set; }

        /// <summary>
        /// Object to use to report progress on Exchangers
        /// </summary>
        public ProgressReporter ReportProgress { get; private set; }

        protected void Initialise(TSourceRepository source, TTargetRepository target)
        {
            TargetRepository = target;
            SourceRepository = source;
        }

        public TMapping GetOrCreateMappings<TMapping>() where TMapping : IXbimMappings<TSourceRepository, TTargetRepository>, new()
        {
            var mappings = new TMapping {Exchanger = this};
            ConcurrentDictionary<Type, IXbimMappings<TSourceRepository, TTargetRepository>> toMappings;
            if (_mappings.TryGetValue(mappings.MapFromType, out toMappings))
            {
                IXbimMappings<TSourceRepository, TTargetRepository> imappings;
                if (toMappings.TryGetValue(mappings.MapToType, out imappings))
                    return (TMapping)imappings;
                toMappings.TryAdd(mappings.MapToType, mappings);
                return mappings;
            }
            toMappings = new ConcurrentDictionary<Type, IXbimMappings<TSourceRepository, TTargetRepository>>();
            toMappings.TryAdd(mappings.MapToType, mappings);
            _mappings.TryAdd(mappings.MapFromType, toMappings);
            return mappings;
        }

        public abstract TTargetRepository Convert();

        private int _strId;
        public virtual string GetStringIdentifier() {
            return (_strId++).ToString(CultureInfo.InvariantCulture);
        }

        private int _intId;
        public virtual int GetIntIdentifier()
        {
            return _intId++;
        }
        public virtual Guid GetGuidIdentifier()
        {
            return Guid.NewGuid();
        }       
    }
}
