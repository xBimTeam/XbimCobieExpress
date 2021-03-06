﻿using System;
using System.Collections.Generic;

namespace Xbim.CobieExpress.Exchanger
{
    public interface IXbimMappings<TSourceRepository, TTargetRepository>
    {

        Type MapFromType { get; }
        Type MapToType { get; }

        Type MapKeyType { get; }

        XbimExchanger<TSourceRepository, TTargetRepository> Exchanger { get; set; }

        IDictionary<object, object> Mappings { get; }

        object CreateTargetObject();
        bool GetTargetObject(object key, out object targetObject);
        object GetOrCreateTargetObject(object key);
        object AddMapping(object source, object target);

    }
}
