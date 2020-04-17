﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class GraphConnectionContextProvider : IConnectionContextProvider
    {
        public object GetConnectionContext(KeyedCollection<string, ConfigParameter> configParameters, ConnectionContextOperationType contextOperationType)
        {
            return GraphConnectionContext.GetConnectionContext(configParameters);
        }
    }
}
