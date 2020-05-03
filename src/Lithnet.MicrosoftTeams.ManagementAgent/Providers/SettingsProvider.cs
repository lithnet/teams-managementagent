using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;
using NLog.Config;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class SettingsProvider : ISettingsProvider
    {
        public LoggingConfiguration GetCustomLogConfiguration(KeyedCollection<string, ConfigParameter> configParameters)
        {
            throw new NotImplementedException();
        }

        public string ManagementAgentName => "Lithnet Microsoft Teams Management Agent";

        public bool HandleOwnLogConfiguration => false;
    }
}
