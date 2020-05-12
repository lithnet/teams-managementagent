using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using Microsoft.MetadirectoryServices.DetachedObjectModel;
using System.Security;

namespace Lithnet.MicrosoftTeams.ManagementAgent.Tests
{
    public class ConfigurationParameters : KeyedCollection<string, ConfigParameter>
    {
        public ConfigurationParameters()
        {
            this.Add(new ConfigParameter(ConfigParameterNames.TenantDomain, ConfigurationManager.AppSettings["tenant-domain"]));
            this.Add(new ConfigParameter(ConfigParameterNames.ClientId, ConfigurationManager.AppSettings["client-id"]));

            SecureString ss = new SecureString();
            Array.ForEach(ConfigurationManager.AppSettings["secret"].ToArray(), ss.AppendChar);
            ss.MakeReadOnly();

            this.Add(new ConfigParameter(ConfigParameterNames.Secret, ss));
            this.Add(new ConfigParameter(ConfigParameterNames.FilterQuery, ConfigurationManager.AppSettings["filter-query"]));
            this.Add(new ConfigParameter(ConfigParameterNames.ChannelNameFilter, ConfigurationManager.AppSettings["channel-name-filter"]));
            this.Add(new ConfigParameter(ConfigParameterNames.RandomizeChannelNameOnDelete, ConfigurationManager.AppSettings["randomize-channel-name-on-delete"] == "true" ? "1" : "0"));
            this.Add(new ConfigParameter(ConfigParameterNames.UsersToIgnore, ConfigurationManager.AppSettings["users-to-ignore"]));
        }

        protected override string GetKeyForItem(ConfigParameter item)
        {
            return item.Name;
        }
    }
}
