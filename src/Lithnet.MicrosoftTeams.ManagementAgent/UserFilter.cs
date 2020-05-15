using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class UserFilter
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private bool initialized;

        private HashSet<string> usersToIgnore;

        private KeyedCollection<string, ConfigParameter> configParameters;

        private GraphServiceClient client;

        public UserFilter(GraphServiceClient client, KeyedCollection<string, ConfigParameter> configParameters)
        {
            this.configParameters = configParameters;
            this.client = client;
            this.initialized = false;
            this.usersToIgnore = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        public bool ShouldExclude(string userid, CancellationToken token)
        {
            if (!this.initialized)
            {
                lock (this.usersToIgnore)
                {
                    if (!this.initialized)
                    {
                        this.Initialize(token);
                    }
                }
            }

            return this.usersToIgnore.Contains(userid);
        }

        private void Initialize(CancellationToken token)
        {
            this.BuildUsersToIgnore();
            this.BuildGuestList(token).GetAwaiter().GetResult();
            this.initialized = true;
        }

        private void BuildUsersToIgnore()
        {
            this.usersToIgnore.Clear();
            logger.Trace("Initializing user filter list");

            if (this.configParameters.Contains(ConfigParameterNames.UsersToIgnore))
            {
                string raw = this.configParameters[ConfigParameterNames.UsersToIgnore].Value;

                if (!string.IsNullOrWhiteSpace(raw))
                {
                    foreach (string user in raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        this.usersToIgnore.Add(user.ToLower().Trim());
                    }
                }
            }

            if (this.usersToIgnore.Count > 0)
            {
                logger.Trace($"Added {this.usersToIgnore.Count} users to ignore from the MA config");

            }
        }

        private async Task BuildGuestList(CancellationToken token)
        {
            int count = 0;

            if (!MicrosoftTeamsMAConfigSection.Configuration.ManageGuests)
            {
                var result = await GraphHelperUsers.GetGuestUsers(this.client, token);

                foreach (var item in result)
                {
                    this.usersToIgnore.Add(item);
                    count++;
                }
            }

            if (count > 0)
            {
                logger.Trace($"Added {count} guest users to the ignore list");

            }
        }
    }
}
