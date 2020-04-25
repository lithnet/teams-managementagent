using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Net.Http;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.MetadirectoryServices;
using NLog;
using NLog.Config;
using Logger = NLog.Logger;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class GraphConnectionContext : IConnectionContext
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public GraphServiceClient Client { get; private set; }

        internal static GraphConnectionContext GetConnectionContext(KeyedCollection<string, ConfigParameter> configParameters)
        {
            GraphConnectionContext.logger.Info($"Setting up connection to {configParameters[ConfigParameterNames.TenantDomain].Value}");

            System.Net.ServicePointManager.DefaultConnectionLimit = MicrosoftTeamsMAConfigSection.Configuration.ConnectionLimit;
            GlobalSettings.ExportThreadCount = MicrosoftTeamsMAConfigSection.Configuration.ExportThreads;

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(configParameters[ConfigParameterNames.ClientId].Value)
                .WithTenantId(configParameters[ConfigParameterNames.TenantDomain].Value)
                .WithClientSecret(configParameters[ConfigParameterNames.Secret].SecureValue.ConvertToUnsecureString())
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            return new GraphConnectionContext()
            {
                Client = new GraphServiceClient(authProvider)
            };
        }
    }
}
