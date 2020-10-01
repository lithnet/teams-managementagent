extern alias BetaLib;
using Beta = BetaLib.Microsoft.Graph;
using System.Collections.ObjectModel;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Microsoft.MetadirectoryServices;
using Newtonsoft.Json;
using NLog;
using Logger = NLog.Logger;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class GraphConnectionContext : IConnectionContext
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public GraphServiceClient Client { get; private set; }

        public  BetaLib::Microsoft.Graph.GraphServiceClient BetaClient { get; private set; }

        internal UserFilter UserFilter { get; private set; }

        internal static GraphConnectionContext GetConnectionContext(KeyedCollection<string, ConfigParameter> configParameters)
        {
            logger.Info($"Setting up connection to {configParameters[ConfigParameterNames.TenantDomain].Value}");

            System.Net.ServicePointManager.DefaultConnectionLimit = MicrosoftTeamsMAConfigSection.Configuration.ConnectionLimit;
            GlobalSettings.ExportThreadCount = MicrosoftTeamsMAConfigSection.Configuration.ExportThreads;

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(configParameters[ConfigParameterNames.ClientId].Value)
                .WithTenantId(configParameters[ConfigParameterNames.TenantDomain].Value)
                .WithClientSecret(configParameters[ConfigParameterNames.Secret].SecureValue.ConvertToUnsecureString())
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            JsonConvert.DefaultSettings = () => new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore
            };

            var client = new GraphServiceClient(authProvider);
            var betaClient = new  BetaLib::Microsoft.Graph.GraphServiceClient(authProvider);

            return new GraphConnectionContext()
            {
                Client = client,
                BetaClient = betaClient,
                UserFilter = new UserFilter(client, configParameters)
            };
        }
    }
}
