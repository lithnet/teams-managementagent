using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Security;
using Microsoft.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public static class ConfigParameterNames
    {
        internal static readonly string LogFileName = "Log file";

        internal static readonly string ClientId = "Client ID";

        internal static readonly string Secret  = "Secret";

        internal static readonly string TenantDomain = "Tenant Domain";

        internal static readonly string FilterQuery = "Teams query filter";

    }
}