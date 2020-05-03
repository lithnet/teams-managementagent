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

        internal static readonly string FilterQueryDescription = "You can optionally specify a filter to only import teams for groups that match a particular figure. For example, to filter on groups that have a mailNickname starting with 'mim-', use the filter startswith(mailNickname,'mim-'). See the Microsoft Graph API documentation for support filter queries on Azure groups. Leave this blank to see and manage all teams";

        internal static readonly string ChannelNameFilter = "Channel name filter";

        internal static readonly string ChannelNameFilterDescription = "You can use any optional regular expression to filter based on channel names. Leave this blank to include all channels. For example, to ignore all the general channel for each team, use the filter such as ^(?!general$).*$";

        internal static readonly string RandomizeChannelNameOnDelete = "Randomize channel names before deleting";

        internal static readonly string RandomizeChannelNameOnDeleteDecription = "If you delete a channel, you cannot create a new channel with the same name. The MA can rename the channel before it is deleted, so the old name can be used against if needed";

        internal static readonly string UsersToIgnore = "User IDs to ignore";

        internal static readonly string UsersToIgnoreDescription = "An optional comma-separated list of user IDs to ignore from a group's membership. If these user IDs are found as group owners or members, they will not be reported back to the sync engine. Note these values must be azure object IDs, not user principal names";
    }
}