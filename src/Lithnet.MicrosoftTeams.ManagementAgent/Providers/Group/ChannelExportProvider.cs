extern alias BetaLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Beta = BetaLib.Microsoft.Graph;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.MetadirectoryServices;
using NLog;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class ChannelExportProvider : IObjectExportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private HashSet<string> usersToIgnore = new HashSet<string>();

        private IExportContext context;

        private GraphServiceClient client;

        private Beta.GraphServiceClient betaClient;

        private CancellationToken token;

        public void Initialize(IExportContext context)
        {
            this.context = context;
            this.token = context.Token;
            this.betaClient = ((GraphConnectionContext)context.ConnectionContext).BetaClient;
            this.client = ((GraphConnectionContext)context.ConnectionContext).Client;
            this.BuildUsersToIgnore();
        }

        public bool CanExport(CSEntryChange csentry)
        {
            return csentry.ObjectType == "publicChannel" || csentry.ObjectType == "privateChannel";
        }

        private void BuildUsersToIgnore()
        {
            this.usersToIgnore.Clear();

            if (this.context.ConfigParameters.Contains(ConfigParameterNames.UsersToIgnore))
            {
                string raw = this.context.ConfigParameters[ConfigParameterNames.UsersToIgnore].Value;

                if (!string.IsNullOrWhiteSpace(raw))
                {
                    foreach (string user in raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        this.usersToIgnore.Add(user.ToLower().Trim());
                    }
                }
            }
        }

        public async Task<CSEntryChangeResult> PutCSEntryChangeAsync(CSEntryChange csentry)
        {
            return await this.PutCSEntryChangeObject(csentry);
        }

        public async Task<CSEntryChangeResult> PutCSEntryChangeObject(CSEntryChange csentry)
        {
            switch (csentry.ObjectModificationType)
            {
                case ObjectModificationType.Add:
                case ObjectModificationType.Update when csentry.AnchorAttributes.Count == 0:
                    return await this.PutCSEntryChangeAdd(csentry);

                case ObjectModificationType.Delete:
                    return await this.PutCSEntryChangeDelete(csentry);

                case ObjectModificationType.Update:
                    return await this.PutCSEntryChangeUpdate(csentry);

                default:
                    throw new InvalidOperationException($"Unknown or unsupported modification type: {csentry.ObjectModificationType} on object {csentry.DN}");
            }
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeDelete(CSEntryChange csentry)
        {
            string teamid = csentry.GetAnchorValueOrDefault<string>("teamid");

            try
            {
                if (this.context.ConfigParameters[ConfigParameterNames.RandomizeChannelNameOnDelete].Value == "1")
                {
                    string newname = $"deleted-{Guid.NewGuid():N}";
                    Beta.Channel c = new Beta.Channel();
                    c.DisplayName = newname;
                    await GraphHelperTeams.UpdateChannel(this.betaClient, teamid, csentry.DN, c, this.token);
                    logger.Info($"Renamed channel {csentry.DN} on team {teamid} to {newname}");
                }

                await GraphHelperTeams.DeleteChannel(this.betaClient, teamid, csentry.DN, this.token);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.NotFound)
                {
                    logger.Warn($"The request to delete the channel {csentry.DN} failed because it doesn't exist");
                }
                else
                {
                    throw;
                }
            }

            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeAdd(CSEntryChange csentry)
        {
            string teamid = csentry.GetValueAdd<string>("team");

            if (teamid == null)
            {
                logger.Error($"Export of item {csentry.DN} failed as no 'team' attribute was provided\r\n{string.Join(",", csentry.ChangedAttributeNames)}");
                return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.ExportActionRetryReferenceAttribute);
            }

            Beta.Channel c = new Beta.Channel();
            c.DisplayName = csentry.GetValueAdd<string>("displayName");
            c.Description = csentry.GetValueAdd<string>("description");

            // 2020-05-15 This currently doesn't come in with a GET request, so for now, it needs to be initial-flow only
            // https://github.com/microsoftgraph/microsoft-graph-docs/issues/6792
            c.IsFavoriteByDefault = csentry.HasAttributeChange("isFavoriteByDefault") && csentry.GetValueAdd<bool>("isFavoriteByDefault");

            if (csentry.ObjectType == "privateChannel")
            {
                if (!csentry.HasAttributeChangeAdd("owner"))
                {
                    throw new UnexpectedDataException("At least one owner must be specified when creating a channel");
                }

                string ownerID = csentry.GetValueAdds<string>("owner").First();
                c.MembershipType = Beta.ChannelMembershipType.Private;
                Beta.AadUserConversationMember a = new Beta.AadUserConversationMember() { Roles = new[] { "owner" } };
                a.AdditionalData = new Dictionary<string, object>();
                a.AdditionalData.Add("user@odata.bind", $"https://graph.microsoft.com/v1.0/user('{ownerID}')");
                c.AdditionalData = new Dictionary<string, object>();
                c.AdditionalData.Add("members", new Beta.AadUserConversationMember[] {a});
            }

            logger.Trace($"{csentry.DN}: Channel data: {JsonConvert.SerializeObject(c)}");

            var channel = await GraphHelperTeams.CreateChannel(this.betaClient, teamid, c, this.token);

            logger.Trace($"Created channel {channel.Id} for team {teamid}");

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", channel.Id));
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("teamid", teamid));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeUpdate(CSEntryChange csentry)
        {
            await this.PutCSEntryChangeUpdateChannel(csentry);
            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task PutCSEntryChangeUpdateChannel(CSEntryChange csentry)
        {
            string teamid = csentry.GetAnchorValueOrDefault<string>("teamid");
            string channelid = csentry.GetAnchorValueOrDefault<string>("id");

            bool changed = false;
            Beta.Channel channel = new Beta.Channel();

            foreach (AttributeChange change in csentry.AttributeChanges)
            {
                if (change.DataType == AttributeType.Boolean && change.ModificationType == AttributeModificationType.Delete)
                {
                    throw new UnknownOrUnsupportedModificationTypeException($"The property {change.Name} cannot be deleted. If it is a boolean value, set it to false");
                }

                if (change.Name == "team")
                {
                    throw new UnexpectedDataException("The team parameter can only be supplied during an 'add' operation");
                }
                else if (change.Name == "isFavoriteByDefault")
                {
                    channel.IsFavoriteByDefault = change.GetValueAdd<bool>();
                }
                else if (change.Name == "displayName")
                {
                    if (change.ModificationType == AttributeModificationType.Delete)
                    {
                        throw new UnknownOrUnsupportedModificationTypeException($"The property {change.Name} cannot be deleted");
                    }

                    channel.DisplayName = change.GetValueAdd<string>();
                }
                else if (change.Name == "description")
                {
                    if (change.ModificationType == AttributeModificationType.Delete)
                    {
                        channel.AssignNullToProperty("description");
                    }
                    else
                    {
                        channel.Description = change.GetValueAdd<string>();
                    }
                }
                else
                {
                    continue;
                }

                changed = true;
            }

            if (changed)
            {
                logger.Trace($"{csentry.DN}:Updating channel data: {JsonConvert.SerializeObject(channel)}");
                await GraphHelperTeams.UpdateChannel(this.betaClient, teamid, channelid, channel, this.token);
                logger.Info($"{csentry.DN}: Updated channel");
            }

            if (csentry.ObjectType == "privateChannel")
            {
                await this.PutMemberChanges(csentry);
            }
        }

        private async Task PutMemberChanges(CSEntryChange csentry)
        {
            if (!csentry.HasAttributeChange("member") && !csentry.HasAttributeChange("owner"))
            {
                return;
            }

            string teamid = csentry.GetAnchorValueOrDefault<string>("teamid");
            string channelid = csentry.GetAnchorValueOrDefault<string>("id");

            IList<string> memberAdds = csentry.GetValueAdds<string>("member");
            IList<string> memberDeletes = csentry.GetValueDeletes<string>("member");
            IList<string> ownerAdds = csentry.GetValueAdds<string>("owner");
            IList<string> ownerDeletes = csentry.GetValueDeletes<string>("owner");

            if (csentry.HasAttributeChangeDelete("member") || csentry.HasAttributeChangeDelete("owner"))
            {
                var existingMembership = await GraphHelperTeams.GetChannelMembers(this.betaClient, teamid, channelid, this.token);

                if (csentry.HasAttributeChangeDelete("member"))
                {
                    memberDeletes = existingMembership.Where(t => t.Roles.All(u => !string.Equals(u, "owner", StringComparison.OrdinalIgnoreCase))).Select(t => t.Id).ToList();
                }

                if (csentry.HasAttributeChangeDelete("owner"))
                {
                    ownerDeletes = existingMembership.Where(t => t.Roles.Any(u => string.Equals(u, "owner", StringComparison.OrdinalIgnoreCase))).Select(t => t.Id).ToList();
                }
            }

            var memberUpgradesToOwners = memberDeletes.Intersect(ownerAdds).ToList();

            foreach (var m in memberUpgradesToOwners)
            {
                memberDeletes.Remove(m);
                ownerAdds.Remove(m);
            }

            var ownerDowngradeToMembers = ownerDeletes.Intersect(memberAdds).ToList();

            foreach (var m in ownerDowngradeToMembers)
            {
                memberAdds.Remove(m);
                ownerDeletes.Remove(m);
            }

            List<Beta.AadUserConversationMember> cmToAdd = new List<Beta.AadUserConversationMember>();
            List<Beta.AadUserConversationMember> cmToDelete = new List<Beta.AadUserConversationMember>();
            List<Beta.AadUserConversationMember> cmToUpdate = new List<Beta.AadUserConversationMember>();

            foreach (var m in memberDeletes)
            {
                cmToDelete.Add(new Beta.AadUserConversationMember() { UserId = m });
            }

            // If we try to delete the last owner on a channel, the operation will fail. If we are swapping out the full set of owners (eg an add/delete of 100 owners), this will never succeed if we do a 'delete' operation first.
            // If we do an 'add' operation first, and the channel already has the maximum number of owners, the call will fail.
            // So the order of events should be to
            //    1) Process all membership removals except for one owner (100-99 = 1 owner)
            //    2) Process all membership adds except for one owner (1 + 99 = 100 owners)
            //    3) Remove the final owner (100 - 1 = 99 owners)
            //    4) Add the final owner (99 + 1 = 100 owners)

            string lastOwnerToRemove = null;
            if (ownerDeletes.Count > 0)
            {
                lastOwnerToRemove = ownerDeletes[0];
                ownerDeletes.RemoveAt(0);
            }

            string lastOwnerToAdd = null;
            if (ownerAdds.Count > 0)
            {
                lastOwnerToAdd = ownerAdds[0];
                ownerAdds.RemoveAt(0);
            }

            foreach (var m in ownerDeletes)
            {
                cmToDelete.Add(new Beta.AadUserConversationMember() { UserId = m });
            }

            foreach (var m in memberAdds)
            {
                cmToAdd.Add(new Beta.AadUserConversationMember { UserId = m, AdditionalData = new Dictionary<string, object>() { { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{m}" } } });
            }

            foreach (var m in ownerAdds)
            {
                cmToAdd.Add(new Beta.AadUserConversationMember { UserId = m, Roles = new[] { "owner" }, AdditionalData = new Dictionary<string, object>() { { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{m}" } } });
            }

            foreach (var m in memberUpgradesToOwners)
            {
                cmToUpdate.Add(new Beta.AadUserConversationMember { UserId = m, Roles = new[] { "owner" }, AdditionalData = new Dictionary<string, object>() { { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{m}" } } });
            }

            foreach (var m in ownerDowngradeToMembers)
            {
                cmToUpdate.Add(new Beta.AadUserConversationMember { UserId = m, Roles = new string[] { }, AdditionalData = new Dictionary<string, object>() { { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{m}" } } });
            }

            await GraphHelperTeams.RemoveChannelMembers(this.betaClient, teamid, channelid, cmToDelete, true, this.token);
            await GraphHelperTeams.UpdateChannelMembers(this.betaClient, teamid, channelid, cmToUpdate, this.token);
            await GraphHelperTeams.AddChannelMembers(this.betaClient, teamid, channelid, cmToAdd, true, this.token);

            if (lastOwnerToRemove != null)
            {
                cmToDelete.Clear();
                cmToDelete.Add(new Beta.AadUserConversationMember() { UserId = lastOwnerToRemove });
                await GraphHelperTeams.RemoveChannelMembers(this.betaClient, teamid, channelid, cmToDelete, true, this.token);
            }

            if (lastOwnerToAdd != null)
            {
                cmToAdd.Clear();
                cmToAdd.Add(new Beta.AadUserConversationMember { UserId = lastOwnerToAdd, Roles = new[] { "owner" }, AdditionalData = new Dictionary<string, object>() { { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{lastOwnerToAdd}" } } });
                await GraphHelperTeams.AddChannelMembers(this.betaClient, teamid, channelid, cmToAdd, true, this.token);
            }

            logger.Info($"Membership modification for channel {teamid}:{channelid} completed. Members added: {memberAdds.Count}, members removed: {memberDeletes.Count}, owners added: {ownerAdds.Count}, owners removed: {ownerDeletes.Count}, owners downgraded to members: {ownerDowngradeToMembers.Count}, members upgraded to owners: {memberUpgradesToOwners.Count}");
        }
    }
}
