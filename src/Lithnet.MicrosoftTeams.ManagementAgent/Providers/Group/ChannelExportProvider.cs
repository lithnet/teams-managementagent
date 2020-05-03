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
    internal class ChannelExportProvider : IObjectExportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public bool CanExport(CSEntryChange csentry)
        {
            return csentry.ObjectType == "publicChannel" || csentry.ObjectType == "privateChannel";
        }

        public async Task<CSEntryChangeResult> PutCSEntryChangeAsync(CSEntryChange csentry, ExportContext context)
        {
            return await this.PutCSEntryChangeObject(csentry, context);
        }

        public async Task<CSEntryChangeResult> PutCSEntryChangeObject(CSEntryChange csentry, ExportContext context)
        {
            switch (csentry.ObjectModificationType)
            {
                case ObjectModificationType.Add:
                    return await this.PutCSEntryChangeAdd(csentry, context);

                case ObjectModificationType.Delete:
                    return await this.PutCSEntryChangeDelete(csentry, context);

                case ObjectModificationType.Update:
                    return await this.PutCSEntryChangeUpdate(csentry, context);

                default:
                    throw new InvalidOperationException($"Unknown or unsupported modification type: {csentry.ObjectModificationType} on object {csentry.DN}");
            }
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeDelete(CSEntryChange csentry, ExportContext context)
        {
            var client = ((GraphConnectionContext)context.ConnectionContext).BetaClient;
            string teamid = csentry.GetAnchorValueOrDefault<string>("teamid");

            try
            {
                if (context.ConfigParameters[ConfigParameterNames.RandomizeChannelNameOnDelete].Value == "1")
                {
                    string newname = $"deleted-{Guid.NewGuid():N}";
                    Beta.Channel c = new Beta.Channel();
                    c.DisplayName = newname;
                    await GraphHelperTeams.UpdateChannel(client, teamid, csentry.DN, c, context.Token);
                    logger.Info($"Renamed channel {csentry.DN} on team {teamid} to {newname}");
                }

                await GraphHelperTeams.DeleteChannel(client, teamid, csentry.DN, context.Token);
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

        private async Task<CSEntryChangeResult> PutCSEntryChangeAdd(CSEntryChange csentry, ExportContext context)
        {
            var client = ((GraphConnectionContext)context.ConnectionContext).BetaClient;

            string teamid = csentry.GetValueAdd<string>("team");

            if (teamid == null)
            {
                logger.Error($"Export of item {csentry.DN} failed as no 'team' attribute was provided\r\n{string.Join(",", csentry.ChangedAttributeNames)}");
                return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.ExportActionRetryReferenceAttribute);
            }

            Beta.Channel c = new Beta.Channel();
            c.DisplayName = csentry.GetValueAdd<string>("displayName");

            var channel = await GraphHelperTeams.CreateChannel(client, teamid, c, context.Token);

            logger.Trace($"Created channel {channel.Id} for team {teamid}");

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", channel.Id));
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("teamid", teamid));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeUpdate(CSEntryChange csentry, ExportContext context)
        {
            if (csentry.GetAnchorValueOrDefault<string>("id") == null)
            {
                logger.Warn($"Resubmitting object update to add queue as no anchor was present\r\n{string.Join(",", csentry.ChangedAttributeNames)}");
                return await this.PutCSEntryChangeAdd(csentry, context);
            }

            await this.PutCSEntryChangeUpdateChannel(csentry, context);
            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task PutCSEntryChangeUpdateChannel(CSEntryChange csentry, ExportContext context)
        {
            var client = ((GraphConnectionContext)context.ConnectionContext).BetaClient;

            bool changed = false;

            foreach (AttributeChange change in csentry.AttributeChanges)
            {
                if (!SchemaProvider.TeamsProperties.Contains(change.Name))
                {
                    continue;
                }

                if (change.ModificationType == AttributeModificationType.Delete)
                {
                    throw new UnknownOrUnsupportedModificationTypeException($"The property {change.Name} cannot be deleted. If it is a boolean value, set it to false");
                }

                if (change.Name == "visibility")
                {
                    throw new UnexpectedDataException("The visibility parameter can only be supplied during an 'add' operation");
                }
                else if (change.Name == "template")
                {
                    throw new UnexpectedDataException("The template parameter can only be supplied during an 'add' operation");
                }
                else if (change.Name == "isArchived")
                {
                
                  
                }
                else
                {
                    continue;
                }

                changed = true;
            }

            if (changed)
            {
               // logger.Trace($"{csentry.DN}:Updating team data: {JsonConvert.SerializeObject(team)}");

               // await GraphHelperTeams.UpdateTeam(client, csentry.DN, team, context.Token);

                //logger.Info($"{csentry.DN}: Updated team");
            }
        }

        private async Task PutCSEntryChangeUpdateGroup(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Group group = new Group();
            bool changed = false;

            foreach (AttributeChange change in csentry.AttributeChanges)
            {
                if (SchemaProvider.GroupMemberProperties.Contains(change.Name))
                {
                    await this.PutAttributeChangeMembers(csentry.DN, change, context);
                    continue;
                }

                if (!SchemaProvider.GroupFromTeamProperties.Contains(change.Name))
                {
                    continue;
                }

                if (change.ModificationType == AttributeModificationType.Delete)
                {
                    this.AssignNullToProperty(change.Name, group);
                    continue;
                }

                if (change.Name == "displayName")
                {
                    group.DisplayName = change.GetValueAdd<string>();
                }
                else if (change.Name == "description")
                {
                    group.Description = change.GetValueAdd<string>();
                }
                else if (change.Name == "mailNickname")
                {
                    group.MailNickname = change.GetValueAdd<string>();
                }
                else
                {
                    continue;
                }

                changed = true;
            }

            if (changed)
            {
                logger.Trace($"{csentry.DN}:Updating group data: {JsonConvert.SerializeObject(group)}");
                await GraphHelperGroups.UpdateGroup(client, csentry.DN, group, context.Token);
                logger.Info($"{csentry.DN}: Updated group {csentry.DN}");
            }
        }

        private async Task PutAttributeChangeMembers(string groupid, AttributeChange change, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            IList<string> valueDeletes = change.GetValueDeletes<string>();
            IList<string> valueAdds = change.GetValueAdds<string>();

            if (change.ModificationType == AttributeModificationType.Delete)
            {
                if (change.Name == "member")
                {
                    List<DirectoryObject> result = await GraphHelperGroups.GetGroupMembers(client, groupid, context.Token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
                else
                {
                    List<DirectoryObject> result = await GraphHelperGroups.GetGroupOwners(client, groupid, context.Token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
            }

            if (change.Name == "member")
            {
                await GraphHelperGroups.AddGroupMembers(client, groupid, valueAdds, true, context.Token);
                await GraphHelperGroups.RemoveGroupMembers(client, groupid, valueDeletes, true, context.Token);
                logger.Info($"Membership modification for group {groupid} completed. Members added: {valueAdds.Count}, members removed: {valueDeletes.Count}");
            }
            else
            {
                await GraphHelperGroups.AddGroupOwners(client, groupid, valueAdds, true, context.Token);
                await GraphHelperGroups.RemoveGroupOwners(client, groupid, valueDeletes, true, context.Token);
                logger.Info($"Owner modification for group {groupid} completed. Owners added: {valueAdds.Count}, owners removed: {valueDeletes.Count}");
            }
        }

        private void AssignNullToProperty(string name, Group group)
        {
            if (group.AdditionalData == null)
            {
                group.AdditionalData = new Dictionary<string, object>();
            }

            group.AdditionalData.Add(name, null);
        }
    }
}
