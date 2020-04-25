using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using Newtonsoft.Json;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GroupExportProvider : IObjectExportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private const int MaxReferencesPerCreateRequest = 20;

        public bool CanExport(CSEntryChange csentry)
        {
            return csentry.ObjectType == "group";
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
                case ObjectModificationType.None:
                case ObjectModificationType.Replace:
                case ObjectModificationType.Unconfigured:
                    throw new InvalidOperationException($"Unknown or unsupported modification type: {csentry.ObjectModificationType} on object {csentry.DN}");
            }
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeDelete(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            try
            {
                await GraphHelper.DeleteGroup(client, csentry.DN, context.Token);
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.NotFound)
                {
                    logger.Warn($"The request to delete the group {csentry.DN} failed because the group doesn't exist");
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
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Group result = null;

            try
            {
                result = await this.CreateGroup(csentry, context, client);
                await this.CreateTeam(csentry, client, result.Id, context);
            }
            catch
            {
                try
                {
                    if (result != null)
                    {
                        GroupExportProvider.logger.Error($"{csentry.DN}: An exception occurred while creating the team, rolling back the group by deleting it");
                        await Task.Delay(TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.PostGroupCreateDelay));
                        await GraphHelper.DeleteGroup(client, result.Id, CancellationToken.None);
                        GroupExportProvider.logger.Info($"{csentry.DN}: The group was deleted");
                    }
                }
                catch (Exception ex2)
                {
                    GroupExportProvider.logger.Error(ex2, $"{csentry.DN}: An exception occurred while rolling back the team");
                }

                throw;
            }

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", result.Id));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task CreateTeam(CSEntryChange csentry, GraphServiceClient client, string groupId, ExportContext context)
        {
            Team team = new Team
            {
                MemberSettings = new TeamMemberSettings(),
                GuestSettings = new TeamGuestSettings(),
                MessagingSettings = new TeamMessagingSettings(),
                FunSettings = new TeamFunSettings(),
                ODataType = null
            };

            team.MemberSettings.ODataType = null;
            team.GuestSettings.ODataType = null;
            team.MessagingSettings.ODataType = null;
            team.FunSettings.ODataType = null;
            team.AdditionalData = new Dictionary<string, object>();

            string template = csentry.GetValueAdd<string>("template") ?? "https://graph.microsoft.com/beta/teamsTemplates('standard')";

            // if (!string.IsNullOrWhiteSpace(template))
            // {
            //     team.AdditionalData.Add("template@odata.bind", template); //"https://graph.microsoft.com/beta/teamsTemplates('standard')"
            // }

            //team.AdditionalData.Add("group@odata.bind", $"https://graph.microsoft.com/v1.0/groups('{groupId}')");

            team.MemberSettings.AllowCreateUpdateChannels = csentry.HasAttributeChange("memberSettings_allowCreateUpdateChannels") ? csentry.GetValueAdd<bool>("memberSettings_allowCreateUpdateChannels") : default(bool?);
            team.MemberSettings.AllowDeleteChannels = csentry.HasAttributeChange("memberSettings_allowDeleteChannels") ? csentry.GetValueAdd<bool>("memberSettings_allowDeleteChannels") : default(bool?);
            team.MemberSettings.AllowAddRemoveApps = csentry.HasAttributeChange("memberSettings_allowAddRemoveApps") ? csentry.GetValueAdd<bool>("memberSettings_allowAddRemoveApps") : default(bool?);
            team.MemberSettings.AllowCreateUpdateRemoveTabs = csentry.HasAttributeChange("memberSettings_allowCreateUpdateRemoveTabs") ? csentry.GetValueAdd<bool>("memberSettings_allowCreateUpdateRemoveTabs") : default(bool?);
            team.MemberSettings.AllowCreateUpdateRemoveConnectors = csentry.HasAttributeChange("memberSettings_allowCreateUpdateRemoveConnectors") ? csentry.GetValueAdd<bool>("memberSettings_allowCreateUpdateRemoveConnectors") : default(bool?);
            team.GuestSettings.AllowCreateUpdateChannels = csentry.HasAttributeChange("guestSettings_allowCreateUpdateChannels") ? csentry.GetValueAdd<bool>("guestSettings_allowCreateUpdateChannels") : default(bool?);
            team.GuestSettings.AllowCreateUpdateChannels = csentry.HasAttributeChange("guestSettings_allowDeleteChannels") ? csentry.GetValueAdd<bool>("guestSettings_allowDeleteChannels") : default(bool?);
            team.MessagingSettings.AllowUserEditMessages = csentry.HasAttributeChange("messagingSettings_allowUserEditMessages") ? csentry.GetValueAdd<bool>("messagingSettings_allowUserEditMessages") : default(bool?);
            team.MessagingSettings.AllowUserDeleteMessages = csentry.HasAttributeChange("messagingSettings_allowUserDeleteMessages") ? csentry.GetValueAdd<bool>("messagingSettings_allowUserDeleteMessages") : default(bool?);
            team.MessagingSettings.AllowOwnerDeleteMessages = csentry.HasAttributeChange("messagingSettings_allowOwnerDeleteMessages") ? csentry.GetValueAdd<bool>("messagingSettings_allowOwnerDeleteMessages") : default(bool?);
            team.MessagingSettings.AllowTeamMentions = csentry.HasAttributeChange("messagingSettings_allowTeamMentions") ? csentry.GetValueAdd<bool>("messagingSettings_allowTeamMentions") : default(bool?);
            team.MessagingSettings.AllowChannelMentions = csentry.HasAttributeChange("messagingSettings_allowChannelMentions") ? csentry.GetValueAdd<bool>("messagingSettings_allowChannelMentions") : default(bool?);
            team.FunSettings.AllowGiphy = csentry.HasAttributeChange("funSettings_allowGiphy") ? csentry.GetValueAdd<bool>("funSettings_allowGiphy") : default(bool?);
            team.FunSettings.AllowStickersAndMemes = csentry.HasAttributeChange("funSettings_allowStickersAndMemes") ? csentry.GetValueAdd<bool>("funSettings_allowStickersAndMemes") : default(bool?);
            team.FunSettings.AllowCustomMemes = csentry.HasAttributeChange("funSettings_allowCustomMemes") ? csentry.GetValueAdd<bool>("funSettings_allowCustomMemes") : default(bool?);

            string gcr = csentry.GetValueAdd<string>("funSettings_giphyContentRating");
            if (!string.IsNullOrWhiteSpace(gcr))
            {
                if (!Enum.TryParse(gcr, false, out GiphyRatingType grt))
                {
                    throw new UnexpectedDataException($"The value '{gcr}' was not a supported value for funSettings_giphyContentRating. Supported values are (case sensitive) 'Strict' or 'Moderate'");
                }

                team.FunSettings.GiphyContentRating = grt;
            }

            GroupExportProvider.logger.Info($"{csentry.DN}: Creating team for group {groupId} using template {template ?? "standard"}");
            GroupExportProvider.logger.Trace($"{csentry.DN}: Team data: {JsonConvert.SerializeObject(team)}");

            Team tresult = await GraphHelper.CreateTeam(client, groupId, team, context.Token);

            GroupExportProvider.logger.Info($"{csentry.DN}: Created team {tresult?.Id ?? "<unknown id>"} for group {groupId}");
        }

        private async Task<Group> CreateGroup(CSEntryChange csentry, ExportContext context, GraphServiceClient client)
        {
            Group group = new Group();
            group.DisplayName = csentry.GetValueAdd<string>("displayName") ?? throw new UnexpectedDataException("The group must have a displayName");
            group.GroupTypes = new[] { "Unified" };
            group.MailEnabled = true;
            group.Description = csentry.GetValueAdd<string>("description");
            group.MailNickname = csentry.GetValueAdd<string>("mailNickname") ?? throw new UnexpectedDataException("The group must have a mailNickname");
            group.SecurityEnabled = false;
            group.AdditionalData = new Dictionary<string, object>();
            group.Id = csentry.DN;
            group.Visibility = csentry.GetValueAdd<string>("visibility");

            IList<string> members = csentry.GetValueAdds<string>("member") ?? new List<string>();
            IList<string> owners = csentry.GetValueAdds<string>("owner") ?? new List<string>();

            IList<string> deferredMembers = new List<string>();
            IList<string> createOpMembers = new List<string>();

            IList<string> deferredOwners = new List<string>();
            IList<string> createOpOwners = new List<string>();

            int memberCount = 0;

            if (owners.Count > 100)
            {
                throw new UnexpectedDataException($"The group creation request {csentry.DN} contained more than 100 owners");
            }

            foreach (string owner in owners)
            {
                if (memberCount >= GroupExportProvider.MaxReferencesPerCreateRequest)
                {
                    deferredOwners.Add(owner);
                }
                else
                {
                    createOpOwners.Add($"https://graph.microsoft.com/v1.0/users/{owner}");
                    memberCount++;
                }
            }

            foreach (string member in members)
            {
                if (memberCount >= GroupExportProvider.MaxReferencesPerCreateRequest)
                {
                    deferredMembers.Add(member);
                }
                else
                {
                    createOpMembers.Add($"https://graph.microsoft.com/v1.0/users/{member}");
                    memberCount++;
                }
            }

            if (createOpMembers.Count > 0)
            {
                group.AdditionalData.Add("members@odata.bind", createOpMembers.ToArray());
            }

            if (createOpOwners.Count > 0)
            {
                group.AdditionalData.Add("owners@odata.bind", createOpOwners.ToArray());
            }

            GroupExportProvider.logger.Trace($"{csentry.DN}: Creating group {group.MailNickname} with {createOpMembers.Count} members and {createOpOwners.Count} owners (Deferring {deferredMembers.Count} members and {deferredOwners.Count} owners until after group creation)");

            GroupExportProvider.logger.Trace($"{csentry.DN}: Group data: {JsonConvert.SerializeObject(group)}");

            Group result;
            try
            {
                result = await GraphHelper.CreateGroup(client, group, context.Token);
            }
            catch (ServiceException ex)
            {
                if (MicrosoftTeamsMAConfigSection.Configuration.DeleteAddConflictingGroup && ex.StatusCode == HttpStatusCode.BadRequest && ex.Message.IndexOf("mailNickname", 0, StringComparison.Ordinal) > 0)
                {
                    string mailNickname = csentry.GetValueAdd<string>("mailNickname");
                    logger.Warn($"{csentry.DN}: Delete/Adding conflicting group with mailNickname '{mailNickname}'");

                    string existingGroup = await GraphHelper.GetGroupIdByMailNickname(client, mailNickname, context.Token);
                    await GraphHelper.DeleteGroup(client, existingGroup, context.Token);
                    await Task.Delay(TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.PostGroupCreateDelay));
                    result = await GraphHelper.CreateGroup(client, group, context.Token);
                }
                else
                {
                    throw;
                }
            }

            GroupExportProvider.logger.Info($"{csentry.DN}: Created group {group.Id}");

            await Task.Delay(TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.PostGroupCreateDelay), context.Token);

            try
            {
                await GroupExportProvider.ProcessDeferredMembership(context, client, deferredMembers, result, deferredOwners, csentry.DN);
            }
            catch (Exception ex)
            {
                GroupExportProvider.logger.Error(ex, $"{csentry.DN}: An exception occurred while modifying the membership, rolling back the group by deleting it");
                await Task.Delay(TimeSpan.FromSeconds(5));
                await GraphHelper.DeleteGroup(client, csentry.DN, CancellationToken.None);
                GroupExportProvider.logger.Info($"{csentry.DN}: The group was deleted");
                throw;
            }

            return result;
        }

        private static async Task ProcessDeferredMembership(ExportContext context, GraphServiceClient client, IList<string> deferredMembers, Group result, IList<string> deferredOwners, string csentryDN)
        {
            bool success = false;

            while (!success)
            {
                if (deferredMembers.Count > 0)
                {
                    GroupExportProvider.logger.Trace($"{csentryDN}: Adding {deferredMembers.Count} deferred members");
                    await GraphHelper.AddGroupMembers(client, result.Id, deferredMembers, true, context.Token);
                }

                if (deferredOwners.Count > 0)
                {
                    GroupExportProvider.logger.Trace($"{csentryDN}: Adding {deferredOwners.Count} deferred owners");
                    await GraphHelper.AddGroupOwners(client, result.Id, deferredOwners, true, context.Token);
                }

                success = true;
            }
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeUpdate(CSEntryChange csentry, ExportContext context)
        {
            await this.PutCSEntryChangeUpdateGroup(csentry, context);
            await this.PutCSEntryChangeUpdateTeam(csentry, context);
            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task PutCSEntryChangeUpdateTeam(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Team team = new Team();
            team.MemberSettings = new TeamMemberSettings();
            team.MemberSettings.ODataType = null;
            team.GuestSettings = new TeamGuestSettings();
            team.GuestSettings.ODataType = null;
            team.MessagingSettings = new TeamMessagingSettings();
            team.MessagingSettings.ODataType = null;
            team.FunSettings = new TeamFunSettings();
            team.FunSettings.ODataType = null;

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

                if (change.Name == "template")
                {
                    throw new UnexpectedDataException("The template parameter can only be supplied during an 'add' operation");
                }
                else if (change.Name == "isArchived")
                {
                    team.IsArchived = change.GetValueAdd<bool>();
                }
                else if (change.Name == "memberSettings_allowCreateUpdateChannels")
                {
                    team.MemberSettings.AllowCreateUpdateChannels = change.GetValueAdd<bool>();
                }
                else if (change.Name == "memberSettings_allowDeleteChannels")
                {
                    team.MemberSettings.AllowDeleteChannels = change.GetValueAdd<bool>();
                }
                else if (change.Name == "memberSettings_allowAddRemoveApps")
                {
                    team.MemberSettings.AllowAddRemoveApps = change.GetValueAdd<bool>();
                }
                else if (change.Name == "memberSettings_allowCreateUpdateRemoveTabs")
                {
                    team.MemberSettings.AllowCreateUpdateRemoveTabs = change.GetValueAdd<bool>();
                }
                else if (change.Name == "memberSettings_allowCreateUpdateRemoveConnectors")
                {
                    team.MemberSettings.AllowCreateUpdateRemoveConnectors = change.GetValueAdd<bool>();
                }
                else if (change.Name == "guestSettings_allowCreateUpdateChannels")
                {
                    team.GuestSettings.AllowCreateUpdateChannels = change.GetValueAdd<bool>();
                }
                else if (change.Name == "guestSettings_allowDeleteChannels")
                {
                    team.GuestSettings.AllowDeleteChannels = change.GetValueAdd<bool>();
                }
                else if (change.Name == "messagingSettings_allowUserEditMessages")
                {
                    team.MessagingSettings.AllowUserEditMessages = change.GetValueAdd<bool>();
                }
                else if (change.Name == "messagingSettings_allowUserDeleteMessages")
                {
                    team.MessagingSettings.AllowUserDeleteMessages = change.GetValueAdd<bool>();
                }
                else if (change.Name == "messagingSettings_allowOwnerDeleteMessages")
                {
                    team.MessagingSettings.AllowOwnerDeleteMessages = change.GetValueAdd<bool>();
                }
                else if (change.Name == "messagingSettings_allowTeamMentions")
                {
                    team.MessagingSettings.AllowTeamMentions = change.GetValueAdd<bool>();
                }
                else if (change.Name == "messagingSettings_allowChannelMentions")
                {
                    team.MessagingSettings.AllowChannelMentions = change.GetValueAdd<bool>();
                }
                else if (change.Name == "funSettings_allowGiphy")
                {
                    team.FunSettings.AllowGiphy = change.GetValueAdd<bool>();
                }
                else if (change.Name == "funSettings_giphyContentRating")
                {
                    string value = change.GetValueAdd<string>();
                    if (!Enum.TryParse<GiphyRatingType>(value, false, out GiphyRatingType result))
                    {
                        throw new UnexpectedDataException($"The value '{value}' was not a supported value for funSettings_giphyContentRating. Supported values are (case sensitive) 'Strict' or 'Moderate'");
                    }

                    team.FunSettings.GiphyContentRating = result;
                }
                else if (change.Name == "funSettings_allowStickersAndMemes")
                {
                    team.FunSettings.AllowStickersAndMemes = change.GetValueAdd<bool>();
                }
                else if (change.Name == "funSettings_allowCustomMemes")
                {
                    team.FunSettings.AllowCustomMemes = change.GetValueAdd<bool>();
                }
                else
                {
                    continue;
                }

                changed = true;
            }

            if (changed)
            {
                GroupExportProvider.logger.Trace($"{csentry.DN}:Updating team data: {JsonConvert.SerializeObject(team)}");

                await GraphHelper.UpdateTeam(client, csentry.DN, team, context.Token);

                GroupExportProvider.logger.Info($"{csentry.DN}: Updated team");
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
                    await this.PutAttributeChangeMembers(csentry, change, context);
                    continue;
                }

                if (!SchemaProvider.GroupProperties.Contains(change.Name))
                {
                    continue;
                }

                if (change.ModificationType == AttributeModificationType.Delete)
                {
                    this.AssignNullToProperty(change.Name, group);
                    continue;
                }

                if (change.Name == "visibility")
                {
                    throw new UnexpectedDataException("The visibility parameter can only be supplied during an 'add' operation");
                }
                else if (change.Name == "displayName")
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
                else if (change.Name == "isArchived")
                {
                    group.IsArchived = change.GetValueAdd<bool>();
                }
                else
                {
                    continue;
                }

                changed = true;
            }

            if (changed)
            {
                GroupExportProvider.logger.Trace($"{csentry.DN}:Updating group data: {JsonConvert.SerializeObject(group)}");
                await GraphHelper.UpdateGroup(client, csentry.DN, group, context.Token);
                GroupExportProvider.logger.Info($"{csentry.DN}: Updated group {csentry.DN}");
            }
        }

        private async Task PutAttributeChangeMembers(CSEntryChange c, AttributeChange change, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            IList<string> valueDeletes = change.GetValueDeletes<string>();
            IList<string> valueAdds = change.GetValueAdds<string>();

            if (change.ModificationType == AttributeModificationType.Delete)
            {
                if (change.Name == "member")
                {
                    List<DirectoryObject> result = await GraphHelper.GetGroupMembers(client, c.DN, context.Token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
                else
                {
                    List<DirectoryObject> result = await GraphHelper.GetGroupOwners(client, c.DN, context.Token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
            }

            if (change.Name == "member")
            {
                await GraphHelper.AddGroupMembers(client, c.DN, valueAdds, true, context.Token);
                await GraphHelper.RemoveGroupMembers(client, c.DN, valueDeletes, true, context.Token);
                logger.Info($"Membership modification for group {c.DN} completed. Members added: {valueAdds.Count}, members removed: {valueDeletes.Count}");
            }
            else
            {
                await GraphHelper.AddGroupOwners(client, c.DN, valueAdds, true, context.Token);
                await GraphHelper.RemoveGroupOwners(client, c.DN, valueDeletes, true, context.Token);
                logger.Info($"Owner modification for group {c.DN} completed. Owners added: {valueAdds.Count}, owners removed: {valueDeletes.Count}");
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
