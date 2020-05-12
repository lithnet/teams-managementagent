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
    internal class TeamExportProvider : IObjectExportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private IExportContext context;

        private GraphServiceClient client;

        private Beta.GraphServiceClient betaClient;

        private CancellationToken token;

        public void Initialize(IExportContext context)
        {
            this.context = context;
            this.token = context.Token;
            this.betaClient = ((GraphConnectionContext) context.ConnectionContext).BetaClient;
            this.client = ((GraphConnectionContext) context.ConnectionContext).Client;
        }

        public bool CanExport(CSEntryChange csentry)
        {
            return csentry.ObjectType == "team";
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
            try
            {
                await GraphHelperGroups.DeleteGroup(this.client, csentry.DN, this.token);
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

        private async Task<CSEntryChangeResult> PutCSEntryChangeAdd(CSEntryChange csentry)
        {
            string teamid = null;

            try
            {
                IList<string> owners = csentry.GetValueAdds<string>("owner") ?? new List<string>();
                teamid = await this.CreateTeam(csentry, this.betaClient, owners.First());
                await this.PutGroupMembersMailNickname(csentry, teamid);
            }
            catch
            {
                try
                {
                    if (teamid != null)
                    {
                        logger.Error($"{csentry.DN}: An exception occurred while creating the team, rolling back by deleting it");
                        await Task.Delay(TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.PostGroupCreateDelay));
                        await GraphHelperGroups.DeleteGroup(this.client, teamid, CancellationToken.None);
                        logger.Info($"{csentry.DN}: The group was deleted");
                    }
                }
                catch (Exception ex2)
                {
                    logger.Error(ex2, $"{csentry.DN}: An exception occurred while rolling back the team");
                }

                throw;
            }

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", teamid));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task<string> CreateTeam(CSEntryChange csentry, Beta.GraphServiceClient client, string ownerId)
        {
            var team = new Beta.Team
            {
                MemberSettings = new Beta.TeamMemberSettings(),
                GuestSettings = new Beta.TeamGuestSettings(),
                MessagingSettings = new Beta.TeamMessagingSettings(),
                FunSettings = new Beta.TeamFunSettings(),
                ODataType = null
            };

            team.MemberSettings.ODataType = null;
            team.GuestSettings.ODataType = null;
            team.MessagingSettings.ODataType = null;
            team.FunSettings.ODataType = null;
            team.AdditionalData = new Dictionary<string, object>();
            team.DisplayName = csentry.GetValueAdd<string>("displayName");
            team.Description = csentry.GetValueAdd<string>("description");

            if (csentry.HasAttributeChange("visibility"))
            {
                string visibility = csentry.GetValueAdd<string>("visibility");

                if (Enum.TryParse(visibility, out Beta.TeamVisibilityType result))
                {
                    team.Visibility = result;
                }
                else
                {
                    throw new UnexpectedDataException($"The 'visibility' value was not supported {visibility}");
                }
            }

            string template = csentry.GetValueAdd<string>("template") ?? "https://graph.microsoft.com/beta/teamsTemplates('standard')";

            if (!string.IsNullOrWhiteSpace(template))
            {
                team.AdditionalData.Add("template@odata.bind", template); //"https://graph.microsoft.com/beta/teamsTemplates('standard')"
            }

            team.AdditionalData.Add("owners@odata.bind", new string[]
                {$"https://graph.microsoft.com/v1.0/users('{ownerId}')"});

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
                if (!Enum.TryParse(gcr, false, out Beta.GiphyRatingType grt))
                {
                    throw new UnexpectedDataException($"The value '{gcr}' was not a supported value for funSettings_giphyContentRating. Supported values are (case sensitive) 'Strict' or 'Moderate'");
                }

                team.FunSettings.GiphyContentRating = grt;
            }

            logger.Info($"{csentry.DN}: Creating team using template {template ?? "standard"}");
            logger.Trace($"{csentry.DN}: Team data: {JsonConvert.SerializeObject(team)}");

            var tresult = await GraphHelperTeams.CreateTeam(client, team, this.token);

            logger.Info($"{csentry.DN}: Created team {tresult}");

            return tresult;
        }

        private async Task PutGroupMembersMailNickname(CSEntryChange csentry, string teamID)
        {
            if (csentry.HasAttributeChange("mailNickname"))
            {
                Group group = new Group();
                group.MailNickname = csentry.GetValueAdd<string>("mailNickname");

                try
                {
                    await GraphHelperGroups.UpdateGroup(this.client, teamID, group, this.token);
                    logger.Info($"{csentry.DN}: Updated group {group.Id}");
                }
                catch (ServiceException ex)
                {
                    if (MicrosoftTeamsMAConfigSection.Configuration.DeleteAddConflictingGroup && ex.StatusCode == HttpStatusCode.BadRequest && ex.Message.IndexOf("mailNickname", 0, StringComparison.Ordinal) > 0)
                    {
                        string mailNickname = csentry.GetValueAdd<string>("mailNickname");
                        logger.Warn($"{csentry.DN}: Deleting group with conflicting mailNickname '{mailNickname}'");
                        string existingGroup = await GraphHelperGroups.GetGroupIdByMailNickname(this.client, mailNickname, this.token);
                        await GraphHelperGroups.DeleteGroup(this.client, existingGroup, this.token);
                        await Task.Delay(TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.PostGroupCreateDelay));

                        await GraphHelperGroups.UpdateGroup(this.client, teamID, group, this.token);
                        logger.Info($"{csentry.DN}: Updated group {group.Id}");
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            IList<string> members = csentry.GetValueAdds<string>("member") ?? new List<string>();
            IList<string> owners = csentry.GetValueAdds<string>("owner") ?? new List<string>();

            owners.RemoveAt(0);

            IList<string> deferredMembers = new List<string>();
            IList<string> deferredOwners = new List<string>();

            if (owners.Count > 100)
            {
                throw new UnexpectedDataException($"The group creation request {csentry.DN} contained more than 100 owners");
            }

            foreach (string owner in owners)
            {
                deferredOwners.Add(owner);
            }

            foreach (string member in members)
            {
                deferredMembers.Add(member);
            }

            await this.ProcessDeferredMembership(deferredMembers, teamID, deferredOwners, csentry.DN);
        }

        private async Task ProcessDeferredMembership(IList<string> deferredMembers, string groupid, IList<string> deferredOwners, string csentryDN)
        {
            bool success = false;

            while (!success)
            {
                if (deferredMembers.Count > 0)
                {
                    logger.Trace($"{csentryDN}: Adding {deferredMembers.Count} deferred members");
                    await GraphHelperGroups.AddGroupMembers(this.client, groupid, deferredMembers, true, this.token);
                }

                if (deferredOwners.Count > 0)
                {
                    logger.Trace($"{csentryDN}: Adding {deferredOwners.Count} deferred owners");
                    await GraphHelperGroups.AddGroupOwners(this.client, groupid, deferredOwners, true, this.token);
                }

                success = true;
            }
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeUpdate(CSEntryChange csentry)
        {
            await this.PutCSEntryChangeUpdateTeam(csentry);
            await this.PutCSEntryChangeUpdateGroup(csentry);
            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task PutCSEntryChangeUpdateTeam(CSEntryChange csentry)
        {
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
                logger.Trace($"{csentry.DN}:Updating team data: {JsonConvert.SerializeObject(team)}");

                await GraphHelperTeams.UpdateTeam(this.client, csentry.DN, team, this.token);

                logger.Info($"{csentry.DN}: Updated team");
            }
        }

        private async Task PutCSEntryChangeUpdateGroup(CSEntryChange csentry)
        {
            Group group = new Group();
            bool changed = false;

            foreach (AttributeChange change in csentry.AttributeChanges)
            {
                if (SchemaProvider.GroupMemberProperties.Contains(change.Name))
                {
                    await this.PutAttributeChangeMembers(csentry.DN, change);
                    continue;
                }

                if (!SchemaProvider.GroupFromTeamProperties.Contains(change.Name))
                {
                    continue;
                }

                if (change.ModificationType == AttributeModificationType.Delete)
                {
                    group.AssignNullToProperty(change.Name);
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
                await GraphHelperGroups.UpdateGroup(this.client, csentry.DN, group, this.token);
                logger.Info($"{csentry.DN}: Updated group {csentry.DN}");
            }
        }

        private async Task PutAttributeChangeMembers(string groupid, AttributeChange change)
        {
            IList<string> valueDeletes = change.GetValueDeletes<string>();
            IList<string> valueAdds = change.GetValueAdds<string>();

            if (change.ModificationType == AttributeModificationType.Delete)
            {
                if (change.Name == "member")
                {
                    List<DirectoryObject> result = await GraphHelperGroups.GetGroupMembers(this.client, groupid, this.token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
                else
                {
                    List<DirectoryObject> result = await GraphHelperGroups.GetGroupOwners(this.client, groupid, this.token);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
            }

            if (change.Name == "member")
            {
                await GraphHelperGroups.AddGroupMembers(this.client, groupid, valueAdds, true, this.token);
                await GraphHelperGroups.RemoveGroupMembers(this.client, groupid, valueDeletes, true, this.token);
                logger.Info($"Membership modification for group {groupid} completed. Members added: {valueAdds.Count}, members removed: {valueDeletes.Count}");
            }
            else
            {
                await GraphHelperGroups.AddGroupOwners(this.client, groupid, valueAdds, true, this.token);
                await GraphHelperGroups.RemoveGroupOwners(this.client, groupid, valueDeletes, true, this.token);
                logger.Info($"Owner modification for group {groupid} completed. Owners added: {valueAdds.Count}, owners removed: {valueDeletes.Count}");
            }
        }
    }
}
