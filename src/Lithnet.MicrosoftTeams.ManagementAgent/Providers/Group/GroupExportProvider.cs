using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GroupExportProvider : IObjectExportProvider
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public bool CanExport(CSEntryChange csentry)
        {
            return csentry.ObjectType == "group";
        }

        public CSEntryChangeResult PutCSEntryChange(CSEntryChange csentry, ExportContext context)
        {
            return AsyncHelper.RunSync(this.PutCSEntryChangeObject(csentry, context));
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

            await client.Groups[csentry.DN].Request().DeleteAsync();

            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeAdd(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Group result = await GroupExportProvider.CreateGroup(csentry, context, client);

            await GroupExportProvider.CreateTeam(csentry, client, result.Id);

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", result.Id));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private static async Task CreateTeam(CSEntryChange csentry, GraphServiceClient client, string groupId)
        {
            Team team = new Team
            {
                MemberSettings = new TeamMemberSettings(),
                GuestSettings = new TeamGuestSettings(),
                MessagingSettings = new TeamMessagingSettings(),
                FunSettings = new TeamFunSettings(),
                ODataType = null
            };

            team.MemberSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("memberSettings_allowCreateUpdateChannels");
            team.MemberSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("memberSettings_allowDeleteChannels");
            team.MemberSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("memberSettings_allowAddRemoveApps");
            team.MemberSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("memberSettings_allowCreateUpdateRemoveTabs");
            team.MemberSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("memberSettings_allowCreateUpdateRemoveConnectors");
            team.GuestSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("guestSettings_allowCreateUpdateChannels");
            team.GuestSettings.AllowCreateUpdateChannels = csentry.GetValueAdd<bool?>("guestSettings_allowDeleteChannels");
            team.MessagingSettings.AllowUserEditMessages = csentry.GetValueAdd<bool?>("messagingSettings_allowUserEditMessages");
            team.MessagingSettings.AllowUserDeleteMessages = csentry.GetValueAdd<bool?>("messagingSettings_allowUserDeleteMessages");
            team.MessagingSettings.AllowOwnerDeleteMessages = csentry.GetValueAdd<bool?>("messagingSettings_allowOwnerDeleteMessages");
            team.MessagingSettings.AllowTeamMentions = csentry.GetValueAdd<bool?>("messagingSettings_allowTeamMentions");
            team.MessagingSettings.AllowChannelMentions = csentry.GetValueAdd<bool?>("messagingSettings_allowChannelMentions");
            team.FunSettings.AllowGiphy = csentry.GetValueAdd<bool?>("funSettings_allowGiphy");
            team.FunSettings.AllowStickersAndMemes = csentry.GetValueAdd<bool?>("funSettings_allowStickersAndMemes");
            team.FunSettings.AllowCustomMemes = csentry.GetValueAdd<bool?>("funSettings_allowCustomMemes");

            string gcr = csentry.GetValueAdd<string>("funSettings_giphyContentRating");
            if (!string.IsNullOrWhiteSpace(gcr))
            {
                if (!Enum.TryParse(gcr, false, out GiphyRatingType grt))
                {
                    throw new UnexpectedDataException($"The value '{gcr}' was not a supported value for funSettings_giphyContentRating. Supported values are (case sensitive) 'Strict' or 'Moderate'");
                }

                team.FunSettings.GiphyContentRating = grt;
            }

            Team tresult = await client.Groups[groupId].Team
                .Request()
                .PutAsync(team);

            GroupExportProvider.logger.Info($"Created team {tresult.Id} for group {groupId}");
        }

        private static async Task<Group> CreateGroup(CSEntryChange csentry, ExportContext context, GraphServiceClient client)
        {
            Group group = new Group();

            group.DisplayName = csentry.GetValueAdd<string>("displayName") ?? throw new UnexpectedDataException("The group must have a displayName");
            group.GroupTypes = new[] {"Unified"};
            group.MailEnabled = true;
            group.Description = csentry.GetValueAdd<string>("description");
            group.MailNickname = csentry.GetValueAdd<string>("mailNickname") ?? throw new UnexpectedDataException("The group must have a mailNickname");
            group.SecurityEnabled = false;
            group.AdditionalData = new Dictionary<string, object>();
            group.Id = csentry.DN;

            IList<string> members = csentry.GetValueAdds<string>("member") ?? new List<string>();
            IList<string> owners = csentry.GetValueAdds<string>("owner") ?? new List<string>();

            if (members.Count > 0)
            {
                group.AdditionalData.Add("members@odata.bind", members.Select(t => $"https://graph.microsoft.com/v1.0/users/{t}").ToArray());
            }

            if (owners.Count > 0)
            {
                group.AdditionalData.Add("owners@odata.bind", owners.Select(t => $"https://graph.microsoft.com/v1.0/users/{t}").ToArray());
            }

            GroupExportProvider.logger.Trace($"Creating group {group.MailNickname} with {members.Count} members and {owners.Count} owners");

            Group result = await client.Groups
                .Request()
                .AddAsync(group, context.CancellationTokenSource.Token);

            GroupExportProvider.logger.Info($"Created group {group.Id}");
            return result;
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
            team.GuestSettings = new TeamGuestSettings();
            team.MessagingSettings = new TeamMessagingSettings();
            team.FunSettings = new TeamFunSettings();

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

                if (change.Name == "memberSettings_allowCreateUpdateChannels")
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
                await client.Teams[csentry.DN]
                    .Request()
                    .UpdateAsync(team);
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
                await client.Groups[csentry.DN].Request().UpdateAsync(group);
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
                    List<DirectoryObject> result = await GraphHelper.GetGroupMembers(client, c.DN);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
                else
                {
                    List<DirectoryObject> result = await GraphHelper.GetGroupOwners(client, c.DN);
                    valueDeletes = result.Select(t => t.Id).ToList();
                }
            }

            foreach (string add in valueAdds)
            {
                DirectoryObject directoryObject = new DirectoryObject
                {
                    Id = add
                };

                if (change.Name == "member")
                {
                    await client.Groups[$"{c.DN}"].Members.References
                        .Request()
                        .AddAsync(directoryObject);
                }
                else
                {
                    await client.Groups[$"{c.DN}"].Owners.References
                        .Request()
                        .AddAsync(directoryObject);
                }
            }

            foreach (string delete in valueDeletes)
            {
                if (change.Name == "member")
                {
                    await client.Groups[c.DN].Members[delete].Reference
                         .Request()
                         .DeleteAsync();
                }
                else
                {
                    await client.Groups[c.DN].Owners[delete].Reference
                        .Request()
                        .DeleteAsync();
                }
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
