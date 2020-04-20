using System;
using System.Collections.Generic;
using System.Linq;
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

        private const int MaxJsonBatchRequests = 20;

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

            await client.Groups[csentry.DN].Request().DeleteAsync();

            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeAdd(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Group result = null;

            try
            {
                result = await this.CreateGroup(csentry, context, client);
                await this.CreateTeam(csentry, client, result.Id);
            }
            catch (Exception ex)
            {
                try
                {
                    GroupExportProvider.logger.Error(ex, "An exception occurred while creating the team, rolling back the group by deleting it");
                    await client.Groups[result?.Id ?? csentry.DN].Request().DeleteAsync();
                    GroupExportProvider.logger.Info("The group was deleted");
                }
                catch (Exception ex2)
                {
                    GroupExportProvider.logger.Error(ex2, "An exception occurred while rolling back the team");
                }

                throw;
            }

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", result.Id));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task CreateTeam(CSEntryChange csentry, GraphServiceClient client, string groupId)
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


            string template = csentry.GetValueAdd<string>("template") ?? "https://graph.microsoft.com/beta/teamsTemplates('standard')";

            if (string.IsNullOrWhiteSpace(template))
            {
                team.AdditionalData.Add("template@odata.bind", template); //"https://graph.microsoft.com/beta/teamsTemplates('standard')"
            }

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

            GroupExportProvider.logger.Info($"Creating team for group {groupId} using template {template ?? "standard"}");
            GroupExportProvider.logger.Trace($"Team data: {JsonConvert.SerializeObject(team)}");

            Team tresult = await client.Groups[groupId].Team
                .Request()
                .PutAsync(team);

            GroupExportProvider.logger.Info($"Created team {tresult.Id} for group {groupId}");
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

            IList<string> members = csentry.GetValueAdds<string>("member") ?? new List<string>();
            IList<string> owners = csentry.GetValueAdds<string>("owner") ?? new List<string>();

            IList<string> deferredMembers = new List<string>();
            IList<string> createOpMembers = new List<string>();

            IList<string> deferredOwners = new List<string>();
            IList<string> createOpOwners = new List<string>();

            int memberCount = 0;

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

            GroupExportProvider.logger.Trace($"Creating group {group.MailNickname} with {createOpMembers.Count} members and {createOpOwners.Count} owners (Deferring {deferredMembers.Count} members and {deferredOwners.Count} owners until after group creation)");

            GroupExportProvider.logger.Trace($"Group data: {JsonConvert.SerializeObject(group)}");

            Group result = await client.Groups
                .Request()
                .AddAsync(group, context.CancellationTokenSource.Token);

            if (deferredMembers.Count > 0)
            {
                GroupExportProvider.logger.Trace($"Processing {deferredMembers.Count} deferred members");
                await this.AddMembers(client, result.Id, deferredMembers, context.CancellationTokenSource.Token);
            }

            if (deferredOwners.Count > 0)
            {
                GroupExportProvider.logger.Trace($"Processing {deferredOwners.Count} deferred owners");
                await this.AddOwners(client, result.Id, deferredOwners, context.CancellationTokenSource.Token);
            }

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
                GroupExportProvider.logger.Trace($"Updating team data: {JsonConvert.SerializeObject(team)}");

                await client.Teams[csentry.DN]
                    .Request()
                    .UpdateAsync(team);

                GroupExportProvider.logger.Info($"Updated team {csentry.DN}");
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
                GroupExportProvider.logger.Trace($"Updating group data: {JsonConvert.SerializeObject(group)}");
                await client.Groups[csentry.DN].Request().UpdateAsync(group);
                GroupExportProvider.logger.Info($"Updated group {csentry.DN}");
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

            if (change.Name == "member")
            {
                await this.AddMembers(client, c.DN, valueAdds, context.CancellationTokenSource.Token);
                await this.RemoveMembers(client, c.DN, valueDeletes, context.CancellationTokenSource.Token);
                GroupExportProvider.logger.Info($"Membership modification for group {c.DN} completed. Members added: {valueAdds.Count}, members removed: {valueDeletes.Count}");
            }
            else
            {
                await this.AddOwners(client, c.DN, valueAdds, context.CancellationTokenSource.Token);
                await this.RemoveOwners(client, c.DN, valueDeletes, context.CancellationTokenSource.Token);
                GroupExportProvider.logger.Info($"Owner modification for group {c.DN} completed. Owners added: {valueAdds.Count}, owners removed: {valueDeletes.Count}");
            }
        }

        private async Task AddMembers(GraphServiceClient client, string groupid, IList<string> members, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                HttpRequestMessage createEventMessage = new HttpRequestMessage(HttpMethod.Post, client.Groups[groupid].Members.References.Request().RequestUrl);
                createEventMessage.Content = GroupExportProvider.CreateStringContentForMemberId(member);

                requests.Add(new BatchRequestStep(member, createEventMessage));
            }

            GroupExportProvider.logger.Trace($"Adding {requests.Count} members in batch request for group {groupid}");

            await this.SubmitAsBatches(client, requests, ValueModificationType.Add, token);
        }

        private async Task AddOwners(GraphServiceClient client, string groupid, IList<string> members, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                HttpRequestMessage createEventMessage = new HttpRequestMessage(HttpMethod.Post, client.Groups[groupid].Owners.References.Request().RequestUrl);
                createEventMessage.Content = GroupExportProvider.CreateStringContentForMemberId(member);

                requests.Add(new BatchRequestStep(member, createEventMessage));
            }

            GroupExportProvider.logger.Trace($"Adding {requests.Count} owners in batch request for group {groupid}");
            await this.SubmitAsBatches(client, requests, ValueModificationType.Add, token);
        }

        private async Task RemoveMembers(GraphServiceClient client, string groupid, IList<string> members, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                requests.Add(this.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Members[member].Reference.Request().RequestUrl));
            }

            GroupExportProvider.logger.Trace($"Removing {requests.Count} members in batch request for group {groupid}");
            await this.SubmitAsBatches(client, requests, ValueModificationType.Delete, token);
        }

        private async Task RemoveOwners(GraphServiceClient client, string groupid, IList<string> members, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                requests.Add(this.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Owners[member].Reference.Request().RequestUrl));
            }

            GroupExportProvider.logger.Trace($"Removing {requests.Count} owners in batch request for group {groupid}");
            await this.SubmitAsBatches(client, requests, ValueModificationType.Delete, token);
        }

        private BatchRequestStep GenerateBatchRequestStep(HttpMethod method, string id, string requestUrl)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
            return new BatchRequestStep(id, request);
        }

        private async Task SubmitAsBatches(GraphServiceClient client, List<BatchRequestStep> requests, ValueModificationType mode, CancellationToken token)
        {
            BatchRequestContent content = new BatchRequestContent();
            int count = 0;

            foreach (BatchRequestStep r in requests)
            {
                if (count == GroupExportProvider.MaxJsonBatchRequests)
                {
                    await this.SubmitBatchContent(client, content, mode, token);
                    count = 0;
                    content = new BatchRequestContent();
                }

                content.AddBatchRequestStep(r);
                count++;
            }

            if (count > 0)
            {
                await this.SubmitBatchContent(client, content, mode, token);
            }
        }

        private async Task SubmitBatchContent(GraphServiceClient client, BatchRequestContent content, ValueModificationType mode, CancellationToken token)
        {
            BatchResponseContent response = await client.Batch.Request().PostAsync(content, token);

            List<Exception> exceptions = new List<Exception>();

            foreach (KeyValuePair<string, HttpResponseMessage> r in await response.GetResponsesAsync())
            {
                if (!r.Value.IsSuccessStatusCode)
                {
                    if (mode == ValueModificationType.Delete && r.Value.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        logger.Warn($"The request ({r.Key}) to remove object failed because it did not exist");
                        continue;
                    }

                    ErrorResponse er;
                    try
                    {
                        string econtent = await r.Value.Content.ReadAsStringAsync();
                        logger.Trace(econtent);

                        er = JsonConvert.DeserializeObject<ErrorResponse>(econtent);
                    }
                    catch (Exception ex)
                    {
                        logger.Trace(ex, "The error response could not be deserialized");
                        er = new ErrorResponse
                        {
                            Error = new Error
                            {
                                Code = r.Value.StatusCode.ToString(),
                                Message = r.Value.ReasonPhrase
                            }
                        };
                    }

                    if (mode == ValueModificationType.Add && r.Value.StatusCode == System.Net.HttpStatusCode.BadRequest && er.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                        continue;
                    }
                    
                    exceptions.Add(new ServiceException(er.Error, r.Value.Headers, r.Value.StatusCode));

                }
            }

            if (exceptions.Count == 1)
            {
                throw exceptions[0];
            }
            if (exceptions.Count > 1)
            {
                throw new AggregateException("Multiple member operations failed", exceptions);
            }
        }

        private static StringContent CreateStringContentForMemberId(string member)
        {
            return new StringContent("{\"@odata.id\":\"https://graph.microsoft.com/beta/users/" + member + "\"}", System.Text.Encoding.UTF8, "application/json");
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
