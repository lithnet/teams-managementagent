using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Lithnet.Ecma2Framework;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GroupImportProvider : IObjectImportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public async Task GetCSEntryChangesAsync(ImportContext context, SchemaType type)
        {
            try
            {
                IGraphServiceGroupsCollectionRequest groups = this.GetGroupEnumerationRequest(context.InDelta, context.IncomingWatermark, ((GraphConnectionContext)context.ConnectionContext).Client, context);

                BufferBlock<Group> groupQueue = new BufferBlock<Group>(new DataflowBlockOptions() { CancellationToken = context.Token });

                Task consumerTask = this.ConsumeQueue(context, type, groupQueue);

                await this.ProduceObjects(groups, groupQueue, context.Token);

                await consumerTask;
            }
            catch (Exception ex)
            {
                GroupImportProvider.logger.Error(ex, "There was an error importing the group data");
                throw;
            }
        }

        private async Task ProduceObjects(IGraphServiceGroupsCollectionRequest request, ITargetBlock<Group> target, CancellationToken cancellationToken)
        {
            await GraphHelper.GetGroups(request, target, cancellationToken);
            target.Complete();
        }

        private async Task ConsumeQueue(ImportContext context, SchemaType type, ISourceBlock<Group> source)
        {
            var edfo = new ExecutionDataflowBlockOptions
            {
                MaxDegreeOfParallelism = MicrosoftTeamsMAConfigSection.Configuration.ImportThreads,
                CancellationToken = context.Token,
            };

            ActionBlock<Group> action = new ActionBlock<Group>(async group =>
            {
                try
                {
                    CSEntryChange c = this.GroupToCSEntryChange(group, type, context);

                    if (c != null)
                    {
                        await this.GroupMemberToCSEntryChange(c, type, context);
                        await this.TeamToCSEntryChange(c, type, context).ConfigureAwait(false);
                        context.ImportItems.Add(c, context.Token);
                        await this.CreateChannelCSEntryChanges(group, type, context);
                    }
                }
                catch (Exception ex)
                {
                    GroupImportProvider.logger.Error(ex);
                    CSEntryChange csentry = CSEntryChange.Create();
                    csentry.DN = group.Id;
                    csentry.ErrorCodeImport = MAImportError.ImportErrorCustomContinueRun;
                    csentry.ErrorDetail = ex.StackTrace;
                    csentry.ErrorName = ex.Message;
                    context.ImportItems.Add(csentry, context.Token);
                }
            }, edfo);

            source.LinkTo(action, new DataflowLinkOptions() { PropagateCompletion = true });

            await action.Completion;
        }

        private CSEntryChange GroupToCSEntryChange(Group group, SchemaType schemaType, ImportContext context)
        {
            GroupImportProvider.logger.Trace($"Creating CSEntryChange for {group.Id}/{group.DisplayName}");

            CSEntryChange c = CSEntryChange.Create();
            c.ObjectType = "group";
            c.ObjectModificationType = ObjectModificationType.Add;
            c.AnchorAttributes.Add(AnchorAttribute.Create("id", group.Id));
            c.DN = group.Id;

            foreach (SchemaAttribute type in schemaType.Attributes)
            {
                switch (type.Name)
                {
                    case "displayName":
                        if (!string.IsNullOrWhiteSpace(group.DisplayName))
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.DisplayName));
                        }
                        break;

                    case "description":
                        if (!string.IsNullOrWhiteSpace(group.Description))
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.Description));
                        }
                        break;

                    case "id":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.Id));
                        break;

                    case "mailNickname":
                        if (!string.IsNullOrWhiteSpace(group.MailNickname))
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.MailNickname));
                        }

                        break;

                    case "visibility":
                        if (!string.IsNullOrWhiteSpace(group.Visibility))
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.Visibility));
                        }

                        break;
                }
            }

            return c;
        }

        private async Task GroupMemberToCSEntryChange(CSEntryChange c, SchemaType schemaType, ImportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            if (schemaType.Attributes.Contains("member"))
            {
                List<DirectoryObject> members = await GraphHelper.GetGroupMembers(client, c.DN, context.Token);
                if (members.Count > 0)
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("member", members.Select(t => t.Id).ToList<object>()));
                }
            }

            if (schemaType.Attributes.Contains("owner"))
            {
                List<DirectoryObject> owners = await GraphHelper.GetGroupOwners(client, c.DN, context.Token);

                if (owners.Count > 0)
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("owner", owners.Select(t => t.Id).ToList<object>()));
                }
            }
        }

        private async Task CreateChannelCSEntryChanges(Group g, SchemaType schemaType, ImportContext context)
        {
            if (!context.Types.Types.Contains("channel"))
            {
                return;
            }

            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            var channels = await GraphHelper.GetChannels(client, g.Id, context.Token);

            foreach (var channel in channels)
            {
                var members = await GraphHelper.GetChannelMembers(client, g.Id, channel.Id, context.Token);

                CSEntryChange c = CSEntryChange.Create();
                c.ObjectType = "channel";
                c.ObjectModificationType = ObjectModificationType.Add;
                c.AnchorAttributes.Add(AnchorAttribute.Create("id", channel.Id));
                c.DN = channel.Id;

                c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("displayName", channel.DisplayName));

                if (!string.IsNullOrWhiteSpace(channel.Description))
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("description", channel.Description));
                }

                if (members.Count > 0)
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("member", members.Select(t => t.Id).ToList<object>()));
                }

                context.ImportItems.Add(c, context.Token);
            }
        }

        private async Task TeamToCSEntryChange(CSEntryChange c, SchemaType schemaType, ImportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Team team = await GraphHelper.GetTeam(client, c.DN, context.Token);

            foreach (SchemaAttribute type in schemaType.Attributes)
            {
                switch (type.Name)
                {
                    case "isArchived":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.IsArchived ?? false));
                        break;

                    case "memberSettings_allowCreateUpdateChannels":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MemberSettings.AllowCreateUpdateChannels ?? false));
                        break;

                    case "memberSettings_allowDeleteChannels":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MemberSettings.AllowDeleteChannels ?? false));
                        break;

                    case "memberSettings_allowAddRemoveApps":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MemberSettings.AllowAddRemoveApps ?? false));
                        break;

                    case "memberSettings_allowCreateUpdateRemoveTabs":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MemberSettings.AllowCreateUpdateRemoveTabs ?? false));
                        break;

                    case "memberSettings_allowCreateUpdateRemoveConnectors":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MemberSettings.AllowCreateUpdateRemoveConnectors ?? false));
                        break;

                    case "guestSettings_allowCreateUpdateChannels":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.GuestSettings.AllowCreateUpdateChannels ?? false));
                        break;

                    case "guestSettings_allowDeleteChannels":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.GuestSettings.AllowDeleteChannels ?? false));
                        break;

                    case "messagingSettings_allowUserEditMessages":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowUserEditMessages ?? false));
                        break;

                    case "messagingSettings_allowUserDeleteMessages":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowUserDeleteMessages ?? false));
                        break;

                    case "messagingSettings_allowOwnerDeleteMessages":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowOwnerDeleteMessages ?? false));
                        break;

                    case "messagingSettings_allowTeamMentions":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowTeamMentions ?? false));
                        break;

                    case "messagingSettings_allowChannelMentions":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowChannelMentions ?? false));
                        break;

                    case "funSettings_allowGiphy":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.FunSettings.AllowGiphy ?? false));
                        break;

                    case "funSettings_giphyContentRating":
                        if (team.FunSettings.GiphyContentRating != null)
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.FunSettings.GiphyContentRating.ToString()));
                        }

                        break;

                    case "funSettings_allowStickersAndMemes":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.FunSettings.AllowStickersAndMemes ?? false));
                        break;

                    case "funSettings_allowCustomMemes":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.FunSettings.AllowCustomMemes ?? false));
                        break;
                }
            }
        }

        private IGraphServiceGroupsCollectionRequest GetGroupEnumerationRequest(bool inDelta, WatermarkKeyedCollection importState, GraphServiceClient client, ImportContext context)
        {
            string filter = "resourceProvisioningOptions/Any(x:x eq 'Team')";

            if (!string.IsNullOrWhiteSpace(context.ConfigParameters[ConfigParameterNames.FilterQuery].Value))
            {
                filter += $" and {context.ConfigParameters[ConfigParameterNames.FilterQuery].Value}";
            }

            GroupImportProvider.logger.Trace($"Enumerating groups with filter {filter}");

            return client.Groups.Request()
                .Select(e => new
                {
                    e.DisplayName,
                    e.Id,
                    e.ResourceProvisioningOptions,
                    e.MailNickname,
                    e.Description,
                    e.Visibility,
                })
                .Filter(filter);
        }

        public bool CanImport(SchemaType type)
        {
            return type.Name == "group";
        }
    }
}