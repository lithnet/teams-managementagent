using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GroupImportProvider : IObjectImportProvider
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public void GetCSEntryChanges(ImportContext context, SchemaType type)
        {
            AsyncHelper.RunSync(this.GetCSEntryChangesAsync(context, type), context.CancellationTokenSource.Token);
        }

        public async Task GetCSEntryChangesAsync(ImportContext context, SchemaType type)
        {
            try
            {
                IGraphServiceGroupsCollectionPage groups = await this.GetGroupEnumerable(context.InDelta, context.IncomingWatermark, ((GraphConnectionContext)context.ConnectionContext).Client, context);
                BufferBlock<Group> queue = new BufferBlock<Group>();

                Task consumer = this.ConsumeObjects(context, type, queue);

                // Post source data to the dataflow block.
                await this.ProduceObjects(groups, queue).ConfigureAwait(false);

                // Wait for the consumer to process all data.
                await consumer.ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "There was an error importing the group data");
                throw;
            }
        }

        private async Task ProduceObjects(IGraphServiceGroupsCollectionPage page, ITargetBlock<Group> target)
        {
            foreach (Group group in page.CurrentPage)
            {
                target.Post(group);
            }

            while (page.NextPageRequest != null)
            {
                page = await page.NextPageRequest.GetAsync();

                foreach (Group group in page.CurrentPage)
                {
                    target.Post(group);
                }
            }

            target.Complete();
        }

        private async Task ConsumeObjects(ImportContext context, SchemaType type, ISourceBlock<Group> source)
        {
            while (await source.OutputAvailableAsync())
            {
                Group group = source.Receive();

                try
                {
                    CSEntryChange c = await this.GroupToCSEntryChange(context.InDelta, type, group, context).ConfigureAwait(false);

                    if (c != null)
                    {
                        await this.TeamToCSEntryChange(context.InDelta, type, group, context, c).ConfigureAwait(false);
                        context.ImportItems.Add(c, context.CancellationTokenSource.Token);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    CSEntryChange csentry = CSEntryChange.Create();
                    csentry.DN = group.Id;
                    csentry.ErrorCodeImport = MAImportError.ImportErrorCustomContinueRun;
                    csentry.ErrorDetail = ex.StackTrace;
                    csentry.ErrorName = ex.Message;
                    context.ImportItems.Add(csentry, context.CancellationTokenSource.Token);
                }

                context.CancellationTokenSource.Token.ThrowIfCancellationRequested();
            }
        }

        private async Task<CSEntryChange> GroupToCSEntryChange(bool inDelta, SchemaType schemaType, Group group, ImportContext context)
        {
            logger.Trace($"Creating CSEntryChange for {group.Id}/{group.DisplayName}");

            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

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

                    case "id":
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, group.Id));
                        break;

                    case "member":
                        List<DirectoryObject> members = await this.GetMembers(client.Groups[group.Id].Members);
                        if (members.Count > 0)
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, members.Select(t => t.Id).ToList<object>()));
                        }

                        break;

                    case "owner":
                        List<DirectoryObject> owners = await this.GetMembers(client.Groups[group.Id].Owners);

                        if (owners.Count > 0)
                        {
                            c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, owners.Select(t => t.Id).ToList<object>()));
                        }

                        break;

                    default:
                        break;
                      //  throw new NoSuchAttributeInObjectTypeException($"The attribute {type.Name} was not found");
                }
            }

            return c;
        }

        private async Task<IGraphServiceGroupsCollectionPage> GetGroupEnumerable(bool inDelta, WatermarkKeyedCollection importState, GraphServiceClient client, ImportContext context)
        {
            return await client.Groups.Request()
                .Select(e => new
                {
                    e.DisplayName,
                    e.Id,
                    e.ResourceProvisioningOptions,
                })
                .Filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
                .GetAsync(context.CancellationTokenSource.Token);
        }

        private async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupMembersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await this.GetMembers(page);
        }

        private async Task<List<DirectoryObject>> GetMembers(IGroupOwnersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupOwnersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await this.GetMembers(page);
        }

        private async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesPage page)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync();

                    members.AddRange(page.CurrentPage);
                }
            }

            return members;
        }

        private async Task<List<DirectoryObject>> GetMembers(IGroupOwnersCollectionWithReferencesPage page)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync();

                    members.AddRange(page.CurrentPage);
                }
            }

            return members;
        }

        private async Task TeamToCSEntryChange(bool inDelta, SchemaType schemaType, Group group, ImportContext context, CSEntryChange c)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            Team team = await client.Teams[group.Id].Request().GetAsync();

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
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, team.MessagingSettings.AllowUserEditMessages?? false));
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

        public bool CanImport(SchemaType type)
        {
            return type.Name == "group";
        }
    }
}