extern alias BetaLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Beta = BetaLib.Microsoft.Graph;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;
using NLog;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class GroupImportProvider : IObjectImportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public async Task GetCSEntryChangesAsync(ImportContext context, SchemaType type)
        {
            try
            {
                BufferBlock<Beta.Group> groupQueue = new BufferBlock<Beta.Group>(new DataflowBlockOptions { CancellationToken = context.Token });

                Task consumerTask = this.ConsumeQueue(context, type, groupQueue);

                await this.ProduceObjects(context, groupQueue);

                await consumerTask;
            }
            catch (Exception ex)
            {
                GroupImportProvider.logger.Error(ex, "There was an error importing the group data");
                throw;
            }
        }

        private async Task ProduceObjects(ImportContext context, ITargetBlock<Beta.Group> target)
        {
            var client = ((GraphConnectionContext)context.ConnectionContext).BetaClient;
            await GraphHelper.GetGroups(client, target, context.ConfigParameters[ConfigParameterNames.FilterQuery].Value, context.Token, "displayName", "resourceProvisioningOptions", "id", "mailNickname", "description", "visibility");
            target.Complete();
        }

        /// <summary>
        /// Group delta imports have a few problems that need to be resolved.
        /// 1. You can't currently filter on teams-type groups only. You have to get all groups and then filter yourself
        /// 2. This isn't so much of a problem apart from that you have to specify your attribute selection on the initial query to include members and owners. This means you get all members and owners for all groups
        /// 3. Membership information comes in chunks of 20 members, however chunks for each group can be returned in any order. This breaks the way FIM works, as we would have to hold all group objects in memory, wait to see if duplicates arrive in the stream, merge them, and only once we have all groups with all members, confidently pass them back to the sync engine in one massive batch
        /// </summary>
        /// <param name="context"></param>
        /// <param name="target"></param>
        /// <returns></returns>
        private async Task ProduceObjectsDelta(ImportContext context, ITargetBlock<Group> target)
        {
            var client = ((GraphConnectionContext)context.ConnectionContext).Client;

            string newDeltaLink;

            if (context.InDelta)
            {
                if (!context.IncomingWatermark.Contains("group"))
                {
                    throw new WarningNoWatermarkException();
                }

                Watermark watermark = context.IncomingWatermark["group"];

                if (watermark.Value == null)
                {
                    throw new WarningNoWatermarkException();
                }

                newDeltaLink = await GraphHelper.GetGroups(client, watermark.Value, target, context.Token);
            }
            else
            {
                newDeltaLink = await GraphHelper.GetGroups(client, target, context.Token, "displayName", "resourceProvisioningOptions", "id", "mailNickname", "description", "visibility", "members", "owners");
            }

            if (newDeltaLink != null)
            {
                logger.Trace($"Got delta link {newDeltaLink}");
                context.OutgoingWatermark.Add(new Watermark("group", newDeltaLink, "string"));
            }

            target.Complete();
        }

        private async Task ConsumeQueue(ImportContext context, SchemaType type, ISourceBlock<Beta.Group> source)
        {
            var edfo = new ExecutionDataflowBlockOptions
            {
                MaxDegreeOfParallelism = MicrosoftTeamsMAConfigSection.Configuration.ImportThreads,
                CancellationToken = context.Token,
            };

            ActionBlock<Beta.Group> action = new ActionBlock<Beta.Group>(async group =>
            {
                try
                {
                    //if (this.ShouldFilterDelta(group, context))
                    //{
                    //    return;
                    //}

                    CSEntryChange c = this.GroupToCSEntryChange(group, type, context);

                    if (c != null)
                    {
                        await this.GroupMemberToCSEntryChange(c, type, context);
                        await this.TeamToCSEntryChange(c, type, context).ConfigureAwait(false);
                        context.ImportItems.Add(c, context.Token);
                        await this.CreateChannelCSEntryChanges(group.Id, type, context);
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

        private bool ShouldFilterDelta(Beta.Group group, ImportContext context)
        {
            string filter = context.ConfigParameters[ConfigParameterNames.FilterQuery].Value;

            if (!string.IsNullOrWhiteSpace(filter))
            {
                if (group.MailNickname == null || !group.MailNickname.StartsWith(filter, StringComparison.OrdinalIgnoreCase))
                {
                    logger.Trace($"Filtering group {group.Id} with nickname {group.MailNickname}");
                    return true;
                }
            }

            if (group.AdditionalData == null)
            {
                logger.Trace($"Filtering non-team group {group.Id} with nickname {group.MailNickname}");
                return true;
            }

            if (!context.InDelta && group.AdditionalData.ContainsKey("@removed"))
            {
                logger.Trace($"Filtering deleted group {group.Id} with nickname {group.MailNickname}");
                return true;
            }

            if (!group.AdditionalData.ContainsKey("resourceProvisioningOptions"))
            {
                logger.Trace($"Filtering non-team group {group.Id} with nickname {group.MailNickname}");
                return true;
            }

            var rpo = group.AdditionalData["resourceProvisioningOptions"] as JArray;

            if (!rpo.Values().Any(t => string.Equals(t.Value<string>(), "team", StringComparison.OrdinalIgnoreCase)))
            {
                logger.Trace($"Filtering non-team group {group.Id} with nickname {group.MailNickname}");
                return true;
            }

            return false;
        }

        private CSEntryChange GroupToCSEntryChange(Beta.Group group, SchemaType schemaType, ImportContext context)
        {
            CSEntryChange c = CSEntryChange.Create();
            c.ObjectType = "group";
            c.AnchorAttributes.Add(AnchorAttribute.Create("id", group.Id));
            c.DN = group.Id;

            bool isRemoved = group.AdditionalData?.ContainsKey("@removed") ?? false;

            logger.Trace(JsonConvert.SerializeObject(group));

            if (context.InDelta)
            {
                if (isRemoved)
                {
                    c.ObjectModificationType = ObjectModificationType.Delete;
                    return c;
                }
                else
                {
                    c.ObjectModificationType = ObjectModificationType.Replace;
                }
            }
            else
            {
                if (isRemoved)
                {
                    return null;
                }

                c.ObjectModificationType = ObjectModificationType.Add;
            }

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

        private async Task CreateChannelCSEntryChanges(string groupid, SchemaType schemaType, ImportContext context)
        {
            if (!context.Types.Types.Contains("channel"))
            {
                return;
            }

            Beta.GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).BetaClient;

            var channels = await GraphHelper.GetChannels(client, groupid, context.Token);

            foreach (var channel in channels)
            {
                var members = await GraphHelper.GetChannelMembers(client, groupid, channel.Id, context.Token);

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

        public bool CanImport(SchemaType type)
        {
            return type.Name == "group";
        }
    }
}