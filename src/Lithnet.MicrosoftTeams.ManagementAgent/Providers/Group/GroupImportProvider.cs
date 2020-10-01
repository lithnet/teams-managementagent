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

        private IImportContext context;

        private GraphServiceClient client;

        private Beta.GraphServiceClient betaClient;

        private CancellationToken token;

        private UserFilter userFilter;

        public void Initialize(IImportContext context)
        {
            this.context = context;
            this.token = context.Token;
            this.betaClient = ((GraphConnectionContext)context.ConnectionContext).BetaClient;
            this.client = ((GraphConnectionContext)context.ConnectionContext).Client;
            this.userFilter = ((GraphConnectionContext)context.ConnectionContext).UserFilter;
        }

        public async Task GetCSEntryChangesAsync(SchemaType type)
        {
            try
            {
                BufferBlock<Beta.Group> groupQueue = new BufferBlock<Beta.Group>(new DataflowBlockOptions { CancellationToken = this.token });

                Task consumerTask = this.ConsumeQueue(type, groupQueue);

                await this.ProduceObjects(groupQueue);

                await consumerTask;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "There was an error importing the group data");
                throw;
            }
        }

        private async Task ProduceObjects(ITargetBlock<Beta.Group> target)
        {
            await GraphHelperGroups.GetGroups(this.betaClient, target, this.context.ConfigParameters[ConfigParameterNames.FilterQuery].Value,this.token, "displayName", "resourceProvisioningOptions", "id", "mailNickname", "description", "visibility");
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
        private async Task ProduceObjectsDelta(ITargetBlock<Group> target)
        {
            string newDeltaLink;

            if (this.context.InDelta)
            {
                if (!this.context.IncomingWatermark.Contains("group"))
                {
                    throw new WarningNoWatermarkException();
                }

                Watermark watermark = this.context.IncomingWatermark["group"];

                if (watermark.Value == null)
                {
                    throw new WarningNoWatermarkException();
                }

                newDeltaLink = await GraphHelperGroups.GetGroups(this.client, watermark.Value, target,this.token);
            }
            else
            {
                newDeltaLink = await GraphHelperGroups.GetGroups(this.client, target,this.token, "displayName", "resourceProvisioningOptions", "id", "mailNickname", "description", "visibility", "members", "owners");
            }

            if (newDeltaLink != null)
            {
                logger.Trace($"Got delta link {newDeltaLink}");
                this.context.OutgoingWatermark.Add(new Watermark("group", newDeltaLink, "string"));
            }

            target.Complete();
        }

        private async Task ConsumeQueue(SchemaType type, ISourceBlock<Beta.Group> source)
        {
            var edfo = new ExecutionDataflowBlockOptions
            {
                MaxDegreeOfParallelism = MicrosoftTeamsMAConfigSection.Configuration.ImportThreads,
                CancellationToken = this.token,
            };

            ActionBlock<Beta.Group> action = new ActionBlock<Beta.Group>(async group =>
            {
                try
                {
                    //if (this.ShouldFilterDelta(group, context))
                    //{
                    //    return;
                    //}

                    CSEntryChange c = this.GroupToCSEntryChange(group, type);

                    if (c != null)
                    {
                        await this.GroupMemberToCSEntryChange(c, type);
                        await this.TeamToCSEntryChange(c, type).ConfigureAwait(false);
                        this.context.ImportItems.Add(c,this.token);
                        await this.CreateChannelCSEntryChanges(group.Id, type);
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
                    this.context.ImportItems.Add(csentry,this.token);
                }
            }, edfo);

            source.LinkTo(action, new DataflowLinkOptions() { PropagateCompletion = true });

            await action.Completion;
        }

        private bool ShouldFilterDelta(Beta.Group group)
        {
            string filter = this.context.ConfigParameters[ConfigParameterNames.FilterQuery].Value;

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

            if (!this.context.InDelta && group.AdditionalData.ContainsKey("@removed"))
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

        private CSEntryChange GroupToCSEntryChange(Beta.Group group, SchemaType schemaType)
        {
            CSEntryChange c = CSEntryChange.Create();
            c.ObjectType = "group";
            c.AnchorAttributes.Add(AnchorAttribute.Create("id", group.Id));
            c.DN = group.Id;

            bool isRemoved = group.AdditionalData?.ContainsKey("@removed") ?? false;

            logger.Trace(JsonConvert.SerializeObject(group));

            if (this.context.InDelta)
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

        private async Task GroupMemberToCSEntryChange(CSEntryChange c, SchemaType schemaType)
        {
            if (schemaType.Attributes.Contains("member"))
            {
                List<DirectoryObject> members = await GraphHelperGroups.GetGroupMembers(this.client, c.DN, this.token);
                List<object> memberIds = members.Where(u => !this.userFilter.ShouldExclude(u.Id, this.token)).Select(t => t.Id).ToList<object>();

                if (memberIds.Count > 0)
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("member", memberIds));
                }
            }

            if (schemaType.Attributes.Contains("owner"))
            {
                List<DirectoryObject> owners = await GraphHelperGroups.GetGroupOwners(this.client, c.DN, this.token);
                List<object> ownerIds = owners.Where(u => !this.userFilter.ShouldExclude(u.Id, this.token)).Select(t => t.Id).ToList<object>();

                if (ownerIds.Count > 0)
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("owner", ownerIds));
                }
            }
        }

        private async Task CreateChannelCSEntryChanges(string groupid, SchemaType schemaType)
        {
            if (!this.context.Types.Types.Contains("channel"))
            {
                return;
            }

            var channels = await GraphHelperTeams.GetChannels(this.betaClient, groupid,this.token);

            foreach (var channel in channels)
            {
                var members = await GraphHelperTeams.GetChannelMembers(this.betaClient, groupid, channel.Id,this.token);

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

                this.context.ImportItems.Add(c,this.token);
            }
        }

        private async Task TeamToCSEntryChange(CSEntryChange c, SchemaType schemaType)
        {
            Beta.Team team = await GraphHelperTeams.GetTeam(this.betaClient, c.DN,this.token);

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