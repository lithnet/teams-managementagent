using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;
using NLog;
using System.Threading.Tasks.Dataflow;
using Microsoft.Graph;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class UserImportProvider : IObjectImportProviderAsync
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private HashSet<string> usersToIgnore = new HashSet<string>();

        private IImportContext context;

        private GraphServiceClient client;

        private CancellationToken token;

        public void Initialize(IImportContext context)
        {
            this.context = context;
            this.token = context.Token;
            this.client = ((GraphConnectionContext)context.ConnectionContext).Client;
            this.BuildUsersToIgnore();
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

        public async Task GetCSEntryChangesAsync(SchemaType type)
        {
            try
            {
                BufferBlock<User> queue = new BufferBlock<User>(new DataflowBlockOptions() { CancellationToken = this.token });

                Task consumerTask = this.ConsumeObjects(type, queue);

                await this.ProduceObjects(queue);
                await consumerTask;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "There was an error importing the user data");
                throw;
            }
        }

        private async Task ProduceObjects(ITargetBlock<User> target)
        {
            string newDeltaLink = null;

            if (this.context.InDelta)
            {
                if (!this.context.IncomingWatermark.Contains("user"))
                {
                    throw new WarningNoWatermarkException();
                }

                Watermark watermark = this.context.IncomingWatermark["user"];

                if (watermark.Value == null)
                {
                    throw new WarningNoWatermarkException();
                }

                newDeltaLink = await GraphHelperUsers.GetUsersWithDelta(this.client, watermark.Value, target, this.token);
            }
            else
            {
                //await GraphHelperUsers.GetUsers(client, target, context.Token, "displayName", "onPremisesSamAccountName", "id", "userPrincipalName");
                newDeltaLink = await GraphHelperUsers.GetUsersWithDelta(this.client, target, this.token, "displayName", "onPremisesSamAccountName", "id", "userPrincipalName");
            }

            if (newDeltaLink != null)
            {
                logger.Trace($"Got delta link {newDeltaLink}");
                this.context.OutgoingWatermark.Add(new Watermark("user", newDeltaLink, "string"));
            }

            target.Complete();
        }

        private async Task ConsumeObjects(SchemaType type, ISourceBlock<User> source)
        {
            var edfo = new ExecutionDataflowBlockOptions
            {
                MaxDegreeOfParallelism = MicrosoftTeamsMAConfigSection.Configuration.ImportThreads,
                CancellationToken = this.token,
            };

            ConcurrentDictionary<string, byte> seenUserIDs = new ConcurrentDictionary<string, byte>();

            ActionBlock<User> action = new ActionBlock<User>(user =>
            {
                try
                {
                    if (!seenUserIDs.TryAdd(user.Id, 0))
                    {
                        logger.Trace($"Skipping duplicate entry returned from graph for user {user.Id}");
                        return;
                    }

                    if (this.usersToIgnore.Contains(user.Id.ToLower()))
                    {
                        logger.Trace($"Ignoring user {user.Id}");
                        return;
                    }

                    CSEntryChange c = this.UserToCSEntryChange(this.context.InDelta, type, user);

                    if (c != null)
                    {
                        this.context.ImportItems.Add(c, this.token);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    CSEntryChange csentry = CSEntryChange.Create();
                    csentry.DN = user.Id;
                    csentry.ErrorCodeImport = MAImportError.ImportErrorCustomContinueRun;
                    csentry.ErrorDetail = ex.StackTrace;
                    csentry.ErrorName = ex.Message;
                    this.context.ImportItems.Add(csentry, this.token);
                }
            }, edfo);

            source.LinkTo(action, new DataflowLinkOptions() { PropagateCompletion = true });

            await action.Completion;
        }

        private CSEntryChange UserToCSEntryChange(bool inDelta, SchemaType schemaType, User user)
        {
            CSEntryChange c = CSEntryChange.Create();
            c.ObjectType = "user";
            c.AnchorAttributes.Add(AnchorAttribute.Create("id", user.Id));
            c.DN = user.Id;

            bool isRemoved = user.AdditionalData?.ContainsKey("@removed") ?? false;

            if (inDelta)
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
                if (type.Name == "onPremisesSamAccountName")
                {
                    if (!string.IsNullOrWhiteSpace(user.OnPremisesSamAccountName))
                    {
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, user.OnPremisesSamAccountName));
                    }
                }
                else if (type.Name == "upn")
                {
                    if (!string.IsNullOrWhiteSpace(user.UserPrincipalName))
                    {
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, user.UserPrincipalName));
                    }
                }
                else if (type.Name == "displayName")
                {
                    if (!string.IsNullOrWhiteSpace(user.DisplayName))
                    {
                        c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, user.DisplayName));
                    }
                }
                else if (type.Name == "id")
                {
                    c.AttributeChanges.Add(AttributeChange.CreateAttributeAdd(type.Name, user.Id));
                }
                else
                {
                    throw new NoSuchAttributeInObjectTypeException($"The attribute {type.Name} was not found");
                }
            }

            return c;
        }

        public bool CanImport(SchemaType type)
        {
            return type.Name == "user";
        }
    }
}
