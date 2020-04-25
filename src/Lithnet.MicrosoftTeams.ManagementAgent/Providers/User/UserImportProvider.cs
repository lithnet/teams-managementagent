using System;
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

        public async Task GetCSEntryChangesAsync(ImportContext context, SchemaType type)
        {
            try
            {
                IGraphServiceUsersCollectionRequest users = this.GetUserEnumerationRequest(context.InDelta, context.IncomingWatermark, ((GraphConnectionContext)context.ConnectionContext).Client, context);

                BufferBlock<User> queue = new BufferBlock<User>(new DataflowBlockOptions() { CancellationToken = context.Token });

                Task consumerTask = this.ConsumeObjects(context, type, queue);

                await this.ProduceObjects(users, queue, context.Token);
                await consumerTask;
            }
            catch (Exception ex)
            {
                UserImportProvider.logger.Error(ex, "There was an error importing the user data");
                throw;
            }
        }

        private async Task ProduceObjects(IGraphServiceUsersCollectionRequest request, ITargetBlock<User> target, CancellationToken token)
        {
            await GraphHelper.GetUsers(request, target, token);
            target.Complete();
        }

        private async Task ConsumeObjects(ImportContext context, SchemaType type, ISourceBlock<User> source)
        {
            var edfo = new ExecutionDataflowBlockOptions
            {
                MaxDegreeOfParallelism = MicrosoftTeamsMAConfigSection.Configuration.ImportThreads,
                CancellationToken = context.Token,
            };

            ActionBlock<User> action = new ActionBlock<User>(user =>
            {
                try
                {
                    CSEntryChange c = this.UserToCSEntryChange(context.InDelta, type, user, context);

                    if (c != null)
                    {
                        context.ImportItems.Add(c, context.Token);
                    }
                }
                catch (Exception ex)
                {
                    UserImportProvider.logger.Error(ex);
                    CSEntryChange csentry = CSEntryChange.Create();
                    csentry.DN = user.Id;
                    csentry.ErrorCodeImport = MAImportError.ImportErrorCustomContinueRun;
                    csentry.ErrorDetail = ex.StackTrace;
                    csentry.ErrorName = ex.Message;
                    context.ImportItems.Add(csentry, context.Token);
                }
            }, edfo);

            source.LinkTo(action, new DataflowLinkOptions() { PropagateCompletion = true });

            await action.Completion;
        }

        private CSEntryChange UserToCSEntryChange(bool inDelta, SchemaType schemaType, User user, ImportContext context)
        {
            CSEntryChange c = CSEntryChange.Create();
            c.ObjectType = "user";
            c.ObjectModificationType = ObjectModificationType.Add;
            c.AnchorAttributes.Add(AnchorAttribute.Create("id", user.Id));
            c.DN = user.Id;

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

        private IGraphServiceUsersCollectionRequest GetUserEnumerationRequest(bool inDelta, WatermarkKeyedCollection importState, GraphServiceClient client, ImportContext context)
        {
            return client.Users.Request().Select(e => new
            {
                e.DisplayName,
                e.OnPremisesSamAccountName,
                e.Id,
                e.UserPrincipalName,
            });
        }

        public bool CanImport(SchemaType type)
        {
            return type.Name == "user";
        }
    }
}
