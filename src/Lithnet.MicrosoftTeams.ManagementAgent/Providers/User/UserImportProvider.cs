using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.MetadirectoryServices;
using NLog;
using System.Threading.Tasks.Dataflow;
using Microsoft.Graph;
using Logger = NLog.Logger;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class UserImportProvider : IObjectImportProvider
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
                IGraphServiceUsersCollectionPage users = await this.GetUserEnumerable(context.InDelta, context.IncomingWatermark, ((GraphConnectionContext)context.ConnectionContext).Client, context);
                BufferBlock<User> queue = new BufferBlock<User>();

                Task consumer = this.ConsumeObjects(context, type, queue);

                // Post source data to the dataflow block.
                await this.ProduceObjects(users, queue).ConfigureAwait(false);

                // Wait for the consumer to process all data.
                await consumer.ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "There was an error importing the user data");
                throw;
            }
        }

        private async Task ProduceObjects(IGraphServiceUsersCollectionPage page, ITargetBlock<User> target)
        {
            foreach (User user in page.CurrentPage)
            {
                target.Post(user);
            }

            while (page.NextPageRequest != null)
            {
                page = await page.NextPageRequest.GetAsync();

                foreach (User user in page.CurrentPage)
                {
                    target.Post(user);
                }
            }

            target.Complete();
        }

        private async Task ConsumeObjects(ImportContext context, SchemaType type, ISourceBlock<User> source)
        {
            while (await source.OutputAvailableAsync())
            {
                User user = source.Receive();

                try
                {
                    CSEntryChange c = await this.UserToCSEntryChange(context.InDelta, type, user, context).ConfigureAwait(false);

                    if (c != null)
                    {
                        context.ImportItems.Add(c, context.CancellationTokenSource.Token);
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
                    context.ImportItems.Add(csentry, context.CancellationTokenSource.Token);
                }

                context.CancellationTokenSource.Token.ThrowIfCancellationRequested();
            }
        }

        private async Task<CSEntryChange> UserToCSEntryChange(bool inDelta, SchemaType schemaType, User user, ImportContext context)
        {
            logger.Trace($"Creating CSEntryChange for {user.Id}/{user.OnPremisesSamAccountName}");

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

        private async Task<IGraphServiceUsersCollectionPage> GetUserEnumerable(bool inDelta, WatermarkKeyedCollection importState, GraphServiceClient client, ImportContext context)
        {
            return await client.Users.Request().Select(e => new
            {
                e.DisplayName,
                e.OnPremisesSamAccountName,
                e.Id,
                e.UserPrincipalName
            }).GetAsync(context.CancellationTokenSource.Token);
        }

        public bool CanImport(SchemaType type)
        {
            return type.Name == "user";
        }
    }
}
