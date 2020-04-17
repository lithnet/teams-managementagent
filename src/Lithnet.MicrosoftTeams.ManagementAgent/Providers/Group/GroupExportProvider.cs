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

            Group group = new Group();

            group.DisplayName = csentry.AttributeChanges["displayName"].GetValueAdd<string>();
            group.GroupTypes = new List<String>() { "Unified" };
            group.MailEnabled = true;
            group.Description = "test";
            group.MailNickname = "mx-" + Guid.NewGuid();
            group.SecurityEnabled = false;
            group.AdditionalData = new Dictionary<string, object>();
            group.Id = csentry.DN;

            IList<string> members = csentry.AttributeChanges["member"].GetValueAdds<string>();
            IList<string> owners = csentry.AttributeChanges["owner"].GetValueAdds<string>();

            if (members.Count > 0)
            {
                group.AdditionalData.Add("members@odata.bind", members.Select(t => $"https://graph.microsoft.com/v1.0/users/{t}").ToArray());
            }

            if (owners.Count > 0)
            {
                group.AdditionalData.Add("owners@odata.bind", owners.Select(t => $"https://graph.microsoft.com/v1.0/users/{t}").ToArray());
            }

            logger.Trace($"Creating group {group.MailNickname} with {members.Count} members and {owners.Count} owners");

            Group result = await client.Groups
                .Request()
                .AddAsync(group, context.CancellationTokenSource.Token);

            logger.Info($"Created group {group.Id}");

            Team team = new Team();
            team.ODataType = null;

            var tresult = await client.Groups[$"{result.Id}"].Team
                .Request()
                .PutAsync(team);

            logger.Info($"Created team {tresult.Id} for group {result.Id}");

            List<AttributeChange> anchorChanges = new List<AttributeChange>();
            anchorChanges.Add(AttributeChange.CreateAttributeAdd("id", result.Id));

            return CSEntryChangeResult.Create(csentry.Identifier, anchorChanges, MAExportError.Success);
        }

        private async Task<CSEntryChangeResult> PutCSEntryChangeUpdate(CSEntryChange csentry, ExportContext context)
        {
            GraphServiceClient client = ((GraphConnectionContext)context.ConnectionContext).Client;

            //Group group = null;

            //foreach (AttributeChange change in csentry.AttributeChanges)
            //{
            //    if (change.Name == "name")
            //    {
            //        if (group == null)
            //        {
            //            group = AsyncHelper.RunSync(client.Groups.GetGroupAsync(csentry.DN, null, context.CancellationTokenSource.Token), context.CancellationTokenSource.Token);
            //        }

            //        group.Profile.Name = change.GetValueAdd<string>();
            //    }
            //    else if (change.Name == "description")
            //    {
            //        if (group == null)
            //        {
            //            group = AsyncHelper.RunSync(client.Groups.GetGroupAsync(csentry.DN, null, context.CancellationTokenSource.Token), context.CancellationTokenSource.Token);
            //        }

            //        group.Profile.Description = change.GetValueAdd<string>();
            //    }
            //    else if (change.Name == "member")
            //    {
            //        foreach (string add in change.GetValueAdds<string>())
            //        {
            //            AsyncHelper.RunSync(client.Groups.AddUserToGroupAsync(csentry.DN, add, context.CancellationTokenSource.Token), context.CancellationTokenSource.Token);
            //        }

            //        foreach (string delete in change.GetValueDeletes<string>())
            //        {
            //            AsyncHelper.RunSync(client.Groups.RemoveGroupUserAsync(csentry.DN, delete, context.CancellationTokenSource.Token), context.CancellationTokenSource.Token);
            //        }
            //    }
            //}

            //if (group != null)
            //{
            //    AsyncHelper.RunSync(client.Groups.UpdateGroupAsync(group, csentry.DN, context.CancellationTokenSource.Token), context.CancellationTokenSource.Token);
            //}

            return CSEntryChangeResult.Create(csentry.Identifier, null, MAExportError.Success);
        }
    }
}
