extern alias BetaLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Lithnet.Ecma2Framework;
using Beta = BetaLib.Microsoft.Graph;
using Microsoft.Graph;
using Newtonsoft.Json;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class GraphHelperGroups
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static async Task DeleteGroup(GraphServiceClient client, string id, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().DeleteAsync(token), token, 1);
        }

        public static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Members.Request().GetAsync(token), token, 0);

            return await GetMembers(result, token);
        }

        public static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Owners.Request().GetAsync(token), token, 0);

            return await GetOwners(result, token);
        }

        public static async Task AddGroupMembers(GraphServiceClient client, string groupid, IList<string> members, bool ignoreMemberExists, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            Dictionary<string, Func<BatchRequestStep>> requests = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (string member in members)
            {
                requests.Add(member, () => GraphHelper.GenerateBatchRequestStepJsonContent(HttpMethod.Post, member, client.Groups[groupid].Members.References.Request().RequestUrl, GraphHelperGroups.GetUserODataId(member)));
            }

            logger.Trace($"Adding {requests.Count} members in batch request for group {groupid}");

            await GraphHelper.SubmitAsBatches(client, requests, false, ignoreMemberExists, token);
        }

        public static async Task UpdateGroupOwners(GraphServiceClient client, string groupid, IList<string> adds, IList<string> deletes, bool ignoreMemberExists, bool ignoreNotFound, CancellationToken token)
        {
            // If we try to delete the last owner on a channel, the operation will fail. If we are swapping out the full set of owners (eg an add/delete of 100 owners), this will never succeed if we do a 'delete' operation first.
            // If we do an 'add' operation first, and the channel already has the maximum number of owners, the call will fail.
            // So the order of events should be to
            //    1) Process all membership removals except for one owner (100-99 = 1 owner)
            //    2) Process all membership adds except for one owner (1 + 99 = 100 owners)
            //    3) Remove the final owner (100 - 1 = 99 owners)
            //    4) Add the final owner (99 + 1 = 100 owners)

            string lastOwnerToRemove = null;

            if (deletes.Count > 0)
            {
                if (adds.Count > 0)
                {
                    // We only need to deal with the above condition if we are processing deletes and adds at the same time

                    lastOwnerToRemove = deletes[0];
                    deletes.RemoveAt(0);
                }

                await GraphHelperGroups.RemoveGroupOwners(client, groupid, deletes, true, token);
            }

            string lastOwnerToAdd = null;
            if (adds.Count > 0)
            {
                if (deletes.Count > 0)
                {
                    // We only need to deal with the above condition if we are processing deletes and adds at the same time

                    lastOwnerToAdd = adds[0];
                    adds.RemoveAt(0);
                }

                await GraphHelperGroups.AddGroupOwners(client, groupid, adds, true, token);
            }

            if (lastOwnerToRemove != null)
            {
                await GraphHelperGroups.RemoveGroupOwners(client, groupid, new List<string>() { lastOwnerToRemove }, ignoreNotFound, token);
            }

            if (lastOwnerToAdd != null)
            {
                await GraphHelperGroups.AddGroupOwners(client, groupid, new List<string>() { lastOwnerToAdd }, ignoreMemberExists, token);
            }
        }

        public static async Task AddGroupOwners(GraphServiceClient client, string groupid, IList<string> members, bool ignoreMemberExists, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            Dictionary<string, Func<BatchRequestStep>> requests = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (string member in members)
            {
                requests.Add(member, () => GraphHelper.GenerateBatchRequestStepJsonContent(HttpMethod.Post, member, client.Groups[groupid].Owners.References.Request().RequestUrl, GraphHelperGroups.GetUserODataId(member)));
            }

            logger.Trace($"Adding {requests.Count} owners in batch request for group {groupid}");
            await GraphHelper.SubmitAsBatches(client, requests, false, ignoreMemberExists, token);
        }

        public static async Task RemoveGroupMembers(GraphServiceClient client, string groupid, IList<string> members, bool ignoreNotFound, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            Dictionary<string, Func<BatchRequestStep>> requests = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (string member in members)
            {
                requests.Add(member, () => GraphHelper.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Members[member].Reference.Request().RequestUrl));
            }

            logger.Trace($"Removing {requests.Count} members in batch request for group {groupid}");
            await GraphHelper.SubmitAsBatches(client, requests, ignoreNotFound, false, token);
        }

        public static async Task RemoveGroupOwners(GraphServiceClient client, string groupid, IList<string> members, bool ignoreNotFound, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            Dictionary<string, Func<BatchRequestStep>> requests = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (string member in members)
            {
                requests.Add(member, () => GraphHelper.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Owners[member].Reference.Request().RequestUrl));
            }

            logger.Trace($"Removing {requests.Count} owners in batch request for group {groupid}");
            await GraphHelper.SubmitAsBatches(client, requests, ignoreNotFound, false, token);
        }

        public static async Task GetGroups( BetaLib::Microsoft.Graph.GraphServiceClient client, ITargetBlock< BetaLib::Microsoft.Graph.Group> target, string filter, CancellationToken token, params string[] selectProperties)
        {
            string constructedFilter = "resourceProvisioningOptions/Any(x:x eq 'Team')";

            if (!string.IsNullOrWhiteSpace(filter))
            {
                constructedFilter += $" and {filter}";
            }

            logger.Trace($"Enumerating groups with filter {constructedFilter}");

            var request = client.Groups.Request().Select(string.Join(",", selectProperties)).Filter(constructedFilter);

            await GetGroups(request, target, token);
        }

        public static async Task<string> GetGroups(GraphServiceClient client, ITargetBlock<Group> target, CancellationToken token, params string[] selectProperties)
        {
            var request = client.Groups.Delta().Request().Select(string.Join(",", selectProperties));
            return await GetGroups(request, target, token);
        }

        public static async Task<Group> CreateGroup(GraphServiceClient client, Group group, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups.Request().AddAsync(group, token), token, 1);
        }

        public static async Task<string> GetGroupIdByMailNickname(GraphServiceClient client, string mailNickname, CancellationToken token)
        {
            var collectionPage = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups.Request().Filter($"mailNickname eq '{mailNickname}'").Select("id").GetAsync(token), token, 0);

            if (collectionPage.Count == 0)
            {
                return null;
            }

            if (collectionPage.Count > 1)
            {
                throw new InvalidOperationException();
            }

            return collectionPage[0].Id;
        }

        public static async Task<Group> GetGroup(GraphServiceClient client, string id, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().GetAsync(token), token, 1);
        }

        public static async Task UpdateGroup(GraphServiceClient client, string id, Group group, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().UpdateAsync(group, token), token, 1);
        }

        public static async Task<string> GetGroups(GraphServiceClient client, string deltaLink, ITargetBlock<Group> target, CancellationToken token)
        {
            IGroupDeltaCollectionPage page = new GroupDeltaCollectionPage();
            page.InitializeNextPageRequest(client, deltaLink);
            return await GetGroups(page, target, token);
        }

        private static async Task<string> GetGroups(IGroupDeltaRequest request, ITargetBlock<Group> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);
            return await GetGroups(page, target, token);
        }

        private static async Task GetGroups( BetaLib::Microsoft.Graph.IGraphServiceGroupsCollectionRequest request, ITargetBlock< BetaLib::Microsoft.Graph.Group> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach ( BetaLib::Microsoft.Graph.Group group in page.CurrentPage)
            {
                target.Post(group);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach ( BetaLib::Microsoft.Graph.Group group in page.CurrentPage)
                {
                    target.Post(group);
                }
            }
        }

        private static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesPage page, CancellationToken token)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                    members.AddRange(page.CurrentPage);
                }
            }

            return members;
        }

        private static async Task<List<DirectoryObject>> GetOwners(IGroupOwnersCollectionWithReferencesPage page, CancellationToken token)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                    members.AddRange(page.CurrentPage);
                }
            }

            return members;
        }

        private static async Task<string> GetGroups(IGroupDeltaCollectionPage page, ITargetBlock<Group> target, CancellationToken token)
        {
            Group lastGroup = null;

            if (page.Count > 0)
            {
                foreach (Group group in page.CurrentPage)
                {
                    if (lastGroup?.Id == group.Id)
                    {
                        logger.Trace($"Merging group {group.Id}");
                        MergeGroup(lastGroup, group);
                    }
                    else
                    {
                        if (lastGroup != null)
                        {
                            logger.Trace($"posting group {lastGroup.Id}. next group {group.Id}");
                            target.Post(lastGroup);
                        }
                    }

                    lastGroup = group;
                }
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach (Group group in page.CurrentPage)
                {
                    if (lastGroup?.Id == group.Id)
                    {
                        logger.Trace($"Merging group {group.Id}");
                        MergeGroup(lastGroup, group);
                    }
                    else
                    {
                        if (lastGroup != null)
                        {
                            logger.Trace($"posting group {lastGroup.Id}. next group {group.Id}");
                            target.Post(lastGroup);
                        }
                    }

                    lastGroup = group;
                }
            }

            if (lastGroup != null)
            {
                target.Post(lastGroup);
            }

            page.AdditionalData.TryGetValue("@odata.deltaLink", out object link);

            return link as string;
        }

        private static void MergeGroup(Group source, Group target)
        {
            target.AdditionalData.Add($"mergedGroup-{Guid.NewGuid()}", source.AdditionalData);
        }

        private static string GetUserODataId(string id)
        {
            return $"{{\"@odata.id\":\"https://graph.microsoft.com/v1.0/users/{id}\"}}";
        }
    }
}
