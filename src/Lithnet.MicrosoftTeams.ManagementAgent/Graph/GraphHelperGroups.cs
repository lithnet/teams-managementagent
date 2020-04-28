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

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                HttpRequestMessage createEventMessage = new HttpRequestMessage(HttpMethod.Post, client.Groups[groupid].Members.References.Request().RequestUrl);
                createEventMessage.Content = CreateStringContentForMemberId(member);

                requests.Add(new BatchRequestStep(member, createEventMessage));
            }

            logger.Trace($"Adding {requests.Count} members in batch request for group {groupid}");

            await GraphHelper.SubmitAsBatches(client, requests, false, ignoreMemberExists, token);
        }

        public static async Task AddGroupOwners(GraphServiceClient client, string groupid, IList<string> members, bool ignoreMemberExists, CancellationToken token)
        {
            if (members.Count == 0)
            {
                return;
            }

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                HttpRequestMessage createEventMessage = new HttpRequestMessage(HttpMethod.Post, client.Groups[groupid].Owners.References.Request().RequestUrl);
                createEventMessage.Content = CreateStringContentForMemberId(member);

                requests.Add(new BatchRequestStep(member, createEventMessage));
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

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                requests.Add(GraphHelper.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Members[member].Reference.Request().RequestUrl));
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

            List<BatchRequestStep> requests = new List<BatchRequestStep>();

            foreach (string member in members)
            {
                requests.Add(GraphHelper.GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Owners[member].Reference.Request().RequestUrl));
            }

            logger.Trace($"Removing {requests.Count} owners in batch request for group {groupid}");
            await GraphHelper.SubmitAsBatches(client, requests, ignoreNotFound, false, token);
        }

        public static async Task GetGroups(Beta.GraphServiceClient client, ITargetBlock<Beta.Group> target, string filter, CancellationToken token, params string[] selectProperties)
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
        
        private static async Task GetGroups(Beta.IGraphServiceGroupsCollectionRequest request, ITargetBlock<Beta.Group> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach (Beta.Group group in page.CurrentPage)
            {
                target.Post(group);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach (Beta.Group group in page.CurrentPage)
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

        private static StringContent CreateStringContentForMemberId(string member)
        {
            return new StringContent("{\"@odata.id\":\"https://graph.microsoft.com/beta/users/" + member + "\"}", Encoding.UTF8, "application/json");
        }
    }
}
