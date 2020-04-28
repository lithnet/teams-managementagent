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
    internal static class GraphHelper
    {
        private const int MaxJsonBatchRequests = 20;

        private const int MaxRetry = 7;

        private static TokenBucket rateLimiter = new TokenBucket("graph", MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit, TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestWindowSeconds), MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit);

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Members.Request().GetAsync(token), token, 0);

            return await GraphHelper.GetMembers(result, token);
        }

        public static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Owners.Request().GetAsync(token), token, 0);

            return await GraphHelper.GetOwners(result, token);
        }

        public static async Task<Team> GetTeam(GraphServiceClient client, string teamid, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[teamid].Request().GetAsync(token), token, 0);
        }

        public static async Task<List<Beta.Channel>> GetChannels(Beta.GraphServiceClient client, string groupid, CancellationToken token)
        {
            List<Beta.Channel> channels = new List<Beta.Channel>();

            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[groupid].Channels.Request().GetAsync(token), token, 0);

            if (page?.Count > 0)
            {
                channels.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                    channels.AddRange(page.CurrentPage);
                }
            }

            return channels;
        }

        public static async Task<List<Beta.ConversationMember>> GetChannelMembers(Beta.GraphServiceClient client, string groupId, string channelId, CancellationToken token)
        {
            List<Beta.ConversationMember> members = new List<Beta.ConversationMember>();

            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[groupId].Channels[channelId].Members.Request().GetAsync(token), token, 0);

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

            await SubmitAsBatches(client, requests, false, ignoreMemberExists, token);
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
            await SubmitAsBatches(client, requests, false, ignoreMemberExists, token);
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
                requests.Add(GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Members[member].Reference.Request().RequestUrl));
            }

            logger.Trace($"Removing {requests.Count} members in batch request for group {groupid}");
            await SubmitAsBatches(client, requests, ignoreNotFound, false, token);
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
                requests.Add(GenerateBatchRequestStep(HttpMethod.Delete, member, client.Groups[groupid].Owners[member].Reference.Request().RequestUrl));
            }

            logger.Trace($"Removing {requests.Count} owners in batch request for group {groupid}");
            await SubmitAsBatches(client, requests, ignoreNotFound, false, token);
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

        private static async Task<string> GetUsers(IUserDeltaRequest request, ITargetBlock<User> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach (User user in page.CurrentPage)
            {
                target.Post(user);
            }

            return await GetUsers(page, target, token);
        }

        public static async Task<string> GetUsers(GraphServiceClient client, string deltaLink, ITargetBlock<User> target, CancellationToken token)
        {
            IUserDeltaCollectionPage page = new UserDeltaCollectionPage();
            page.InitializeNextPageRequest(client, deltaLink);
            return await GetUsers(page, target, token);
        }

        public static async Task<string> GetUsers(GraphServiceClient client, ITargetBlock<User> target, CancellationToken token, params string[] selectProperties)
        {
            var request = client.Users.Delta().Request().Select(string.Join(",", selectProperties));
            return await GraphHelper.GetUsers(request, target, token);
        }

        private static async Task<string> GetUsers(IUserDeltaCollectionPage page, ITargetBlock<User> target, CancellationToken token)
        {
            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach (User user in page.CurrentPage)
                {
                    target.Post(user);
                }
            }

            page.AdditionalData.TryGetValue("@odata.deltaLink", out object link);

            return link as string;
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
            return await GraphHelper.GetGroups(request, target, token);
        }

        private static async Task<string> GetGroups(IGroupDeltaRequest request, ITargetBlock<Group> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);
            return await GetGroups(page, target, token);
        }

        public static async Task<string> GetGroups(GraphServiceClient client, string deltaLink, ITargetBlock<Group> target, CancellationToken token)
        {
            IGroupDeltaCollectionPage page = new GroupDeltaCollectionPage();
            page.InitializeNextPageRequest(client, deltaLink);
            return await GetGroups(page, target, token);
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

        private static async Task GetUsers(IGraphServiceUsersCollectionRequest request, ITargetBlock<User> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach (User user in page.CurrentPage)
            {
                target.Post(user);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach (User user in page.CurrentPage)
                {
                    target.Post(user);
                }
            }
        }

        public static async Task DeleteGroup(GraphServiceClient client, string id, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().DeleteAsync(token), token, 1);
        }

        public static async Task<Team> CreateTeam(GraphServiceClient client, string groupid, Team team, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Team.Request().PutAsync(team, token), token, 1);
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

        public static async Task UpdateTeam(GraphServiceClient client, string id, Team team, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[id].Request().UpdateAsync(team, token), token, 1);
        }

        public static async Task UpdateGroup(GraphServiceClient client, string id, Group group, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().UpdateAsync(group, token), token, 1);
        }

        private static async Task SubmitAsBatches(GraphServiceClient client, List<BatchRequestStep> requests, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token)
        {
            BatchRequestContent content = new BatchRequestContent();
            int count = 0;

            foreach (BatchRequestStep r in requests)
            {
                if (count == MaxJsonBatchRequests)
                {
                    await SubmitBatchContent(client, content, ignoreNotFound, ignoreRefAlreadyExists, token);
                    count = 0;
                    content = new BatchRequestContent();
                }

                content.AddBatchRequestStep(r);
                count++;
            }

            if (count > 0)
            {
                await SubmitBatchContent(client, content, ignoreNotFound, ignoreRefAlreadyExists, token);
            }
        }

        private static async Task SubmitBatchContent(GraphServiceClient client, BatchRequestContent content, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token, int retryCount = 1)
        {
            BatchResponseContent response = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Batch.Request().PostAsync(content, token), token, 21);

            List<Exception> exceptions = new List<Exception>();
            List<BatchRequestStep> stepsToRetry = new List<BatchRequestStep>();
            int retryInterval = 0;

            var responses = await response.GetResponsesAsync();

            foreach (KeyValuePair<string, HttpResponseMessage> r in responses)
            {
                using (r.Value)
                {
                    if (!r.Value.IsSuccessStatusCode)
                    {
                        if (ignoreNotFound && r.Value.StatusCode == HttpStatusCode.NotFound)
                        {
                            logger.Warn($"The request ({r.Key}) to remove object failed because it did not exist");
                            continue;
                        }

                        ErrorResponse er;
                        try
                        {
                            string econtent = await r.Value.Content.ReadAsStringAsync();
                            logger.Trace(econtent);

                            er = JsonConvert.DeserializeObject<ErrorResponse>(econtent);
                        }
                        catch (Exception ex)
                        {
                            logger.Trace(ex, "The error response could not be deserialized");
                            er = new ErrorResponse
                            {
                                Error = new Error
                                {
                                    Code = r.Value.StatusCode.ToString(),
                                    Message = r.Value.ReasonPhrase
                                }
                            };
                        }

                        if (r.Value.StatusCode == (HttpStatusCode)429 && retryCount <= 5)
                        {
                            if (retryInterval == 0 && r.Value.Headers.TryGetValues("Retry-After", out IEnumerable<string> outvalues))
                            {
                                string tryAfter = outvalues.FirstOrDefault() ?? "0";
                                retryInterval = int.Parse(tryAfter);
                                logger.Warn($"Rate limit encountered, backoff interval of {retryInterval} found");
                            }
                            else
                            {
                                logger.Warn("Rate limit encountered, but no backoff interval specified");
                            }

                            var step = content.BatchRequestSteps.FirstOrDefault(t => t.Key == r.Key);
                            stepsToRetry.Add(step.Value);
                            continue;
                        }

                        if (ignoreRefAlreadyExists && r.Value.StatusCode == HttpStatusCode.BadRequest && er.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                        {
                            logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                            continue;
                        }

                        exceptions.Add(new ServiceException(er.Error, r.Value.Headers, r.Value.StatusCode));
                    }
                }
            }

            if (stepsToRetry.Count > 0 && retryCount <= 5)
            {
                BatchRequestContent newContent = new BatchRequestContent();

                foreach (var stepToRetry in stepsToRetry)
                {
                    newContent.AddBatchRequestStep(stepToRetry);
                }

                if (retryInterval == 0)
                {
                    retryInterval = 30;
                }

                logger.Info($"Sleeping for {retryInterval} before retrying after attempt {retryCount}");
                await Task.Delay(TimeSpan.FromSeconds(retryInterval), token);
                await SubmitBatchContent(client, newContent, ignoreNotFound, ignoreRefAlreadyExists, token, ++retryCount);
            }

            if (exceptions.Count == 1)
            {
                throw exceptions[0];
            }
            if (exceptions.Count > 1)
            {
                throw new AggregateException("Multiple operations failed", exceptions);
            }
        }

        private static BatchRequestStep GenerateBatchRequestStep(HttpMethod method, string id, string requestUrl)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
            return new BatchRequestStep(id, request);
        }

        private static StringContent CreateStringContentForMemberId(string member)
        {
            return new StringContent("{\"@odata.id\":\"https://graph.microsoft.com/beta/users/" + member + "\"}", Encoding.UTF8, "application/json");
        }

        private static bool IsRetryable(Exception ex)
        {
            return ex is TimeoutException || ex is ServiceException se && (se.StatusCode == HttpStatusCode.NotFound || se.StatusCode == HttpStatusCode.BadGateway);
        }

        private static T ExecuteWithRetry<T>(Func<T> task, CancellationToken token)
        {
            return ExecuteWithRetryAndRateLimit(task, token, 0);
        }

        private static async Task<T> ExecuteWithRetry<T>(Func<Task<T>> task, CancellationToken token)
        {
            return await ExecuteWithRetryAndRateLimit(task, token, 0);
        }

        private static T ExecuteWithRetryAndRateLimit<T>(Func<T> task, CancellationToken token, int requests)
        {
            T result = default(T);

            bool success = false;
            int retryCount = 0;

            while (!success)
            {
                try
                {
                    rateLimiter.Consume(requests, token);
                    result = task();
                    success = true;
                }
                catch (ServiceException ex)
                {
                    if (IsRetryable(ex) && retryCount <= MaxRetry)
                    {
                        retryCount++;
                        logger.Warn(ex, $"A retryable error was detected (attempt: {retryCount})");
                        Task.Delay(TimeSpan.FromSeconds(5 * retryCount), token).Wait(token);
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            return result;
        }

        private static async Task<T> ExecuteWithRetryAndRateLimit<T>(Func<Task<T>> task, CancellationToken token, int requests)
        {
            T result = default(T);

            bool success = false;
            int retryCount = 0;

            while (!success)
            {
                try
                {
                    rateLimiter.Consume(requests, token);
                    result = await task();
                    success = true;
                }
                catch (ServiceException ex)
                {
                    if (IsRetryable(ex) && retryCount <= MaxRetry)
                    {
                        retryCount++;
                        logger.Warn(ex, $"A retryable error was detected (attempt: {retryCount})");
                        Task.Delay(TimeSpan.FromSeconds(5 * retryCount), token).Wait(token);
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            return result;
        }
    }
}
