using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
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

        internal static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Members.Request().GetAsync(token), token, 0 );

            return await GraphHelper.GetMembers(result, token);
        }

        internal static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Owners.Request().GetAsync(token), token, 0);

            return await GraphHelper.GetOwners(result, token);
        }

        internal static async Task<Team> GetTeam(GraphServiceClient client, string teamid, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[teamid].Request().GetAsync(token), token, 0);
        }

        internal static async Task<List<Channel>> GetChannels(GraphServiceClient client, string groupid, CancellationToken token)
        {
            List<Channel> channels = new List<Channel>();

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

        internal static async Task<List<ConversationMember>> GetChannelMembers(GraphServiceClient client, string groupId, string channelId, CancellationToken token)
        {
            List<ConversationMember> members = new List<ConversationMember>();

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

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesPage page, CancellationToken token)
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

        internal static async Task<List<DirectoryObject>> GetOwners(IGroupOwnersCollectionWithReferencesPage page, CancellationToken token)
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

        internal static async Task AddGroupMembers(GraphServiceClient client, string groupid, IList<string> members, bool ignoreMemberExists, CancellationToken token)
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

        internal static async Task AddGroupOwners(GraphServiceClient client, string groupid, IList<string> members, bool ignoreMemberExists, CancellationToken token)
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

        internal static async Task RemoveGroupMembers(GraphServiceClient client, string groupid, IList<string> members, bool ignoreNotFound, CancellationToken token)
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

        internal static async Task RemoveGroupOwners(GraphServiceClient client, string groupid, IList<string> members, bool ignoreNotFound, CancellationToken token)
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

        internal static async Task GetGroups(IGraphServiceGroupsCollectionRequest request, ITargetBlock<Group> target, CancellationToken cancellationToken)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(cancellationToken), cancellationToken, 0);

            foreach (Group group in page.CurrentPage)
            {
                target.Post(group);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(cancellationToken), cancellationToken, 0);

                foreach (Group group in page.CurrentPage)
                {
                    target.Post(group);
                }
            }
        }

        internal static async Task GetUsers(IGraphServiceUsersCollectionRequest request, ITargetBlock<User> target, CancellationToken cancellationToken)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(cancellationToken), cancellationToken,0 );

            foreach (User user in page.CurrentPage)
            {
                target.Post(user);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(cancellationToken), cancellationToken, 0);

                foreach (User user in page.CurrentPage)
                {
                    target.Post(user);
                }
            }
        }

        internal static async Task DeleteGroup(GraphServiceClient client, string id, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[id].Request().DeleteAsync(token), token, 1);
        }

        internal static async Task<Team> CreateTeam(GraphServiceClient client, string groupid, Team team, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Team.Request().PutAsync(team, token), token, 1);
        }

        internal static async Task<Group> CreateGroup(GraphServiceClient client, Group group, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups.Request().AddAsync(group, token), token, 1);
        }

        internal static async Task<string> GetGroupIdByMailNickname(GraphServiceClient client, string mailNickname, CancellationToken token)
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

        internal static async Task UpdateTeam(GraphServiceClient client, string id, Team team, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[id].Request().UpdateAsync(team, token), token, 1);
        }

        internal static async Task UpdateGroup(GraphServiceClient client, string id, Group group, CancellationToken token)
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

        internal static T ExecuteWithRetry<T>(Func<T> task, CancellationToken token)
        {
            return ExecuteWithRetryAndRateLimit(task, token, 0);
        }

        internal static async Task<T> ExecuteWithRetry<T>(Func<Task<T>> task, CancellationToken token)
        {
            return await ExecuteWithRetryAndRateLimit(task, token, 0);
        }

        internal static T ExecuteWithRetryAndRateLimit<T>(Func<T> task, CancellationToken token, int requests)
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

        internal static async Task<T> ExecuteWithRetryAndRateLimit<T>(Func<Task<T>> task, CancellationToken token, int requests)
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
