using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class GraphHelper
    {
        private const int MaxJsonBatchRequests = 20;

        private static TokenBucket rateLimiter = new TokenBucket("batch", MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit, TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestWindowSeconds), MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit);

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        internal static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage result = await client.Groups[groupid].Members
                .Request()
                .GetAsync(token);

            return await GraphHelper.GetMembers(result, token);
        }

        internal static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage result = await client.Groups[groupid].Owners
                .Request()
                .GetAsync(token);

            return await GraphHelper.GetOwners(result, token);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesRequestBuilder builder, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync(token);

            return await GraphHelper.GetMembers(page, token);
        }

        internal static async Task<List<DirectoryObject>> GetOwners(IGroupOwnersCollectionWithReferencesRequestBuilder builder, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync(token);

            return await GraphHelper.GetOwners(page, token);
        }

        internal static async Task<List<Channel>> GetChannels(GraphServiceClient client, string groupid, CancellationToken token)
        {
            List<Channel> channels = new List<Channel>();

            var page = await client.Teams[groupid].Channels.Request().GetAsync(token);

            if (page?.Count > 0)
            {
                channels.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync(token);

                    channels.AddRange(page.CurrentPage);
                }
            }

            return channels;
        }

        internal static async Task<List<ConversationMember>> GetChannelMembers(GraphServiceClient client, string groupId, string channelId, CancellationToken token)
        {
            List<ConversationMember> members = new List<ConversationMember>();

            var page = await client.Teams[groupId].Channels[channelId].Members.Request().GetAsync(token);

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync(token);

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
                    page = await page.NextPageRequest.GetAsync(token);

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
                    page = await page.NextPageRequest.GetAsync(token);

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
            rateLimiter.Consume(content.BatchRequestSteps.Count + 1, token);

            BatchResponseContent response = await client.Batch.Request().PostAsync(content, token);

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
    }
}
