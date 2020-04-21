using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using Newtonsoft.Json;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class GraphHelper
    {
        private const int MaxJsonBatchRequests = 20;

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        internal static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupMembersCollectionWithReferencesPage result = await client.Groups[groupid].Members
                .Request()
                .GetAsync(token);

            return await GraphHelper.GetMembers(result);
        }

        internal static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid, CancellationToken token)
        {
            IGroupOwnersCollectionWithReferencesPage result = await client.Groups[groupid].Owners
                .Request()
                .GetAsync(token);

            return await GraphHelper.GetMembers(result);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupMembersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await GraphHelper.GetMembers(page);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupOwnersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupOwnersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await GraphHelper.GetMembers(page);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesPage page)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync();

                    members.AddRange(page.CurrentPage);
                }
            }

            return members;
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupOwnersCollectionWithReferencesPage page)
        {
            List<DirectoryObject> members = new List<DirectoryObject>();

            if (page?.Count > 0)
            {
                members.AddRange(page.CurrentPage);

                while (page.NextPageRequest != null)
                {
                    page = await page.NextPageRequest.GetAsync();

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


        private static async Task SubmitBatchContent(GraphServiceClient client, BatchRequestContent content, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token)
        {
            BatchResponseContent response = await client.Batch.Request().PostAsync(content, token);

            List<Exception> exceptions = new List<Exception>();

            foreach (KeyValuePair<string, HttpResponseMessage> r in await response.GetResponsesAsync())
            {
                if (!r.Value.IsSuccessStatusCode)
                {
                    if (ignoreNotFound && r.Value.StatusCode == System.Net.HttpStatusCode.NotFound)
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

                    if (ignoreRefAlreadyExists && r.Value.StatusCode == System.Net.HttpStatusCode.BadRequest && er.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                        continue;
                    }

                    exceptions.Add(new ServiceException(er.Error, r.Value.Headers, r.Value.StatusCode));
                }
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
