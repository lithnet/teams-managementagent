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

        private static readonly TokenBucket RateLimiter = new TokenBucket("graph", MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit, TimeSpan.FromSeconds(MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestWindowSeconds), MicrosoftTeamsMAConfigSection.Configuration.RateLimitRequestLimit);

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        internal static async Task SubmitAsBatches(IBaseClient client, Dictionary<string, Func<BatchRequestStep>> requests, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token)
        {
            foreach (var batch in GetBatchPartitions(requests))
            {
                await SubmitBatchContent(client, batch, ignoreNotFound, ignoreRefAlreadyExists, token);
            }
        }

        private static IEnumerable<Dictionary<string, Func<BatchRequestStep>>> GetBatchPartitions(Dictionary<string, Func<BatchRequestStep>> requests)
        {
            Dictionary<string, Func<BatchRequestStep>> batch = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (KeyValuePair<string, Func<BatchRequestStep>> r in requests)
            {
                if (batch.Count == MaxJsonBatchRequests)
                {
                    yield return batch;
                    batch = new Dictionary<string, Func<BatchRequestStep>>();
                }

                batch.Add(r.Key, r.Value);
            }

            if (batch.Count > 0)
            {
                yield return batch;
            }
        }

        private static async Task SubmitBatchContent(IBaseClient client, Dictionary<string, Func<BatchRequestStep>> requests, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token, int attemptCount = 1)
        {
            BatchRequestContent content = GraphHelper.BuildBatchRequest(requests);

            BatchResponseContent response = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Batch.Request().PostAsync(content, token), token, content.BatchRequestSteps.Count + 1);

            List<GraphBatchResult> results = await GetBatchResults(await response.GetResponsesAsync(), ignoreNotFound, ignoreRefAlreadyExists, attemptCount <= MaxRetry);

            GraphHelper.ThrowOnExceptions(results);

            int retryInterval = 8 * attemptCount;
            Dictionary<string, Func<BatchRequestStep>> stepsToRetry = new Dictionary<string, Func<BatchRequestStep>>();

            foreach (var result in results.Where(t => t.IsRetryable))
            {
                retryInterval = Math.Max(result.RetryInterval, retryInterval);
                stepsToRetry.Add(result.ID, requests[result.ID]);
            }

            if (stepsToRetry.Count > 0)
            {
                logger.Info($"Sleeping for {retryInterval} before retrying after attempt {attemptCount}");
                await Task.Delay(TimeSpan.FromSeconds(retryInterval), token);
                await GraphHelper.SubmitBatchContent(client, stepsToRetry, ignoreNotFound, ignoreRefAlreadyExists, token, ++attemptCount);
            }
        }

        private static BatchRequestContent BuildBatchRequest(Dictionary<string, Func<BatchRequestStep>> requests)
        {
            BatchRequestContent content = new BatchRequestContent();

            foreach (KeyValuePair<string, Func<BatchRequestStep>> item in requests)
            {
                content.AddBatchRequestStep(item.Value.Invoke());
            }

            logger.Trace(JsonConvert.SerializeObject(content));
            return content;
        }

        private static void ThrowOnExceptions(List<GraphBatchResult> results)
        {
            List<Exception> exceptions = results.Where(t => t.IsFailed).Select(t => t.Exception).ToList();

            if (exceptions.Count > 0)
            {
                if (exceptions.Count == 1)
                {
                    throw exceptions[0];
                }

                if (exceptions.Count > 1)
                {
                    throw new AggregateException("Multiple operations failed", exceptions);
                }
            }
        }

        private static async Task<List<GraphBatchResult>> GetBatchResults(Dictionary<string, HttpResponseMessage> responses, bool ignoreNotFound, bool ignoreRefAlreadyExists, bool canRetry)
        {
            List<GraphBatchResult> results = new List<GraphBatchResult>();

            foreach (KeyValuePair<string, HttpResponseMessage> r in responses)
            {
                using (r.Value)
                {
                    GraphBatchResult result = new GraphBatchResult();
                    result.ID = r.Key;
                    result.IsSuccess = r.Value.IsSuccessStatusCode;
                    results.Add(result);

                    if (result.IsSuccess)
                    {
                        continue;
                    }

                    if (ignoreNotFound && r.Value.StatusCode == HttpStatusCode.NotFound)
                    {
                        result.IsSuccess = true;
                        GraphHelper.logger.Warn($"The request ({r.Key}) to remove object failed because it did not exist");
                        continue;
                    }

                    result.ErrorResponse = await GraphHelper.GetErrorResponseFromHttpResponseMessage(r);

                    if (ignoreRefAlreadyExists && r.Value.StatusCode == HttpStatusCode.BadRequest && result.ErrorResponse.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        result.IsSuccess = true;
                        GraphHelper.logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                        continue;
                    }

                    if (canRetry && r.Value.StatusCode == (HttpStatusCode)429)
                    {
                        if (r.Value.Headers.TryGetValues("Retry-After", out IEnumerable<string> outvalues))
                        {
                            string tryAfter = outvalues.FirstOrDefault() ?? "0";
                            result.RetryInterval = int.Parse(tryAfter);
                            GraphHelper.logger.Warn($"Rate limit encountered, backoff interval of {result.RetryInterval} found");
                        }
                        else
                        {
                            GraphHelper.logger.Warn("Rate limit encountered, but no backoff interval specified");
                        }

                        result.IsRetryable = true;
                        continue;
                    }

                    if (canRetry && r.Value.StatusCode == HttpStatusCode.NotFound && string.Equals(result.ErrorResponse.Error.Code, "Request_ResourceNotFound", StringComparison.OrdinalIgnoreCase))
                    {
                        result.IsRetryable = true;
                        continue;
                    }

                    result.IsFailed = true;
                    result.Exception = new ServiceException(result.ErrorResponse.Error, r.Value.Headers, r.Value.StatusCode);
                }
            }

            return results;
        }

        private static async Task<ErrorResponse> GetErrorResponseFromHttpResponseMessage(KeyValuePair<string, HttpResponseMessage> r)
        {
            ErrorResponse er;
            try
            {
                string econtent = await r.Value.Content.ReadAsStringAsync();
                GraphHelper.logger.Trace(econtent);

                er = JsonConvert.DeserializeObject<ErrorResponse>(econtent);
            }
            catch (Exception ex)
            {
                GraphHelper.logger.Trace(ex, "The error response could not be deserialized");

                er = new ErrorResponse
                {
                    Error = new Error
                    {
                        Code = r.Value.StatusCode.ToString(),
                        Message = r.Value.ReasonPhrase
                    }
                };
            }

            return er;
        }

        internal static BatchRequestStep GenerateBatchRequestStep(HttpMethod method, string id, string requestUrl)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
            return new BatchRequestStep(id, request);
        }

        internal static BatchRequestStep GenerateBatchRequestStepJsonContent(HttpMethod method, string id, string requestUrl, string jsonbody)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
            request.Content = new StringContent(jsonbody, Encoding.UTF8, "application/json");
            return new BatchRequestStep(id, request);
        }

        private static bool IsRetryable(Exception ex)
        {
            return ex is TimeoutException || ex is ServiceException se && (se.StatusCode == HttpStatusCode.NotFound || se.StatusCode == HttpStatusCode.BadGateway);
        }

        internal static T ExecuteWithRetry<T>(Func<T> task, CancellationToken token)
        {
            return ExecuteWithRetryAndRateLimit(task, token, 0, IsRetryable);
        }

        internal static async Task<T> ExecuteWithRetry<T>(Func<Task<T>> task, CancellationToken token)
        {
            return await ExecuteWithRetryAndRateLimit(task, token, 0, IsRetryable);
        }

        internal static T ExecuteWithRetryAndRateLimit<T>(Func<T> task, CancellationToken token, int requests)
        {
            return ExecuteWithRetryAndRateLimit(task, token, requests, IsRetryable);
        }

        internal static async Task<T> ExecuteWithRetryAndRateLimit<T>(Func<Task<T>> task, CancellationToken token, int requests)
        {
            return await ExecuteWithRetryAndRateLimit(task, token, requests, IsRetryable);
        }

        internal static T ExecuteWithRetryAndRateLimit<T>(Func<T> task, CancellationToken token, int requests, Func<Exception, bool> isRetryable)
        {
            T result = default(T);

            bool success = false;
            int retryCount = 0;

            while (!success)
            {
                try
                {
                    GraphHelper.RateLimiter.Consume(requests, token);
                    result = task();
                    success = true;
                }
                catch (ServiceException ex)
                {
                    if (isRetryable(ex) && retryCount <= MaxRetry)
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

        internal static async Task<T> ExecuteWithRetryAndRateLimit<T>(Func<Task<T>> task, CancellationToken token, int requests, Func<Exception, bool> isRetryable)
        {
            T result = default(T);

            bool success = false;
            int retryCount = 0;

            while (!success)
            {
                try
                {
                    GraphHelper.RateLimiter.Consume(requests, token);
                    result = await task();
                    success = true;
                }
                catch (ServiceException ex)
                {
                    if (isRetryable(ex) && retryCount <= MaxRetry)
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

        public static void AssignNullToProperty(this Entity e, string name)
        {
            if (e.AdditionalData == null)
            {
                e.AdditionalData = new Dictionary<string, object>();
            }

            e.AdditionalData.Add(name, null);
        }

        public static void AssignNullToProperty(this  BetaLib::Microsoft.Graph.Entity e, string name)
        {
            if (e.AdditionalData == null)
            {
                e.AdditionalData = new Dictionary<string, object>();
            }

            e.AdditionalData.Add(name, null);
        }
    }
}
