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

        internal static async Task SubmitAsBatches(GraphServiceClient client, List<BatchRequestStep> requests, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token)
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

        private static async Task SubmitBatchContent(GraphServiceClient client, BatchRequestContent content, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token, int attemptCount = 1)
        {
            BatchResponseContent response = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Batch.Request().PostAsync(content, token), token, content.BatchRequestSteps.Count + 1);

            List<Exception> exceptions = new List<Exception>();
            List<BatchRequestStep> stepsToRetry = new List<BatchRequestStep>();
            int retryInterval = 0;

            var responses = await response.GetResponsesAsync();

            foreach (KeyValuePair<string, HttpResponseMessage> r in responses)
            {
                using (r.Value)
                {
                    if (r.Value.IsSuccessStatusCode)
                    {
                        continue;
                    }

                    if (ignoreNotFound && r.Value.StatusCode == HttpStatusCode.NotFound)
                    {
                        GraphHelper.logger.Warn($"The request ({r.Key}) to remove object failed because it did not exist");
                        continue;
                    }

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

                    if (r.Value.StatusCode == (HttpStatusCode)429 && attemptCount <= 5)
                    {
                        if (retryInterval == 0 && r.Value.Headers.TryGetValues("Retry-After", out IEnumerable<string> outvalues))
                        {
                            string tryAfter = outvalues.FirstOrDefault() ?? "0";
                            retryInterval = int.Parse(tryAfter);
                            GraphHelper.logger.Warn($"Rate limit encountered, backoff interval of {retryInterval} found");
                        }
                        else
                        {
                            GraphHelper.logger.Warn("Rate limit encountered, but no backoff interval specified");
                        }

                        var step = content.BatchRequestSteps.FirstOrDefault(t => t.Key == r.Key);
                        stepsToRetry.Add(step.Value);
                        continue;
                    }

                    if (ignoreRefAlreadyExists && r.Value.StatusCode == HttpStatusCode.BadRequest && er.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        GraphHelper.logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                        continue;
                    }

                    if (r.Value.StatusCode == HttpStatusCode.NotFound && string.Equals(er.Error.Code, "Request_ResourceNotFound", StringComparison.OrdinalIgnoreCase))
                    {
                        var step = content.BatchRequestSteps.FirstOrDefault(t => t.Key == r.Key);
                        stepsToRetry.Add(step.Value);
                        continue;
                    }

                    exceptions.Add(new ServiceException(er.Error, r.Value.Headers, r.Value.StatusCode));
                }
            }

            if (stepsToRetry.Count > 0 && attemptCount <= 5)
            {
                BatchRequestContent newContent = new BatchRequestContent();

                foreach (var stepToRetry in stepsToRetry)
                {
                    newContent.AddBatchRequestStep(stepToRetry);
                }

                if (retryInterval == 0)
                {
                    retryInterval = 8 * attemptCount;
                }

                logger.Info($"Sleeping for {retryInterval} before retrying after attempt {attemptCount}");
                await Task.Delay(TimeSpan.FromSeconds(retryInterval), token);
                await GraphHelper.SubmitBatchContent(client, newContent, ignoreNotFound, ignoreRefAlreadyExists, token, ++attemptCount);
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

        internal static async Task SubmitAsBatches(Beta.GraphServiceClient client, List<BatchRequestStep> requests, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token)
        {
            BatchRequestContent content = new BatchRequestContent();
            int count = 0;

            foreach (BatchRequestStep r in requests)
            {
                if (count == GraphHelper.MaxJsonBatchRequests)
                {
                    await GraphHelper.SubmitBatchContent(client, content, ignoreNotFound, ignoreRefAlreadyExists, token);
                    count = 0;
                    content = new BatchRequestContent();
                }

                content.AddBatchRequestStep(r);
                count++;
            }

            if (count > 0)
            {
                await GraphHelper.SubmitBatchContent(client, content, ignoreNotFound, ignoreRefAlreadyExists, token);
            }
        }

        private static async Task SubmitBatchContent(Beta.GraphServiceClient client, BatchRequestContent content, bool ignoreNotFound, bool ignoreRefAlreadyExists, CancellationToken token, int attemptCount = 1)
        {
            BatchResponseContent response = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Batch.Request().PostAsync(content, token), token, content.BatchRequestSteps.Count + 1);

            List<Exception> exceptions = new List<Exception>();
            List<BatchRequestStep> stepsToRetry = new List<BatchRequestStep>();
            int retryInterval = 0;

            var responses = await response.GetResponsesAsync();

            foreach (KeyValuePair<string, HttpResponseMessage> r in responses)
            {
                using (r.Value)
                {
                    if (r.Value.IsSuccessStatusCode)
                    {
                        continue;
                    }

                    if (ignoreNotFound && r.Value.StatusCode == HttpStatusCode.NotFound)
                    {
                        GraphHelper.logger.Warn($"The request ({r.Key}) to remove object failed because it did not exist");
                        continue;
                    }

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

                    if (r.Value.StatusCode == (HttpStatusCode)429 && attemptCount <= 5)
                    {
                        if (retryInterval == 0 && r.Value.Headers.TryGetValues("Retry-After", out IEnumerable<string> outvalues))
                        {
                            string tryAfter = outvalues.FirstOrDefault() ?? "0";
                            retryInterval = int.Parse(tryAfter);
                            GraphHelper.logger.Warn($"Rate limit encountered, backoff interval of {retryInterval} found");
                        }
                        else
                        {
                            GraphHelper.logger.Warn("Rate limit encountered, but no backoff interval specified");
                        }

                        var step = content.BatchRequestSteps.FirstOrDefault(t => t.Key == r.Key);
                        stepsToRetry.Add(step.Value);
                        continue;
                    }

                    if (ignoreRefAlreadyExists && r.Value.StatusCode == HttpStatusCode.BadRequest && er.Error.Message.IndexOf("object references already exist", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        GraphHelper.logger.Warn($"The request ({r.Key}) to add object failed because it already exists");
                        continue;
                    }

                    if (r.Value.StatusCode == HttpStatusCode.NotFound && string.Equals(er.Error.Code, "Request_ResourceNotFound", StringComparison.OrdinalIgnoreCase))
                    {
                        var step = content.BatchRequestSteps.FirstOrDefault(t => t.Key == r.Key);
                        stepsToRetry.Add(step.Value);
                        continue;
                    }

                    exceptions.Add(new ServiceException(er.Error, r.Value.Headers, r.Value.StatusCode));
                }
            }

            if (stepsToRetry.Count > 0 && attemptCount <= 5)
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

                logger.Info($"Sleeping for {retryInterval} before retrying after attempt {attemptCount}");
                await Task.Delay(TimeSpan.FromSeconds(retryInterval), token);
                await SubmitBatchContent(client, newContent, ignoreNotFound, ignoreRefAlreadyExists, token, ++attemptCount);
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

        internal static BatchRequestStep GenerateBatchRequestStep(HttpMethod method, string id, string requestUrl)
        {
            HttpRequestMessage request = new HttpRequestMessage(method, requestUrl);
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

        public static void AssignNullToProperty(this Beta.Entity e, string name)
        {
            if (e.AdditionalData == null)
            {
                e.AdditionalData = new Dictionary<string, object>();
            }

            e.AdditionalData.Add(name, null);
        }
    }
}
