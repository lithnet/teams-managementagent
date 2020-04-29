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
    internal static class GraphHelperTeams
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

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

        public static async Task<Team> CreateTeamFromGroup(GraphServiceClient client, string groupid, Team team, CancellationToken token)
        {
            return await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Groups[groupid].Team.Request().PutAsync(team, token), token, 1);
        }

        public static async Task<string> CreateTeam(Beta.GraphServiceClient client, Beta.Team team, CancellationToken token)
        {
            TeamsAsyncOperation result;

            string location = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await RequestCreateTeam(client, team, token), token, 1);

            result = await GraphHelperTeams.WaitForTeamsAsyncOperation(client, token, location);

            if (result.Status == TeamsAsyncOperationStatus.Succeeded)
            {
                return result.TargetResourceId;
            }

            string serializedResponse = JsonConvert.SerializeObject(result);
            logger.Error($"Team creation failed\r\n{serializedResponse}");

            throw new ServiceException(new Error() { Code = result.Error.Code, AdditionalData = result.Error.AdditionalData, Message = result.Error.Message });
        }

        private static async Task<TeamsAsyncOperation> WaitForTeamsAsyncOperation(Beta.GraphServiceClient client, CancellationToken token, string location)
        {
            TeamsAsyncOperation result;
            int waitCount = 0;

            do
            {
                await Task.Delay(TimeSpan.FromSeconds(5 * waitCount), token);
                waitCount++;

                var b = new TeamsAsyncOperationRequestBuilder($"{client.BaseUrl}{location}", client);

                result = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await b.Request().GetAsync(token), token, 1);
                GraphHelperTeams.logger.Trace($"Result of async operation is {result.Status}: Count : {waitCount}");
            } while (result.Status == TeamsAsyncOperationStatus.InProgress || result.Status == TeamsAsyncOperationStatus.NotStarted);

            return result;
        }

        private static async Task<string> RequestCreateTeam(Beta.GraphServiceClient client, Beta.Team team, CancellationToken token)
        {
            var message = client.Teams.Request().GetHttpRequestMessage();
            message.Method = HttpMethod.Post;
            message.Content = new StringContent(JsonConvert.SerializeObject(team), Encoding.UTF8, "application/json");

            using (var response = await client.HttpProvider.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, token))
            {
                if (response.StatusCode == HttpStatusCode.Accepted)
                {
                    var location = response.Headers.Location;

                    if (location == null)
                    {
                        throw new InvalidOperationException("The location header was null");
                    }

                    logger.Trace($"Create operation was successful at {location}");
                    return location.ToString();
                }

                logger.Trace($"{response.StatusCode}");

                throw new InvalidOperationException("The request was not successful");
            }
        }

        public static async Task UpdateTeam(GraphServiceClient client, string id, Team team, CancellationToken token)
        {
            await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await client.Teams[id].Request().UpdateAsync(team, token), token, 1);
        }

    }
}
