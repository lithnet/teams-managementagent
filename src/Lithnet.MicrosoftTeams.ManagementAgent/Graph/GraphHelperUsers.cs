using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Microsoft.Graph;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class GraphHelperUsers
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static async Task<string> GetUsers(GraphServiceClient client, string deltaLink, ITargetBlock<User> target, CancellationToken token)
        {
            IUserDeltaCollectionPage page = new UserDeltaCollectionPage();
            page.InitializeNextPageRequest(client, deltaLink);
            return await GetUsers(page, target, token);
        }

        public static async Task<string> GetUsers(GraphServiceClient client, ITargetBlock<User> target, CancellationToken token, params string[] selectProperties)
        {
            var request = client.Users.Delta().Request().Select(string.Join(",", selectProperties));
            return await GetUsers(request, target, token);
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
    }
}
