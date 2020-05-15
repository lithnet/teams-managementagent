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

        public static async Task<string> GetUsersWithDelta(GraphServiceClient client, string deltaLink, ITargetBlock<User> target, CancellationToken token)
        {
            IUserDeltaCollectionPage page = new UserDeltaCollectionPage();
            page.InitializeNextPageRequest(client, deltaLink);
            return await GraphHelperUsers.GetUsersWithDelta(page, target, token);
        }

        public static async Task<string> GetUsersWithDelta(GraphServiceClient client, ITargetBlock<User> target, CancellationToken token, params string[] selectProperties)
        {
            var request = client.Users.Delta().Request().Select(string.Join(",", selectProperties));
            return await GraphHelperUsers.GetUsersWithDelta(request, target, token);
        }

        public static async Task GetUsers(GraphServiceClient client, ITargetBlock<User> target, CancellationToken token, params string[] selectProperties)
        {
            var request = client.Users.Request().Select(string.Join(",", selectProperties));
            await GraphHelperUsers.GetUsers(request, target, token);
        }

        public static async Task<List<string>> GetGuestUsers(GraphServiceClient client, CancellationToken token)
        {
            var request = client.Users.Request().Filter("userType eq 'guest'");
            List<string> guests = new List<string>();

            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach (User user in page.CurrentPage)
            {
                guests.Add(user.Id);
            }

            while (page.NextPageRequest != null)
            {
                page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await page.NextPageRequest.GetAsync(token), token, 0);

                foreach (User user in page.CurrentPage)
                {
                    guests.Add(user.Id);
                }
            }

            return guests;
        }

        private static async Task<string> GetUsersWithDelta(IUserDeltaRequest request, ITargetBlock<User> target, CancellationToken token)
        {
            var page = await GraphHelper.ExecuteWithRetryAndRateLimit(async () => await request.GetAsync(token), token, 0);

            foreach (User user in page.CurrentPage)
            {
                target.Post(user);
            }

            return await GraphHelperUsers.GetUsersWithDelta(page, target, token);
        }

        private static async Task<string> GetUsersWithDelta(IUserDeltaCollectionPage page, ITargetBlock<User> target, CancellationToken token)
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
