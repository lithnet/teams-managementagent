using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal static class GraphHelper
    {
        internal static async Task<List<DirectoryObject>> GetGroupMembers(GraphServiceClient client, string groupid)
        {
            IGroupMembersCollectionWithReferencesPage result = await client.Groups[groupid].Members
                .Request()
                .GetAsync();

            return await GetMembers(result);
        }

        internal static async Task<List<DirectoryObject>> GetGroupOwners(GraphServiceClient client, string groupid)
        {
            IGroupOwnersCollectionWithReferencesPage result = await client.Groups[groupid].Owners
                .Request()
                .GetAsync();

            return await GetMembers(result);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupMembersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupMembersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await GetMembers(page);
        }

        internal static async Task<List<DirectoryObject>> GetMembers(IGroupOwnersCollectionWithReferencesRequestBuilder builder)
        {
            IGroupOwnersCollectionWithReferencesPage page = await builder.Request().Select(t => t.Id)
                .GetAsync();

            return await GetMembers(page);
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
    }
}
