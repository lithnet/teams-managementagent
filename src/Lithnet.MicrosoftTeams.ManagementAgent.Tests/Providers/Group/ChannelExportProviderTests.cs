extern alias BetaLib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Lithnet.MicrosoftTeams.ManagementAgent;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using Microsoft.MetadirectoryServices.DetachedObjectModel;
using System.Threading;
using Beta = BetaLib.Microsoft.Graph;
using Microsoft.Graph;
using Newtonsoft.Json;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent.Tests
{
    [TestClass()]
    public class ChannelExportProviderTests
    {
        static IExportContext context;
        static List<string> teamsToDelete;
        static ChannelExportProvider channelExportProvider;
        static TeamExportProvider teamExportProvider;


        [ClassInitialize]
        public static void InitializeTests(TestContext d)
        {
            context = new TestExportContext();
            teamsToDelete = new List<string>();

            channelExportProvider = new ChannelExportProvider();
            channelExportProvider.Initialize(context);

            teamExportProvider = new TeamExportProvider();
            teamExportProvider.Initialize(context);
        }

        [ClassCleanup]
        public static async Task Cleanup()
        {
            List<Exception> exceptions = new List<Exception>();

            foreach (string teamid in teamsToDelete)
            {
                try
                {
                    await GraphHelperGroups.DeleteGroup(UnitTestControl.Client, teamid, context.Token);
                }
                catch (Exception ex)
                {
                    exceptions.Add(ex);
                }
            }

            if (exceptions.Count > 0)
            {
                throw new AggregateException("One or more groups could not be cleaned up", exceptions);
            }
        }

        [TestMethod()]
        public async Task CreatePublicChannelTest()
        {
            string teamid = await CreateTeamWithMembers();

            CSEntryChange channel = CSEntryChange.Create();
            channel.ObjectType = "publicChannel";
            channel.ObjectModificationType = ObjectModificationType.Add;
            channel.CreateAttributeAdd("team", teamid);
            channel.CreateAttributeAdd("displayName", "my test channel");
            channel.CreateAttributeAdd("description", "my description");
            channel.CreateAttributeAdd("isFavoriteByDefault", true);

            CSEntryChangeResult channelResult = await channelExportProvider.PutCSEntryChangeAsync(channel);
            string channelid = channelResult.GetAnchorValueOrDefault<string>("id");
            Assert.IsTrue(channelResult.ErrorCode == MAExportError.Success);

            Beta.Channel newChannel = await GetChannelFromTeam(teamid, channelid);

            Assert.AreEqual("my test channel", newChannel.DisplayName);
            Assert.AreEqual("my description", newChannel.Description);
            Assert.AreEqual(true, newChannel.IsFavoriteByDefault);
        }

        private static async Task<string> CreateTeamWithMembers()
        {
            CSEntryChange team = CSEntryChange.Create();
            team.ObjectType = "team";
            team.ObjectModificationType = ObjectModificationType.Add;
            team.CreateAttributeAdd("displayName", "mytestteam");
            team.CreateAttributeAdd("member", UnitTestControl.Users.Take(20).Select(t => t.Id).ToList<object>());
            team.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(20, 20).Select(t => t.Id).ToList<object>());

            CSEntryChangeResult teamResult = await teamExportProvider.PutCSEntryChangeAsync(team);
            AddTeamToCleanupTask(teamResult);

            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            return teamResult.GetAnchorValueOrDefault<string>("id");
        }

        private static async Task<Beta.Channel> GetChannelFromTeam(string teamid, string channelid)
        {
            List<Beta.Channel> channels = await GraphHelperTeams.GetChannels(UnitTestControl.BetaClient, teamid, CancellationToken.None);
            return channels.First(t => t.Id == channelid);
        }

        private static void AddTeamToCleanupTask(CSEntryChangeResult c)
        {
            string teamid = c.GetAnchorValueOrDefault<string>("id");
            AddTeamToCleanupTask(teamid);
        }

        private static void AddTeamToCleanupTask(string id)
        {
            if (!string.IsNullOrWhiteSpace(id))
            {
                teamsToDelete.Add(id);
            }
        }
    }
}