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

        [ClassInitialize]
        public static void InitializeTests(TestContext d)
        {
            context = new TestExportContext();
        }

        [TestMethod()]
        public async Task PutCSEntryChangeAsyncTest()
        {
            string teamid = null;
            string channelid = null;

            try
            {
                ChannelExportProvider channelExportProvider = new ChannelExportProvider();
                channelExportProvider.Initialize(context);

                TeamExportProvider teamExportProvider = new TeamExportProvider();
                teamExportProvider.Initialize(context);

                CSEntryChange team = CSEntryChange.Create();
                team.ObjectType = "team";
                team.ObjectModificationType = ObjectModificationType.Add;
                team.CreateAttributeAdd("displayName", "mytestteam");
                team.CreateAttributeAdd("member", UnitTestControl.Users.Take(20).Select(t => t.Id).ToList<object>());
                team.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(20, 20).Select(t => t.Id).ToList<object>());

                var teamResult = await teamExportProvider.PutCSEntryChangeAsync(team);

                Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

                teamid = teamResult.AnchorAttributes[0].GetValueAdd<string>();

                CSEntryChange channel = CSEntryChange.Create();
                channel.ObjectType = "publiChannel";
                channel.ObjectModificationType = ObjectModificationType.Add;
                channel.CreateAttributeAdd("team", teamid);
                channel.CreateAttributeAdd("displayName", "My test channel");

                var channelResult = await channelExportProvider.PutCSEntryChangeAsync(channel);
                channelid = channelResult.AnchorAttributes["id"].GetValueAdd<string>();

                Assert.IsTrue(channelResult.ErrorCode == MAExportError.Success);
            }
            finally
            {
                if (teamid != null)
                {
                    await GraphHelperGroups.DeleteGroup(UnitTestControl.Client, teamid, context.Token);
                }
            }
        }
    }
}