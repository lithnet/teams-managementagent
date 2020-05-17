extern alias BetaLib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.MetadirectoryServices;
using System.Threading;
using Beta = BetaLib.Microsoft.Graph;
using Lithnet.Ecma2Framework;
using Lithnet.MetadirectoryServices;
using Microsoft.Graph;

namespace Lithnet.MicrosoftTeams.ManagementAgent.Tests
{
    [TestClass]
    public class TeamTests
    {
        static IExportContext context;
        static List<string> teamsToDelete;
        static TeamExportProvider teamExportProvider;

        [ClassInitialize]
        public static void InitializeTests(TestContext d)
        {
            context = new TestExportContext();
            teamsToDelete = new List<string>();

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

        [TestMethod]
        public async Task DeleteTeam()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Delete;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            await TeamTests.SubmitCSEntryChange(cs);

            try
            {
                var team = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);
                Assert.Fail("The team was not deleted");
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode != System.Net.HttpStatusCode.NotFound)
                {
                    Assert.Fail("The team was not deleted");
                }
            }
        }


        [TestMethod]
        public async Task ArchiveTeam()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));
            cs.CreateAttributeAdd("isArchived", true);
            await TeamTests.SubmitCSEntryChange(cs);

            var team = await GraphHelperTeams.GetTeam(UnitTestControl.BetaClient, teamid, CancellationToken.None);

            Assert.IsTrue(team.IsArchived ?? false);
        }

        [TestMethod]
        public async Task UnarchiveTeam()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            await GraphHelperTeams.ArchiveTeam(UnitTestControl.BetaClient, teamid, false, CancellationToken.None);

            cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));
            cs.CreateAttributeAdd("isArchived", false);
            await TeamTests.SubmitCSEntryChange(cs);

            var team = await GraphHelperTeams.GetTeam(UnitTestControl.BetaClient, teamid, CancellationToken.None);

            Assert.IsFalse(team.IsArchived ?? false);
        }

        [TestMethod]
        public async Task CreateTeamFromEduTemplate()
        {
            string template = "https://graph.microsoft.com/beta/teamsTemplates('educationClass')";

            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());
            cs.CreateAttributeAdd("template",template);

            string teamid = await TeamTests.SubmitCSEntryChange(cs);
            var team = await GraphHelperTeams.GetTeam(UnitTestControl.BetaClient, teamid, CancellationToken.None);

            //Assert.AreEqual(template, team.Template);
            // Template information is not being returned from the API, so the following check looks for an attribute only present on classroom templates
            Assert.IsNotNull(team.AdditionalData);
            Assert.IsTrue(team.AdditionalData.ContainsKey("classSettings"));
        }

        [TestMethod]
        public async Task CreateTeamTestAllowedPermissions()
        {
            string mailNickname = $"ut-{Guid.NewGuid()}";

            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("description", "my description");
            cs.CreateAttributeAdd("mailNickname", mailNickname);
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateChannels", true);
            cs.CreateAttributeAdd("memberSettings_allowDeleteChannels", true);
            cs.CreateAttributeAdd("memberSettings_allowAddRemoveApps", true);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveTabs", true);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveConnectors", true);
            cs.CreateAttributeAdd("guestSettings_allowCreateUpdateChannels", true);
            cs.CreateAttributeAdd("guestSettings_allowDeleteChannels", true);
            cs.CreateAttributeAdd("messagingSettings_allowUserEditMessages", true);
            cs.CreateAttributeAdd("messagingSettings_allowUserDeleteMessages", true);
            cs.CreateAttributeAdd("messagingSettings_allowOwnerDeleteMessages", true);
            cs.CreateAttributeAdd("messagingSettings_allowTeamMentions", true);
            cs.CreateAttributeAdd("messagingSettings_allowChannelMentions", true);
            cs.CreateAttributeAdd("funSettings_allowGiphy", true);
            cs.CreateAttributeAdd("funSettings_giphyContentRating", "Strict");
            cs.CreateAttributeAdd("funSettings_allowStickersAndMemes", true);
            cs.CreateAttributeAdd("funSettings_allowCustomMemes", true);
            cs.CreateAttributeAdd("visibility", "Private");

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            var team = await GraphHelperTeams.GetTeam(UnitTestControl.Client, teamid, CancellationToken.None);
            var group = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual("mytestteam", group.DisplayName);
            Assert.AreEqual("my description", group.Description);
            Assert.AreEqual(mailNickname, group.MailNickname);
            Assert.AreEqual("Private", group.Visibility, true);

            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(true, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.GuestSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserEditMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(true, team.FunSettings.AllowGiphy);
            Assert.AreEqual(GiphyRatingType.Strict, team.FunSettings.GiphyContentRating);
            Assert.AreEqual(true, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(true, team.FunSettings.AllowCustomMemes);
        }

        [TestMethod]
        public async Task CreateTeamTestDeniedPermissions()
        {
            string mailNickname = $"ut-{Guid.NewGuid()}";

            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("description", "my description");
            cs.CreateAttributeAdd("mailNickname", mailNickname);
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateChannels", false);
            cs.CreateAttributeAdd("memberSettings_allowDeleteChannels", false);
            cs.CreateAttributeAdd("memberSettings_allowAddRemoveApps", false);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveTabs", false);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveConnectors", false);
            cs.CreateAttributeAdd("guestSettings_allowCreateUpdateChannels", false);
            cs.CreateAttributeAdd("guestSettings_allowDeleteChannels", false);
            cs.CreateAttributeAdd("messagingSettings_allowUserEditMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowUserDeleteMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowOwnerDeleteMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowTeamMentions", false);
            cs.CreateAttributeAdd("messagingSettings_allowChannelMentions", false);
            cs.CreateAttributeAdd("funSettings_allowGiphy", false);
            cs.CreateAttributeAdd("funSettings_giphyContentRating", "Moderate");
            cs.CreateAttributeAdd("funSettings_allowStickersAndMemes", false);
            cs.CreateAttributeAdd("funSettings_allowCustomMemes", false);
            cs.CreateAttributeAdd("visibility", "Public");

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            var team = await GraphHelperTeams.GetTeam(UnitTestControl.Client, teamid, CancellationToken.None);
            var group = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual("mytestteam", group.DisplayName);
            Assert.AreEqual("my description", group.Description);
            Assert.AreEqual(mailNickname, group.MailNickname);
            Assert.AreEqual("Public", group.Visibility, true);

            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, team.MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(false, team.MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(false, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, team.GuestSettings.AllowDeleteChannels);
            Assert.AreEqual(false, team.MessagingSettings.AllowUserEditMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(false, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(false, team.FunSettings.AllowGiphy);
            Assert.AreEqual(GiphyRatingType.Moderate, team.FunSettings.GiphyContentRating);
            Assert.AreEqual(false, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(false, team.FunSettings.AllowCustomMemes);
        }

        [TestMethod]
        public async Task CreateTeamTestUpdatePermissions()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", "mytestteam");
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateChannels", false);
            cs.CreateAttributeAdd("memberSettings_allowDeleteChannels", false);
            cs.CreateAttributeAdd("memberSettings_allowAddRemoveApps", false);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveTabs", false);
            cs.CreateAttributeAdd("memberSettings_allowCreateUpdateRemoveConnectors", false);
            cs.CreateAttributeAdd("guestSettings_allowCreateUpdateChannels", false);
            cs.CreateAttributeAdd("guestSettings_allowDeleteChannels", false);
            cs.CreateAttributeAdd("messagingSettings_allowUserEditMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowUserDeleteMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowOwnerDeleteMessages", false);
            cs.CreateAttributeAdd("messagingSettings_allowTeamMentions", false);
            cs.CreateAttributeAdd("messagingSettings_allowChannelMentions", false);
            cs.CreateAttributeAdd("funSettings_allowGiphy", false);
            cs.CreateAttributeAdd("funSettings_giphyContentRating", "Moderate");
            cs.CreateAttributeAdd("funSettings_allowStickersAndMemes", false);
            cs.CreateAttributeAdd("funSettings_allowCustomMemes", false);

            string teamid = await TeamTests.SubmitCSEntryChange(cs);

            var team = await GraphHelperTeams.GetTeam(UnitTestControl.Client, teamid, CancellationToken.None);
            var group = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual("mytestteam", group.DisplayName);
            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, team.MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(false, team.MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(false, team.MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(false, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, team.GuestSettings.AllowDeleteChannels);
            Assert.AreEqual(false, team.MessagingSettings.AllowUserEditMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(false, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(false, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(false, team.FunSettings.AllowGiphy);
            Assert.AreEqual(GiphyRatingType.Moderate, team.FunSettings.GiphyContentRating);
            Assert.AreEqual(false, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(false, team.FunSettings.AllowCustomMemes);

            cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));
            cs.CreateAttributeReplace("memberSettings_allowCreateUpdateChannels", true);
            cs.CreateAttributeReplace("memberSettings_allowDeleteChannels", true);
            cs.CreateAttributeReplace("memberSettings_allowAddRemoveApps", true);
            cs.CreateAttributeReplace("memberSettings_allowCreateUpdateRemoveTabs", true);
            cs.CreateAttributeReplace("memberSettings_allowCreateUpdateRemoveConnectors", true);
            cs.CreateAttributeReplace("guestSettings_allowCreateUpdateChannels", true);
            cs.CreateAttributeReplace("guestSettings_allowDeleteChannels", true);
            cs.CreateAttributeReplace("messagingSettings_allowUserEditMessages", true);
            cs.CreateAttributeReplace("messagingSettings_allowUserDeleteMessages", true);
            cs.CreateAttributeReplace("messagingSettings_allowOwnerDeleteMessages", true);
            cs.CreateAttributeReplace("messagingSettings_allowTeamMentions", true);
            cs.CreateAttributeReplace("messagingSettings_allowChannelMentions", true);
            cs.CreateAttributeReplace("funSettings_allowGiphy", true);
            cs.CreateAttributeReplace("funSettings_giphyContentRating", "Strict");
            cs.CreateAttributeReplace("funSettings_allowStickersAndMemes", true);
            cs.CreateAttributeReplace("funSettings_allowCustomMemes", true);

            await TeamTests.SubmitCSEntryChange(cs);

            team = await GraphHelperTeams.GetTeam(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(true, team.MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(true, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.GuestSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserEditMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(true, team.FunSettings.AllowGiphy);
            Assert.AreEqual(GiphyRatingType.Strict, team.FunSettings.GiphyContentRating);
            Assert.AreEqual(true, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(true, team.FunSettings.AllowCustomMemes);
        }

        [TestMethod]
        public async Task ThrowOnVisibilityModification()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeAdd("visibility", "Public");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (InitialFlowAttributeModificationException)
            {
            }
        }

        [TestMethod]
        public async Task ThrowOnTemplateModification()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeAdd("template", "xxx");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (InitialFlowAttributeModificationException)
            {
            }
        }

        [TestMethod]
        public async Task ThrowOnIsArchivedDelete()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeDelete("isArchived");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (UnsupportedBooleanAttributeDeleteException)
            {
            }
        }

        [TestMethod]
        public async Task ThrowOnTemplateDelete()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeDelete("template");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (InitialFlowAttributeModificationException)
            {
            }
        }

        private static async Task<string> SubmitCSEntryChange(CSEntryChange cs)
        {
            string teamid;

            CSEntryChangeResult teamResult = await TeamTests.teamExportProvider.PutCSEntryChangeAsync(cs);
            if (cs.ObjectModificationType == ObjectModificationType.Add)
            {
                TeamTests.AddTeamToCleanupTask(teamResult);
                teamid = teamResult.GetAnchorValueOrDefault<string>("id");
            }
            else
            {
                teamid = cs.GetAnchorValueOrDefault<string>("id");
            }

            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            return teamid;
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