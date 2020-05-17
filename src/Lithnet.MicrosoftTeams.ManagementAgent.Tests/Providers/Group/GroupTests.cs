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
    public class GroupTests
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
        public async Task UpdateGroupAttributes()
        {
            string displayName1 = "displayName 1";
            string description1 = "description 1";
            string mailNickname1 = $"ut-{Guid.NewGuid()}";

            string displayName2 = "displayName 2";
            string description2 = "description 2";
            string mailNickname2 = $"ut-{Guid.NewGuid()}";

            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", displayName1);
            cs.CreateAttributeAdd("description", description1);
            cs.CreateAttributeAdd("mailNickname", mailNickname1);
            cs.CreateAttributeAdd("owner", UnitTestControl.Users.GetRange(0, 1).Select(t => t.Id).ToList<object>());

            string teamid = await SubmitCSEntryChange(cs);
            await Task.Delay(TimeSpan.FromSeconds(30));

            var team = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual(displayName1, team.DisplayName);
            Assert.AreEqual(description1, team.Description);
            Assert.AreEqual(mailNickname1, team.MailNickname);

            cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.CreateAttributeReplace("displayName", displayName2);
            cs.CreateAttributeReplace("description", description2);
            cs.CreateAttributeReplace("mailNickname", mailNickname2);
            await SubmitCSEntryChange(cs);

            team = await GraphHelperGroups.GetGroup(UnitTestControl.Client, teamid, CancellationToken.None);

            Assert.AreEqual(displayName2, team.DisplayName);
            Assert.AreEqual(description2, team.Description);
            Assert.AreEqual(mailNickname2, team.MailNickname);
        }

        [TestMethod]
        public async Task ThrowOnDisplayNameDelete()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeDelete("displayName");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (UnsupportedAttributeDeleteException)
            {
            }
        }

        [TestMethod]
        public async Task ThrowOnMailNicknameDelete()
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", "1234"));
            cs.CreateAttributeDelete("mailNickname");

            try
            {
                await teamExportProvider.PutCSEntryChangeAsync(cs);
                Assert.Fail("The expected exception was not thrown");
            }
            catch (UnsupportedAttributeDeleteException)
            {
            }
        }

        [TestMethod]
        public async Task CreateTeamWithMembers()
        {
            List<string> members = UnitTestControl.Users.Take(50).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(50, 10).Select(t => t.Id).ToList();

            await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);
        }

        [TestMethod]
        public async Task UpdateTeamMembers()
        {
            List<string> members = UnitTestControl.Users.Take(50).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(50, 10).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            var toDelete = members.GetRange(0, 10);
            var toAdd = UnitTestControl.Users.GetRange(60, 10).Select(t => t.Id).ToList();

            cs.CreateAttributeUpdate("member", toAdd.ToList<object>(), toDelete.ToList<object>());

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var expected = members.Concat(toAdd).ToList();
            expected = expected.Except(toDelete).ToList();

            var actual = (await GraphHelperGroups.GetGroupMembers(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            CollectionAssert.AreEquivalent(expected, actual);
        }

        [TestMethod]
        public async Task UpdateTeamOwners()
        {
            List<string> members = UnitTestControl.Users.Take(1).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(1, 50).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            var toDelete = owners.GetRange(0, 10);
            var toAdd = UnitTestControl.Users.GetRange(51, 10).Select(t => t.Id).ToList();

            cs.CreateAttributeUpdate("owner", toAdd.ToList<object>(), toDelete.ToList<object>());

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var expected = owners.Concat(toAdd).ToList();
            expected = expected.Except(toDelete).ToList();

            var actual = (await GraphHelperGroups.GetGroupOwners(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            CollectionAssert.AreEquivalent(expected, actual);
        }

        [TestMethod]
        public async Task ReplaceTeamMembers()
        {
            List<string> members = UnitTestControl.Users.Take(100).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(100, 1).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            var membersToDelete = members;
            var membersToAdd = UnitTestControl.Users.GetRange(101, 100).Select(t => t.Id).ToList();

            cs.CreateAttributeUpdate("member", membersToAdd.ToList<object>(), membersToDelete.ToList<object>());

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var expected = members.Concat(membersToAdd).ToList();
            expected = expected.Except(membersToDelete).ToList();

            var actual = (await GraphHelperGroups.GetGroupMembers(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            CollectionAssert.AreEquivalent(expected, actual);
        }

        [TestMethod]
        public async Task ReplaceTeamOwners()
        {
            List<string> members = UnitTestControl.Users.Take(1).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(1, 100).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            var ownersToDelete = owners;
            var ownersToAdd = UnitTestControl.Users.GetRange(101, 100).Select(t => t.Id).ToList();

            cs.CreateAttributeUpdate("owner", ownersToAdd.ToList<object>(), ownersToDelete.ToList<object>());

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var actual = (await GraphHelperGroups.GetGroupOwners(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            CollectionAssert.AreEquivalent(ownersToAdd, actual);
        }

        [TestMethod]
        public async Task DeleteTeamOwners()
        {
            List<string> members = UnitTestControl.Users.Take(1).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(1, 100).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));
            cs.CreateAttributeDelete("owner");

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var actual = (await GraphHelperGroups.GetGroupOwners(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            Assert.AreEqual(0, actual.Count);
        }

        [TestMethod]
        public async Task DeleteTeamMembers()
        {
            List<string> members = UnitTestControl.Users.Take(100).Select(t => t.Id).ToList();
            List<string> owners = UnitTestControl.Users.GetRange(100, 1).Select(t => t.Id).ToList();

            string teamid = await GroupTests.CreateAndValidateTeam("mytestteam", members, owners);

            var cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Update;
            cs.AnchorAttributes.Add(AnchorAttribute.Create("id", teamid));

            cs.CreateAttributeDelete("member");

            var teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            Assert.IsTrue(teamResult.ErrorCode == MAExportError.Success);

            var actual = (await GraphHelperGroups.GetGroupMembers(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();

            Assert.AreEqual(0, actual.Count);
        }

        private static async Task<string> CreateAndValidateTeam(string displayName, int memberCount, int ownerCount)
        {
            List<string> members = null;
            if (memberCount > 0)
            {
                members = UnitTestControl.Users.Take(memberCount).Select(t => t.Id).ToList();
            }

            List<string> owners = null;
            if (ownerCount > 0)
            {
                owners = UnitTestControl.Users.GetRange(memberCount, ownerCount).Select(t => t.Id).ToList();
            }

            return await GroupTests.CreateAndValidateTeam(displayName, members, owners);
        }

        private static async Task<string> CreateAndValidateTeam(string displayName, List<string> members = null, List<string> owners = null)
        {
            CSEntryChange cs = CSEntryChange.Create();
            cs.ObjectType = "team";
            cs.ObjectModificationType = ObjectModificationType.Add;
            cs.CreateAttributeAdd("displayName", displayName);

            if (members != null && members.Count > 0)
            {
                cs.CreateAttributeAdd("member", members.ToList<object>());
            }

            if (owners != null && owners.Count > 0)
            {
                cs.CreateAttributeAdd("owner", owners.ToList<object>());
            }

            string teamid = await SubmitCSEntryChange(cs);

            if (members != null)
            {
                var actualMembers = (await GraphHelperGroups.GetGroupMembers(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();
                CollectionAssert.AreEquivalent(members, actualMembers);
            }

            if (owners != null)
            {
                var actualOwners = (await GraphHelperGroups.GetGroupOwners(UnitTestControl.Client, teamid, CancellationToken.None)).Select(t => t.Id).ToList();
                CollectionAssert.AreEquivalent(owners, actualOwners);
            }

            return teamid;
        }

        private static async Task<string> SubmitCSEntryChange(CSEntryChange cs)
        {
            string teamid;

            CSEntryChangeResult teamResult = await teamExportProvider.PutCSEntryChangeAsync(cs);
            if (cs.ObjectModificationType == ObjectModificationType.Add)
            {
                AddTeamToCleanupTask(teamResult);
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