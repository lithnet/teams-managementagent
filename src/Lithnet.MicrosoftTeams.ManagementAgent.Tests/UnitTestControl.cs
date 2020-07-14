extern alias BetaLib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using Lithnet.Ecma2Framework;
using Microsoft.Graph;
using Microsoft.MetadirectoryServices;
using Newtonsoft.Json;
using NLog;
using NLog.Config;
using NLog.Targets;
using Beta = BetaLib.Microsoft.Graph;

namespace Lithnet.MicrosoftTeams.ManagementAgent.Tests
{
    internal static class UnitTestControl
    {
        private static List<User> users;

        public static GraphConnectionContext GraphConnectionContext { get; }

        public static GraphServiceClient Client { get; }

        public static Beta.GraphServiceClient BetaClient { get; }

        public static KeyedCollection<string, ConfigParameter> ConfigParameters { get; }

        public static List<User> Users
        {
            get
            {
                if (users == null)
                {
                    users = GetUserList().GetAwaiter().GetResult();
                }

                return users;
            }
        }

        static UnitTestControl()
        {
            ConfigParameters = new ConfigurationParameters();
            GraphConnectionContext = GraphConnectionContext.GetConnectionContext(ConfigParameters);
            Client = GraphConnectionContext.Client;
            BetaClient = GraphConnectionContext.BetaClient;
            SetupLogging();
        }

        private static void SetupLogging()
        {
            LoggingConfiguration logConfiguration = new LoggingConfiguration();

            LogLevel level = LogLevel.Trace;

            TraceTarget ttarget = new TraceTarget();
            logConfiguration.AddTarget("tt", ttarget);
            LoggingRule ttRule = new LoggingRule("*", level, ttarget);
            logConfiguration.LoggingRules.Add(ttRule);

            FileTarget fileTarget = new FileTarget();
            logConfiguration.AddTarget("file", fileTarget);
            fileTarget.FileName = "debug-test.log";
            fileTarget.Layout = "${longdate}|[${threadid:padding=4}]|${level:uppercase=true:padding=5}|${message}${exception:format=ToString}";
            fileTarget.ArchiveEvery = FileArchivePeriod.Day;
            fileTarget.ArchiveNumbering = ArchiveNumberingMode.Date;
            fileTarget.MaxArchiveFiles = 7;

            LoggingRule rule2 = new LoggingRule("*", level, fileTarget);
            logConfiguration.LoggingRules.Add(rule2);

            LogManager.Configuration = logConfiguration;
            LogManager.ReconfigExistingLoggers();
        }

        private static async Task<List<User>> GetUserList()
        {
            BufferBlock<User> target = new BufferBlock<User>();
            List<User> newList = new List<User>();

            ActionBlock<User> action = new ActionBlock<User>(user =>
            {
                newList.Add(user);
            });

            target.LinkTo(action, new DataflowLinkOptions() { PropagateCompletion = true });

            await GraphHelperUsers.GetUsers(Client, target, CancellationToken.None, "displayName", "onPremisesSamAccountName", "id", "userPrincipalName");
            target.Complete();

            await action.Completion;

            return newList;
        }
    }
}
