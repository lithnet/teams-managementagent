using System.Configuration;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class MicrosoftTeamsMAConfigSection : ConfigurationSection
    {
        private const string SectionName = "lithnet-microsoftteams-ma";
        private const string PropHttpDebugEnabled = "http-debug-enabled";
        private const string PropProxyUrl = "proxy-url";
        private const string PropExportThreads = "export-threads";
        private const string PropImportThreads = "import-threads";
        private const string PropUserListPageSize = "user-list-page-size";
        private const string PropGroupListPageSize = "group-list-page-size";
        private const string PropConnectionLimit = "connection-limit";
        private const string PropHttpClientTimeout = "http-client-timeout";

        internal static MicrosoftTeamsMAConfigSection GetConfiguration()
        {
            MicrosoftTeamsMAConfigSection section = (MicrosoftTeamsMAConfigSection)ConfigurationManager.GetSection(MicrosoftTeamsMAConfigSection.SectionName);

            if (section == null)
            {
                section = new MicrosoftTeamsMAConfigSection();
            }

            return section;
        }

        internal static MicrosoftTeamsMAConfigSection Configuration { get; private set; }

        static MicrosoftTeamsMAConfigSection()
        {
            MicrosoftTeamsMAConfigSection.Configuration = MicrosoftTeamsMAConfigSection.GetConfiguration();
        }

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropHttpDebugEnabled, IsRequired = false, DefaultValue = false)]
        public bool HttpDebugEnabled => (bool)this[MicrosoftTeamsMAConfigSection.PropHttpDebugEnabled];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropProxyUrl, IsRequired = false, DefaultValue = null)]
        public string ProxyUrl => (string)this[MicrosoftTeamsMAConfigSection.PropProxyUrl];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropExportThreads, IsRequired = false, DefaultValue = 1)]
        public int ExportThreads => (int)this[MicrosoftTeamsMAConfigSection.PropExportThreads];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropImportThreads, IsRequired = false, DefaultValue = 10)]
        public int ImportThreads => (int)this[MicrosoftTeamsMAConfigSection.PropImportThreads];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropConnectionLimit, IsRequired = false, DefaultValue = 1000)]
        public int ConnectionLimit => (int)this[MicrosoftTeamsMAConfigSection.PropConnectionLimit];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropHttpClientTimeout, IsRequired = false, DefaultValue = 120)]
        public int HttpClientTimeout => (int)this[MicrosoftTeamsMAConfigSection.PropHttpClientTimeout];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropUserListPageSize, IsRequired = false, DefaultValue = 200)]
        public int UserListPageSize => (int)this[MicrosoftTeamsMAConfigSection.PropUserListPageSize];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropGroupListPageSize, IsRequired = false, DefaultValue = -1)]
        public int GroupListPageSize => (int)this[MicrosoftTeamsMAConfigSection.PropGroupListPageSize];
    }
}