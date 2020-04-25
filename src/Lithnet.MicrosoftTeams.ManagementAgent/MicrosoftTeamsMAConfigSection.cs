using System.Configuration;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class MicrosoftTeamsMAConfigSection : ConfigurationSection
    {
        private const string SectionName = "lithnet-microsoftteams-ma";
        private const string PropExportThreads = "export-threads";
        private const string PropImportThreads = "import-threads";
        private const string PropConnectionLimit = "connection-limit";
        private const string PropPostGroupCreateDelay = "post-group-create-delay";
        private const string PropRateLimitRequestWindow = "rate-limit-request-window-seconds";
        private const string PropRateLimitRequestLimit = "rate-limit-request-limit";
        private const string PropDeleteAddConflictingGroup = "delete-add-conflicting-group";

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

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropExportThreads, IsRequired = false, DefaultValue = 5)]
        public int ExportThreads => (int)this[MicrosoftTeamsMAConfigSection.PropExportThreads];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropImportThreads, IsRequired = false, DefaultValue = 10)]
        public int ImportThreads => (int)this[MicrosoftTeamsMAConfigSection.PropImportThreads];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropConnectionLimit, IsRequired = false, DefaultValue = 1000)]
        public int ConnectionLimit => (int)this[MicrosoftTeamsMAConfigSection.PropConnectionLimit];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropPostGroupCreateDelay, IsRequired = false, DefaultValue = 8)]
        public int PostGroupCreateDelay => (int)this[MicrosoftTeamsMAConfigSection.PropPostGroupCreateDelay];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropRateLimitRequestWindow, IsRequired = false, DefaultValue = 150)]
        public int RateLimitRequestWindowSeconds => (int)this[MicrosoftTeamsMAConfigSection.PropRateLimitRequestWindow];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropRateLimitRequestLimit, IsRequired = false, DefaultValue = 3000)]
        public int RateLimitRequestLimit => (int)this[MicrosoftTeamsMAConfigSection.PropRateLimitRequestLimit];

        [ConfigurationProperty(MicrosoftTeamsMAConfigSection.PropDeleteAddConflictingGroup, IsRequired = false, DefaultValue = false)]
        public bool DeleteAddConflictingGroup => (bool)this[MicrosoftTeamsMAConfigSection.PropDeleteAddConflictingGroup];
    }
}