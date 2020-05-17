using System.Collections.Generic;
using System.Collections.ObjectModel;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class ConfigParametersProvider : IConfigParametersProvider
    {
        public void GetConfigParameters(KeyedCollection<string, ConfigParameter> existingParameters, IList<ConfigParameterDefinition> newDefinitions,  ConfigParameterPage page)
        {
            switch (page)
            {
                case ConfigParameterPage.Connectivity:
                    newDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.TenantDomain, string.Empty));
                    newDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.ClientId, string.Empty));
                    newDefinitions.Add(ConfigParameterDefinition.CreateEncryptedStringParameter(ConfigParameterNames.Secret, string.Empty));
                    break;

                case ConfigParameterPage.Global:
                    newDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter(ConfigParameterNames.FilterQueryDescription));
                    newDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.FilterQuery, string.Empty));
                    newDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    newDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter(ConfigParameterNames.ChannelNameFilterDescription));
                    newDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.ChannelNameFilter, string.Empty));
                    newDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    newDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter(ConfigParameterNames.RandomizeChannelNameOnDeleteDecription));
                    newDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(ConfigParameterNames.RandomizeChannelNameOnDelete, false));
                    newDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    newDefinitions.Add(ConfigParameterDefinition.CreateLabelParameter(ConfigParameterNames.UsersToIgnoreDescription));
                    newDefinitions.Add(ConfigParameterDefinition.CreateTextParameter(ConfigParameterNames.UsersToIgnore));
                    newDefinitions.Add(ConfigParameterDefinition.CreateDividerParameter());
                    newDefinitions.Add(ConfigParameterDefinition.CreateCheckBoxParameter(ConfigParameterNames.SetSpoReadOnly, false));
                    break;

                case ConfigParameterPage.Partition:
                    break;

                case ConfigParameterPage.RunStep:
                    break;
            }
        }

        public ParameterValidationResult ValidateConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            return new ParameterValidationResult();
        }
    }
}
