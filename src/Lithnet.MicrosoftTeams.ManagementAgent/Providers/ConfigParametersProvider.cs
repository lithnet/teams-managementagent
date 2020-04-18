using System.Collections.Generic;
using System.Collections.ObjectModel;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    public class ConfigParametersProvider : IConfigParametersProvider
    {
        public IList<ConfigParameterDefinition> GetConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            List<ConfigParameterDefinition> configParametersDefinitions = new List<ConfigParameterDefinition>();

            switch (page)
            {
                case ConfigParameterPage.Connectivity:
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.TenantDomain, string.Empty));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.ClientId, string.Empty));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateEncryptedStringParameter(ConfigParameterNames.Secret, string.Empty));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.LogFileName, string.Empty));
                    break;

                case ConfigParameterPage.Global:
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter(ConfigParameterNames.FilterQuery, string.Empty));
                    break;

                case ConfigParameterPage.Partition:
                    break;

                case ConfigParameterPage.RunStep:
                    break;
            }

            return configParametersDefinitions;
        }

        public ParameterValidationResult ValidateConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            return new ParameterValidationResult();
        }
    }
}
