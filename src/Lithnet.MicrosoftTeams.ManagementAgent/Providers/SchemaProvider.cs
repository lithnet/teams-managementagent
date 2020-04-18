using System;
using System.Collections.Generic;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class SchemaProvider : ISchemaProvider
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        internal static HashSet<string> GroupProperties = new HashSet<string>();
        internal static HashSet<string> GroupMemberProperties = new HashSet<string>();
        internal static HashSet<string> TeamsProperties = new HashSet<string>();

        public Schema GetMmsSchema(SchemaContext context)
        {
            Schema mmsSchema = new Schema();
            SchemaType mmsType = SchemaProvider.GetSchemaTypeUser();
            mmsSchema.Types.Add(mmsType);

            mmsType = SchemaProvider.GetSchemaTypeGroup();
            mmsSchema.Types.Add(mmsType);

            return mmsSchema;
        }

        private static SchemaType GetSchemaTypeUser()
        {
            SchemaType mmsType = SchemaType.Create("user", true);
            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("onPremisesSamAccountName", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("upn", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            return mmsType;
        }

        private static SchemaType GetSchemaTypeGroup()
        {
            SchemaType mmsType = SchemaType.Create("group", true);
            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);
            GroupProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupProperties.Add(mmsAttribute.Name);
            
            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("description", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("mailNickname", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupMemberProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("owner", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupMemberProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("isArchived", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            GroupProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowAddRemoveApps", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveTabs", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveConnectors", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserEditMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowOwnerDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowTeamMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowChannelMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowGiphy", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_giphyContentRating", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowStickersAndMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowCustomMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);
            TeamsProperties.Add(mmsAttribute.Name);

            return mmsType;
        }
    }
}