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

        internal static HashSet<string> GroupProperties = new HashSet<string>() { "id", "displayName", "description", "mailNickname", "visibility" };

        internal static HashSet<string> GroupMemberProperties = new HashSet<string>() { "member", "owner" };

        internal static HashSet<string> TeamsFromGroupProperties = new HashSet<string>() {
            "template",
            "isArchived",
            "memberSettings_allowCreateUpdateChannels",
            "memberSettings_allowDeleteChannels",
            "memberSettings_allowAddRemoveApps",
            "memberSettings_allowCreateUpdateRemoveTabs",
            "memberSettings_allowCreateUpdateRemoveConnectors",
            "guestSettings_allowCreateUpdateChannels",
            "guestSettings_allowDeleteChannels",
            "messagingSettings_allowUserEditMessages",
            "messagingSettings_allowUserDeleteMessages",
            "messagingSettings_allowOwnerDeleteMessages",
            "messagingSettings_allowTeamMentions",
            "messagingSettings_allowChannelMentions",
            "funSettings_allowGiphy",
            "funSettings_giphyContentRating",
            "funSettings_allowStickersAndMemes",
            "funSettings_allowCustomMemes",
        };

        internal static HashSet<string> TeamsProperties = new HashSet<string>() {
            "template",
            "isArchived",
            "memberSettings_allowCreateUpdateChannels",
            "memberSettings_allowDeleteChannels",
            "memberSettings_allowAddRemoveApps",
            "memberSettings_allowCreateUpdateRemoveTabs",
            "memberSettings_allowCreateUpdateRemoveConnectors",
            "guestSettings_allowCreateUpdateChannels",
            "guestSettings_allowDeleteChannels",
            "messagingSettings_allowUserEditMessages",
            "messagingSettings_allowUserDeleteMessages",
            "messagingSettings_allowOwnerDeleteMessages",
            "messagingSettings_allowTeamMentions",
            "messagingSettings_allowChannelMentions",
            "funSettings_allowGiphy",
            "funSettings_giphyContentRating",
            "funSettings_allowStickersAndMemes",
            "funSettings_allowCustomMemes",
            "visibility"
        };

        internal static HashSet<string> GroupFromTeamProperties = new HashSet<string>() { "id", "mailNickname", "displayName", "description" };

        public Schema GetMmsSchema(SchemaContext context)
        {
            Schema mmsSchema = new Schema();
            mmsSchema.Types.Add(SchemaProvider.GetSchemaTypeUser());
            mmsSchema.Types.Add(SchemaProvider.GetSchemaTypeGroup());
            mmsSchema.Types.Add(SchemaProvider.GetSchemaTypeTeam());
            mmsSchema.Types.Add(SchemaProvider.GetSchemaTypeChannel());

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

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            return mmsType;
        }

        private static SchemaType GetSchemaTypeChannel()
        {
            SchemaType mmsType = SchemaType.Create("channel", true);

            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("teamid", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("description", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("email", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("webUrl", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("owner", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("isFavoriteByDefault", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("membershipType", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            return mmsType;
        }

        private static SchemaType GetSchemaTypeGroup()
        {
            SchemaType mmsType = SchemaType.Create("group", true);

            //Group

            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("description", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("mailNickname", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("visibility", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            // Group member
            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("owner", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            // Teams

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("isArchived", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("template", AttributeType.String, AttributeOperation.ExportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowAddRemoveApps", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveTabs", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveConnectors", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserEditMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowOwnerDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowTeamMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowChannelMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowGiphy", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_giphyContentRating", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowStickersAndMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowCustomMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            return mmsType;
        }

        private static SchemaType GetSchemaTypeTeam()
        {
            SchemaType mmsType = SchemaType.Create("team", true);

            //Group

            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("description", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("mailNickname", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("visibility", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            // Group member
            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("owner", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            // Teams

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("isArchived", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("template", AttributeType.String, AttributeOperation.ExportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowAddRemoveApps", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveTabs", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("memberSettings_allowCreateUpdateRemoveConnectors", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowCreateUpdateChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("guestSettings_allowDeleteChannels", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserEditMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowUserDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowOwnerDeleteMessages", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowTeamMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("messagingSettings_allowChannelMentions", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowGiphy", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_giphyContentRating", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowStickersAndMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("funSettings_allowCustomMemes", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            return mmsType;
        }
    }
}