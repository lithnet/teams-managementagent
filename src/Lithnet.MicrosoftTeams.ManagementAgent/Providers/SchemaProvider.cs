﻿using System;
using System.Collections.Generic;
using Lithnet.Ecma2Framework;
using Microsoft.MetadirectoryServices;
using NLog;

namespace Lithnet.MicrosoftTeams.ManagementAgent
{
    internal class SchemaProvider : ISchemaProvider
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        internal static HashSet<string> GroupProperties = new HashSet<string>() { "id", "displayName", "description", "mailNickname", "isArchived" };
        internal static HashSet<string> GroupMemberProperties = new HashSet<string>() {"member", "owner"};
        internal static HashSet<string> TeamsProperties = new HashSet<string>() {
            "template",
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

            //Group

            SchemaAttribute mmsAttribute = SchemaAttribute.CreateAnchorAttribute("id", AttributeType.String, AttributeOperation.ImportOnly);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("displayName", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("description", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("mailNickname", AttributeType.String, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateSingleValuedAttribute("isArchived", AttributeType.Boolean, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            // Group member
            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("member", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

            mmsAttribute = SchemaAttribute.CreateMultiValuedAttribute("owner", AttributeType.Reference, AttributeOperation.ImportExport);
            mmsType.Attributes.Add(mmsAttribute);

        // Teams
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