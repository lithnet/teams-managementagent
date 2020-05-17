- Graph issue: Setting IsFavoriteByDefault may not work https://github.com/microsoftgraph/microsoft-graph-docs/issues/6792
- Graph issue: IsFavoriteByDefault does not appear up on import, and exported-change-not-reimported errors may appear on this attribute
- Graph issue: Getting private channel members may fail intermittently with error code Forbidden: Failed to execute Skype backend request GetThreadRosterS2SRequest
- Graph issue: Creating a private channel may fail with the following error message "CreateChannel_Private: Cannot create private channel, user <id> is not part of team <id>". A verification of the group shows that the user is indeed a member of the specified team. 
- Graph issue: Creating a team with visibility=HiddenMembership fails with the error 
  
  ```ErrorMessage : {"errors":[{"message":"Team Visibility must be one of known types: [Private,Public,HiddenMembership]."}```
