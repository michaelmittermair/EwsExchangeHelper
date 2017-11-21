using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper
{
    public partial class MsExchangeServices
    {
        /// <summary>
        /// Add the given user to the folder with the specified rights
        /// </summary>
        /// <param name="user">Name of the user, which should get the rights</param>
        /// <param name="permission">Predefined rights</param>
        public void EnableFolderPermissions(string user, WellKnownFolderName permission)
        {
            // Create a property set to use for folder binding.
            var propSet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.Permissions);

            // Specify the SMTP address of the new user and the folder permissions level.
            var fldperm = new FolderPermission(user, FolderPermissionLevel.Owner);

            // Bind to the folder and get the current permissions. 
            // This call results in a GetFolder call to EWS.
            var sentItemsFolder = Folder.Bind(ExchangeService, permission, propSet);

            // Add the permissions for the new user to the Sent Items DACL.
            sentItemsFolder.Permissions.Add(fldperm);

            // This call results in a UpdateFolder call to EWS.
            sentItemsFolder.Update();
        }

        /// <summary>
        /// Removing the given user from the folder with the specified rights
        /// </summary>
        /// <param name="user">Name of the user, which should get the rights</param>
        /// <param name="permission">Predefined rights</param>
        public void RemoveFolderPermissions(string user, WellKnownFolderName permission)
        {
            // Create a property set to use for folder binding.
            var propSet = new PropertySet(BasePropertySet.FirstClassProperties, FolderSchema.Permissions);

            // Bind to the folder and get the current permissions. 
            // This call results in a GetFolder call to EWS.
            var sentItemsFolder = Folder.Bind(ExchangeService, permission, propSet);

            // Iterate through the collection of permissions and remove permissions for any 
            // user with a display name or SMTP address. This leaves the anonymous and 
            // default user permissions unchanged. 
            if (sentItemsFolder.Permissions.Count != 0)
            {
                for (var t = 0; t < sentItemsFolder.Permissions.Count; t++)
                {
                    // Find any permissions associated with the specified user and remove them from the DACL
                    if (sentItemsFolder.Permissions[t].UserId.DisplayName != null || sentItemsFolder.Permissions[t].UserId.PrimarySmtpAddress != null)
                    {
                        sentItemsFolder.Permissions.Remove(sentItemsFolder.Permissions[t]);
                    }
                }
            }

            // This call results in an UpdateFolder call to EWS.
            sentItemsFolder.Update();
        }
    }
}
