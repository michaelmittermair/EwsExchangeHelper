using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper
{
    public partial class MsExchangeServices
    {
        public ExtendedPropertyDefinition CustomerId = new ExtendedPropertyDefinition(14922, MapiPropertyType.String);

        /// <summary>
        /// Retriving a list of contacts with the property customer id
        /// </summary>
        /// <param name="folder">folder, in which the module should search for contacts</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>list of contacts</returns>
        public Collection<Contact> GetContactsWithCustomerId(Folder folder, PropertySet propSet)
        {
            return GetContactWithCustomerId(folder.Id, propSet);
        }

        /// <summary>
        /// Retriving a list of contacts with the property customer id
        /// </summary>
        /// <param name="folder">WellKnowFolderName for the folder</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>list of contacts</returns>
        public Collection<Contact> GetContactsWithCustomerId(WellKnownFolderName folder, PropertySet propSet)
        {
            return GetContactWithCustomerId(folder, propSet);
        }

        /// <summary>
        /// loading the contact from the specified folder
        /// </summary>
        /// <param name="id">id of the folder</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>list of contacts</returns>
        private Collection<Contact> GetContactWithCustomerId(FolderId id, PropertySet propSet)
        {
            // creating a view. This view specifies, how many contacts should be displayed. The maximum is 1000 contacts at once
            var itemView = new ItemView(int.MaxValue, 0, OffsetBasePoint.Beginning)
            {
                PropertySet = propSet
            };

            // the filtered list, where the contacts are in
            FindItemsResults<Item> searchResults;

            // list, with all contacts on the public folder
            var resultSet = new List<Contact>();

            do
            {
                // getting the first list with contacts
                searchResults = GetItems(id, itemView);

                // adding them to my list
                resultSet.AddRange(searchResults.Select(item => item as Contact).Where(contact => contact != null));

                // modified the offset, to get the next list (if > 1000)
                if (searchResults.NextPageOffset.HasValue)
                    itemView.Offset = searchResults.NextPageOffset.Value;

                // checking, if there are more contacts available
            } while (searchResults.MoreAvailable);

             
            return new Collection<Contact>(resultSet);
        }


        /// <summary>
        /// Retriving a list of contacts with the property customer id  and contact picture
        /// </summary>
        /// <param name="folder">folder, in which the module should search for contacts</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>list of contacts</returns>
        public Collection<Contact> BatchGetContactItemsWithAttachments(Folder folder, PropertySet propSet)
        {
            return BatchGetContactsWithAttachments(folder.Id, propSet);
        }

        /// <summary>
        /// Retriving a list of contacts with the property customer id and contact picture
        /// </summary>
        /// <param name="folder">WellKnowFolderName for the folder</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>list of contacts</returns>
        public Collection<Contact> BatchGetContactItemsWithAttachments(WellKnownFolderName folder, PropertySet propSet)
        {
            return BatchGetContactsWithAttachments(folder, propSet);
        }

        /// <summary>
        /// loading the contacts from a folder with the contact picture attachments
        /// </summary>
        /// <param name="id">id of the folder</param>
        /// <param name="propSet">properties, that should also be loaded</param>
        /// <returns>collection of contacts</returns>
        private Collection<Contact> BatchGetContactsWithAttachments(FolderId id, PropertySet propSet)
        {
            // creating a view. This view specifies, how many contacts should be displayed. The maximum is 1000 contacts at once
            var itemView = new ItemView(int.MaxValue, 0, OffsetBasePoint.Beginning)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly)
            };

            // the filtered list, where the contacts are in
            FindItemsResults<Item> searchResults;

            // list, with all contacts on the public folder
            var resultSet = new List<Contact>();

            do
            {
                // getting the first list with contacts
                searchResults = GetItems(id, itemView);

                // adding them to my list
                resultSet.AddRange(searchResults.Select(item => item as Contact).Where(contact => contact != null));

                // modified the offset, to get the next list (if > 1000)
                if (searchResults.NextPageOffset.HasValue)
                    itemView.Offset = searchResults.NextPageOffset.Value;

                // checking, if there are more contacts available
            } while (searchResults.MoreAvailable);

            var list = new Collection<ItemId>();
            foreach (var contact in resultSet)
            {
                list.Add(contact.Id);
            }

            // Get the items from the server.
            // This method call results in a GetItem call to EWS.
            var response = ExchangeService.BindToItems(list, propSet);

            // Instantiate a collection of Contact objects to populate from the values that are returned by the Exchange server.
            var contactItems = new Collection<Contact>();

            var attachmentIds = new Collection<Attachment>();

            foreach (var getItemResponse in response)
            {
                try
                {
                    var item = getItemResponse.Item;

                    var contact = item as Contact;
                    if (contact == null)
                        continue;

                    // adding the contact to the contactlist
                    contactItems.Add(contact);

                    // checking the attachments for contact pictures and add them to the attachment list
                    foreach (var attachment in item.Attachments)
                    {
                        if (attachment is FileAttachment && (attachment as FileAttachment).IsContactPhoto)
                        {
                            // Load the attachment to access the content.
                            attachmentIds.Add(attachment);
                        }
                    }
                }
                catch (Exception)
                {
                    // ignored
                }
            }

            if (attachmentIds.Count > 0)
                // loading all attachments at once
                ExchangeService.GetAttachments(attachmentIds.ToArray(), null, null);

            return contactItems;
        }
    }
}
