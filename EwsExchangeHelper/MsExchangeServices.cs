using System;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;

namespace EwsExchangeHelper
{
    public partial class MsExchangeServices : IDisposable
    {
        public ExchangeService ExchangeService { get; private set; }

        #region Constructors

        /// <summary>
        /// Creating a new Serverconnection based on the usernamen and passwort
        /// </summary>
        /// <param name="user">e-mail-address of the user</param>
        /// <param name="password">corresponding password</param>
        /// <param name="serverName">url or name of the server</param>
        public MsExchangeServices(string user, string password, string serverName)
        {
            ExchangeCredentials credentials = new WebCredentials(user, password);

            Initialize(serverName, credentials);
        }

        private void Initialize(string serverName, ExchangeCredentials credential = null, WebProxy proxy = null)
        {
            if (string.IsNullOrWhiteSpace(UserPrincipal.Current.EmailAddress))
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
                {
                    Url = new Uri($"https://{serverName}/EWS/Exchange.asmx")
                };
            }
            else
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
                ExchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2) { UseDefaultCredentials = true };
                ExchangeService.AutodiscoverUrl(UserPrincipal.Current.EmailAddress, RedirectionUrlValidationCallback);
            }

            if (ExchangeService == null) return;

            if (credential != null)
                ExchangeService.Credentials = credential;

            if (proxy != null)
                ExchangeService.WebProxy = proxy;
        }


        public MsExchangeServices(string serverName)
        {
            Initialize(serverName);
        }


        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            var result = false;

            var redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
                result = true;

            return result;
        }

        private static bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == 0)
            {
                // In all other cases, return false.
                return false;
            }
            if (chain != null)
            {
                foreach (var status in chain.ChainStatus)
                {
                    if (certificate.Subject == certificate.Issuer &&
                        status.Status == X509ChainStatusFlags.UntrustedRoot)
                    {
                        // Self-signed certificates with an untrusted root are valid. 
                    }
                    else
                    {
                        if (status.Status != X509ChainStatusFlags.NoError)
                        {
                            // If there are any other errors in the certificate chain, the certificate is invalid,
                            // so the method returns false.
                            return false;
                        }
                    }
                }
            }

            // When processing reaches this line, the only errors in the certificate chain are 
            // untrusted root errors for self-signed certificates. These certificates are valid
            // for default Exchange server installations, so return true.
            return true;
        }

        #endregion

        public void ImpersonateUser(ConnectingIdType idType, string id)
        {
            ExchangeService.ImpersonatedUserId = new ImpersonatedUserId(idType, id);
        }

        #region Items

        /// <summary>
        /// Getting the items of a specific folder associated with a view
        /// </summary>
        /// <param name="folderId">id of the folder</param>
        /// <param name="itemView">the associated view</param>
        /// <returns>list of elements (maximum 1000 items). To get more, use @GetNextItem </returns>
        public FindItemsResults<Item> GetItems(FolderId folderId, ItemView itemView)
        {
            return ExchangeService.FindItems(folderId, itemView);
        }

        public FindItemsResults<Item> GetNextItems(Folder folder, int pagesize)
        {
            var itemView = new ItemView(pagesize, 0, OffsetBasePoint.Beginning);
            return ExchangeService.FindItems(folder.Id, itemView);
        }

        #endregion

        #region Folders

        public FolderId GetFolderId(string ewsFolderPath, WellKnownFolderName wellKnownFolderName)
        {
            return ExistsFolder(ewsFolderPath, wellKnownFolderName)
                ? GetFolderByPath(ewsFolderPath, wellKnownFolderName).Id
                : null;
        }

        public bool ExistsFolder(string ewsFolderPath, WellKnownFolderName wellKnownFolderName)
        {
            try
            {
                GetFolderByPath(ewsFolderPath, wellKnownFolderName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void CreateContactFolder(string folderName, WellKnownFolderName wellKnownFolderName)
        {
            var folder = new ContactsFolder(ExchangeService) { DisplayName = folderName };
            folder.Save(wellKnownFolderName);
        }

        public Folder GetFolderByPath(string ewsFolderPath, WellKnownFolderName wellKnownFolderName)
        {
            var folders = ewsFolderPath.Split('\\');

            Folder parentFolderId = null;
            Folder actualFolder = null;

            for (var i = 0; i < folders.Length; i++)
            {
                if (0 == i)
                {
                    parentFolderId = GetTopLevelFolder(folders[i], wellKnownFolderName);
                    actualFolder = parentFolderId;
                }
                else
                {
                    if (parentFolderId != null) actualFolder = GetFolder(parentFolderId.Id, folders[i]);
                    parentFolderId = actualFolder;
                }
            }
            return actualFolder;
        }

        private Folder GetTopLevelFolder(string folderName, WellKnownFolderName wellKnownFolderName)
        {
            var findFolderResults = ExchangeService.FindFolders(wellKnownFolderName, new FolderView(int.MaxValue));
            foreach (
                var folder in
                    findFolderResults.Where(
                        folder => folderName.Equals(folder.DisplayName, StringComparison.InvariantCultureIgnoreCase)))
                return folder;

            throw new Exception("Top Level Folder not found: " + folderName);
        }

        private Folder GetFolder(FolderId parentFolderId, string folderName)
        {
            var findFolderResults = ExchangeService.FindFolders(parentFolderId, new FolderView(int.MaxValue));
            foreach (
                var folder in
                    findFolderResults.Where(
                        folder => folderName.Equals(folder.DisplayName, StringComparison.InvariantCultureIgnoreCase)))
                return folder;

            throw new Exception("Folder not found: " + folderName);
        }

        #endregion

        public void Dispose()
        {
            ExchangeService = null;
        }
    }
}