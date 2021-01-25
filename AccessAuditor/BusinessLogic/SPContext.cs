using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AccessAuditor.BusinessLogic
{
    public class SPContext
    {
        private static string username;
        private static string password;
        private static ClientContext _instance;
        private SPContext() { }

       

        public ClientContext CurrentContext
        {
            get
            {
                string username = "mujassir@chaqua.org";
                string password = "lM6oJ9xX9jK2sY0y";
                string hostSiteURL = "https://chaqua.sharepoint.com/";

                if (_instance == null)
                    _instance = GetSPOContext(hostSiteURL, username, password);

                return _instance;
            }
        }

        private static ClientContext GetSPOContext(string spSiteURL, string username, string password)
        {

            var secure = new SecureString();
            foreach (char c in password)
            {
                secure.AppendChar(c);
            }
            ClientContext spoContext = new ClientContext(spSiteURL);
            spoContext.Credentials = new SharePointOnlineCredentials(username, secure);
            return spoContext;
        }
    }
}
