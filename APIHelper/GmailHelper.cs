using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GmailAPI.APIHelper
{
    public static class GmailHelper
    {
        static string[] Scopes = { GmailService.Scope.MailGoogleCom };
        static string ApplicationName = "Gmail API Application";

        public static GmailService GetService()
        {
            UserCredential credential;
            using (FileStream stream = new FileStream(Convert.ToString(ConfigurationManager.AppSettings["ClientInfo"]), FileMode.Open, FileAccess.Read))
            {

            }

            return null;
        }

    }
}
