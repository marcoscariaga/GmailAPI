using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

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
                String FolderPath = Convert.ToString(ConfigurationManager.AppSettings["CredentialsInfo"]);
                String FilePath = Path.Combine(FolderPath, "APITokenCredentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(FilePath, true)).Result;
            }

            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        public static void MsgMarkAsread(string hostEmailAddress, string msgId)
        {
            //MESSAGE MARKS AS READ AFTER READING MESSAGE
            ModifyMessageRequest mods = new ModifyMessageRequest();
            mods.AddLabelIds = null;
            mods.RemoveLabelIds = new List<string> { "UNREAD" };
            GetService().Users.Messages.Modify(mods, hostEmailAddress, msgId).Execute();
        }

        public static List<string> GetAttachments(string userId, string messageId, String outputDir)
        {
            try
            {
                List<string> fileName = new List<string>();
                GmailService gService = GetService();
                Message message = gService.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> parts = message.Payload.Parts;

                foreach(MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        string attId = part.Body.AttachmentId;
                        MessagePartBody attachPart = gService.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        byte[] data = Convert.FromBase64String(attachPart.Data); //Base64ToByte(attachPart.Data);
                        File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                        fileName.Add(part.Filename);
                    }
                }
                return fileName;
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
        }

        public static string MsgNestedParts(IList<MessagePart> parts)
        {
            string str = string.Empty;
            if (parts.Count() < 0)
            {
                return string.Empty;
            }
            else
            {
                IList<MessagePart> plainTestMail = parts.Where(x => x.MimeType == "text/plain").ToList();
                IList<MessagePart> attachmentMail = parts.Where(x => x.MimeType == "multipart/alternative").ToList();

                if (plainTestMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in plainTestMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }

                if (attachmentMail.Count() > 0)
                {
                    foreach (MessagePart eachPart in attachmentMail)
                    {
                        if (eachPart.Parts == null)
                        {
                            if (eachPart.Body != null && eachPart.Body.Data != null)
                            {
                                str += eachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(eachPart.Parts);
                        }
                    }
                }

                return str;
            }
        }

        public static string Base64Decode(string base64Test)
        {
            string encodeText = string.Empty;

            //STEP-1: Replace all special character od base64Test
            encodeText = base64Test.Replace("-", "+");
            encodeText = base64Test.Replace("_", "/");
            encodeText = base64Test.Replace(" ", "+");
            encodeText = base64Test.Replace("=", "+");

            //STEP-2: Fixed invalid length of base64Test
            if (encodeText.Length % 4 > 0) { encodeText += new string('=', 4 - encodeText.Length % 4); }
            else if (encodeText.Length % 4 == 0)
            {
                encodeText = encodeText.Substring(0, encodeText.Length - 1);
                if (encodeText.Length % 4 > 0) { encodeText += new string('+', 4 - encodeText.Length % 4); }
            }

            //STEP-3: Convert to Byte array
            byte[] byteArray = Convert.FromBase64String(encodeText);

            //STEP-4: Encoding to UTF-8 format
            return Encoding.UTF8.GetString(byteArray);
        }
    }
}
