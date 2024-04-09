using GmailAPI.APIHelper;
using GmailAPI.APIHelpér;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace GmailAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Gmail> mailLists = GetAllEmailsAndSaveToDB(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static List<Gmail> GetAllEmailsAndSaveToDB(string HostEmailAddress)
        {
            List<Gmail> emailList = new List<Gmail>();

            try
            {
                GmailService gmailService = GmailHelper.GetService();
                UsersResource.MessagesResource.ListRequest listRequest = gmailService.Users.Messages.List(HostEmailAddress);
                listRequest.LabelIds = "INBOX";
                listRequest.IncludeSpamTrash = false;
                listRequest.Q = "is:unread"; //ONLY FOR UNREAD EMAILS...
                //listRequest.Q = "in:inbox";

                bool hasNext = true;

                while (hasNext)
                {
                    //GET ALL EMAILS
                    ListMessagesResponse listResponse = listRequest.Execute();
                    hasNext = listResponse.NextPageToken != null;

                    if (listResponse != null && listResponse.Messages != null)
                    {
                        //LOOP THROUGH EACH EMAIL AND GET WHAT FILEDS I WANT
                        foreach (Message msg in listResponse.Messages)
                        {
                            //MESSAGE MARKS AS READ AFTER READING MESSAGE
                            GmailHelper.MsgMarkAsread(HostEmailAddress, msg.Id);

                            UsersResource.MessagesResource.GetRequest message = gmailService.Users.Messages.Get(HostEmailAddress, msg.Id);
                            Console.WriteLine("\n-----------------NEW MAIL-----------------");
                            Console.WriteLine("STEP-1: Message ID: " + msg.Id);

                            //MAKE ANOTHER REQUEST FOR THAT EMAIL ID...
                            Message msgContent = message.Execute();

                            if (msgContent != null)
                            {
                                string fromAddress = string.Empty;
                                string date = string.Empty;
                                string subject = string.Empty;
                                string mailBody = string.Empty;
                                string readableText = string.Empty;

                                //LOOP THROUGH THE HEADERS AND GET THE FIELDS WE NEED (SUBJECT, MAIL)
                                foreach (var messageParts in msgContent.Payload.Headers)
                                {
                                    if (messageParts.Name == "From")
                                    {
                                        fromAddress = messageParts.Value;
                                    }
                                    else if (messageParts.Name == "Date")
                                    {
                                        date = messageParts.Value;
                                    }
                                    else if (messageParts.Name == "Subject")
                                    {
                                        subject = messageParts.Value;
                                    }
                                }

                                if (subject != "Inscrição")
                                    continue;

                                #region GET ATTACHMENTS

                                //READ MAIL BODY
                                //Console.WriteLine("STEP-2: Read Mail Body");
                                //List<string> fileName = GmailHelper.GetAttachments(HostEmailAddress, msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));

                                //if (fileName.Count() > 0)
                                //{
                                //    foreach (var eachFile in fileName)
                                //    {
                                //        //GET USER ID USING FROM EMAIL ADDRESS
                                //        string[] rectifyFromAddress = fromAddress.Split(' ');
                                //        string fromAdd = rectifyFromAddress[rectifyFromAddress.Length - 1];

                                //        if (!string.IsNullOrEmpty(fromAdd))
                                //        {
                                //            fromAdd = fromAdd.Replace("<", string.Empty);
                                //            fromAdd = fromAdd.Replace(">", string.Empty);
                                //        }
                                //    }
                                //}
                                //else
                                //{
                                //    Console.WriteLine("STEP-3: Mail has no attatchments.");
                                //}

                                #endregion

                                //READ MAIL BODY
                                mailBody = string.Empty;
                                if (msgContent.Payload.Parts == null && msgContent.Payload.Body != null)
                                {
                                    mailBody = msgContent.Payload.Body.Data;
                                }
                                else
                                {
                                    mailBody = GmailHelper.MsgNestedParts(msgContent.Payload.Parts);
                                }

                                //BASE64 TO READABLE TEXT
                                readableText = string.Empty;
                                readableText = GmailHelper.Base64Decode(mailBody);

                                Console.WriteLine("STEP-4: Identify & Configure Mails.");

                                if (!string.IsNullOrEmpty(readableText))
                                {
                                    Gmail gmail = new Gmail();
                                    gmail.From = fromAddress;
                                    gmail.Body = readableText;
                                    gmail.MailDateTime = Convert.ToDateTime(date).ToUniversalTime();
                                    emailList.Add(gmail);

                                    GmailHelper.SeparateBodyAndSaveToDB(gmail.Body, gmail.MailDateTime);
                                }
                            }
                        }
                        if (hasNext)
                            listRequest.PageToken = listResponse.NextPageToken;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return emailList;
        }
    }
}
