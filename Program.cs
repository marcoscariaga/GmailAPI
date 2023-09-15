using GmailAPI.APIHelper;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GmailAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Gmail> MailLists = GetAllEmails(Convert.ToString(ConfigurationManager.Appsettings["HostAddress"]));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }

        public static List<Gmail> GetAllEmails(string HostEmailAddress)
        {
            try
            {
                GmailService gmailService = GmailHelper.GetService();
                List<Gmail> emailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest listRequest = gmailService.Users.Messages.List(HostEmailAddress);
                listRequest.LabelIds = "INBOX";
                listRequest.IncludeSpamTrash = false;
                //listRequest.Q = "is:unread"; //ONLY FOR UNREAD EMAILS...

                //GET ALL EMAILS
                ListMessagesResponse listResponse = listRequest.Execute();

                if (listResponse != null && listResponse != null)
                {
                    //LOOP THROUGH EACH EMAIL AND GET WHAT FILEDS I WANT
                    foreach (Message msg in listResponse.Messages)
                    {
                        //MESSAGE MARKS AS READ AFTER READING MESSAGE
                        GmailHelper.MsgMarkAsread(HostEmailAddress, msg.Id);

                        UsersResource.MessagesResource.GetRequest message = gmailService.Users.Messages.Get(HostEmailAddress, msg);
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
                            sting readableText = string.Empty;

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

                            //READ MAIL BODY
                            Console.WriteLine("STEP-2: Read Mail Body");
                            List<string> fileName = GmailHelper.GetAttachments(HostEmailAddress, msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]);

                            if (fileName.Count() > 0)
                            {
                                foreach (var eachFile in fileName)
                                {
                                    //GET SUER ID USING FROM EMAIL ADDRESS
                                    string[] rectifyFromAddress = fromAddress.Split(' ');
                                    string fromAdd = rectifyFromAddress[rectifyFromAddress.Length - 1];

                                    if(!string.IsNullOrEmpty(fromAdd))
                                    {
                                        fromAdd = fromAdd.Replace("<", string.Empty);
                                        fromAdd = fromAdd.Replace(">", string.Empty);
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("STEP-3: Mail has no attatchments.");
                            }
                        }
                    }
                }
            }
        }
    }
}
