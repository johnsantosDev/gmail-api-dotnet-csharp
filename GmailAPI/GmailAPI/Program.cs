using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GmailAPI.APIHelper;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;

namespace GmailAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                List<Gmail> MailLists = GetAllEmails(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
                
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
                GmailService GmailService = GmailAPIHelper.GetService();
                List<Gmail> EmailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List(HostEmailAddress);
                ListRequest.LabelIds = "INBOX";
                ListRequest.IncludeSpamTrash = false;
                ListRequest.Q = "is:unread"; //ONLY FOR UNDREAD EMAIL'S...

                //GET ALL EMAILS
                ListMessagesResponse ListResponse = ListRequest.Execute();
                
                if (ListResponse != null && ListResponse.Messages != null)
                {
                    //LOOP THROUGH EACH EMAIL AND GET WHAT FIELDS I WANT
                    foreach (Message Msg in ListResponse.Messages)
                    {
                        //MESSAGE MARKS AS READ AFTER READING MESSAGE
                        GmailAPIHelper.MsgMarkAsRead(HostEmailAddress, Msg.Id);

                        UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get(HostEmailAddress, Msg.Id);
                        Console.WriteLine("\n-----------------NEW MAIL----------------------");
                        Console.WriteLine("STEP-1: Message ID:" + Msg.Id);

                        //MAKE ANOTHER REQUEST FOR THAT EMAIL ID...
                        Message MsgContent = Message.Execute();

                        if (MsgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string Date = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;

                            //LOOP THROUGH THE HEADERS AND GET THE FIELDS WE NEED (SUBJECT, MAIL)
                            foreach (var MessageParts in MsgContent.Payload.Headers)
                            {
                                if (MessageParts.Name == "From")
                                {
                                    FromAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Date")
                                {
                                    Date = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Subject")
                                {
                                    Subject = MessageParts.Value;
                                }
                            }
                            //READ MAIL BODY
                            Console.WriteLine("STEP-2: Read Mail Body");
                            List<string> FileName = GmailAPIHelper.GetAttachments(HostEmailAddress, Msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));

                            if (FileName.Count() > 0)
                            {
                                foreach (var EachFile in FileName)
                                {
                                    //GET USER ID USING FROM EMAIL ADDRESS-------------------------------------------------------
                                    string[] RectifyFromAddress = FromAddress.Split(' ');
                                    string FromAdd = RectifyFromAddress[RectifyFromAddress.Length - 1];

                                    if (!string.IsNullOrEmpty(FromAdd))
                                    {
                                        FromAdd = FromAdd.Replace("<", string.Empty);
                                        FromAdd = FromAdd.Replace(">", string.Empty);
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("STEP-3: Mail has no attachments.");
                            }

                            //READ MAIL BODY-------------------------------------------------------------------------------------
                            MailBody = string.Empty;
                            if (MsgContent.Payload.Parts == null && MsgContent.Payload.Body != null)
                            {
                                MailBody = MsgContent.Payload.Body.Data;
                            }
                            else
                            {
                                MailBody = GmailAPIHelper.MsgNestedParts(MsgContent.Payload.Parts);
                            }

                            //BASE64 TO READABLE TEXT--------------------------------------------------------------------------------
                            ReadableText = string.Empty;
                            ReadableText = GmailAPIHelper.Base64Decode(MailBody);

                            Console.WriteLine("STEP-4: Identifying & Configure Mails.");

                            if (!string.IsNullOrEmpty(ReadableText))
                            {
                                Gmail GMail = new Gmail();
                                GMail.From = FromAddress;
                                GMail.Body = ReadableText;
                                GMail.MailDateTime = Convert.ToDateTime(Date);
                                EmailList.Add(GMail);
                            }
                        }
                    }
                }
                return EmailList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);

                return null;
            }
        }
    }
}
