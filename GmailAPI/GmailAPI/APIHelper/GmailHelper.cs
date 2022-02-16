using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Configuration;

namespace GmailAPI.APIHelper
{
    public static class GmailAPIHelper
    {
        static string[] Scopes = { GmailService.Scope.MailGoogleCom };
        static string ApplicationName = "Gmail API Application";

        public static GmailService GetService()
        {
            UserCredential credential;
            using (FileStream stream = new FileStream(Convert.ToString(ConfigurationManager.AppSettings["ClientInfo"]), 
                FileMode.Open, FileAccess.Read))
            {
                String FolderPath = Convert.ToString(ConfigurationManager.AppSettings["CredentialsInfo"]);
                String FilePath = Path.Combine(FolderPath, "APITokenCredentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(FilePath, true)).Result;
            }
            // Create Gmail API service.
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }
        public static string MsgNestedParts(IList<MessagePart> Parts)
        {
            string str = string.Empty;
            if (Parts.Count() < 0)
            {
                return string.Empty;
            }
            else
            {
                IList<MessagePart> PlainTestMail = Parts.Where(x => x.MimeType == "text/plain").ToList();
                IList<MessagePart> AttachmentMail = Parts.Where(x => x.MimeType == "multipart/alternative").ToList();

                if (PlainTestMail.Count() > 0)
                {
                    foreach (MessagePart EachPart in PlainTestMail)
                    {
                        if (EachPart.Parts == null)
                        {
                            if (EachPart.Body != null && EachPart.Body.Data != null)
                            {
                                str += EachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(EachPart.Parts);
                        }
                    }
                }
                if (AttachmentMail.Count() > 0)
                {
                    foreach (MessagePart EachPart in AttachmentMail)
                    {
                        if (EachPart.Parts == null)
                        {
                            if (EachPart.Body != null && EachPart.Body.Data != null)
                            {
                                str += EachPart.Body.Data;
                            }
                        }
                        else
                        {
                            return MsgNestedParts(EachPart.Parts);
                        }
                    }
                }
                return str;
            }
        }
        public static List<string> GetAttachments(string userId, string messageId, String outputDir)
        {
            try
            {
                List<string> FileName = new List<string>();
                GmailService GServices = GetService();
                Message message = GServices.Users.Messages.Get(userId, messageId).Execute();
                IList<MessagePart> parts = message.Payload.Parts;

                foreach (MessagePart part in parts)
                {
                    if (!String.IsNullOrEmpty(part.Filename))
                    {
                        string attId = part.Body.AttachmentId;
                        MessagePartBody attachPart = GServices.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        byte[] data = Base64ToByte(attachPart.Data);
                        File.WriteAllBytes(Path.Combine(outputDir, part.Filename), data);
                        FileName.Add(part.Filename);
                    }
                }
                return FileName;
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                return null;
            }
        }
        public static string Base64Decode(string Base64Test)
        {
            string EncodTxt = string.Empty;
            //STEP-1: Replace all special Character of Base64Test
            EncodTxt = Base64Test.Replace("-", "+");
            EncodTxt = EncodTxt.Replace("_", "/");
            EncodTxt = EncodTxt.Replace(" ", "+");
            EncodTxt = EncodTxt.Replace("=", "+");

            //STEP-2: Fixed invalid length of Base64Test
            if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('=', 4 - EncodTxt.Length % 4); }
            else if (EncodTxt.Length % 4 == 0)
            {
                EncodTxt = EncodTxt.Substring(0, EncodTxt.Length - 1);
                if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('+', 4 - EncodTxt.Length % 4); }
            }

            //STEP-3: Convert to Byte array
            byte[] ByteArray = Convert.FromBase64String(EncodTxt);

            //STEP-4: Encoding to UTF8 Format
            return Encoding.UTF8.GetString(ByteArray);
        }
        public static byte[] Base64ToByte(string Base64Test)
        {
            string EncodTxt = string.Empty;
            //STEP-1: Replace all special Character of Base64Test
            EncodTxt = Base64Test.Replace("-", "+");
            EncodTxt = EncodTxt.Replace("_", "/");
            EncodTxt = EncodTxt.Replace(" ", "+");
            EncodTxt = EncodTxt.Replace("=", "+");

            //STEP-2: Fixed invalid length of Base64Test
            if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('=', 4 - EncodTxt.Length % 4); }
            else if (EncodTxt.Length % 4 == 0)
            {
                EncodTxt = EncodTxt.Substring(0, EncodTxt.Length - 1);
                if (EncodTxt.Length % 4 > 0) { EncodTxt += new string('+', 4 - EncodTxt.Length % 4); }
            }

            //STEP-3: Convert to Byte array
            return Convert.FromBase64String(EncodTxt);
        }
        public static void MsgMarkAsRead(string HostEmailAddress, string MsgId)
        {
            //MESSAGE MARKS AS READ AFTER READING MESSAGE
            ModifyMessageRequest mods = new ModifyMessageRequest();
            mods.AddLabelIds = null;
            mods.RemoveLabelIds = new List<string> { "UNREAD" };
            GetService().Users.Messages.Modify(mods, HostEmailAddress, MsgId).Execute();
        }
    }
}
