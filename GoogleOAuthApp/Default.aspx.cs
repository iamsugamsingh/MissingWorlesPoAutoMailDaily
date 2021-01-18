using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleOAuthApp
{
    public partial class Default : System.Web.UI.Page
    {
        static string[] Scopes = { GmailService.Scope.GmailReadonly };
        static string ApplicationName = "Gmail API .NET Quickstart";

        protected void Page_Load(object sender, EventArgs e)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(Server.MapPath(@"~/client.com.json"), FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = Server.MapPath(@"~/token.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Response.Write("Credential file saved to: " + credPath +"<br/>");
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            //UsersResource.LabelsResource.ListRequest request = service.Users.Labels.List("me");

            //// List labels.
            //IList<Label> labels = request.Execute().Labels;
            //Response.Write("Labels:" +"<br/>");
            //if (labels != null && labels.Count > 0)
            //{
            //    foreach (var labelItem in labels)
            //    {
            //        Response.Write(labelItem.Name+"<br/>");
            //    }
            //}
            //else
            //{
            //    Response.Write("No labels found.");
            //}


            var inboxlistRequest = service.Users.Messages.List("me");
            inboxlistRequest.LabelIds = "INBOX";
            inboxlistRequest.IncludeSpamTrash = false;
            //get our emails   
            var emailListResponse = inboxlistRequest.Execute();
            if (emailListResponse != null && emailListResponse.Messages != null)
            {
                //loop through each email and get what fields you want...   
                foreach (var email in emailListResponse.Messages)
                {
                    var emailInfoRequest = service.Users.Messages.Get("me", email.Id);
                    var emailInfoResponse = emailInfoRequest.Execute();
                    if (emailInfoResponse != null)
                    {
                        String from = "";
                        String date = "";
                        String subject = "";
                        //loop through the headers to get from,date,subject, body
                        foreach (var mParts in emailInfoResponse.Payload.Headers)
                        {
                            //if (mParts.Name == "Date")
                            //{
                            //    date = mParts.Value;
                            //}
                            //else if (mParts.Name == "From")
                            //{
                            //    from = mParts.Value;
                            //}
                            //else if (mParts.Name == "Subject")
                            //{
                            //    subject = mParts.Value;

                            //}
                            //if (date != "" && from != "")
                            //{

                            //}

                            if (mParts.Name == "Subject")
                            {
                                subject = mParts.Value;
                                if(subject.Contains("SS"))
                                {
                                    if (emailInfoResponse.Payload.Parts != null)
                                    {
                                        foreach (MessagePart p in emailInfoResponse.Payload.Parts)
                                        {
                                            if (emailInfoResponse.Payload.MimeType == "text/html")
                                            {
                                                byte[] data = FromBase64ForUrlString(p.Body.Data);
                                                string decodedString = Encoding.UTF8.GetString(data);
                                                Response.Write(decodedString);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //Console.ReadLine();
        }

        public static byte[] FromBase64ForUrlString(string base64ForUrlInput)
        {
            int padChars = (base64ForUrlInput.Length % 4) == 0 ? 0 : (4 - (base64ForUrlInput.Length % 4));
            StringBuilder result = new StringBuilder(base64ForUrlInput, base64ForUrlInput.Length + padChars);
            result.Append(String.Empty.PadRight(padChars, '='));
            result.Replace('-', '+');
            result.Replace('_', '/');
            return Convert.FromBase64String(result.ToString());
        }
    }
}