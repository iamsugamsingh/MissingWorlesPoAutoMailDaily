using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Data;
using System.Configuration;
using System.Data.OleDb;
using System.Net.Mail;
using System.Net;

namespace GoogleOAuthApp
{
    public partial class Test : System.Web.UI.Page
    {
        DataTable datatable = new DataTable();
        List<String> inDBworlespolist = new List<String>();
        HashSet<String> inEmailworlespolist = new HashSet<String>();
        OleDbConnection con = new OleDbConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString);
        string emailEndDate;
        string[] Scopes = { GmailService.Scope.GmailModify };
        string ApplicationName = "Gmail API .NET Quickstart";
        string Pass, FromEmailid, HostAdd, msg;
        HashSet<string> pendingList = new HashSet<string>();
        protected void Page_Load(object sender, EventArgs e)
        {
            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/gmail-dotnet-quickstart.json     

            String from = "";
            String emailDate = "";
            String subject = "";
            String body = "";
            UserCredential credential;

            string worlesPoRange = MinMaxWorlesPoNumberInDataBase();
            long startRangeValue = Convert.ToInt32(worlesPoRange.Split('-').GetValue(0).ToString());
            long endRangeValue = Convert.ToInt32(worlesPoRange.Split('-').Last().ToString());

            try
            {
                using (var stream = new FileStream(Server.MapPath(@"~/client.com.json"), FileMode.Open, FileAccess.Read))
                {
                    string credPath = Server.MapPath(@"~/GmailAuthorization/");
                    credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart2.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                    //Response.Write("Credential file saved to: " + credPath);
                }

                // Create Gmail API service.
                var service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                var re = service.Users.Messages.List("me");
                re.LabelIds = "INBOX";
                re.MaxResults = 1000;
                //re.Q = "is:unread"; //only get unread;

                var res = re.Execute();

                if (res != null && res.Messages != null)
                {
                    //Response.Write("there are emails. press any key to continue!"+ res.Messages.Count);

                    foreach (var email in res.Messages)
                    {
                        var emailInfoReq = service.Users.Messages.Get("me", email.Id);
                        var emailInfoResponse = emailInfoReq.Execute();

                        if (emailInfoResponse != null)
                        {
                            //loop through the headers and get the fields we need...
                            foreach (var mParts in emailInfoResponse.Payload.Headers)
                            {
                                if (mParts.Name == "Date")
                                {
                                    emailDate = mParts.Value;
                                }
                                else if (mParts.Name == "From")
                                {
                                    from = mParts.Value;
                                }
                                else if (mParts.Name == "Subject")
                                {
                                    subject = mParts.Value;
                                }
                                string currentDateInEmail = "", startDate = null;

                                if (emailDate != "")
                                {
                                    currentDateInEmail = emailDate.Substring(5, 11);
                                    startDate = DateTime.Today.AddDays(-2).ToShortDateString();

                                    if (Convert.ToDateTime(currentDateInEmail) <= Convert.ToDateTime(startDate) && Convert.ToDateTime(currentDateInEmail) >= Convert.ToDateTime(emailEndDate))
                                    {
                                        if (subject.Contains("PO") && emailDate != "")
                                        {
                                            if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                                                body = DecodeBase64String(emailInfoResponse.Payload.Body.Data);
                                            else
                                                body = GetNestedBodyParts(emailInfoResponse.Payload.Parts, "");   

                                            string[] emailsubject = subject.Split(' ');

                                            foreach(var isPoNum in emailsubject)
                                            {
                                                if(isPoNum.Contains("PO") && isPoNum.Length>6)
                                                {
                                                    if (isPoNum.Contains("-"))
                                                    {
                                                        long to = Convert.ToInt32(isPoNum.Split('-').GetValue(0).ToString().Substring(2, 5));
                                                        long last = Convert.ToInt32(isPoNum.Split('-').Last());

                                                        for (long i = to; i <= last; i++)
                                                        {


                                                            if (startRangeValue <= i && endRangeValue >= i)
                                                            {
                                                                inEmailworlespolist.Add(i.ToString());
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (startRangeValue <= Convert.ToInt32(isPoNum.Substring(2,5)) && endRangeValue >= Convert.ToInt32(isPoNum.Substring(2, 5)))
                                                        {
                                                            inEmailworlespolist.Add(isPoNum.Substring(2, 5));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                String DecodeBase64String(string s)
                {
                    var ts = s.Replace("-", "+");
                    ts = ts.Replace("_", "/");
                    var bc = Convert.FromBase64String(ts);
                    var tts = Encoding.UTF8.GetString(bc);

                    return tts;
                }

                String GetNestedBodyParts(IList<MessagePart> part, string curr)
                {
                    string str = curr;
                    if (part == null)
                    {
                        return str;
                    }
                    else
                    {
                        foreach (var parts in part)
                        {
                            if (parts.Parts == null)
                            {
                                if (parts.Body != null && parts.Body.Data != null)
                                {
                                    var ts = DecodeBase64String(parts.Body.Data);
                                    str += ts;
                                }
                            }
                            else
                            {
                                return GetNestedBodyParts(parts.Parts, str);
                            }
                        }

                        return str;
                    }
                }

                msg = "Dear Sir, <br/><br/> These are the Missing Worles POs, that still not put in our system by technical department. <br/><br/>";
                string isPendingPo = "";
                foreach (var emailPoNumber in inEmailworlespolist)
                {
                    //Response.Write(emailPoNumber+"..........................<br/>");
                    con.Open();
                    OleDbCommand cmd = new OleDbCommand("Select WorlesPo from [Pedidos de clientes] where WorlesPo IN ('"+ emailPoNumber+"')", con);
                    OleDbDataReader dr = cmd.ExecuteReader();
                    if(dr.HasRows && dr.Read())
                    {

                    }
                    else
                    {
                        msg += emailPoNumber + ", ";
                        isPendingPo = "exist";
                    }
                    con.Close();
                }

                if (isPendingPo == "exist")
                {
                    //technical@worles.com
                    Email_With_CCandBCC("technical@worles.com", "", "it2@anugroup.net", "Missing Worles PO Reminder", msg);
                }
            }
            catch (Exception ex)
            {
                Response.Write("Error in page load <br>" + ex.Message);
            }
        }
        public void Email_With_CCandBCC(String ToEmail, string cc, string bcc, string Subj, string Message)
        {
            try
            {
                //Reading sender Email credential from web.config file
                HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
                FromEmailid = ConfigurationManager.AppSettings["FromMail"].ToString();
                Pass = ConfigurationManager.AppSettings["Password"].ToString();

                //creating the object of MailMessage
                MailMessage mailMessage = new MailMessage();

                mailMessage.From = new MailAddress(FromEmailid); //From Email Id
                mailMessage.Subject = Subj; //Subject of Email
                mailMessage.Body = Message; //body or message of Email
                mailMessage.IsBodyHtml = true;

                foreach (var address in ToEmail.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mailMessage.To.Add(address);
                }

                foreach (var address in cc.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mailMessage.CC.Add(address);
                }

                foreach (var address in bcc.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mailMessage.Bcc.Add(address);
                }

                SmtpClient smtp = new SmtpClient(); // creating object of smptpclient
                smtp.Host = HostAdd; //host of emailaddress for example smtp.gmail.com etc

                //network and security related credentials

                smtp.EnableSsl = true;
                NetworkCredential NetworkCred = new NetworkCredential();
                NetworkCred.UserName = mailMessage.From.Address;
                NetworkCred.Password = Pass;
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = 587;
                smtp.Send(mailMessage);

                Response.Write("Email Send Succesfully!...");
            }
            catch (Exception ex)
            {
                Response.Write("Can't Send Mail..... " + ex.Message);
            }
        }
        public string MinMaxWorlesPoNumberInDataBase()
        {
            long orderSeries = 0;
            long checkRange = 0;
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Top 1 WorlesPo from [Pedidos de clientes] where WorlesPo IS Not Null And worlesPo not like('X%') and  worlesPo not like('V%') and  worlesPo not like('e%') and  worlesPo not like('E%') and  worlesPo not like('S%') and  worlesPo not like('P%') and  worlesPo not like('A%') and  worlesPo not like('u%') and  worlesPo not like('T%') and  worlesPo not like('R%') and  worlesPo not like('N%') and  worlesPo not like('M%') and  worlesPo not like('I%') and (Len(WorlesPo)=5) Order By WorlesPo Desc", con);
            OleDbDataReader dr = cmd.ExecuteReader();
            if (dr.HasRows && dr.Read())
            {
                string rr = dr["WorlesPo"].ToString();
                orderSeries = Convert.ToInt32(dr["WorlesPo"].ToString().Substring(0, 1)) * 10000;
                checkRange = (Convert.ToInt32(dr["WorlesPo"].ToString().Substring(0, 1)) + 1) * 10000;
            }
            con.Close();
            return orderSeries + "-" + checkRange;
        }

    }
}