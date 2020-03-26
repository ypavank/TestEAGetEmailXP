using EAGetMail;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TestEAEmailImperApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //ReceiveMailByOauth();
        }

        // Generate an unqiue email file name based on date time
        static string _generateFileName(int sequence)
        {
            DateTime currentDateTime = DateTime.Now;
            return string.Format("{0}-{1:000}-{2:000}.eml",
                currentDateTime.ToString("yyyyMMddHHmmss", new CultureInfo("en-US")),
                currentDateTime.Millisecond,
                sequence);
        }

        static string _postString(string uri, string requestData)
        {
            HttpWebRequest httpRequest = WebRequest.Create(uri) as HttpWebRequest;
            httpRequest.Method = "POST";
            httpRequest.ContentType = "application/x-www-form-urlencoded";

            using (Stream requestStream = httpRequest.GetRequestStream())
            {
                byte[] requestBuffer = Encoding.UTF8.GetBytes(requestData);
                requestStream.Write(requestBuffer, 0, requestBuffer.Length);
                requestStream.Close();
            }

            try
            {
                HttpWebResponse httpResponse = httpRequest.GetResponse() as HttpWebResponse;
                var responseText = new StreamReader(httpResponse.GetResponseStream()).ReadToEnd();
                Console.WriteLine(responseText);
                return responseText;
            }
            catch (WebException ep)
            {
                if (ep.Status == WebExceptionStatus.ProtocolError)
                {
                    var responseText = new StreamReader(ep.Response.GetResponseStream()).ReadToEnd();
                    Console.WriteLine(responseText);
                }

                throw ep;
            }
        }

        public void ReceiveMailByOauth()
        {
            try
            {
                string client_id = "3ac1c2b8-9518-4f7d-9520-f886515364d6";// "8f54719b-4070-41ae-91ad-f48e3c793c5f";
                string client_secret = "s3y39]smt_rQ5D.-WXbol0RV6=9mFUaz";// "cbmYyGQjz[d29wL2ArcgoO7HLwJXL/-.";

                // If your application is not created by Office365 administrator, 
                // please use Office365 directory tenant id, you should ask Offic365 administrator to send it to you.
                // Office365 administrator can query tenant id in https://portal.azure.com/ - Azure Active Directory.
                string tenant = "b5b8b483-5597-4ae7-8e27-fcc464a3b584";// "79a42c6f-5a9a-439b-a2ca-7aa1b0ed9776";
                string scopes = "https://outlook.office.com/EWS.AccessAsUser.All%20offline_access%20email%20openid";                
                //string scopes = "https://outlook.office365.com/EWS.AccessUserEmail";

                string requestData =
                    string.Format("client_id={0}&client_secret={1}&scope={2}",
                        client_id, client_secret, scopes);
                

                string tokenUri = string.Format("https://login.microsoftonline.com/{0}/oauth2/v2.0/token", tenant);
                //string tokenUri = string.Format("https://login.microsoftonline.com/{0}", tenant);
                string responseText = _postString(tokenUri, requestData);

                OAuthResponseParser parser = new OAuthResponseParser();
                parser.Load(responseText);

                // Create a folder named "inbox" under current directory
                // to save the email retrieved.
                string localInbox = string.Format("{0}\\inbox", Directory.GetCurrentDirectory());
                // If the folder is not existed, create it.
                if (!Directory.Exists(localInbox))
                {
                    Directory.CreateDirectory(localInbox);
                }

                string officeUser = "pavan.yarlagadda@bsci.com";
                string token = parser.AccessToken;
                //token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6Imhhc2VaaEZUQ3gtY1RmVFM4ZGtUN1lydVEzQ3BwUFl5ZE1xUDdjT3d1RFkiLCJhbGciOiJSUzI1NiIsIng1dCI6IllNRUxIVDBndmIwbXhvU0RvWWZvbWpxZmpZVSIsImtpZCI6IllNRUxIVDBndmIwbXhvU0RvWWZvbWpxZmpZVSJ9.eyJhdWQiOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsImlzcyI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0L2I1YjhiNDgzLTU1OTctNGFlNy04ZTI3LWZjYzQ2NGEzYjU4NC8iLCJpYXQiOjE1ODUxNDA4NzgsIm5iZiI6MTU4NTE0MDg3OCwiZXhwIjoxNTg1MTQ0Nzc4LCJhY2N0IjowLCJhY3IiOiIxIiwiYWlvIjoiQVRRQXkvOE9BQUFBaTFCQTFnQmw0dXpCRkVIOTlVWnEzTjIwRFFUUjkwZjM2eU1DNkoxeWRqNzkxRjZSZzJ5YnlFNUs1RzZ0WWVRUCIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiVHJhYzIgRVdTIENvbm5lY3Rpb24iLCJhcHBpZCI6IjNhYzFjMmI4LTk1MTgtNGY3ZC05NTIwLWY4ODY1MTUzNjRkNiIsImFwcGlkYWNyIjoiMCIsImRldmljZWlkIjoiNmNlMTc2MDgtMTUyOC00N2I5LWIwODItOTc3ZDJhNmE1N2E2IiwiZW5mcG9saWRzIjpbXSwiZmFtaWx5X25hbWUiOiJZYXJsYWdhZGRhIiwiZ2l2ZW5fbmFtZSI6IlBhdmFuIiwiaXBhZGRyIjoiMTY1LjIyNS4wLjc5IiwibmFtZSI6IllhcmxhZ2FkZGEsIFBhdmFuIiwib2lkIjoiZDNiMTMzMzYtMGFlZC00MzQ1LWE1ZmQtMmNjODM4YTZiNWYxIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTI3MjQxMTM3OTctNDI0MTE3MDAxNi0yNTY2NzgzOTgwLTIxNzU3NSIsInB1aWQiOiIxMDAzQkZGRDlBQ0VENjExIiwic2NwIjoiRVdTLkFjY2Vzc0FzVXNlci5BbGwgVXNlci5SZWFkIiwic2lkIjoiYzZhZDQwYWQtMTg4OC00ZGU0LTgxYTgtMmI5OTM0MzYxYzhjIiwic2lnbmluX3N0YXRlIjpbImR2Y19tbmdkIiwiZHZjX2RtamQiLCJrbXNpIl0sInN1YiI6IlRrdnVRU3NvV3p3MlZIaWtucVJpVnFJMXAtclpqaUJkc2JYbm81ZWRtUEEiLCJ0aWQiOiJiNWI4YjQ4My01NTk3LTRhZTctOGUyNy1mY2M0NjRhM2I1ODQiLCJ1bmlxdWVfbmFtZSI6InBhdmFuLnlhcmxhZ2FkZGFAYnNjaS5jb20iLCJ1cG4iOiJwYXZhbi55YXJsYWdhZGRhQGJzY2kuY29tIiwidXRpIjoiYUNsWXBZZzM4MGl2bkVMSlZZUk5BQSIsInZlciI6IjEuMCJ9.Yd4qV8E3jQAOzDKWU-H5e7__XvJvPXOa4QwT_ryauAVPLG8rV_3tIwZdVHdsYUb3T4dSjkivq7kpl2AvbJGiBF1tZSBhrKEjlpPtYNZyE_oqTtiB0G06ZsoDac37ZatPqKq3OhraYh4OO7VYM8WNCvoeebwHIxe3PZZYRuFCZ-Vu2E3HtkiXFduxXQlcsKXFvEI136xQXMPYw1onSGEybMiz5HYsPPTVHx0PW4WEAyb4tiTNtby2vfFxbMFUROkZhwvRzeQgsHXYuafer2ebiWHapPOmkIbrXHnIvyjEJH8vSg4gGRW2HWvDyLxwR5UhZDcn_if0wa16qdDkFz7qzg";
                // use SSL EWS + OAUTH 2.0
                MailServer oServer = new MailServer("outlook.office365.com", officeUser, token, true,
                    ServerAuthType.AuthXOAUTH2, ServerProtocol.ExchangeEWS);

                Console.WriteLine("Connecting server ...");

                MailClient oClient = new MailClient("TryIt");
                oClient.Connect(oServer);

                Console.WriteLine("Retreiving email list ...");
                MailInfo[] infos = oClient.GetMailInfos();
                Console.WriteLine("Total {0} email(s)", infos.Length);

                for (int i = 0; i < infos.Length; i++)
                {
                    Console.WriteLine("Checking {0}/{1} ...", i + 1, infos.Length);
                    MailInfo info = infos[i];

                    // Generate an unqiue email file name based on date time.
                    string fileName = _generateFileName(i + 1);
                    string fullPath = string.Format("{0}\\{1}", localInbox, fileName);

                    Console.WriteLine("Downloading {0}/{1} ...", i + 1, infos.Length);
                    Mail oMail = oClient.GetMail(info);

                    // Save mail to local file
                    oMail.SaveAs(fullPath, true);

                    // Mark the email as deleted on server.
                    Console.WriteLine("Deleting ... {0}/{1}", i + 1, infos.Length);
                    oClient.Delete(info);
                }

                Console.WriteLine("Disconnecting ...");

                // Delete method just mark the email as deleted, 
                // Quit method expunge the emails from server permanently.
                oClient.Quit();

                Console.WriteLine("Completed!");
            }
            catch (Exception ep)
            {
                Console.WriteLine("Error: {0}", ep.Message);
            }
        }

        private void Authenticate_Click(object sender, RoutedEventArgs e)
        {
            ReceiveMailByOauth();
        }
    }
}
