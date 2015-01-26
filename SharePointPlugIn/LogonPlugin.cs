using System;
using System.Net;
using System.Security;
using Microsoft.VisualStudio.TestTools.WebTesting;
using Microsoft.SharePoint.Client;

namespace SharePointPlugin
{
    [System.ComponentModel.DisplayName("SharePoint Login Plugin")]
    [System.ComponentModel.Description("Logs in a user to perform web tests")]
    public class LogonPlugin : WebTestPlugin
    {
        string _siteUrl;

        public override void PreWebTest(object sender, PreWebTestEventArgs e)
        {

            using (ClientContext clientContext = new ClientContext(_siteUrl))
            {
                Uri siteUri = new Uri(_siteUrl);
                string username = e.WebTest.UserName;
                string password = e.WebTest.Password;
                try
                {
                    if (username.Contains("{{"))
                    {
                        //looks to be a databound field
                        var usernamekey = e.WebTest.UserName.Replace("{{", "").Replace("}}", "");
                        username = (string)e.WebTest.Context[usernamekey];
                    }
                    if (password.Contains("{{"))
                    {
                        //looks to be a databound field
                        var passwordkey = e.WebTest.Password.Replace("{{", "").Replace("}}", "");
                        password = (string)e.WebTest.Context[passwordkey];
                    }
                }
                catch(Exception ex)
                {
                    e.WebTest.Outcome = Outcome.Fail;
                    e.WebTest.AddCommentToResult(ex.Message);
                }

                SecureString securePassword = new SecureString();
                foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);
                var credentials = new SharePointOnlineCredentials(username, securePassword);
                try
                {
                    string authCookieValue = credentials.GetAuthenticationCookie(siteUri);
                    authCookieValue = authCookieValue.Replace("SPOIDCRL=", "");
                    var cc = e.WebTest.Context.CookieContainer;
                    cc.Add(new Cookie(
                    "FedAuth",
                    authCookieValue,
                    String.Empty,
                    siteUri.Authority));
                    Console.WriteLine(authCookieValue);
                }
                catch(IdcrlException ex)
                {
                    e.WebTest.Outcome = Outcome.Fail;
                    e.WebTest.AddCommentToResult(ex.Message);
                }
                
            }

            base.PreWebTest(sender, e);
        }

        public override void PostWebTest(object sender, PostWebTestEventArgs e)
        {
            base.PostWebTest(sender, e);
        }

        public string SiteUrl
        {
            get
            {
                return _siteUrl;
            }

            set
            {
                _siteUrl = value;
            }
        }
    }
}

